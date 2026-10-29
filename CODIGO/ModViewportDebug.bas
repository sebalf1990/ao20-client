Attribute VB_Name = "ModViewportDebug"
'    Argentum 20 - Game Client Program
'
'    Modulo provisorio de debug del PoC de viewport fullscreen.
'    Plan: ia/archivo/plans/2026/2026-05-20-viewport-fullscreen-poc.md
'
'    Responsabilidades:
'      - Ring buffer en memoria con las ultimas N lineas que pasaron por la consola
'        del juego (hookeado en AddtoRichTextBox).
'      - Overlay DirectX con info tecnica + crosshair + hover/click marker + borde
'        + mirror de la consola en zona abajo-izquierda.
'      - Log a archivo App.path\logs\viewport_debug.log para post-mortem.
'      - Tecla local de toggle (F2) y flag local del modo viewport (HUD vs FULL).
'
Option Explicit

Private Const RING_SIZE              As Integer = 80
Private Const CONSOLE_RENDER_LINES   As Integer = 10
Private Const CLICK_MARKER_MS        As Long = 500
Private Const OVERLAY_INFO_BOX_W     As Integer = 420
Private Const OVERLAY_INFO_X         As Integer = 5
Private Const OVERLAY_INFO_Y         As Integer = 190
Private Const OVERLAY_INFO_LINE_H    As Integer = 18
Private Const OVERLAY_CONSOLE_BOX_W  As Integer = 380
Private Const OVERLAY_CONSOLE_BOX_H  As Integer = 170
Private Const OVERLAY_CONSOLE_LINE_H As Integer = 16
Private Const WM_MOUSEWHEEL          As Long = &H20A
Private Const WM_DESTROY             As Long = &H2

Private Type t_ViewportDebugPoint
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As t_ViewportDebugPoint) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As t_ViewportDebugPoint) As Long
Private Type t_DebugConsoleLine
    text     As String
    red      As Integer
    green    As Integer
    blue     As Integer
    bold     As Boolean
    italic   As Boolean
    ts       As Long
    category As Byte
End Type

Private g_ring(0 To 79)         As t_DebugConsoleLine
Private g_ring_head             As Integer
Private g_ring_count            As Integer

Private g_viewport_full         As Boolean
Private g_overlay_visible       As Boolean

Private g_last_click_tx         As Byte
Private g_last_click_ty         As Byte
Private g_last_click_ts         As Long
Private g_last_click_valid      As Boolean

Private g_log_initialized       As Boolean
Private g_log_path              As String
Private g_console_scroll_lines  As Long
Private g_console_total_visual   As Long
Private g_subclassed_renderer    As Boolean
Private g_subclassed_form        As Boolean

' Plan 21.002 V9.12: filtros togglables de consola FULL.
Public Const e_ConsoleCat_General As Byte = 0
Public Const e_ConsoleCat_Combat  As Byte = 1
Public Const e_ConsoleCat_Global  As Byte = 2
Private g_filter_loaded   As Boolean
Private g_show_general    As Boolean
Private g_show_combat     As Boolean
Private g_show_global     As Boolean
Private Const BTN_W           As Long = 38
Private Const BTN_H           As Long = 16
Private Const BTN_PAD_X       As Long = 4

' ============== API: estado ==============

Public Sub ViewportDebug_ToggleOverlay()
    g_overlay_visible = Not g_overlay_visible
    If g_overlay_visible Then
        Call SetMask(FeatureToggles, eShowViewportDebug)
        Call ViewportDebug_AppendLog("[OVERLAY] ON", 0, 255, 0)
    Else
        Call UnsetMask(FeatureToggles, eShowViewportDebug)
        Call ViewportDebug_AppendLog("[OVERLAY] OFF", 0, 255, 0)
    End If
End Sub

Public Sub ViewportDebug_SetViewportMode(ByVal IsFullscreen As Boolean)
    g_viewport_full = IsFullscreen
    g_console_scroll_lines = 0
    If IsFullscreen Then
        Call ViewportDebug_EnableMouseWheel
    Else
        Call ViewportDebug_DisableMouseWheel
    End If
    Call ViewportDebug_AppendLog("[MODE] " & IIf(IsFullscreen, "FULL", "HUD"), 0, 255, 0)
End Sub

Public Function ViewportDebug_IsFullscreen() As Boolean
    ViewportDebug_IsFullscreen = g_viewport_full
End Function

Public Function ViewportDebug_IsOverlayVisible() As Boolean
    ViewportDebug_IsOverlayVisible = g_overlay_visible
End Function
Public Sub ViewportDebug_EnableMouseWheel()
    On Error Resume Next
    If Not g_subclassed_renderer Then g_subclassed_renderer = (Subclass(frmMain.renderer.hWnd, AddressOf ViewportDebug_WndProc) <> 0)
    If Not g_subclassed_form Then g_subclassed_form = (Subclass(frmMain.hWnd, AddressOf ViewportDebug_WndProc) <> 0)
End Sub

Public Sub ViewportDebug_DisableMouseWheel()
    On Error Resume Next
    If g_subclassed_renderer Then Call UnSubclass(frmMain.renderer.hWnd, AddressOf ViewportDebug_WndProc)
    If g_subclassed_form Then Call UnSubclass(frmMain.hWnd, AddressOf ViewportDebug_WndProc)
    g_subclassed_renderer = False
    g_subclassed_form = False
End Sub

Public Function ViewportDebug_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    On Error GoTo ViewportDebug_WndProc_Err
    If uMsg = WM_MOUSEWHEEL Then
        Dim wheelX As Long
        Dim wheelY As Long
        If g_viewport_full And ViewportDebug_GetCursorRendererPos(wheelX, wheelY) Then
            If IsMouseOverConsole(wheelX, wheelY) Then
                If ViewportDebug_SignedHiWord(wParam) > 0 Then
                    Call ViewportDebug_ScrollConsole(-3)
                Else
                    Call ViewportDebug_ScrollConsole(3)
                End If
                ViewportDebug_WndProc = 0
                Exit Function
            End If
        End If
    ElseIf uMsg = WM_DESTROY Then
        If hWnd = frmMain.renderer.hWnd Then g_subclassed_renderer = False
        If hWnd = frmMain.hWnd Then g_subclassed_form = False
    End If
    ViewportDebug_WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    Exit Function
ViewportDebug_WndProc_Err:
    ViewportDebug_WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

Private Function ViewportDebug_GetCursorRendererPos(ByRef x As Long, ByRef y As Long) As Boolean
    Dim pt As t_ViewportDebugPoint

    If GetCursorPos(pt) = 0 Then Exit Function
    Call ScreenToClient(frmMain.renderer.hWnd, pt)
    x = pt.x
    y = pt.y
    ViewportDebug_GetCursorRendererPos = True
End Function

Private Function ViewportDebug_SignedHiWord(ByVal value As Long) As Long
    Dim hi As Long

    If value < 0 Then
        hi = ((value And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        hi = (value \ &H10000) And &HFFFF&
    End If

    If (hi And &H8000&) <> 0 Then hi = hi - &H10000
    ViewportDebug_SignedHiWord = hi
End Function

Public Function ViewportDebug_HandleMouseDown(ByVal x As Long, ByVal y As Long) As Boolean
    On Error Resume Next
    If Not g_viewport_full Then Exit Function
    If Not IsMouseOverConsole(x, y) Then Exit Function

    Dim consX As Long, consY As Long, consW As Long, consH As Long
    Call GetConsoleRect(consX, consY, consW, consH)
    If ViewportDebug_HitConsoleFilter(x, y, consX, consY) Then
        ViewportDebug_HandleMouseDown = True
        Exit Function
    End If
    If x >= consX + consW - 12 And x <= consX + consW And y >= consY + 22 And y <= consY + consH Then
        If g_console_total_visual > 0 Then
            Dim trackH As Long
            trackH = consH - 32
            If trackH < 1 Then trackH = 1
            g_console_scroll_lines = (g_console_total_visual * (y - (consY + 24))) \ trackH
            Call ClampConsoleScroll
        End If
        ViewportDebug_HandleMouseDown = True
    End If
End Function

Public Sub ViewportDebug_ScrollConsole(ByVal deltaLines As Long)
    g_console_scroll_lines = g_console_scroll_lines + deltaLines
    Call ClampConsoleScroll
End Sub

Private Sub ClampConsoleScroll()
    If g_console_scroll_lines < 0 Then g_console_scroll_lines = 0
    If g_console_total_visual > 0 Then
        If g_console_scroll_lines > g_console_total_visual Then g_console_scroll_lines = g_console_total_visual
    End If
End Sub

Private Sub GetConsoleRect(ByRef consX As Long, ByRef consY As Long, ByRef consW As Long, ByRef consH As Long)
    consX = 5
    consW = hotkey_render_posX - consX - 10
    If consW < 280 Then consW = OVERLAY_CONSOLE_BOX_W
    consH = OVERLAY_CONSOLE_BOX_H
    consY = frmMain.renderer.ScaleHeight - consH - 5
End Sub

Private Function IsMouseOverConsole(ByVal x As Long, ByVal y As Long) As Boolean
    Dim consX As Long, consY As Long, consW As Long, consH As Long
    Call GetConsoleRect(consX, consY, consW, consH)
    IsMouseOverConsole = (x >= consX And x <= consX + consW And y >= consY And y <= consY + consH)
End Function

' ============== API: eventos ==============

Public Sub ViewportDebug_PushConsoleLine(ByVal text As String, _
                                         ByVal red As Integer, _
                                         ByVal green As Integer, _
                                         ByVal blue As Integer, _
                                         ByVal bold As Boolean, _
                                         ByVal italic As Boolean, _
                                         Optional ByVal appendToPrevious As Boolean = False)
    On Error GoTo ViewportDebug_PushConsoleLine_Err
    Dim cleanText As String
    cleanText = Replace(text, vbCrLf, " ")
    cleanText = Replace(cleanText, vbLf, " ")
    cleanText = Replace(cleanText, vbCr, " ")
    Dim cat As Byte
    cat = ClassifyConsoleCategory(red, green, blue, CBool(bold), CBool(italic))
    Call PushOneRingEntry(cleanText, red, green, blue, bold, italic, appendToPrevious, cat)
    Call ViewportDebug_AppendLog(text, red, green, blue)
    Exit Sub
ViewportDebug_PushConsoleLine_Err:
    ' No reportar via RegistrarError para evitar recursion (RegistrarError podria llamar consola).
    Resume Next
End Sub

Private Function WrapConsoleText(ByVal text As String, ByVal maxWidth As Long, ByVal bold As Boolean, ByRef outLines() As String) As Integer
    On Error GoTo WrapConsoleText_Err
    ReDim outLines(0 To 0) As String
    Dim count As Integer
    Dim rest As String
    rest = Trim$(text)
    If LenB(rest) = 0 Then
        outLines(0) = " "
        WrapConsoleText = 1
        Exit Function
    End If

    ' Engine_Text_Render usa multi_line como parametro 6; no pasar bold aca.
    Do While Len(rest) > 0
        Dim takeLen As Long
        takeLen = Len(rest)
        Do While takeLen > 1 And Engine_Text_Width(Left$(rest, takeLen), False) > maxWidth
            takeLen = takeLen - 1
        Loop

        If takeLen < Len(rest) Then
            Dim lastSpace As Long
            lastSpace = InStrRev(Left$(rest, takeLen), " ")
            If lastSpace > 8 Then takeLen = lastSpace - 1
        End If

        If count > 0 Then ReDim Preserve outLines(0 To count) As String
        outLines(count) = Left$(rest, takeLen)
        count = count + 1
        rest = Trim$(Mid$(rest, takeLen + 1))
    Loop

    WrapConsoleText = count
    Exit Function
WrapConsoleText_Err:
    ReDim outLines(0 To 0) As String
    outLines(0) = text
    WrapConsoleText = 1
End Function

Private Sub RenderConsoleTextLine(ByVal text As String, _
                                  ByVal x As Long, _
                                  ByVal y As Long, _
                                  ByVal width As Long, _
                                  ByVal red As Byte, _
                                  ByVal green As Byte, _
                                  ByVal blue As Byte)
    On Error Resume Next
    Dim textColor(3) As RGBA
    Dim shadowColor(3) As RGBA

    Call RGBAList(textColor, red, green, blue, 255)
    Call RGBAList(shadowColor, CByte(red \ 6), CByte(green \ 6), CByte(blue \ 6), 204)

    ' Renderer de una sola linea: evita Engine_Text_Render y su wrap interno.
    Call RenderConsoleGlyphLine(text, x + 1, y + 1, shadowColor)
    Call RenderConsoleGlyphLine(text, x, y, textColor)
End Sub

Private Sub RenderConsoleGlyphLine(ByVal text As String, _
                                   ByVal x As Long, _
                                   ByVal y As Long, _
                                   ByRef textColor() As RGBA)
    On Error Resume Next
    Dim i As Long
    Dim charCode As Integer
    Dim grhIndex As Long
    Dim drawX As Long
    Dim charW As Long

    drawX = x

    For i = 1 To Len(text)
        charCode = Asc(Mid$(text, i, 1))
        If charCode = 32 Then
            drawX = drawX + 4
        Else
            grhIndex = Fuentes(1).Caracteres(charCode)
            If grhIndex > 12 Then
                Call Draw_GrhFont(grhIndex, drawX, y, textColor)
                charW = 0
                charW = GrhData(GrhData(grhIndex).Frames(1)).pixelWidth
                If charW <= 0 Then charW = GrhData(grhIndex).pixelWidth
                If charW <= 0 Then charW = 6
                drawX = drawX + charW
            End If
        End If
    Next i
End Sub

Private Sub PushOneRingEntry(ByVal text As String, _
                              ByVal red As Integer, _
                              ByVal green As Integer, _
                              ByVal blue As Integer, _
                              ByVal bold As Boolean, _
                              ByVal italic As Boolean, _
                              ByVal appendToPrevious As Boolean, _
                              ByVal category As Byte)
    On Error Resume Next
    If appendToPrevious And g_ring_count > 0 Then
        Dim lastIndex As Integer
        lastIndex = (g_ring_head - 1 + RING_SIZE) Mod RING_SIZE
        g_ring(lastIndex).text = g_ring(lastIndex).text & text
        g_ring(lastIndex).ts = FrameTime
        Exit Sub
    End If
    With g_ring(g_ring_head)
        .text = text
        .red = red
        .green = green
        .blue = blue
        .bold = bold
        .italic = italic
        .ts = FrameTime
        .category = category
    End With
    g_ring_head = (g_ring_head + 1) Mod RING_SIZE
    If g_ring_count < RING_SIZE Then g_ring_count = g_ring_count + 1
End Sub

Private Function ClassifyConsoleCategory(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal bold As Boolean, ByVal italic As Boolean) As Byte
    ' GLOBAL: FONTTYPE_GLOBAL = (255,102,204) italic
    If italic And r = 255 And g = 102 And b = 204 Then
        ClassifyConsoleCategory = e_ConsoleCat_Global
        Exit Function
    End If
    ' COMBAT: FIGHT/CRIMINAL/CRIMINAL_CAOS => red dominante, green/blue bajos, bold
    If bold And r >= 200 And g <= 110 And b <= 110 Then
        ClassifyConsoleCategory = e_ConsoleCat_Combat
        Exit Function
    End If
    ClassifyConsoleCategory = e_ConsoleCat_General
End Function

Private Sub EnsureFiltersLoaded()
    If g_filter_loaded Then Exit Sub
    g_filter_loaded = True
    g_show_general = ReadConsoleFilterSetting("ShowGeneral")
    g_show_combat  = ReadConsoleFilterSetting("ShowCombat")
    g_show_global  = ReadConsoleFilterSetting("ShowGlobal")
End Sub

Private Function ReadConsoleFilterSetting(ByVal key As String) As Boolean
    Dim v As String
    v = GetSetting("ConsoleFull", key)
    If LenB(v) = 0 Then
        ReadConsoleFilterSetting = True
    Else
        ReadConsoleFilterSetting = (v = "1")
    End If
End Function

Public Sub ViewportDebug_ToggleCategory(ByVal category As Byte)
    Call EnsureFiltersLoaded
    Select Case category
        Case e_ConsoleCat_General: g_show_general = Not g_show_general: Call SaveSetting("ConsoleFull", "ShowGeneral", IIf(g_show_general, "1", "0"))
        Case e_ConsoleCat_Combat:  g_show_combat  = Not g_show_combat:  Call SaveSetting("ConsoleFull", "ShowCombat",  IIf(g_show_combat,  "1", "0"))
        Case e_ConsoleCat_Global:  g_show_global  = Not g_show_global:  Call SaveSetting("ConsoleFull", "ShowGlobal",  IIf(g_show_global,  "1", "0"))
    End Select
End Sub

Public Function ViewportDebug_IsCategoryEnabled(ByVal category As Byte) As Boolean
    Call EnsureFiltersLoaded
    Select Case category
        Case e_ConsoleCat_General: ViewportDebug_IsCategoryEnabled = g_show_general
        Case e_ConsoleCat_Combat:  ViewportDebug_IsCategoryEnabled = g_show_combat
        Case e_ConsoleCat_Global:  ViewportDebug_IsCategoryEnabled = g_show_global
    End Select
End Function

Public Sub ViewportDebug_OnClick(ByVal click_tx As Byte, ByVal click_ty As Byte)
    g_last_click_tx = click_tx
    g_last_click_ty = click_ty
    g_last_click_ts = FrameTime
    g_last_click_valid = True
End Sub

' ============== RENDER ==============

Public Sub Render_ViewportDebugOverlay()
    On Error GoTo Render_ViewportDebugOverlay_Err
    If Not g_viewport_full And Not g_overlay_visible Then Exit Sub
    If g_overlay_visible And Not IsSet(FeatureToggles, eShowViewportDebug) Then Exit Sub

    Dim sw As Long, sh As Long
    sw = frmMain.renderer.ScaleWidth
    sh = frmMain.renderer.ScaleHeight

    Dim borderColor As RGBA
    borderColor = RGBA_From_Comp(255, 50, 50, 200)
    Dim crossColor As RGBA
    crossColor = RGBA_From_Comp(255, 255, 255, 160)
    Dim hoverColor As RGBA
    hoverColor = RGBA_From_Comp(80, 255, 255, 60)
    Dim clickColor As RGBA
    clickColor = RGBA_From_Comp(255, 255, 80, 120)
    Dim bgInfo As RGBA
    bgInfo = RGBA_From_Comp(0, 0, 0, 200)
    Dim bgCons As RGBA
    bgCons = RGBA_From_Comp(0, 0, 0, 140)
    Dim white(3) As RGBA
    Call RGBAList(white, 255, 255, 255, 255)

    If g_overlay_visible Then
    ' 1) Borde rojo del renderer (perimetro 1px) - valida que Render_Main_Rect se actualizo
    Call Engine_Draw_Box(0, 0, sw, 1, borderColor)
    Call Engine_Draw_Box(0, sh - 1, sw, 1, borderColor)
    Call Engine_Draw_Box(0, 0, 1, sh, borderColor)
    Call Engine_Draw_Box(sw - 1, 0, 1, sh, borderColor)

    ' 2) Crosshair central - valida que el char queda visualmente centrado
    Dim cx As Long, cy As Long
    cx = g_viewport_logical_left + g_viewport_logical_width \ 2
    cy = g_viewport_logical_top + g_viewport_logical_height \ 2
    Call Engine_Draw_Box(cx - 12, cy, 24, 1, crossColor)
    Call Engine_Draw_Box(cx, cy - 12, 1, 24, crossColor)

    ' 3) Hover marker - tile bajo el mouse
    If mouseX >= g_viewport_logical_left And mouseY >= g_viewport_logical_top And mouseX < g_viewport_logical_left + g_viewport_logical_width And mouseY < g_viewport_logical_top + g_viewport_logical_height Then
        Dim hx As Long, hy As Long
        hx = g_viewport_logical_left + ((mouseX - g_viewport_logical_left) \ 32) * 32
        hy = g_viewport_logical_top + ((mouseY - g_viewport_logical_top) \ 32) * 32
        Call Engine_Draw_Box(hx, hy, 32, 32, hoverColor)
    End If

    ' 4) Click marker - flash del ultimo click izquierdo, 500ms
    If g_last_click_valid Then
        Dim ageMs As Long
        ageMs = FrameTime - g_last_click_ts
        If ageMs >= 0 And ageMs < CLICK_MARKER_MS Then
            Dim px As Long, py As Long
            px = Get_Pixelx_Of_XY(g_last_click_tx)
            py = Get_PixelY_Of_XY(g_last_click_ty)
            Call Engine_Draw_Box(px, py, 32, 32, clickColor)
        Else
            g_last_click_valid = False
        End If
    End If

    ' 5) Caja de texto top-left con info tecnica
    Dim info(8) As String
    info(0) = "[ViewportDebug] F2 toggle  (mode: " & IIf(g_viewport_full, "FULL", "HUD") & ")"
    info(1) = "renderer: L=" & frmMain.renderer.Left & " T=" & frmMain.renderer.Top & " W=" & sw & " H=" & sh
    info(2) = "logical: L=" & g_viewport_logical_left & " T=" & g_viewport_logical_top & " W=" & g_viewport_logical_width & " H=" & g_viewport_logical_height
    info(3) = "Render_Main_Rect: R=" & Render_Main_Rect.Right & " B=" & Render_Main_Rect.Bottom
    info(4) = "HalfWindowTile: w=" & HalfWindowTileWidth & " h=" & HalfWindowTileHeight
    info(5) = "MinBorder x=" & MinXBorder & " y=" & MinYBorder & "  MaxBorder x=" & MaxXBorder & " y=" & MaxYBorder
    info(6) = "UserPos: x=" & UserPos.x & " y=" & UserPos.y & "  Map: " & UserMap
    info(7) = "mouse(px): " & mouseX & "," & mouseY & "   tile: tX=" & tX & " tY=" & tY
    info(8) = "fps: " & engine.fps

    Dim infoH As Long
    infoH = 9 * OVERLAY_INFO_LINE_H + 10
    Call Engine_Draw_Box(OVERLAY_INFO_X, OVERLAY_INFO_Y, OVERLAY_INFO_BOX_W, infoH, bgInfo)

    Dim i As Integer
    For i = 0 To 8
        Call Engine_Text_Render(info(i), OVERLAY_INFO_X + 5, OVERLAY_INFO_Y + 3 + i * OVERLAY_INFO_LINE_H, white, 1, False)
    Next i
    End If

    ' 6) Mirror de consola abajo-izquierda - ajusta texto por ancho hasta antes de hotkeys
    Dim consX As Long, consY As Long, consW As Long, consH As Long
    Call GetConsoleRect(consX, consY, consW, consH)

    Call Engine_Draw_Box(consX, consY, consW, consH, bgCons)
    Call Render_ConsoleFilterButtons(consX, consY, consW)

    If g_ring_count > 0 Then
        Dim bottomY As Long
        bottomY = consY + consH - OVERLAY_CONSOLE_LINE_H - 4
        Dim minY As Long
        minY = consY + 22
        Dim lineColor(3) As RGBA
        Dim idx As Integer
        Dim wrapped() As String
        Dim wrappedCount As Integer
        Dim wrappedIdx As Integer
        Dim renderedLines As Integer
        Dim maxLines As Integer
        maxLines = (consH - 26) \ OVERLAY_CONSOLE_LINE_H

        Dim skippedLines As Long
        Dim totalVisual As Long
        Call EnsureFiltersLoaded
        For i = 0 To g_ring_count - 1
            idx = (g_ring_head - 1 - i + RING_SIZE) Mod RING_SIZE
            With g_ring(idx)
                If IsCategoryVisible(.category) Then
                    wrappedCount = WrapConsoleText(.text, consW - 24, .bold, wrapped)
                    totalVisual = totalVisual + wrappedCount
                End If
            End With
        Next i
        g_console_total_visual = totalVisual
        Call ClampConsoleScroll

        For i = 0 To g_ring_count - 1
            idx = (g_ring_head - 1 - i + RING_SIZE) Mod RING_SIZE
            With g_ring(idx)
                If Not IsCategoryVisible(.category) Then GoTo NextEntry
                Call RGBAList(lineColor, CByte(.red And 255), CByte(.green And 255), CByte(.blue And 255), 255)
                wrappedCount = WrapConsoleText(.text, consW - 24, .bold, wrapped)
                For wrappedIdx = wrappedCount - 1 To 0 Step -1
                    If skippedLines < g_console_scroll_lines Then
                        skippedLines = skippedLines + 1
                    Else
                        If bottomY < minY Then Exit For
                        Call RenderConsoleTextLine(wrapped(wrappedIdx), consX + 5, bottomY, consW - 24, CByte(.red And 255), CByte(.green And 255), CByte(.blue And 255))
                        bottomY = bottomY - OVERLAY_CONSOLE_LINE_H
                        renderedLines = renderedLines + 1
                        If renderedLines >= maxLines Then Exit For
                    End If
                Next wrappedIdx
            End With
NextEntry:
            If bottomY < minY Or renderedLines >= maxLines Then Exit For
        Next i

        ' Scrollbar siempre visible mientras la consola tenga contenido,
        ' aunque la categoria activa no produzca overflow (asi no parpadea al togglear filtros).
        If g_ring_count > 0 Then
            Dim trackX As Long, trackY As Long, trackH As Long
            trackX = consX + consW - 8
            trackY = consY + 24
            trackH = consH - 32
            Dim scrollTrack As RGBA
            scrollTrack = RGBA_From_Comp(220, 220, 220, 80)
            Call Engine_Draw_Box(trackX, trackY, 6, trackH, scrollTrack)

            Dim thumbH As Long, thumbY As Long
            If g_console_total_visual <= 0 Then
                thumbH = trackH
            ElseIf g_console_total_visual <= maxLines Then
                thumbH = trackH
            Else
                thumbH = (maxLines * trackH) \ g_console_total_visual
            End If
            If thumbH < 18 Then thumbH = 18
            If thumbH > trackH Then thumbH = trackH
            thumbY = trackY
            If g_console_total_visual > maxLines Then thumbY = trackY + (g_console_scroll_lines * (trackH - thumbH)) \ (g_console_total_visual - maxLines)
            Dim scrollThumb As RGBA
            scrollThumb = RGBA_From_Comp(235, 235, 235, 220)
            Call Engine_Draw_Box(trackX - 1, thumbY, 8, thumbH, scrollThumb)
        End If
    End If

    Exit Sub
Render_ViewportDebugOverlay_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportDebug.Render_ViewportDebugOverlay", Erl)
    Resume Next
End Sub

' ============== LOG A ARCHIVO ==============

Public Sub ViewportDebug_AppendDiagLog(ByVal text As String)
    Call ViewportDebug_AppendLog(text, 0, 200, 200)
End Sub

Private Sub ViewportDebug_AppendLog(ByVal text As String, ByVal R As Integer, ByVal G As Integer, ByVal B As Integer)
    On Error GoTo ViewportDebug_AppendLog_Err
    If Not g_log_initialized Then
        g_log_path = App.path & "\logs\viewport_debug.log"
        ' Asumimos carpeta logs existe (la convencion del cliente ya escribe Errores.log ahi).
        ' Si no existe el error cae a Resume Next y no logueamos nada.
        g_log_initialized = True
    End If
    Dim f As Integer
    f = FreeFile
    Open g_log_path For Append As #f
    Print #f, format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & R & "," & G & "," & B & "] " & text
    Close #f
    Exit Sub
ViewportDebug_AppendLog_Err:
    ' No re-reportar: el log es best-effort.
    Resume Next
End Sub

Private Function IsCategoryVisible(ByVal cat As Byte) As Boolean
    Select Case cat
        Case e_ConsoleCat_Combat: IsCategoryVisible = g_show_combat
        Case e_ConsoleCat_Global: IsCategoryVisible = g_show_global
        Case Else: IsCategoryVisible = g_show_general
    End Select
End Function

Public Sub Render_ConsoleFilterButtons(ByVal consX As Long, ByVal consY As Long, ByVal consW As Long)
    On Error Resume Next
    Call EnsureFiltersLoaded
    Dim white(3) As RGBA
    Call RGBAList(white, 255, 255, 255, 255)
    Dim dimC(3) As RGBA
    Call RGBAList(dimC, 130, 130, 130, 255)

    Dim btnX As Long
    btnX = consX + 5
    Call DrawFilterButton(btnX, consY + 3, "Gen", g_show_general, white, dimC)
    btnX = btnX + BTN_W + BTN_PAD_X
    Call DrawFilterButton(btnX, consY + 3, "Cmb", g_show_combat, white, dimC)
    btnX = btnX + BTN_W + BTN_PAD_X
    Call DrawFilterButton(btnX, consY + 3, "Glb", g_show_global, white, dimC)
End Sub

Private Sub DrawFilterButton(ByVal x As Long, ByVal y As Long, ByVal label As String, ByVal active As Boolean, ByRef onColor() As RGBA, ByRef offColor() As RGBA)
    On Error Resume Next
    Dim bg As RGBA
    If active Then
        bg = RGBA_From_Comp(60, 140, 200, 200)
    Else
        bg = RGBA_From_Comp(40, 40, 40, 180)
    End If
    Call Engine_Draw_Box(x, y, BTN_W, BTN_H, bg)
    Dim borderC As RGBA
    borderC = RGBA_From_Comp(0, 0, 0, 220)
    Call Engine_Draw_Box(x, y, BTN_W, 1, borderC)
    Call Engine_Draw_Box(x, y + BTN_H - 1, BTN_W, 1, borderC)
    Call Engine_Draw_Box(x, y, 1, BTN_H, borderC)
    Call Engine_Draw_Box(x + BTN_W - 1, y, 1, BTN_H, borderC)
    Dim tw As Long
    tw = Engine_Text_Width(label, False)
    If active Then
        Call Engine_Text_Render(label, x + (BTN_W - tw) \ 2, y + 1, onColor, 1, False)
    Else
        Call Engine_Text_Render(label, x + (BTN_W - tw) \ 2, y + 1, offColor, 1, False)
    End If
End Sub

Public Function ViewportDebug_HitConsoleFilter(ByVal mx As Long, ByVal my As Long, ByVal consX As Long, ByVal consY As Long) As Boolean
    Dim btnX As Long
    Dim btnY As Long
    btnY = consY + 3
    btnX = consX + 5
    If my < btnY Or my > btnY + BTN_H Then Exit Function
    Dim i As Long
    For i = 0 To 2
        If mx >= btnX And mx <= btnX + BTN_W Then
            Select Case i
                Case 0: Call ViewportDebug_ToggleCategory(e_ConsoleCat_General)
                Case 1: Call ViewportDebug_ToggleCategory(e_ConsoleCat_Combat)
                Case 2: Call ViewportDebug_ToggleCategory(e_ConsoleCat_Global)
            End Select
            ViewportDebug_HitConsoleFilter = True
            Exit Function
        End If
        btnX = btnX + BTN_W + BTN_PAD_X
    Next i
End Function
