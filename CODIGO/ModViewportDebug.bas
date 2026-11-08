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
Private Const CONSOLE_ENTRY_MAX_CHARS As Long = 1024
Private Const CONSOLE_HEADER_H       As Long = 26
Private Const CONSOLE_TRACK_TOP      As Long = 24
Private Const CONSOLE_TRACK_PAD      As Long = 32
Private Const CONSOLE_THUMB_MIN_H    As Long = 18
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
    ' E3: cache del wrap (se recalcula solo al agregar texto o cambiar el ancho).
    wrapLines() As String
    wrapCount   As Integer
    wrapWidth   As Long
    ' E6: set when the entry was added or extended since the last recompute; countedLines = wrapped lines
    ' already accounted for (so only the new lines move the scroll anchor).
    isNew        As Boolean
    countedLines As Integer
End Type

Private g_ring(0 To RING_SIZE - 1) As t_DebugConsoleLine
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
Private g_wrap_dirty             As Boolean
Private g_wrap_total_width       As Long
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

' E6: "new messages" indicator shown while the console is scrolled up and lines arrive.
Private Const UNREAD_TEXT     As String = "v Nuevos mensajes"
Private Const UNREAD_PAD_X    As Long = 8
Private Const UNREAD_H        As Long = 18
Private Const UNREAD_TEXT_Y   As Long = 3
Private Const UNREAD_BOTTOM   As Long = 4
Private Const CONSOLE_SCROLLBAR_W As Long = 12
Private g_pill_drawn          As Boolean
Private g_pill_x              As Long
Private g_pill_y              As Long
Private g_pill_w              As Long
Private g_console_unread      As Boolean
Private g_unread_width        As Long

' Log a archivo: rotacion simple (plan 2026-11-07-fullscreen-pulido-y-vecinos E1).
Private Const LOG_MAX_BYTES   As Long = 1048576
Private g_log_disabled        As Boolean

' E0 medicion (plan 2026-11-07-fullscreen-pulido-y-vecinos): solo al log diag.
' Off since plan 2026-11-07-fullscreen-pulido-y-vecinos closed; set True to re-measure ([PERF] lines).
Private Const PERF_ENABLED      As Boolean = False
Private Const PERF_WINDOW_MS    As Double = 5000
Public Const PERF_PIECE_HUD       As Long = 0
Public Const PERF_PIECE_CONSOLE   As Long = 1
Public Const PERF_PIECE_BUFFS     As Long = 2
Public Const PERF_PIECE_GMBAR     As Long = 3
Public Const PERF_PIECE_NEIGHBORS As Long = 4
Public Const PERF_PIECE_BUTTONS   As Long = 5
Private Const PERF_PIECE_MAX      As Long = 5
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private g_perf_freq           As Double
Private g_perf_acc(0 To 5)    As Double
Private g_perf_frames         As Long
Private g_perf_window_start   As Double

' ============== API: estado ==============

Public Sub ViewportDebug_ToggleOverlay()
    g_overlay_visible = Not g_overlay_visible
    ' E5: g_overlay_visible es la unica fuente de verdad (antes se duplicaba en el bit eShowViewportDebug).
    Call ViewportDebug_AppendLog("[OVERLAY] " & IIf(g_overlay_visible, "ON", "OFF"), 0, 255, 0)
End Sub

Public Sub ViewportDebug_SetViewportMode(ByVal IsFullscreen As Boolean)
    g_viewport_full = IsFullscreen
    g_console_scroll_lines = 0
    g_console_unread = False
    g_wrap_dirty = True
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

' isClick = True solo desde renderer_MouseDown; desde MouseMove (arrastre) el scroll sigue
' funcionando pero los filtros NO se alternan (E5).
Public Function ViewportDebug_HandleMouseDown(ByVal x As Long, ByVal y As Long, ByVal isClick As Boolean) As Boolean
    On Error Resume Next
    If Not g_viewport_full Then Exit Function
    If Not IsMouseOverConsole(x, y) Then Exit Function

    Dim consX As Long, consY As Long, consW As Long, consH As Long
    Call GetConsoleRect(consX, consY, consW, consH)
    If ViewportDebug_HitConsoleFilter(x, y, consX, consY, isClick) Then
        ViewportDebug_HandleMouseDown = True
        Exit Function
    End If
    If isClick And g_pill_drawn Then
        If x >= g_pill_x And x < g_pill_x + g_pill_w And y >= g_pill_y And y < g_pill_y + UNREAD_H Then
            g_console_scroll_lines = 0
            g_console_unread = False
            ViewportDebug_HandleMouseDown = True
            Exit Function
        End If
    End If
    If x >= consX + consW - CONSOLE_SCROLLBAR_W And x < consX + consW And y >= consY + 22 And y < consY + consH Then
        Dim maxScroll As Long
        maxScroll = GetConsoleMaxScroll()
        If maxScroll > 0 Then
            ' E5: el centro del pulgar sigue al cursor; extremo de arriba = linea mas vieja (scroll maximo),
            ' extremo de abajo = linea mas nueva (scroll 0).
            Dim trackH As Long, travel As Long, thumbPos As Long
            trackH = consH - CONSOLE_TRACK_PAD
            If trackH < 1 Then trackH = 1
            travel = trackH - ConsoleThumbHeight(trackH)
            If travel > 0 Then
                thumbPos = y - (consY + CONSOLE_TRACK_TOP) - ConsoleThumbHeight(trackH) \ 2
                If thumbPos < 0 Then thumbPos = 0
                If thumbPos > travel Then thumbPos = travel
                g_console_scroll_lines = maxScroll - (maxScroll * thumbPos + travel \ 2) \ travel
                Call ClampConsoleScroll
            End If
        End If
        ViewportDebug_HandleMouseDown = True
    ElseIf isClick Then
        ' E6: any click inside the console box is swallowed so it never reaches the world.
        ViewportDebug_HandleMouseDown = True
    End If
End Function

' E6: True when the point is inside the FULL console box (used to block right/middle click-through).
Public Function ViewportDebug_IsOverConsole(ByVal x As Long, ByVal y As Long) As Boolean
    If Not g_viewport_full Then Exit Function
    ViewportDebug_IsOverConsole = IsMouseOverConsole(x, y)
End Function

Public Sub ViewportDebug_ScrollConsole(ByVal deltaLines As Long)
    g_console_scroll_lines = g_console_scroll_lines + deltaLines
    Call ClampConsoleScroll
    If g_console_scroll_lines = 0 Then g_console_unread = False
End Sub

' E6: rectangle of the "new messages" pill (bottom center of the box, left of the scrollbar). Width is cached.
Private Sub GetUnreadPillRect(ByVal consX As Long, ByVal consY As Long, ByVal consW As Long, ByVal consH As Long, _
                              ByRef pillX As Long, ByRef pillY As Long, ByRef pillW As Long)
    If g_unread_width = 0 Then g_unread_width = Engine_Text_Width(UNREAD_TEXT, False, 1) + UNREAD_PAD_X * 2
    pillW = g_unread_width
    pillX = consX + (consW - CONSOLE_SCROLLBAR_W - pillW) \ 2
    pillY = consY + consH - UNREAD_H - UNREAD_BOTTOM
End Sub

Private Sub Render_ConsoleUnreadPill(ByVal consX As Long, ByVal consY As Long, ByVal consW As Long, ByVal consH As Long)
    Dim pillX As Long, pillY As Long, pillW As Long
    Call GetUnreadPillRect(consX, consY, consW, consH, pillX, pillY, pillW)
    g_pill_x = pillX
    g_pill_y = pillY
    g_pill_w = pillW
    g_pill_drawn = True
    Dim pillBg As RGBA
    Dim pillBorder As RGBA
    Dim pillText(3) As RGBA
    pillBg = RGBA_From_Comp(20, 20, 20, 235)
    pillBorder = RGBA_From_Comp(200, 170, 90, 255)
    Call RGBAList(pillText, 255, 230, 140, 255)
    Call Engine_Draw_Box_Border(pillX, pillY, pillW, UNREAD_H, pillBg, pillBorder)
    Call Engine_Text_Render(UNREAD_TEXT, pillX + UNREAD_PAD_X, pillY + UNREAD_TEXT_Y, pillText, 1, False)
End Sub

' E5: cantidad de lineas que entran en la caja (misma cuenta que usa el dibujo).
Private Function ConsoleVisibleLines(ByVal consH As Long) As Long
    ConsoleVisibleLines = (consH - CONSOLE_HEADER_H) \ OVERLAY_CONSOLE_LINE_H
End Function

' E5: el scroll maximo es total - lineas visibles (con scroll = total quedaba la caja vacia).
Private Function GetConsoleMaxScroll() As Long
    GetConsoleMaxScroll = g_console_total_visual - ConsoleVisibleLines(OVERLAY_CONSOLE_BOX_H)
    If GetConsoleMaxScroll < 0 Then GetConsoleMaxScroll = 0
End Function

Private Sub ClampConsoleScroll()
    If g_console_scroll_lines < 0 Then g_console_scroll_lines = 0
    Dim maxScroll As Long
    maxScroll = GetConsoleMaxScroll()
    If g_console_scroll_lines > maxScroll Then g_console_scroll_lines = maxScroll
End Sub

Private Function ConsoleThumbHeight(ByVal trackH As Long) As Long
    Dim visibleLines As Long
    Dim thumbH As Long
    visibleLines = ConsoleVisibleLines(OVERLAY_CONSOLE_BOX_H)
    If g_console_total_visual <= visibleLines Then
        thumbH = trackH
    Else
        thumbH = (visibleLines * trackH) \ g_console_total_visual
    End If
    If thumbH < CONSOLE_THUMB_MIN_H Then thumbH = CONSOLE_THUMB_MIN_H
    If thumbH > trackH Then thumbH = trackH
    ConsoleThumbHeight = thumbH
End Function

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
    IsMouseOverConsole = (x >= consX And x < consX + consW And y >= consY And y < consY + consH)
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
    ' D5 (2026-11-07): el chat/consola NO se escribe a disco; solo [MODE]/[OVERLAY]/diag/[PERF].
    Call PushOneRingEntry(cleanText, red, green, blue, bold, italic, appendToPrevious, cat)
    Exit Sub
ViewportDebug_PushConsoleLine_Err:
    ' No reportar via RegistrarError para evitar recursion (RegistrarError podria llamar consola).
    Resume Next
End Sub

Private Function WrapConsoleText(ByVal text As String, ByVal maxWidth As Long, ByRef outLines() As String) As Integer
    On Error GoTo WrapConsoleText_Err
    ReDim outLines(0 To 0) As String
    Dim count As Integer
    Dim rest As String
    Dim restLen As Long
    Dim takeLen As Long
    Dim lastSpace As Long
    Dim lineWidth As Long
    Dim ignoredGrh As Long
    Dim i As Long
    rest = Trim$(text)
    If LenB(rest) = 0 Then
        outLines(0) = " "
        WrapConsoleText = 1
        Exit Function
    End If

    ' E3: ancho acumulado una sola vez por linea (antes se re-media el prefijo completo
    ' al achicar de a un caracter). Usa la misma regla de glifos que el dibujo.
    Do While Len(rest) > 0
        restLen = Len(rest)
        takeLen = restLen
        lineWidth = 0
        For i = 1 To restLen
            lineWidth = lineWidth + ConsoleGlyphAdvance(Asc(Mid$(rest, i, 1)), ignoredGrh)
            If lineWidth > maxWidth Then
                takeLen = i - 1
                Exit For
            End If
        Next i
        If takeLen < 1 Then takeLen = 1

        If takeLen < restLen Then
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

' E3: garantiza que la entrada del anillo tenga el wrap vigente para ese ancho.
Private Sub EnsureEntryWrapped(ByVal ringIndex As Integer, ByVal maxWidth As Long)
    On Error GoTo EnsureEntryWrapped_Err
    If g_ring(ringIndex).wrapCount > 0 And g_ring(ringIndex).wrapWidth = maxWidth Then Exit Sub
    Dim tmpLines() As String
    Dim lineCount As Integer
    Dim k As Integer
    lineCount = WrapConsoleText(g_ring(ringIndex).text, maxWidth, tmpLines)
    ReDim g_ring(ringIndex).wrapLines(0 To lineCount - 1) As String
    For k = 0 To lineCount - 1
        g_ring(ringIndex).wrapLines(k) = tmpLines(k)
    Next k
    g_ring(ringIndex).wrapCount = lineCount
    g_ring(ringIndex).wrapWidth = maxWidth
    Exit Sub
EnsureEntryWrapped_Err:
    ' Sin RegistrarError (evita recursion con la consola). Deja la entrada como una sola linea.
    ReDim g_ring(ringIndex).wrapLines(0 To 0) As String
    g_ring(ringIndex).wrapLines(0) = g_ring(ringIndex).text
    g_ring(ringIndex).wrapCount = 1
    g_ring(ringIndex).wrapWidth = maxWidth
End Sub

' E3: regla unica de ancho de glifo de la consola. La usan el wrap (medir) y
' RenderConsoleGlyphLine (dibujar): espacio = 4 px, glifos <= 12 no se dibujan ni avanzan.
' grhIndex devuelve el grh a dibujar (0 = nada que dibujar).
Private Function ConsoleGlyphAdvance(ByVal charCode As Long, ByRef grhIndex As Long) As Long
    On Error Resume Next
    Dim charW As Long
    grhIndex = 0
    If charCode = 32 Then
        ConsoleGlyphAdvance = 4
        Exit Function
    End If
    grhIndex = Fuentes(1).Caracteres(charCode)
    If grhIndex <= 12 Then
        grhIndex = 0
        Exit Function
    End If
    charW = 0
    charW = GrhData(GrhData(grhIndex).Frames(1)).pixelWidth
    If charW <= 0 Then charW = GrhData(grhIndex).pixelWidth
    If charW <= 0 Then charW = 6
    ConsoleGlyphAdvance = charW
End Function

Private Sub RenderConsoleTextLine(ByVal text As String, _
                                  ByVal x As Long, _
                                  ByVal y As Long, _
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
    Dim grhIndex As Long
    Dim drawX As Long
    Dim advance As Long

    drawX = x

    For i = 1 To Len(text)
        advance = ConsoleGlyphAdvance(Asc(Mid$(text, i, 1)), grhIndex)
        If grhIndex > 0 Then Call Draw_GrhFont(grhIndex, drawX, y, textColor)
        drawX = drawX + advance
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
        ' E5: tope de largo por entrada (un mensaje que se sigue anexando sin fin crecia sin limite).
        Dim room As Long
        room = CONSOLE_ENTRY_MAX_CHARS - Len(g_ring(lastIndex).text)
        If room > 0 Then
            g_ring(lastIndex).text = g_ring(lastIndex).text & Left$(text, room)
            If Not g_ring(lastIndex).isNew Then g_ring(lastIndex).countedLines = g_ring(lastIndex).wrapCount
            g_ring(lastIndex).isNew = True
            g_ring(lastIndex).wrapCount = 0
            g_wrap_dirty = True
        End If
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
        .wrapCount = 0
        .isNew = True
        .countedLines = 0
    End With
    g_wrap_dirty = True
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
    g_wrap_dirty = True
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
    g_pill_drawn = False
    If Not g_viewport_full And Not g_overlay_visible Then Exit Sub

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
    ' E5: solo en FULL; con F2 en HUD clasico la consola real es el RichTextBox, no se duplica.
    If Not g_viewport_full Then Exit Sub
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
        Dim wrappedCount As Integer
        Dim wrappedIdx As Integer
        Dim wrapWidth As Long
        Dim renderedLines As Integer
        Dim maxLines As Integer
        maxLines = ConsoleVisibleLines(consH)

        Dim skippedLines As Long
        Dim totalVisual As Long
        Dim newLines As Long
        Call EnsureFiltersLoaded
        wrapWidth = consW - 24
        ' E3: el total de lineas visuales solo se recalcula si cambio el texto, el filtro o el ancho.
        If g_wrap_dirty Or g_wrap_total_width <> wrapWidth Then
            For i = 0 To g_ring_count - 1
                idx = (g_ring_head - 1 - i + RING_SIZE) Mod RING_SIZE
                If IsCategoryVisible(g_ring(idx).category) Then
                    Call EnsureEntryWrapped(idx, wrapWidth)
                    totalVisual = totalVisual + g_ring(idx).wrapCount
                    ' E6: only lines of entries added/extended since the last recompute count as new (a filter or
                    ' width change re-wraps old entries but never sets isNew).
                    If g_ring(idx).isNew Then
                        newLines = newLines + g_ring(idx).wrapCount - g_ring(idx).countedLines
                        g_ring(idx).countedLines = g_ring(idx).wrapCount
                    End If
                End If
                g_ring(idx).isNew = False
            Next i
            g_console_total_visual = totalVisual
            g_wrap_total_width = wrapWidth
            g_wrap_dirty = False
            ' E6: scroll counts from the newest line, so when lines arrive while scrolled up the offset grows by the
            ' same amount to keep the same text on screen, and the indicator is raised. At the bottom nothing changes.
            ' The count comes from the new lines themselves, so a full ring (old lines dropped) still works.
            If newLines > 0 And g_console_scroll_lines > 0 Then
                g_console_scroll_lines = g_console_scroll_lines + newLines
                g_console_unread = True
            End If
        End If
        Call ClampConsoleScroll
        If g_console_scroll_lines = 0 Then g_console_unread = False

        For i = 0 To g_ring_count - 1
            idx = (g_ring_head - 1 - i + RING_SIZE) Mod RING_SIZE
            With g_ring(idx)
                If Not IsCategoryVisible(.category) Then GoTo NextEntry
                Call RGBAList(lineColor, CByte(.red And 255), CByte(.green And 255), CByte(.blue And 255), 255)
                Call EnsureEntryWrapped(idx, wrapWidth)
                wrappedCount = .wrapCount
                For wrappedIdx = wrappedCount - 1 To 0 Step -1
                    If skippedLines < g_console_scroll_lines Then
                        skippedLines = skippedLines + 1
                    Else
                        If bottomY < minY Then Exit For
                        Call RenderConsoleTextLine(.wrapLines(wrappedIdx), consX + 5, bottomY, CByte(.red And 255), CByte(.green And 255), CByte(.blue And 255))
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
            trackY = consY + CONSOLE_TRACK_TOP
            trackH = consH - CONSOLE_TRACK_PAD
            Dim scrollTrack As RGBA
            scrollTrack = RGBA_From_Comp(220, 220, 220, 80)
            Call Engine_Draw_Box(trackX, trackY, 6, trackH, scrollTrack)

            Dim thumbH As Long, thumbY As Long, maxScrollLines As Long
            thumbH = ConsoleThumbHeight(trackH)
            thumbY = trackY
            maxScrollLines = GetConsoleMaxScroll()
            ' E5: scroll 0 = lineas mas nuevas = pulgar abajo; scroll maximo = lineas mas viejas = pulgar arriba.
            If maxScrollLines > 0 Then thumbY = trackY + ((maxScrollLines - g_console_scroll_lines) * (trackH - thumbH)) \ maxScrollLines
            Dim scrollThumb As RGBA
            scrollThumb = RGBA_From_Comp(235, 235, 235, 220)
            Call Engine_Draw_Box(trackX - 1, thumbY, 8, thumbH, scrollThumb)
        End If
        If g_console_unread Then Call Render_ConsoleUnreadPill(consX, consY, consW, consH)
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
    If g_log_disabled Then Exit Sub
    Dim f As Integer
    If Not g_log_initialized Then
        g_log_initialized = True
        Dim logDir As String
        logDir = App.path & "\logs"
        On Error Resume Next
        If LenB(Dir$(logDir, vbDirectory)) = 0 Then MkDir logDir
        If LenB(Dir$(logDir, vbDirectory)) = 0 Then g_log_disabled = True
        On Error GoTo ViewportDebug_AppendLog_Err
        If g_log_disabled Then Exit Sub
        g_log_path = logDir & "\viewport_debug.log"
        Call ViewportDebug_RotateLog(logDir)
    End If
    f = FreeFile
    Open g_log_path For Append As #f
    Print #f, format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & R & "," & G & "," & B & "] " & text
    Close #f
    Exit Sub
ViewportDebug_AppendLog_Err:
    ' No re-reportar: el log es best-effort. Sin Resume Next (evita Print/Close sobre un handle invalido).
    On Error Resume Next
    Close #f
End Sub

' Al primer uso de la sesion: si el log pesa mas de 1 MB, pasa a viewport_debug.old.log (pisa el anterior).
Private Sub ViewportDebug_RotateLog(ByVal logDir As String)
    On Error Resume Next
    If LenB(Dir$(g_log_path)) = 0 Then Exit Sub
    If FileLen(g_log_path) <= LOG_MAX_BYTES Then Exit Sub
    Dim oldPath As String
    oldPath = logDir & "\viewport_debug.old.log"
    If LenB(Dir$(oldPath)) > 0 Then Kill oldPath
    Name g_log_path As oldPath
End Sub

' ============== PERF (E0) ==============

' Milisegundos de un contador de alta resolucion. Devuelve 0 si la medicion esta apagada.
Public Function Perf_Now() As Double
    If Not PERF_ENABLED Then Exit Function
    Dim counter As Currency
    If g_perf_freq = 0 Then
        Dim freq As Currency
        If QueryPerformanceFrequency(freq) = 0 Then
            g_perf_freq = -1
        Else
            g_perf_freq = CDbl(freq)
        End If
    End If
    If g_perf_freq <= 0 Then Exit Function
    If QueryPerformanceCounter(counter) = 0 Then Exit Function
    Perf_Now = CDbl(counter) / g_perf_freq * 1000#
End Function

' Igual que Perf_Now pero solo mide mientras estamos en FULL (0 = no medir).
Public Function Perf_FrameStart() As Double
    If Not PERF_ENABLED Then Exit Function
    If Not g_viewport_full Then Exit Function
    Perf_FrameStart = Perf_Now()
End Function

Public Sub Perf_FrameAdd(ByVal piece As Long, ByVal startMs As Double)
    If startMs = 0 Then Exit Sub
    If piece < 0 Or piece > PERF_PIECE_MAX Then Exit Sub
    g_perf_acc(piece) = g_perf_acc(piece) + (Perf_Now() - startMs)
End Sub

Private Sub Perf_ResetFrameStats()
    Dim i As Long
    For i = 0 To PERF_PIECE_MAX
        g_perf_acc(i) = 0
    Next i
    g_perf_frames = 0
End Sub

' Llamar una vez por frame al final de engine.render.
Public Sub Perf_FrameEnd()
    If Not PERF_ENABLED Then Exit Sub
    If Not g_viewport_full Then
        If g_perf_frames > 0 Or g_perf_window_start <> 0 Then
            Call Perf_ResetFrameStats
            g_perf_window_start = 0
        End If
        Exit Sub
    End If
    Dim nowMs As Double
    nowMs = Perf_Now()
    If g_perf_window_start = 0 Then g_perf_window_start = nowMs
    g_perf_frames = g_perf_frames + 1
    If nowMs - g_perf_window_start < PERF_WINDOW_MS Then Exit Sub
    Dim n As Double
    n = g_perf_frames
    Call ViewportDebug_AppendDiagLog("[PERF] frame avg over " & g_perf_frames & " frames: hud=" & format$(g_perf_acc(PERF_PIECE_HUD) / n, "0.000") & _
        " console=" & format$(g_perf_acc(PERF_PIECE_CONSOLE) / n, "0.000") & _
        " buffs=" & format$(g_perf_acc(PERF_PIECE_BUFFS) / n, "0.000") & _
        " gmbar=" & format$(g_perf_acc(PERF_PIECE_GMBAR) / n, "0.000") & _
        " neighbors=" & format$(g_perf_acc(PERF_PIECE_NEIGHBORS) / n, "0.000") & _
        " buttons=" & format$(g_perf_acc(PERF_PIECE_BUTTONS) / n, "0.000") & _
        " fps=" & engine.fps & " (ms)")
    Call Perf_ResetFrameStats
    g_perf_window_start = nowMs
End Sub

' Una linea por cruce/cambio de mapa. startMs = Perf_Now() tomado al inicio de SwitchMap.
Public Sub Perf_LogSwitchMap(ByVal mapId As Long, ByVal startMs As Double, ByVal cargarMs As Double, ByVal vecinosMs As Double)
    If startMs = 0 Then Exit Sub
    Call ViewportDebug_AppendDiagLog("[PERF] SwitchMap map=" & mapId & " total=" & format$(Perf_Now() - startMs, "0.0") & "ms cargar=" & format$(cargarMs, "0.0") & "ms vecinos=" & format$(vecinosMs, "0.0") & "ms full=" & IIf(g_viewport_full, "1", "0"))
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

' isClick = False: solo informa si el cursor esta sobre un filtro (lo consume) sin alternarlo ni escribir el ini.
Public Function ViewportDebug_HitConsoleFilter(ByVal mx As Long, ByVal my As Long, ByVal consX As Long, ByVal consY As Long, ByVal isClick As Boolean) As Boolean
    Dim btnX As Long
    Dim btnY As Long
    btnY = consY + 3
    btnX = consX + 5
    If my < btnY Or my >= btnY + BTN_H Then Exit Function
    Dim i As Long
    For i = 0 To 2
        If mx >= btnX And mx < btnX + BTN_W Then
            If isClick Then
                Select Case i
                    Case 0: Call ViewportDebug_ToggleCategory(e_ConsoleCat_General)
                    Case 1: Call ViewportDebug_ToggleCategory(e_ConsoleCat_Combat)
                    Case 2: Call ViewportDebug_ToggleCategory(e_ConsoleCat_Global)
                End Select
            End If
            ViewportDebug_HitConsoleFilter = True
            Exit Function
        End If
        btnX = btnX + BTN_W + BTN_PAD_X
    Next i
End Function
