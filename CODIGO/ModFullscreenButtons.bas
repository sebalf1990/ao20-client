Attribute VB_Name = "ModFullscreenButtons"
' Argentum 20 Game Client
'
'    Botonera FULL: 5 columnas x 3 filas.
'    Top row: seguros. Mid+bottom rows: menus.
'    Estados visuales (ver docs/cliente/viewport-fullscreen-hud-flotante.md):
'      - normal: BMP -default (o -over atenuado si el boton no tiene -default).
'      - seguro apagado: BMP -off atenuado; encendido: BMP -on a brillo completo.
'      - hover: BMP -over (o el mismo BMP) a brillo completo + brillo extra + borde dorado.
'    Tooltip DX (nombre del boton) tras una demora, encima del panel. Click en un menu abierto lo cierra.
'    Posicion de los menus: se recuerda en Configuracion.ini [FloatingWindows] (ModFloatingWindows).
'
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md (Etapa 9)
'
Option Explicit

Private Const BUTTON_SIZE As Long = 32
Private Const BUTTON_GAP  As Long = 4
Private Const BUTTON_COLS As Long = 5
Private Const BUTTON_ROWS As Long = 3
Private Const PANEL_PAD   As Long = 2
Private Const PANEL_W     As Long = PANEL_PAD * 2 + BUTTON_COLS * BUTTON_SIZE + (BUTTON_COLS - 1) * BUTTON_GAP
Private Const PANEL_H     As Long = PANEL_PAD * 2 + BUTTON_ROWS * BUTTON_SIZE + (BUTTON_ROWS - 1) * BUTTON_GAP

' Alphas (0-255) de las capas que se mezclan sobre el BMP del boton.
Private Const DIM_ALPHA_NO_DEFAULT As Long = 60  ' normal de los botones que solo tienen BMP -over
Private Const DIM_ALPHA_SAFE_OFF   As Long = 70  ' seguro apagado
Private Const HOVER_LIGHT_ALPHA    As Long = 45  ' brillo extra del hover
Private Const BLEND_ALPHA_SHIFT    As Long = &H10000 ' SourceConstantAlpha es el 3er byte de BLENDFUNCTION

' Tooltip
Private Const TOOLTIP_DELAY_MS As Long = 400
Private Const TOOLTIP_H        As Long = 20
Private Const TOOLTIP_PAD_X    As Long = 6
Private Const TOOLTIP_TEXT_Y   As Long = 4
Private Const TOOLTIP_GAP      As Long = 4
Private Const TOOLTIP_EDGE     As Long = 4

' Posicion de los menus: se guarda cuando lleva este tiempo quieta.
Private Const MENU_SAVE_DELAY_MS As Long = 300

Private Type t_ButtonsPoint
    x As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As t_ButtonsPoint) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As t_ButtonsPoint) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal blendFunction As Long) As Long

' Layout enum: indice del 0 al 14 mapea a posicion row*5+col
Private Enum e_FullButtonId
    efbSafeParty = 0
    efbSafeClan = 1
    efbSafeAttack = 2
    efbSafeResu = 3
    efbSafeLegion = 4
    efbGuild = 5
    efbChallenges = 6
    efbStats = 7
    efbConfig = 8
    efbSkills = 9
    efbInventory = 10
    efbSpells = 11
    efbKeys = 12
    efbParty = 13
    efbQuests = 14
    efbCount = 15
End Enum

Private g_buttons_active As Boolean
Private g_hovered_button As Long

' Tooltip: instante en que arranco el hover actual; un click lo oculta hasta que cambie el hover.
Private g_hover_tick          As Long
Private g_tooltip_suppressed  As Boolean

' Capas de mezcla (DC de 1x1 pixel, negro y blanco) para AlphaBlend sobre el PictureBox.
Private g_dim_dc    As Long
Private g_dim_bmp   As Long
Private g_dim_old   As Long
Private g_light_dc  As Long
Private g_light_bmp As Long
Private g_light_old As Long

' Menus abiertos desde la botonera: ultimo estado visto por Update_FullscreenButtonsFrame.
Private g_menu_open(0 To efbCount - 1)       As Boolean
Private g_menu_last_x(0 To efbCount - 1)     As Long
Private g_menu_last_y(0 To efbCount - 1)     As Long
Private g_menu_dirty(0 To efbCount - 1)      As Boolean
Private g_menu_dirty_tick(0 To efbCount - 1) As Long

' E3: cache de los BMP de los botones (antes LoadInterface releia/descomprimia 15 por repintado).
' Se vacia si cambia el idioma (los BMP son localizados).
Private g_pic_cache      As Collection
Private g_pic_cache_lang As Long

Public Sub Set_FullscreenButtonsActive(ByVal active As Boolean)
    On Error GoTo Err_h
    g_buttons_active = active
    g_hovered_button = -1
    g_tooltip_suppressed = False
    If active Then
        Call ResetMenuTracking
        Call Position_FullscreenButtons
        frmMain.picFullscreenButtons.Visible = True
        frmMain.picFullscreenButtons.ZOrder 0
        Call Render_FullscreenButtons
    Else
        Call FlushMenuPositions
        frmMain.picFullscreenButtons.Visible = False
        Call DestroyBlendLayers
    End If
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.Set_FullscreenButtonsActive", Erl)
    Resume Next
End Sub

Public Sub Refresh_FullscreenButtons()
    On Error Resume Next
    If g_buttons_active Then Call Render_FullscreenButtons
End Sub

Private Sub Position_FullscreenButtons()
    On Error Resume Next
    frmMain.picFullscreenButtons.Width = PANEL_W
    frmMain.picFullscreenButtons.Height = PANEL_H
    frmMain.picFullscreenButtons.Left = g_viewport_logical_left + g_viewport_logical_width - PANEL_W - 10
    frmMain.picFullscreenButtons.Top = g_viewport_logical_top + g_viewport_logical_height - PANEL_H - 8
End Sub

Public Sub Render_FullscreenButtons()
    On Error GoTo Err_h
    If Not g_buttons_active Then Exit Sub
    Dim perfT As Double
    perfT = Perf_FrameStart()
    Call EnsureBlendLayers

    With frmMain.picFullscreenButtons
        .BackColor = RGB(0, 0, 0)
        .Cls
        ' Seguros: ON a brillo completo, OFF atenuado (ademas de la tilde/cruz del BMP).
        Dim partyBmp As String
        partyBmp = IIf(SeguroParty, "boton-seguro-party-on.bmp", "boton-seguro-party-off.bmp")
        Call DrawButton(efbSafeParty, partyBmp, partyBmp, IIf(SeguroParty, 0, DIM_ALPHA_SAFE_OFF))
        Dim clanBmp As String
        clanBmp = IIf(SeguroClanX, "boton-seguro-clan-on.bmp", "boton-seguro-clan-off.bmp")
        Call DrawButton(efbSafeClan, clanBmp, clanBmp, IIf(SeguroClanX, 0, DIM_ALPHA_SAFE_OFF))
        Dim attBmp As String
        attBmp = IIf(SeguroGame, "boton-seguro-ciudadano-on.bmp", "boton-seguro-ciudadano-off.bmp")
        Call DrawButton(efbSafeAttack, attBmp, attBmp, IIf(SeguroGame, 0, DIM_ALPHA_SAFE_OFF))
        Dim resuBmp As String
        resuBmp = IIf(SeguroResuX, "boton-fantasma-on.bmp", "boton-fantasma-off.bmp")
        Call DrawButton(efbSafeResu, resuBmp, resuBmp, IIf(SeguroResuX, 0, DIM_ALPHA_SAFE_OFF))
        Dim legBmp As String
        legBmp = IIf(LegionarySecureX, "boton-demonio-on.bmp", "boton-demonio-off.bmp")
        Call DrawButton(efbSafeLegion, legBmp, legBmp, IIf(LegionarySecureX, 0, DIM_ALPHA_SAFE_OFF))
        ' Menus. Clanes, Grupo, Quests y Llavero no tienen BMP -default (Llavero: el -off es mucho mas oscuro
        ' que el resto): su normal es el -over atenuado y el hover es el -over completo.
        Call DrawButton(efbGuild, "boton-clanes-over.bmp", "boton-clanes-over.bmp", DIM_ALPHA_NO_DEFAULT)
        Call DrawButton(efbChallenges, "boton-retos-default.bmp", "boton-retos-over.bmp", 0)
        Call DrawButton(efbStats, "boton-estadisticas-default.bmp", "boton-estadisticas-over.bmp", 0)
        Call DrawButton(efbConfig, "boton-ajustes-default.bmp", "boton-ajustes-over.bmp", 0)
        Call DrawButton(efbSkills, "boton-estadisticas-default.bmp", "boton-estadisticas-over.bmp", 0)
        Call DrawButton(efbInventory, "boton-inventory-default.bmp", "boton-inventory-over.bmp", 0)
        Call DrawButton(efbSpells, "boton-hechizos-default.bmp", "boton-hechizos-over.bmp", 0)
        Call DrawButton(efbKeys, "boton-llavero-over.bmp", "boton-llavero-over.bmp", DIM_ALPHA_NO_DEFAULT)
        Call DrawButton(efbParty, "boton-grupo-over.bmp", "boton-grupo-over.bmp", DIM_ALPHA_NO_DEFAULT)
        Call DrawButton(efbQuests, "boton-quests-over.bmp", "boton-quests-over.bmp", DIM_ALPHA_NO_DEFAULT)
        .Refresh
    End With
    Call Perf_FrameAdd(PERF_PIECE_BUTTONS, perfT)
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.Render_FullscreenButtons", Erl)
    Resume Next
End Sub

' dimAlpha: atenuado (negro con ese alpha) del estado normal; 0 = sin atenuar. El hover siempre va a brillo completo.
Private Sub DrawButton(ByVal buttonId As e_FullButtonId, ByVal defaultBmp As String, ByVal hoverBmp As String, ByVal dimAlpha As Long)
    On Error Resume Next
    Dim bx As Long, by As Long
    bx = ButtonX(buttonId)
    by = ButtonY(buttonId)
    Dim isHovered As Boolean
    isHovered = (g_hovered_button = buttonId)
    Dim bmp As String
    If isHovered Then
        bmp = hoverBmp
    Else
        bmp = defaultBmp
    End If
    Dim pic As IPictureDisp
    Set pic = GetButtonPicture(bmp)
    If pic Is Nothing Then
        Set pic = GetButtonPicture(defaultBmp)
    End If
    If pic Is Nothing Then Exit Sub
    frmMain.picFullscreenButtons.PaintPicture pic, bx, by, BUTTON_SIZE, BUTTON_SIZE
    If isHovered Then
        ' Si el hover usa el mismo BMP que el normal no hay diferencia de imagen: se suma brillo.
        If StrComp(hoverBmp, defaultBmp, vbTextCompare) = 0 Then
            Call BlendLayer(g_light_dc, bx, by, HOVER_LIGHT_ALPHA)
        End If
        Call DrawHoverBorder(bx, by)
    ElseIf dimAlpha > 0 Then
        Call BlendLayer(g_dim_dc, bx, by, dimAlpha)
    End If
End Sub

Private Sub DrawHoverBorder(ByVal bx As Long, ByVal by As Long)
    On Error Resume Next
    Dim borderColor As Long
    borderColor = RGB(255, 210, 90)
    frmMain.picFullscreenButtons.Line (bx, by)-(bx + BUTTON_SIZE - 1, by + BUTTON_SIZE - 1), borderColor, B
    frmMain.picFullscreenButtons.Line (bx + 1, by + 1)-(bx + BUTTON_SIZE - 2, by + BUTTON_SIZE - 2), borderColor, B
End Sub

' Mezcla la capa (DC de 1 pixel) sobre el cuadro del boton con el alpha indicado. Sin capa o sin alpha no hace nada.
Private Sub BlendLayer(ByVal layerDc As Long, ByVal bx As Long, ByVal by As Long, ByVal layerAlpha As Long)
    On Error Resume Next
    If layerDc = 0 Then Exit Sub
    If layerAlpha <= 0 Then Exit Sub
    Call AlphaBlend(frmMain.picFullscreenButtons.hdc, bx, by, BUTTON_SIZE, BUTTON_SIZE, layerDc, 0, 0, 1, 1, layerAlpha * BLEND_ALPHA_SHIFT)
End Sub

Private Sub EnsureBlendLayers()
    On Error Resume Next
    If g_dim_dc = 0 Then g_dim_dc = CreateSolidDc(RGB(0, 0, 0), g_dim_bmp, g_dim_old)
    If g_light_dc = 0 Then g_light_dc = CreateSolidDc(RGB(255, 255, 255), g_light_bmp, g_light_old)
End Sub

Private Sub DestroyBlendLayers()
    On Error Resume Next
    Call DestroySolidDc(g_dim_dc, g_dim_bmp, g_dim_old)
    Call DestroySolidDc(g_light_dc, g_light_bmp, g_light_old)
End Sub

' Crea un DC en memoria con un bitmap de 1x1 pixel del color dado (0 si falla).
Private Function CreateSolidDc(ByVal fillColor As Long, ByRef hBmp As Long, ByRef hOld As Long) As Long
    Dim screenDc As Long
    Dim memDc As Long
    screenDc = GetDC(0)
    memDc = CreateCompatibleDC(screenDc)
    If memDc <> 0 Then
        hBmp = CreateCompatibleBitmap(screenDc, 1, 1)
        If hBmp <> 0 Then
            hOld = SelectObject(memDc, hBmp)
            Call SetPixelV(memDc, 0, 0, fillColor)
            CreateSolidDc = memDc
        Else
            Call DeleteDC(memDc)
        End If
    End If
    Call ReleaseDC(0, screenDc)
End Function

Private Sub DestroySolidDc(ByRef hDc As Long, ByRef hBmp As Long, ByRef hOld As Long)
    If hDc <> 0 Then
        Call SelectObject(hDc, hOld)
        Call DeleteObject(hBmp)
        Call DeleteDC(hDc)
    End If
    hDc = 0
    hBmp = 0
    hOld = 0
End Sub

Private Function GetButtonPicture(ByVal bmpName As String) As IPictureDisp
    On Error Resume Next
    If g_pic_cache Is Nothing Then Set g_pic_cache = New Collection
    If g_pic_cache_lang <> language Then
        Set g_pic_cache = New Collection
        g_pic_cache_lang = language
    End If
    Dim cached As IPictureDisp
    Set cached = g_pic_cache.Item(bmpName)
    If cached Is Nothing Then
        Err.Clear
        Set cached = LoadInterface(bmpName)
        If Not (cached Is Nothing) Then g_pic_cache.Add cached, bmpName
    End If
    Set GetButtonPicture = cached
End Function

Public Sub Handle_FullscreenButtonsMouseMove(ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    If Not g_buttons_active Then Exit Sub
    Dim newHover As Long
    newHover = HitTestButton(CLng(x), CLng(y))
    If newHover <> g_hovered_button Then
        Call SetHoveredButton(newHover)
        Call Render_FullscreenButtons
    End If
End Sub

' Cambia el boton resaltado y reinicia el temporizador del tooltip.
Private Sub SetHoveredButton(ByVal buttonId As Long)
    g_hovered_button = buttonId
    g_hover_tick = GetTickCount And &H7FFFFFFF
    g_tooltip_suppressed = False
End Sub

' E5: hook por frame (engine.render, antes de Engine_EndScene): detecta menus movidos/abiertos, limpia el hover
' si el cursor salio del panel (el PictureBox no tiene MouseLeave) y dibuja el tooltip DX.
Public Sub Update_FullscreenButtonsFrame()
    On Error Resume Next
    If Not g_buttons_active Then Exit Sub
    Call UpdateMenuPositions
    Call UpdateHoverLeave
    Call Render_ButtonTooltip
End Sub

Private Sub UpdateHoverLeave()
    On Error Resume Next
    If g_hovered_button = -1 Then Exit Sub
    Dim pt As t_ButtonsPoint
    If GetCursorPos(pt) = 0 Then Exit Sub
    Call ScreenToClient(frmMain.picFullscreenButtons.hWnd, pt)
    If pt.x < 0 Or pt.y < 0 Or pt.x >= PANEL_W Or pt.y >= PANEL_H Then Call Reset_FullscreenButtonsHover
End Sub

Public Sub Reset_FullscreenButtonsHover()
    On Error Resume Next
    If Not g_buttons_active Then Exit Sub
    If g_hovered_button <> -1 Then
        Call SetHoveredButton(-1)
        Call Render_FullscreenButtons
    End If
End Sub

Private Sub Render_ButtonTooltip()
    On Error GoTo Err_h
    If g_hovered_button < 0 Then Exit Sub
    If g_tooltip_suppressed Then Exit Sub
    Dim elapsed As Long
    elapsed = (GetTickCount And &H7FFFFFFF) - g_hover_tick
    ' elapsed < 0 solo si el contador dio la vuelta (24 dias de uptime): se muestra igual.
    If elapsed >= 0 And elapsed < TOOLTIP_DELAY_MS Then Exit Sub
    Dim tipText As String
    tipText = ButtonTooltip(g_hovered_button)
    If LenB(tipText) = 0 Then Exit Sub

    Dim tipW As Long
    Dim tipX As Long
    Dim tipY As Long
    Dim maxX As Long
    tipW = Engine_Text_Width(tipText, False, 1) + TOOLTIP_PAD_X * 2
    ' Centrado sobre el boton, encima del panel (el PictureBox GDI tapa lo que se dibuje dentro de su area).
    tipX = frmMain.picFullscreenButtons.Left + ButtonX(g_hovered_button) + BUTTON_SIZE \ 2 - tipW \ 2
    maxX = frmMain.renderer.ScaleWidth - tipW - TOOLTIP_EDGE
    If tipX > maxX Then tipX = maxX
    If tipX < TOOLTIP_EDGE Then tipX = TOOLTIP_EDGE
    tipY = frmMain.picFullscreenButtons.Top - TOOLTIP_H - TOOLTIP_GAP

    Dim bgColor As RGBA
    Dim borderColor As RGBA
    Dim textColor(3) As RGBA
    bgColor = RGBA_From_Comp(15, 15, 15, 235)
    borderColor = RGBA_From_Comp(200, 170, 90, 255)
    Call RGBAList(textColor, 255, 255, 255, 255)
    Call Engine_Draw_Box_Border(tipX, tipY, tipW, TOOLTIP_H, bgColor, borderColor)
    Call Engine_Text_Render(tipText, tipX + TOOLTIP_PAD_X, tipY + TOOLTIP_TEXT_Y, textColor, 1, False)
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.Render_ButtonTooltip", Erl)
    Resume Next
End Sub

Private Function ButtonTooltip(ByVal buttonId As Long) As String
    Select Case buttonId
        Case efbSafeParty
            ButtonTooltip = "Seguro de grupo" & SafeStateText(SeguroParty)
        Case efbSafeClan
            ButtonTooltip = "Seguro de clan" & SafeStateText(SeguroClanX)
        Case efbSafeAttack
            ButtonTooltip = "Seguro de ataque" & SafeStateText(SeguroGame)
        Case efbSafeResu
            ButtonTooltip = "Seguro de resurrección" & SafeStateText(SeguroResuX)
        Case efbSafeLegion
            ButtonTooltip = "Seguro de legión" & SafeStateText(LegionarySecureX)
        Case efbGuild
            ButtonTooltip = "Clanes"
        Case efbChallenges
            ButtonTooltip = "Retos"
        Case efbStats
            ButtonTooltip = "Estadísticas"
        Case efbConfig
            ButtonTooltip = "Configuración"
        Case efbSkills
            ButtonTooltip = "Habilidades"
        Case efbInventory
            ButtonTooltip = "Inventario"
        Case efbSpells
            ButtonTooltip = "Hechizos"
        Case efbKeys
            ButtonTooltip = "Llavero"
        Case efbParty
            ButtonTooltip = "Grupo"
        Case efbQuests
            ButtonTooltip = "Quests"
    End Select
End Function

Private Function SafeStateText(ByVal isOn As Boolean) As String
    If isOn Then
        SafeStateText = " (activado)"
    Else
        SafeStateText = " (desactivado)"
    End If
End Function

Public Sub Handle_FullscreenButtonsClick(ByVal x As Single, ByVal y As Single)
    On Error GoTo Err_h
    If Not g_buttons_active Then Exit Sub

    Dim buttonId As Long
    buttonId = HitTestButton(CLng(x), CLng(y))
    If buttonId < 0 Then Exit Sub
    g_tooltip_suppressed = True

    ' Toggle: segundo click sobre un menu abierto lo cierra (Inventario/Hechizos se alternan en Toggle_FloatingWindow).
    If CloseMenuIfOpen(buttonId) Then
        Call Render_FullscreenButtons
        Exit Sub
    End If

    Select Case buttonId
        Case efbInventory
            Call Toggle_FloatingWindow(efwInventory)
        Case efbSpells
            Call Toggle_FloatingWindow(efwSpells)
        Case efbKeys
            Call OpenKeysWindow
        Case efbParty
            Call WriteRequestGrupo
        Case efbQuests
            If Not pausa Then Call WriteQuestListRequest
        Case efbGuild
            If Not pausa Then Call WriteRequestGuildLeaderInfo
        Case efbChallenges
            Call ParseUserCommand("/RETAR")
        Case efbStats
            LlegaronAtrib = False
            LlegaronStats = False
            Call WriteRequestAtributes
            Call WriteRequestMiniStats
        Case efbSkills
            Call RequestSkills
        Case efbConfig
            frmOpciones.Show , frmMain
        Case efbSafeParty
            Call WriteParyToggle
        Case efbSafeClan
            Call WriteSeguroClan
        Case efbSafeAttack
            Call WriteSafeToggle
        Case efbSafeResu
            Call WriteSeguroResu
        Case efbSafeLegion
            Call WriteLegionarySecure
    End Select

    Call Render_FullscreenButtons
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.Handle_FullscreenButtonsClick", Erl)
    Resume Next
End Sub

Private Sub OpenKeysWindow()
    On Error GoTo Err_h
    FrmKeyInv.Show , frmMain
    ' Posicion recordada si existe; si no, E5: la ventana del llavero se ancla al boton del llavero
    ' (centrada encima), no a 890/555 fijos.
    If Not ApplySavedMenuPos(FrmKeyInv, efbKeys) Then
        Dim keysCenterX As Long, keysTopY As Long
        keysCenterX = frmMain.picFullscreenButtons.Left + ButtonX(efbKeys) + BUTTON_SIZE \ 2
        keysTopY = frmMain.picFullscreenButtons.Top + ButtonY(efbKeys)
        FrmKeyInv.Left = frmMain.Left + keysCenterX * screen.TwipsPerPixelX - FrmKeyInv.Width \ 2
        FrmKeyInv.Top = frmMain.Top + keysTopY * screen.TwipsPerPixelY - FrmKeyInv.Height
    End If
    Call TrackMenuOpened(efbKeys, FrmKeyInv)
    frmMain.cmdLlavero.Tag = "2"
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.OpenKeysWindow", Erl)
    Resume Next
End Sub

' ---------------------------------------------------------------------------------------------
' Menus de la botonera: toggle y posicion recordada (solo en FULL: en HUD clasico los menus
' conservan su posicion original y el layout es otro).
' La posicion se guarda en pixeles relativos a la esquina de frmMain, en
' Configuracion.ini [FloatingWindows] claves <NombreForm>Left/<NombreForm>Top.
' ---------------------------------------------------------------------------------------------

' Nombre del formulario que abre cada boton ("" si el boton no abre un form).
Private Function MenuFormName(ByVal buttonId As Long) As String
    Select Case buttonId
        Case efbGuild
            MenuFormName = "frmGuildLeader"
        Case efbChallenges
            MenuFormName = "frmRetos"
        Case efbStats
            MenuFormName = "frmStatistics"
        Case efbSkills
            MenuFormName = "frmEstadisticas"
        Case efbConfig
            MenuFormName = "frmOpciones"
        Case efbKeys
            MenuFormName = "FrmKeyInv"
        Case efbParty
            MenuFormName = "FrmGrupo"
        Case efbQuests
            MenuFormName = "FrmQuests"
    End Select
End Function

' Busca el form entre los cargados. Referenciar un form por su nombre lo carga; recorrer Forms no.
Private Function FindMenuForm(ByVal buttonId As Long) As Form
    Dim wantedName As String
    wantedName = MenuFormName(buttonId)
    If LenB(wantedName) = 0 Then Exit Function
    Dim candidate As Form
    For Each candidate In Forms
        If StrComp(candidate.Name, wantedName, vbTextCompare) = 0 Then
            Set FindMenuForm = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function CloseMenuIfOpen(ByVal buttonId As Long) As Boolean
    On Error GoTo Err_h
    Dim menuForm As Form
    Set menuForm = FindMenuForm(buttonId)
    If menuForm Is Nothing Then Exit Function
    If Not menuForm.Visible Then Exit Function
    If buttonId = efbKeys Then
        Call frmMain.CerrarLlavero
    Else
        Unload menuForm
    End If
    Set menuForm = Nothing
    CloseMenuIfOpen = True
    Exit Function
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.CloseMenuIfOpen", Erl)
    Resume Next
End Function

Private Sub CurrentMenuPos(ByVal menuForm As Form, ByRef posX As Long, ByRef posY As Long)
    posX = CLng((menuForm.Left - frmMain.Left) / screen.TwipsPerPixelX)
    posY = CLng((menuForm.Top - frmMain.Top) / screen.TwipsPerPixelY)
End Sub

' Mueve el form a la posicion guardada (acotada a frmMain por si cambio el monitor/ventana).
' Devuelve False si no hay posicion guardada.
Private Function ApplySavedMenuPos(ByVal menuForm As Form, ByVal buttonId As Long) As Boolean
    On Error GoTo Err_h
    Dim savedX As Long
    Dim savedY As Long
    If Not FloatingWindows_LoadMenuPos(MenuFormName(buttonId), savedX, savedY) Then Exit Function
    Dim newLeft As Single
    Dim newTop As Single
    Dim maxLeft As Single
    Dim maxTop As Single
    newLeft = frmMain.Left + savedX * screen.TwipsPerPixelX
    newTop = frmMain.Top + savedY * screen.TwipsPerPixelY
    maxLeft = frmMain.Left + frmMain.Width - menuForm.Width
    maxTop = frmMain.Top + frmMain.Height - menuForm.Height
    If newLeft > maxLeft Then newLeft = maxLeft
    If newTop > maxTop Then newTop = maxTop
    If newLeft < frmMain.Left Then newLeft = frmMain.Left
    If newTop < frmMain.Top Then newTop = frmMain.Top
    Call menuForm.Move(newLeft, newTop)
    ApplySavedMenuPos = True
    Exit Function
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.ApplySavedMenuPos", Erl)
    Resume Next
End Function

' Marca el menu como abierto con la posicion actual como base (sin pendiente de guardar).
Private Sub TrackMenuOpened(ByVal buttonId As Long, ByVal menuForm As Form)
    Dim posX As Long
    Dim posY As Long
    Call CurrentMenuPos(menuForm, posX, posY)
    g_menu_open(buttonId) = True
    g_menu_last_x(buttonId) = posX
    g_menu_last_y(buttonId) = posY
    g_menu_dirty(buttonId) = False
End Sub

Private Sub ResetMenuTracking()
    Dim buttonId As Long
    For buttonId = 0 To efbCount - 1
        g_menu_open(buttonId) = False
        g_menu_dirty(buttonId) = False
    Next buttonId
End Sub

Private Sub PersistMenuPos(ByVal buttonId As Long)
    Call FloatingWindows_SaveMenuPos(MenuFormName(buttonId), g_menu_last_x(buttonId), g_menu_last_y(buttonId))
    g_menu_dirty(buttonId) = False
End Sub

' Al salir de FULL: guarda lo que haya quedado pendiente.
Private Sub FlushMenuPositions()
    On Error Resume Next
    Dim buttonId As Long
    For buttonId = 0 To efbCount - 1
        If g_menu_dirty(buttonId) Then Call PersistMenuPos(buttonId)
        g_menu_open(buttonId) = False
    Next buttonId
End Sub

' Una vez por frame: aplica la posicion guardada cuando un menu se abre (lo abra el click o la respuesta del
' server) y guarda la nueva posicion cuando el jugador lo movio y quedo quieto MENU_SAVE_DELAY_MS.
Private Sub UpdateMenuPositions()
    On Error GoTo Err_h
    Dim buttonId As Long
    Dim menuForm As Form
    Dim isOpen As Boolean
    Dim posX As Long
    Dim posY As Long
    Dim nowTick As Long
    Dim stillMs As Long
    nowTick = GetTickCount And &H7FFFFFFF
    For buttonId = 0 To efbCount - 1
        If LenB(MenuFormName(buttonId)) > 0 Then
            Set menuForm = FindMenuForm(buttonId)
            isOpen = False
            If Not (menuForm Is Nothing) Then isOpen = menuForm.Visible
            If isOpen Then
                If Not g_menu_open(buttonId) Then
                    ' Recien abierto: restaura la posicion guardada (si hay).
                    Call ApplySavedMenuPos(menuForm, buttonId)
                    Call TrackMenuOpened(buttonId, menuForm)
                Else
                    Call CurrentMenuPos(menuForm, posX, posY)
                    If posX <> g_menu_last_x(buttonId) Or posY <> g_menu_last_y(buttonId) Then
                        g_menu_last_x(buttonId) = posX
                        g_menu_last_y(buttonId) = posY
                        g_menu_dirty(buttonId) = True
                        g_menu_dirty_tick(buttonId) = nowTick
                    ElseIf g_menu_dirty(buttonId) Then
                        stillMs = nowTick - g_menu_dirty_tick(buttonId)
                        If stillMs < 0 Or stillMs >= MENU_SAVE_DELAY_MS Then Call PersistMenuPos(buttonId)
                    End If
                End If
            ElseIf g_menu_open(buttonId) Then
                ' Cerrado (descargado u oculto): guarda lo pendiente con la ultima posicion vista.
                g_menu_open(buttonId) = False
                If g_menu_dirty(buttonId) Then Call PersistMenuPos(buttonId)
            End If
            Set menuForm = Nothing
        End If
    Next buttonId
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.UpdateMenuPositions", Erl)
    Resume Next
End Sub

Private Function HitTestButton(ByVal x As Long, ByVal y As Long) As Long
    Dim i As Long
    For i = 0 To efbCount - 1
        If x >= ButtonX(i) And x < ButtonX(i) + BUTTON_SIZE Then
            If y >= ButtonY(i) And y < ButtonY(i) + BUTTON_SIZE Then
                HitTestButton = i
                Exit Function
            End If
        End If
    Next i
    HitTestButton = -1
End Function

Private Function ButtonX(ByVal buttonId As Long) As Long
    ButtonX = PANEL_PAD + (buttonId Mod BUTTON_COLS) * (BUTTON_SIZE + BUTTON_GAP)
End Function

Private Function ButtonY(ByVal buttonId As Long) As Long
    ButtonY = PANEL_PAD + (buttonId \ BUTTON_COLS) * (BUTTON_SIZE + BUTTON_GAP)
End Function
