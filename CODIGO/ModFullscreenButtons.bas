Attribute VB_Name = "ModFullscreenButtons"
' Argentum 20 Game Client
'
'    Botonera FULL: 5 columnas x 3 filas.
'    Top row: seguros. Mid+bottom rows: menus.
'    Hover destaca el boton con el BMP -over. Default usa el BMP -default.
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

Public Sub Set_FullscreenButtonsActive(ByVal active As Boolean)
    On Error GoTo Err_h
    g_buttons_active = active
    g_hovered_button = -1
    If active Then
        Call Position_FullscreenButtons
        frmMain.picFullscreenButtons.Visible = True
        frmMain.picFullscreenButtons.ZOrder 0
        Call Render_FullscreenButtons
    Else
        frmMain.picFullscreenButtons.Visible = False
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

    With frmMain.picFullscreenButtons
        .BackColor = RGB(0, 0, 0)
        .Cls
        Dim partyBmp As String
        partyBmp = IIf(SeguroParty, "boton-seguro-party-on.bmp", "boton-seguro-party-off.bmp")
        Call DrawButton(efbSafeParty, partyBmp, partyBmp)
        Dim clanBmp As String
        clanBmp = IIf(SeguroClanX, "boton-seguro-clan-on.bmp", "boton-seguro-clan-off.bmp")
        Call DrawButton(efbSafeClan, clanBmp, clanBmp)
        Dim attBmp As String
        attBmp = IIf(SeguroGame, "boton-seguro-ciudadano-on.bmp", "boton-seguro-ciudadano-off.bmp")
        Call DrawButton(efbSafeAttack, attBmp, attBmp)
        Dim resuBmp As String
        resuBmp = IIf(SeguroResuX, "boton-fantasma-on.bmp", "boton-fantasma-off.bmp")
        Call DrawButton(efbSafeResu, resuBmp, resuBmp)
        Dim legBmp As String
        legBmp = IIf(LegionarySecureX, "boton-demonio-on.bmp", "boton-demonio-off.bmp")
        Call DrawButton(efbSafeLegion, legBmp, legBmp)
        Call DrawButton(efbGuild, "boton-clanes-over.bmp", "boton-clanes-over.bmp")
        Call DrawButton(efbChallenges, "boton-retos-default.bmp", "boton-retos-over.bmp")
        Call DrawButton(efbStats, "boton-estadisticas-default.bmp", "boton-estadisticas-over.bmp")
        Call DrawButton(efbConfig, "boton-ajustes-default.bmp", "boton-ajustes-over.bmp")
        Call DrawButton(efbSkills, "boton-estadisticas-default.bmp", "boton-estadisticas-over.bmp")
        Call DrawButton(efbInventory, "boton-inventory-default.bmp", "boton-inventory-over.bmp")
        Call DrawButton(efbSpells, "boton-hechizos-default.bmp", "boton-hechizos-over.bmp")
        Call DrawButton(efbKeys, "boton-llavero-off.bmp", "boton-llavero-over.bmp")
        Call DrawButton(efbParty, "boton-grupo-over.bmp", "boton-grupo-over.bmp")
        Call DrawButton(efbQuests, "boton-quests-over.bmp", "boton-quests-over.bmp")
        .Refresh
    End With
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.Render_FullscreenButtons", Erl)
    Resume Next
End Sub

Private Sub DrawButton(ByVal buttonId As e_FullButtonId, ByVal defaultBmp As String, ByVal hoverBmp As String)
    On Error Resume Next
    Dim bx As Long, by As Long
    bx = ButtonX(buttonId)
    by = ButtonY(buttonId)
    Dim bmp As String
    If g_hovered_button = buttonId Then
        bmp = hoverBmp
    Else
        bmp = defaultBmp
    End If
    Dim pic As IPictureDisp
    Set pic = LoadInterface(bmp)
    If pic Is Nothing Then
        Set pic = LoadInterface(defaultBmp)
    End If
    If Not (pic Is Nothing) Then
        frmMain.picFullscreenButtons.PaintPicture pic, bx, by, BUTTON_SIZE, BUTTON_SIZE
    End If
End Sub

Public Sub Handle_FullscreenButtonsMouseMove(ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    If Not g_buttons_active Then Exit Sub
    Dim newHover As Long
    newHover = HitTestButton(CLng(x), CLng(y))
    If newHover <> g_hovered_button Then
        g_hovered_button = newHover
        Call Render_FullscreenButtons
    End If
End Sub

Public Sub Reset_FullscreenButtonsHover()
    On Error Resume Next
    If Not g_buttons_active Then Exit Sub
    If g_hovered_button <> -1 Then
        g_hovered_button = -1
        Call Render_FullscreenButtons
    End If
End Sub

Public Sub Handle_FullscreenButtonsClick(ByVal x As Single, ByVal y As Single)
    On Error GoTo Err_h
    If Not g_buttons_active Then Exit Sub

    Dim buttonId As Long
    buttonId = HitTestButton(CLng(x), CLng(y))
    If buttonId < 0 Then Exit Sub

    Select Case buttonId
        Case efbInventory
            Call Toggle_FloatingWindow(efwInventory)
        Case efbSpells
            Call Toggle_FloatingWindow(efwSpells)
        Case efbKeys
            Call ToggleKeysWindow
        Case efbParty
            If FrmGrupo.Visible = False Then Call WriteRequestGrupo
        Case efbQuests
            If Not pausa Then Call WriteQuestListRequest
        Case efbGuild
            If Not pausa Then
                If frmGuildLeader.Visible Then Unload frmGuildLeader
                Call WriteRequestGuildLeaderInfo
            End If
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

Private Sub ToggleKeysWindow()
    On Error GoTo Err_h
    If FrmKeyInv.Visible Then
        Call frmMain.CerrarLlavero
    Else
        FrmKeyInv.Show , frmMain
        FrmKeyInv.Left = frmMain.Left + 890 * screen.TwipsPerPixelX - FrmKeyInv.Width \ 2
        FrmKeyInv.Top = frmMain.Top + 555 * screen.TwipsPerPixelY - FrmKeyInv.Height
        frmMain.cmdLlavero.Tag = "2"
    End If
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenButtons.ToggleKeysWindow", Erl)
    Resume Next
End Sub

Private Function HitTestButton(ByVal x As Long, ByVal y As Long) As Long
    Dim i As Long
    For i = 0 To efbCount - 1
        If x >= ButtonX(i) And x <= ButtonX(i) + BUTTON_SIZE Then
            If y >= ButtonY(i) And y <= ButtonY(i) + BUTTON_SIZE Then
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
