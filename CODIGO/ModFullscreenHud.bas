Attribute VB_Name = "ModFullscreenHud"
' Argentum 20 Game Client
'
'    Render DX del HUD vital en modo viewport fullscreen.
'    V9.6: consola wrap por pixeles y minimapa con posicion.
'    Top-right: mapa y coords centrados con truncado por ancho, minimap con contorno,
'    hora centrada, fuerza/agilidad alineadas a derecha.
'
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md (Etapa 9)
'
Option Explicit

' === Constantes ===
Private Const GMBAR_HEIGHT     As Long = 22

' Top-left HUD
Private Const HUD_LEFT_X       As Long = 10
Private Const HUD_LEFT_Y       As Long = 28
Private Const AVATAR_W         As Long = 60
Private Const AVATAR_H         As Long = 60
Private Const AVATAR_GAP_X     As Long = 10

Private Const LINE_NAME_Y      As Long = 28
Private Const EXP_LINE_Y       As Long = 30
Private Const VIDA_LINE_Y      As Long = 54
Private Const MANA_LINE_Y      As Long = 72

Private Const LABEL_W          As Long = 40
Private Const STAT_BAR_W       As Long = 170
Private Const EXP_BAR_W        As Long = 80
Private Const EXP_BAR_H        As Long = 6
Private Const STAT_BAR_H       As Long = 12

' Top-right HUD
Private Const MINIMAP_X        As Long = 884
Private Const MINIMAP_Y        As Long = 70
Private Const MINIMAP_W        As Long = 100
Private Const MINIMAP_H        As Long = 100
Private Const MAPLABEL_Y       As Long = 40
Private Const MAPLABEL_MAX_W   As Long = 220

' === Estado ===
Private g_fullscreen_hud_active As Boolean

Public Sub Set_FullscreenHudActive(ByVal active As Boolean)
    g_fullscreen_hud_active = active
End Sub

Public Sub Render_FullscreenHud()
    On Error GoTo Render_FullscreenHud_Err
    If Not g_fullscreen_hud_active Then Exit Sub
    If g_game_state Is Nothing Then Exit Sub
    If g_game_state.state <> e_state_gameplay_screen Then Exit Sub
    Call Render_HudTopLeft
    Call Render_HudTopRight
    Exit Sub
Render_FullscreenHud_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenHud.Render_FullscreenHud", Erl)
    Resume Next
End Sub

Private Sub Render_HudTopLeft()
    On Error GoTo Render_HudTopLeft_Err
    Dim white(3) As RGBA
    Call RGBAList(white, 255, 255, 255, 255)
    Dim dimColor(3) As RGBA
    Call RGBAList(dimColor, 200, 200, 200, 255)

    Dim x As Long, y As Long
    x = HUD_LEFT_X
    y = HUD_LEFT_Y

    Call Render_HudAvatar(x, y, AVATAR_W, AVATAR_H)

    Dim tx As Long
    tx = x + AVATAR_W + AVATAR_GAP_X

    Dim barX As Long
    barX = tx + LABEL_W

    ' Linea 1: Nombre + Exp compacta + Lvl
    Call Engine_Text_Render(userName, tx, LINE_NAME_Y, white, 1, True)

    Dim expColor As RGBA
    expColor = RGBA_From_Comp(255, 215, 70, 235)
    Call DrawBar(barX, EXP_LINE_Y, EXP_BAR_W, EXP_BAR_H, UserStats.exp, UserStats.PasarNivel, expColor)

    Dim lvlTxt As String
    lvlTxt = "Lvl " & UserStats.Lvl
    Call Engine_Text_Render(lvlTxt, barX + EXP_BAR_W + 8, LINE_NAME_Y, white, 1, True)

    ' Linea 2: Vida
    Call Engine_Text_Render("Vida", tx, VIDA_LINE_Y - 1, dimColor, 1, False)
    Call DrawStatBar(barX, VIDA_LINE_Y, UserStats.MinHp, UserStats.MaxHp, 220, 30, 30)

    ' Linea 3: Mana
    Call Engine_Text_Render("Mana", tx, MANA_LINE_Y - 1, dimColor, 1, False)
    Call DrawStatBar(barX, MANA_LINE_Y, UserStats.minman, UserStats.maxman, 50, 80, 220)
    Exit Sub
Render_HudTopLeft_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenHud.Render_HudTopLeft", Erl)
    Resume Next
End Sub

Private Sub Render_HudAvatar(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    On Error Resume Next
    Dim cx As Long, cy As Long, r As Long
    cx = x + w \ 2
    cy = y + h \ 2
    r = (w \ 2) - 1

    Dim bgDisc As RGBA
    bgDisc = RGBA_From_Comp(15, 15, 15, 220)
    Call DrawFilledCircle(cx, cy, r, bgDisc)

    Call Render_HudHeadZoomed(cx, cy, r)

    Dim borderColor As RGBA
    borderColor = RGBA_From_Comp(200, 170, 90, 240)
    Call DrawCircleOutline(cx, cy, r, borderColor)
End Sub

Private Sub Render_HudHeadZoomed(ByVal cx As Long, ByVal cy As Long, ByVal r As Long)
    On Error Resume Next
    If UserCharIndex <= 0 Then Exit Sub
    Dim headGrh As Long
    headGrh = charlist(UserCharIndex).Head.Head(E_Heading.south).GrhIndex
    If headGrh <= 0 Or headGrh > MaxGrh Then Exit Sub
    Dim frameGrh As Long
    frameGrh = GrhData(headGrh).Frames(1)
    If frameGrh <= 0 Then Exit Sub
    Dim white(3) As RGBA
    Call RGBAList(white, 255, 255, 255, 255)
    Const ZOOM As Single = 2.6
    With GrhData(frameGrh)
        If .Tx2 = 0 And .FileNum > 0 Then
            Dim tex As Direct3DTexture8
            Dim tw As Long, th As Long
            Set tex = SurfaceDB.GetTexture(.FileNum, tw, th)
            .Tx1 = .sX / tw
            .Tx2 = (.sX + .pixelWidth) / tw
            .Ty1 = .sY / th
            .Ty2 = (.sY + .pixelHeight) / th
        End If
        Dim drawW As Long, drawH As Long
        drawW = CLng(.pixelWidth * ZOOM)
        drawH = CLng(.pixelHeight * ZOOM)
        Dim drawX As Long, drawY As Long
        drawX = cx - drawW \ 2
        drawY = cy - drawH \ 2 + CLng(.pixelHeight * ZOOM * 0.18)
        Call Batch_Textured_Box_Pre(drawX, drawY, .pixelWidth, .pixelHeight, .Tx1, .Ty1, .Tx2, .Ty2, .FileNum, white, True, 0, ZOOM, ZOOM)
    End With
End Sub

Public Function HudAvatar_HitTest(ByVal mx As Long, ByVal my As Long) As Boolean
    If Not g_fullscreen_hud_active Then Exit Function
    Dim cx As Long, cy As Long, r As Long
    cx = HUD_LEFT_X + AVATAR_W \ 2
    cy = HUD_LEFT_Y + AVATAR_H \ 2
    r = (AVATAR_W \ 2) - 1
    Dim dx As Long, dy As Long
    dx = mx - cx
    dy = my - cy
    HudAvatar_HitTest = ((dx * dx + dy * dy) <= r * r)
End Function

Public Sub HudAvatar_OpenStats()
    On Error Resume Next
    LlegaronAtrib = False
    LlegaronStats = False
    Call WriteRequestAtributes
    Call WriteRequestMiniStats
End Sub

Private Sub DrawFilledCircle(ByVal cx As Long, ByVal cy As Long, ByVal r As Long, ByRef color As RGBA)
    On Error Resume Next
    Dim r2 As Long
    r2 = r * r
    Dim iy As Long, half As Long
    For iy = -r To r
        half = CLng(Sqr(r2 - iy * iy))
        If half > 0 Then
            Call Engine_Draw_Box(cx - half, cy + iy, half * 2, 1, color)
        End If
    Next iy
End Sub

Private Sub DrawCircleOutline(ByVal cx As Long, ByVal cy As Long, ByVal r As Long, ByRef color As RGBA)
    On Error Resume Next
    Const PI_VAL As Double = 3.14159265358979
    Dim a As Long
    Dim px As Long, py As Long
    For a = 0 To 359 Step 3
        px = cx + CLng(r * Cos(a * PI_VAL / 180)) - 1
        py = cy + CLng(r * Sin(a * PI_VAL / 180)) - 1
        Call Engine_Draw_Box(px, py, 2, 2, color)
    Next a
End Sub

Private Sub DrawBar(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal current As Long, ByVal maxVal As Long, ByRef color As RGBA)
    On Error Resume Next
    Dim bg As RGBA
    bg = RGBA_From_Comp(30, 30, 30, 220)
    Call Engine_Draw_Box(x, y, w, h, bg)

    Dim fillW As Long
    If maxVal > 0 Then
        Dim cur As Long
        cur = current
        If cur < 0 Then cur = 0
        If cur > maxVal Then cur = maxVal
        fillW = w * cur \ maxVal
    End If
    If fillW > 0 Then
        Call Engine_Draw_Box(x, y, fillW, h, color)
    End If

    Dim borderC As RGBA
    borderC = RGBA_From_Comp(0, 0, 0, 220)
    Call Engine_Draw_Box(x, y, w, 1, borderC)
    Call Engine_Draw_Box(x, y + h - 1, w, 1, borderC)
    Call Engine_Draw_Box(x, y, 1, h, borderC)
    Call Engine_Draw_Box(x + w - 1, y, 1, h, borderC)
End Sub

Private Sub DrawStatBar(ByVal x As Long, ByVal y As Long, ByVal current As Long, ByVal maxVal As Long, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
    On Error Resume Next
    Dim color As RGBA
    color = RGBA_From_Comp(r, g, b, 235)
    Call DrawBar(x, y, STAT_BAR_W, STAT_BAR_H, current, maxVal, color)

    Dim numText As String
    numText = current & "/" & maxVal
    Dim text_w As Long
    text_w = Engine_Text_Width(numText, True)
    Dim white(3) As RGBA
    Call RGBAList(white, 255, 255, 255, 255)
    Call Engine_Text_Render(numText, x + (STAT_BAR_W - text_w) \ 2, y - 1, white, 1, False)
End Sub

Private Function FitTextToWidth(ByVal text As String, ByVal maxWidth As Long, ByVal bold As Boolean) As String
    On Error Resume Next
    If Engine_Text_Width(text, False) <= maxWidth Then
        FitTextToWidth = text
        Exit Function
    End If

    Do While Len(text) > 4
        text = Left(text, Len(text) - 1)
        If Engine_Text_Width(text & "...", False) <= maxWidth Then
            FitTextToWidth = text & "..."
            Exit Function
        End If
    Loop

    FitTextToWidth = Left(text, 1) & "..."
End Function

Private Sub RenderHudTextLine(ByVal text As String, _
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

    Call RenderHudGlyphLine(text, x + 1, y + 1, shadowColor)
    Call RenderHudGlyphLine(text, x, y, textColor)
End Sub

Private Sub RenderHudGlyphLine(ByVal text As String, _
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

Private Sub Render_HudTopRight()
    On Error GoTo Render_HudTopRight_Err
    Dim white(3) As RGBA
    Call RGBAList(white, 255, 255, 255, 255)

    ' Etiqueta sobre el minimapa con truncado para nombres largos
    Dim mapNameTxt As String
    mapNameTxt = vbNullString
    On Error Resume Next
    mapNameTxt = MapDat.map_name
    On Error GoTo Render_HudTopRight_Err
    If LenB(mapNameTxt) = 0 Then mapNameTxt = "Mapa " & UserMap

    mapNameTxt = FitTextToWidth(mapNameTxt, MAPLABEL_MAX_W, False)

    Dim mapW As Long
    mapW = Engine_Text_Width(mapNameTxt, False)
    Dim mapX As Long
    mapX = MINIMAP_X + MINIMAP_W - mapW
    Call RenderHudTextLine(mapNameTxt, mapX, MAPLABEL_Y, MAPLABEL_MAX_W, 255, 255, 255)

    Dim posTxt As String
    posTxt = UserPos.x & ":" & UserPos.y
    Dim posW As Long
    posW = Engine_Text_Width(posTxt, False)
    Call RenderHudTextLine(posTxt, MINIMAP_X + MINIMAP_W - posW, MAPLABEL_Y + 14, 60, 255, 255, 255)
    ' Contorno del minimapa (2px de grosor)
    Dim borderColor As RGBA
    borderColor = RGBA_From_Comp(200, 170, 90, 240)
    Call Engine_Draw_Box(MINIMAP_X - 2, MINIMAP_Y - 2, MINIMAP_W + 4, 2, borderColor)
    Call Engine_Draw_Box(MINIMAP_X - 2, MINIMAP_Y + MINIMAP_H, MINIMAP_W + 4, 2, borderColor)
    Call Engine_Draw_Box(MINIMAP_X - 2, MINIMAP_Y - 2, 2, MINIMAP_H + 4, borderColor)
    Call Engine_Draw_Box(MINIMAP_X + MINIMAP_W, MINIMAP_Y - 2, 2, MINIMAP_H + 4, borderColor)

    ' Fuerza y Agilidad alineadas a la derecha bajo el minimapa
    Dim statsY As Long
    statsY = MINIMAP_Y + MINIMAP_H + 10

    Dim fuerzaTxt As String
    fuerzaTxt = "Fuerza: " & UserStats.str
    Dim fuerzaW As Long
    fuerzaW = Engine_Text_Width(fuerzaTxt, False)
    Call RenderHudTextLine(fuerzaTxt, MINIMAP_X + MINIMAP_W - fuerzaW, statsY, 120, 255, 255, 255)

    Dim agiTxt As String
    agiTxt = "Agilidad: " & UserStats.Agi
    Dim agiW As Long
    agiW = Engine_Text_Width(agiTxt, False)
    Call RenderHudTextLine(agiTxt, MINIMAP_X + MINIMAP_W - agiW, statsY + 16, 120, 255, 255, 255)
    Exit Sub
Render_HudTopRight_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenHud.Render_HudTopRight", Erl)
    Resume Next
End Sub
