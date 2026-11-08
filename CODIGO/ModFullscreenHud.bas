Attribute VB_Name = "ModFullscreenHud"
' Argentum 20 Game Client
'
'    Render DX del HUD vital en modo viewport fullscreen.
'    V9.6: consola wrap por pixeles y minimapa con posicion.
'    Top-right: nombre del mapa y coords alineados a la derecha del minimapa (con truncado por ancho),
'    minimap con contorno, fuerza/agilidad, hora del mundo y oro alineados a derecha.
'
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md (Etapa 9)
'
Option Explicit

' === Constantes ===
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

' Top-right HUD (ancla FULL_MINIMAP_X/Y en ModFullscreenLayout)
Private Const MINIMAP_W        As Long = 100
Private Const MINIMAP_H        As Long = 100
Private Const MAPLABEL_Y       As Long = 40
Private Const MAPLABEL_MAX_W   As Long = 220

' === Estado ===
Private g_fullscreen_hud_active As Boolean

' E3: caches por valor (se reconstruyen solo cuando cambia el dato de origen).
Private g_circle_ready      As Boolean
Private g_circle_cached_r   As Long
Private g_circle_half()     As Long
Private g_circle_out_x()    As Long
Private g_circle_out_y()    As Long
Private g_circle_out_count  As Long
Private Const CIRCLE_OUTLINE_STEP_DEG As Long = 3

Private g_lvl_value         As Long
Private g_lvl_text          As String
Private g_stat_cur(0 To 1)  As Long
Private g_stat_max(0 To 1)  As Long
Private g_stat_text(0 To 1) As String
Private g_stat_width(0 To 1) As Long
Private g_fit_ready         As Boolean
Private g_fit_src           As String
Private g_fit_maxw          As Long
Private g_fit_result        As String
Private g_map_ready         As Boolean
Private g_map_src           As String
Private g_map_fit           As String
Private g_map_width         As Long
Private g_pos_ready         As Boolean
Private g_pos_x             As Long
Private g_pos_y             As Long
Private g_pos_text          As String
Private g_pos_width         As Long
Private g_str_text          As String
Private g_str_value         As Long
Private g_str_width         As Long
Private g_agi_text          As String
Private g_agi_value         As Long
Private g_agi_width         As Long
' E6: hora del mundo (misma fuente que el HUD clasico: frmMain.lblhora) y oro, con texto y ancho en cache.
Private g_time_src          As String
Private g_time_text         As String
Private g_time_width        As Long
Private g_gld_ready         As Boolean
Private g_gld_value         As Long
Private g_gld_over_limit    As Boolean
Private g_gld_text          As String
Private g_gld_width         As Long

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

    If LenB(g_lvl_text) = 0 Or g_lvl_value <> UserStats.Lvl Then
        g_lvl_value = UserStats.Lvl
        g_lvl_text = "Lvl " & g_lvl_value
    End If
    Call Engine_Text_Render(g_lvl_text, barX + EXP_BAR_W + 8, LINE_NAME_Y, white, 1, True)

    ' Linea 2: Vida
    Call Engine_Text_Render("Vida", tx, VIDA_LINE_Y - 1, dimColor, 1, False)
    Call DrawStatBar(barX, VIDA_LINE_Y, UserStats.MinHp, UserStats.MaxHp, 220, 30, 30, 0)

    ' Linea 3: Mana
    Call Engine_Text_Render("Mana", tx, MANA_LINE_Y - 1, dimColor, 1, False)
    Call DrawStatBar(barX, MANA_LINE_Y, UserStats.minman, UserStats.maxman, 50, 80, 220, 1)
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
    If UserCharIndex <= 0 Or UserCharIndex > LastChar Then Exit Sub
    Dim headGrh As Long
    headGrh = charlist(UserCharIndex).Head.Head(E_Heading.south).GrhIndex
    If headGrh <= 0 Or headGrh > MaxGrh Then Exit Sub
    Dim frameGrh As Long
    frameGrh = GrhData(headGrh).Frames(1)
    If frameGrh <= 0 Or frameGrh > MaxGrh Then Exit Sub
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
    Call EnsureCircleCache(r)
    Dim iy As Long, half As Long
    For iy = -r To r
        half = g_circle_half(iy + r)
        If half > 0 Then
            Call Engine_Draw_Box(cx - half, cy + iy, half * 2, 1, color)
        End If
    Next iy
End Sub

' E3: precalcula una sola vez (por radio) el ancho de cada scanline del disco y los
' desplazamientos del contorno (mismos puntos que antes: cada 3 grados, caja de 2x2).
Private Sub EnsureCircleCache(ByVal r As Long)
    On Error Resume Next
    If g_circle_ready And g_circle_cached_r = r Then Exit Sub
    Const PI_VAL As Double = 3.14159265358979
    Dim r2 As Long
    Dim iy As Long
    Dim a As Long
    Dim k As Long
    r2 = r * r
    ReDim g_circle_half(0 To 2 * r) As Long
    For iy = -r To r
        g_circle_half(iy + r) = CLng(Sqr(r2 - iy * iy))
    Next iy
    g_circle_out_count = 360 \ CIRCLE_OUTLINE_STEP_DEG
    ReDim g_circle_out_x(0 To g_circle_out_count - 1) As Long
    ReDim g_circle_out_y(0 To g_circle_out_count - 1) As Long
    For a = 0 To 359 Step CIRCLE_OUTLINE_STEP_DEG
        k = a \ CIRCLE_OUTLINE_STEP_DEG
        g_circle_out_x(k) = CLng(r * Cos(a * PI_VAL / 180)) - 1
        g_circle_out_y(k) = CLng(r * Sin(a * PI_VAL / 180)) - 1
    Next a
    g_circle_cached_r = r
    g_circle_ready = True
End Sub

Private Sub DrawCircleOutline(ByVal cx As Long, ByVal cy As Long, ByVal r As Long, ByRef color As RGBA)
    On Error Resume Next
    Call EnsureCircleCache(r)
    Dim k As Long
    For k = 0 To g_circle_out_count - 1
        Call Engine_Draw_Box(cx + g_circle_out_x(k), cy + g_circle_out_y(k), 2, 2, color)
    Next k
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
        ' E1: en Double (w * cur desbordaba Long con EXP alta).
        fillW = CLng(Int(CDbl(w) * CDbl(cur) / CDbl(maxVal)))
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

Private Sub DrawStatBar(ByVal x As Long, ByVal y As Long, ByVal current As Long, ByVal maxVal As Long, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal barId As Long)
    On Error Resume Next
    Dim color As RGBA
    color = RGBA_From_Comp(r, g, b, 235)
    Call DrawBar(x, y, STAT_BAR_W, STAT_BAR_H, current, maxVal, color)

    ' E3: texto "actual/max" y su ancho solo se rebuscan cuando cambia algun valor.
    If LenB(g_stat_text(barId)) = 0 Or g_stat_cur(barId) <> current Or g_stat_max(barId) <> maxVal Then
        g_stat_cur(barId) = current
        g_stat_max(barId) = maxVal
        g_stat_text(barId) = current & "/" & maxVal
        g_stat_width(barId) = Engine_Text_Width(g_stat_text(barId), True)
    End If
    Dim white(3) As RGBA
    Call RGBAList(white, 255, 255, 255, 255)
    Call Engine_Text_Render(g_stat_text(barId), x + (STAT_BAR_W - g_stat_width(barId)) \ 2, y - 1, white, 1, False)
End Sub

Private Function FitTextToWidth(ByVal text As String, ByVal maxWidth As Long, ByVal bold As Boolean) As String
    On Error Resume Next
    ' E3: cache por (texto, ancho): el nombre del mapa casi nunca cambia.
    If g_fit_ready And maxWidth = g_fit_maxw And text = g_fit_src Then
        FitTextToWidth = g_fit_result
        Exit Function
    End If
    g_fit_src = text
    g_fit_maxw = maxWidth
    g_fit_ready = True

    If HudTextWidth(text) <= maxWidth Then
        g_fit_result = text
        FitTextToWidth = text
        Exit Function
    End If

    Do While Len(text) > 4
        text = Left(text, Len(text) - 1)
        If HudTextWidth(text & "...") <= maxWidth Then
            g_fit_result = text & "..."
            FitTextToWidth = g_fit_result
            Exit Function
        End If
    Loop

    g_fit_result = Left(text, 1) & "..."
    FitTextToWidth = g_fit_result
End Function

' E3: ancho en pixeles con la MISMA regla con la que se dibuja (RenderHudGlyphLine).
' Engine_Text_Width contaba glifos que el dibujo saltea y el texto alineado a la derecha quedaba corrido.
Private Function HudTextWidth(ByVal text As String) As Long
    On Error Resume Next
    Dim i As Long
    Dim ignoredGrh As Long
    For i = 1 To Len(text)
        HudTextWidth = HudTextWidth + HudGlyphAdvance(Asc(Mid$(text, i, 1)), ignoredGrh)
    Next i
End Function

' Regla unica de glifo del HUD: espacio = 4 px; glifos <= 12 o fuera de MaxGrh no se dibujan ni avanzan.
' grhIndex devuelve el grh a dibujar (0 = nada que dibujar).
Private Function HudGlyphAdvance(ByVal charCode As Long, ByRef grhIndex As Long) As Long
    On Error Resume Next
    Dim charW As Long
    Dim frameIdx As Long
    grhIndex = 0
    If charCode = 32 Then
        HudGlyphAdvance = 4
        Exit Function
    End If
    grhIndex = Fuentes(1).Caracteres(charCode)
    If grhIndex <= 12 Or grhIndex > MaxGrh Then
        grhIndex = 0
        Exit Function
    End If
    charW = 0
    frameIdx = GrhData(grhIndex).Frames(1)
    If frameIdx > 0 And frameIdx <= MaxGrh Then charW = GrhData(frameIdx).pixelWidth
    If charW <= 0 Then charW = GrhData(grhIndex).pixelWidth
    If charW <= 0 Then charW = 6
    HudGlyphAdvance = charW
End Function

Private Sub RenderHudTextLine(ByVal text As String, _
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

    Call RenderHudGlyphLine(text, x + 1, y + 1, shadowColor)
    Call RenderHudGlyphLine(text, x, y, textColor)
End Sub

Private Sub RenderHudGlyphLine(ByVal text As String, _
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
        advance = HudGlyphAdvance(Asc(Mid$(text, i, 1)), grhIndex)
        If grhIndex > 0 Then Call Draw_GrhFont(grhIndex, drawX, y, textColor)
        drawX = drawX + advance
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

    ' E3: etiqueta de mapa, coordenadas y stats se rearman solo si cambia el dato de origen.
    If Not g_map_ready Or mapNameTxt <> g_map_src Then
        g_map_src = mapNameTxt
        g_map_fit = FitTextToWidth(mapNameTxt, MAPLABEL_MAX_W, False)
        g_map_width = HudTextWidth(g_map_fit)
        g_map_ready = True
    End If
    Call RenderHudTextLine(g_map_fit, FULL_MINIMAP_X + MINIMAP_W - g_map_width, MAPLABEL_Y, 255, 255, 255)

    If Not g_pos_ready Or UserPos.x <> g_pos_x Or UserPos.y <> g_pos_y Then
        g_pos_x = UserPos.x
        g_pos_y = UserPos.y
        g_pos_text = g_pos_x & ":" & g_pos_y
        g_pos_width = HudTextWidth(g_pos_text)
        g_pos_ready = True
    End If
    Call RenderHudTextLine(g_pos_text, FULL_MINIMAP_X + MINIMAP_W - g_pos_width, MAPLABEL_Y + 14, 255, 255, 255)
    ' Contorno del minimapa (2px de grosor)
    Dim borderColor As RGBA
    borderColor = RGBA_From_Comp(200, 170, 90, 240)
    Call Engine_Draw_Box(FULL_MINIMAP_X - 2, FULL_MINIMAP_Y - 2, MINIMAP_W + 4, 2, borderColor)
    Call Engine_Draw_Box(FULL_MINIMAP_X - 2, FULL_MINIMAP_Y + MINIMAP_H, MINIMAP_W + 4, 2, borderColor)
    Call Engine_Draw_Box(FULL_MINIMAP_X - 2, FULL_MINIMAP_Y - 2, 2, MINIMAP_H + 4, borderColor)
    Call Engine_Draw_Box(FULL_MINIMAP_X + MINIMAP_W, FULL_MINIMAP_Y - 2, 2, MINIMAP_H + 4, borderColor)

    ' Fuerza y Agilidad alineadas a la derecha bajo el minimapa
    Dim statsY As Long
    statsY = FULL_MINIMAP_Y + MINIMAP_H + 10

    If LenB(g_str_text) = 0 Or g_str_value <> UserStats.str Then
        g_str_value = UserStats.str
        g_str_text = "Fuerza: " & g_str_value
        g_str_width = HudTextWidth(g_str_text)
    End If
    Call RenderHudTextLine(g_str_text, FULL_MINIMAP_X + MINIMAP_W - g_str_width, statsY, 255, 255, 255)

    If LenB(g_agi_text) = 0 Or g_agi_value <> UserStats.Agi Then
        g_agi_value = UserStats.Agi
        g_agi_text = "Agilidad: " & g_agi_value
        g_agi_width = HudTextWidth(g_agi_text)
    End If
    Call RenderHudTextLine(g_agi_text, FULL_MINIMAP_X + MINIMAP_W - g_agi_width, statsY + 16, 255, 255, 255)

    ' E6: hora del mundo y oro bajo Agilidad (mismo estilo/alineacion). El oro usa el color del HUD clasico:
    ' amarillo normal, rojo si supera el limite por nivel.
    Dim timeCaption As String
    timeCaption = frmMain.lblhora.Caption
    If LenB(g_time_text) = 0 Or timeCaption <> g_time_src Then
        g_time_src = timeCaption
        g_time_text = "Hora: " & timeCaption
        g_time_width = HudTextWidth(g_time_text)
    End If
    Call RenderHudTextLine(g_time_text, FULL_MINIMAP_X + MINIMAP_W - g_time_width, statsY + 32, 255, 255, 255)

    Dim overLimit As Boolean
    overLimit = (UserStats.GLD > UserStats.Lvl * UserStats.OroPorNivel)
    If Not g_gld_ready Or g_gld_value <> UserStats.GLD Or g_gld_over_limit <> overLimit Then
        g_gld_ready = True
        g_gld_value = UserStats.GLD
        g_gld_over_limit = overLimit
        g_gld_text = "Oro: " & PonerPuntos(g_gld_value)
        g_gld_width = HudTextWidth(g_gld_text)
    End If
    If g_gld_over_limit Then
        Call RenderHudTextLine(g_gld_text, FULL_MINIMAP_X + MINIMAP_W - g_gld_width, statsY + 48, 255, 0, 0)
    Else
        Call RenderHudTextLine(g_gld_text, FULL_MINIMAP_X + MINIMAP_W - g_gld_width, statsY + 48, 255, 255, 128)
    End If
    Exit Sub
Render_HudTopRight_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenHud.Render_HudTopRight", Erl)
    Resume Next
End Sub
