Attribute VB_Name = "ModMinimap"
Option Explicit

' DirectX-rendered dot per slot (0=player, 1-5=allies) used when CenteredMinimap <> 0
Public Type MinimapDotState
    visible  As Boolean
    screenX  As Long
    screenY  As Long
    dotColor(0 To 3) As RGBA
End Type
Public MinimapDots(0 To 5) As MinimapDotState

Public Sub DrawCenteredMinimap()
    Call RenderMinimapCentered(UserMap, UserPos.x, UserPos.y, CenteredMinimapZoom, CenteredMinimapZoom)
End Sub

Public Sub RenderMinimapCentered(ByVal currentMap As Integer, ByVal tileX As Integer, ByVal tileY As Integer, Optional ByVal viewDeltaW As Long = 0, Optional ByVal viewDeltaH As Long = 0)
    On Error GoTo RenderMinimap_Err
    Static colorsInit As Boolean
    If Not colorsInit Then
        Call InitMinimapDotColors
        colorsInit = True
    End If
    Dim i        As Integer
    Dim j        As Byte
    Dim idmap    As Integer
    Dim worldNum As Byte
    Dim mapGridX As Long, mapGridY As Long
    ' Find which world and grid position the current map is in
    worldNum = 0
    idmap = 0
    For j = 1 To TotalWorlds
        For i = 1 To Mundo(j).Ancho * Mundo(j).Alto
            If Mundo(j).MapIndice(i) = currentMap Then
                idmap = i
                worldNum = j
                Exit For
            End If
        Next i
        If idmap > 0 Then Exit For
    Next j
    If idmap = 0 Then Exit Sub
    ' Ensure destination units are pixels
    frmMain.MiniMap.ScaleMode = vbPixels
    ' Load/cached world texture into the DirectX surface manager
    Static lastWorld    As Byte
    Static worldFileNum As Integer
    Static bmpPxW       As Long
    Static bmpPxH       As Long
        Dim minimapFile As String
        Select Case worldNum
            Case 1: minimapFile = "mapa1_200x200.bmp"
            Case 2: minimapFile = "mapa2_200x200.bmp"
            Case Else: minimapFile = ""
        End Select
        If minimapFile = "" Then Exit Sub
        worldFileNum = -CInt(worldNum)
        Call SurfaceDB.GetInterfaceTexture(minimapFile, bmpPxW, bmpPxH)
        If bmpPxW = 0 Or bmpPxH = 0 Then
            worldFileNum = 0
            Exit Sub
        End If
        lastWorld = worldNum
    If worldFileNum = 0 Then Exit Sub
    ' Grid of maps in the world image
    Dim mapCellsX As Long, mapCellsY As Long
    mapCellsX = Mundo(worldNum).Ancho   ' e.g., 100
    mapCellsY = Mundo(worldNum).Alto    ' e.g., 100
    ' Size of one map cell in pixels on the world image
    Dim mapCellPxW As Double: mapCellPxW = bmpPxW / mapCellsX
    Dim mapCellPxH As Double: mapCellPxH = bmpPxH / mapCellsY
    ' Current map's grid coordinates on the world image
    mapGridX = (idmap - 1) Mod mapCellsX
    mapGridY = (idmap - 1) \ mapCellsX
    ' Usable tile ranges inside a map
    ' Usable tile ranges inside a map (using module-level constants)
    Dim tileCountX   As Long, tileCountY As Long
    tileCountX = MINIMAP_TILE_COUNT_X ' 74 tiles
    tileCountY = MINIMAP_TILE_COUNT_Y ' 80 tiles
    ' Clamp incoming tile to valid range, just in case
    If tileX < MINIMAP_MIN_TILE_X Then tileX = MINIMAP_MIN_TILE_X
    If tileX > MINIMAP_MAX_TILE_X Then tileX = MINIMAP_MAX_TILE_X
    If tileY < MINIMAP_MIN_TILE_Y Then tileY = MINIMAP_MIN_TILE_Y
    If tileY > MINIMAP_MAX_TILE_Y Then tileY = MINIMAP_MAX_TILE_Y
    ' Player pixel center on the world image (tile center = offset + 0.5)
    Dim centerPxX As Long: centerPxX = CLng((mapGridX + (tileX - MINIMAP_MIN_TILE_X + 0.5) / tileCountX) * mapCellPxW)
    Dim centerPxY As Long: centerPxY = CLng((mapGridY + (tileY - MINIMAP_MIN_TILE_Y + 0.5) / tileCountY) * mapCellPxH)
    ' Destination size (control size)
    Dim destW As Long, destH As Long
    destW = frmMain.MiniMap.ScaleWidth
    destH = frmMain.MiniMap.ScaleHeight
    ' Clamp the configurable deltas to [-50, +50]
    If viewDeltaW < -50 Then viewDeltaW = -50
    If viewDeltaW > 50 Then viewDeltaW = 50
    If viewDeltaH < -50 Then viewDeltaH = -50
    If viewDeltaH > 50 Then viewDeltaH = 50
    ' Source crop size (zoom). Start from dest size and adjust by deltas.
    ' Smaller src => zoom in; larger src => zoom out. Keep sane minimum.
    Dim srcW As Long, srcH As Long
    srcW = destW + viewDeltaW
    srcH = destH + viewDeltaH
    ' Ensure positive and not exceeding bitmap
    If srcW < 16 Then srcW = 16           ' minimum crop width
    If srcH < 16 Then srcH = 16           ' minimum crop height
    If srcW > bmpPxW Then srcW = bmpPxW
    If srcH > bmpPxH Then srcH = bmpPxH
    ' Source top-left so that the player is centered in the source crop
    Dim srcX As Long, srcY As Long
    srcX = centerPxX - srcW \ 2
    srcY = centerPxY - srcH \ 2
    ' Clamp to bitmap bounds based on source crop size
    If srcX < 0 Then srcX = 0
    If srcY < 0 Then srcY = 0
    If srcX > (bmpPxW - srcW) Then srcX = bmpPxW - srcW
    If srcY > (bmpPxH - srcH) Then srcY = bmpPxH - srcH
    ' Draw: use DirectX 8 rendering instead of slow PaintPicture
    
    Dim DestRect As RECT
    
    
    DestRect.Bottom = destH
    DestRect.Right = destW
    DestRect.Left = 0
    DestRect.Top = 0
    
    MinimapVP_SrcX = srcX
    MinimapVP_SrcY = srcY
    MinimapVP_SrcW = srcW
    MinimapVP_SrcH = srcH
    MinimapVP_DestW = destW
    MinimapVP_DestH = destH
    MinimapVP_MapGridX = mapGridX
    MinimapVP_MapGridY = mapGridY
    MinimapVP_CellPxW = mapCellPxW
    MinimapVP_CellPxH = mapCellPxH
    
    Call Engine_BeginScene
    
    Call Batch_Textured_Box_File(0, 0, srcW, srcH, srcX, srcY, minimapFile, COLOR_WHITE, False, 0, CSng(destW) / CSng(srcW), CSng(destH) / CSng(srcH))
       
    Call RenderMinimapRadar   ' plan radar-v2 E5: debajo del punto propio
    Call RenderMinimapDots
       
    Call Engine_EndScene(DestRect, frmMain.MiniMap.hWnd)
    
    Exit Sub
RenderMinimap_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModMinimap.RenderMinimapCentered", Erl)
End Sub

' AO20: DibujarMiniMapa lives in ModUtils.bas (NPC radar, plan 15.001 phase 3).

Public Sub RenderMinimapDots()
    On Error GoTo RenderMinimapDots_Err
    
    Dim i As Long
    Dim squareSize As Long
    squareSize = 6  ' Size of each square in pixels (adjust as needed)
    
    ' Loop through all minimap dots
    For i = 0 To 5
        If MinimapDots(i).visible Then
            Dim Col As RGBA
            ' Draw the square using the dot's position and color
            Call Batch_Textured_Box_File( _
                MinimapDots(i).screenX - (squareSize \ 2), _
                MinimapDots(i).screenY - (squareSize \ 2), _
                1, 1, 0, 0, _
                "white_pixel.png", _
                MinimapDots(i).dotColor, _
                False, 0, _
                squareSize, squareSize)
        End If
    Next i
    
    Exit Sub
    
RenderMinimapDots_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.RenderMinimapDots", Erl)
End Sub


Public Sub InitMinimapDotColors()
    SetRGBA MinimapDots(0).dotColor(0), 255, 255, 255     ' player  : white
    SetRGBA MinimapDots(0).dotColor(1), 255, 255, 255
    SetRGBA MinimapDots(0).dotColor(2), 255, 255, 255
    SetRGBA MinimapDots(0).dotColor(3), 255, 255, 255
    
    SetRGBA MinimapDots(1).dotColor(0), 255, 255, 0      ' ally 1  : yellow
    SetRGBA MinimapDots(1).dotColor(1), 255, 255, 0
    SetRGBA MinimapDots(1).dotColor(2), 255, 255, 0
    SetRGBA MinimapDots(1).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(2).dotColor(0), 0, 192, 0        ' ally 2  : green
    SetRGBA MinimapDots(2).dotColor(1), 0, 192, 0
    SetRGBA MinimapDots(2).dotColor(2), 0, 192, 0
    SetRGBA MinimapDots(2).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(3).dotColor(0), 255, 128, 0      ' ally 3  : orange
    SetRGBA MinimapDots(3).dotColor(1), 255, 128, 0
    SetRGBA MinimapDots(3).dotColor(2), 255, 128, 0
    SetRGBA MinimapDots(3).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(4).dotColor(0), 255, 0, 255      ' ally 4  : magenta
    SetRGBA MinimapDots(4).dotColor(1), 255, 0, 255
    SetRGBA MinimapDots(4).dotColor(2), 255, 0, 255
    SetRGBA MinimapDots(4).dotColor(3), 255, 255, 0
    
    SetRGBA MinimapDots(5).dotColor(0), 0, 0, 255        ' ally 5  : blue
    SetRGBA MinimapDots(5).dotColor(1), 0, 0, 255
    SetRGBA MinimapDots(5).dotColor(2), 0, 0, 255
    SetRGBA MinimapDots(5).dotColor(3), 255, 255, 0
End Sub

' ============================================================================
' Plan radar-v2 E5: radar en el minimapa centrado (DirectX). Mismas capas, colores y tamano de punto
' que el minimapa clasico (ModUtils.DibujarMiniMapa). Los iconos (quest, sacerdote/banquero, jefe)
' van como puntos del color de su simbolo; el jefe, como punto grande.
' ============================================================================

' Casillero del mapa -> pixel del control en el recorte actual (MinimapVP_*). False si queda afuera.
Public Function MinimapTileToScreen(ByVal tileX As Long, ByVal tileY As Long, ByRef screenX As Long, ByRef screenY As Long) As Boolean
    If MinimapVP_SrcW <= 0 Or MinimapVP_SrcH <= 0 Then Exit Function
    Dim worldX As Double, worldY As Double
    worldX = (MinimapVP_MapGridX + (tileX - MINIMAP_MIN_TILE_X + 0.5) / MINIMAP_TILE_COUNT_X) * MinimapVP_CellPxW
    worldY = (MinimapVP_MapGridY + (tileY - MINIMAP_MIN_TILE_Y + 0.5) / MINIMAP_TILE_COUNT_Y) * MinimapVP_CellPxH
    screenX = CLng((worldX - MinimapVP_SrcX) * MinimapVP_DestW / MinimapVP_SrcW)
    screenY = CLng((worldY - MinimapVP_SrcY) * MinimapVP_DestH / MinimapVP_SrcH)
    MinimapTileToScreen = (screenX >= 0 And screenY >= 0 And screenX < MinimapVP_DestW And screenY < MinimapVP_DestH)
End Function

' Pixel del control -> casillero del mapa (inverso de MinimapTileToScreen). False sin recorte valido.
Public Function MinimapScreenToTile(ByVal screenX As Single, ByVal screenY As Single, ByRef tileX As Single, ByRef tileY As Single) As Boolean
    If MinimapVP_SrcW <= 0 Or MinimapVP_SrcH <= 0 Or MinimapVP_DestW <= 0 Or MinimapVP_DestH <= 0 Then Exit Function
    If MinimapVP_CellPxW <= 0 Or MinimapVP_CellPxH <= 0 Then Exit Function
    Dim worldX As Double, worldY As Double
    worldX = MinimapVP_SrcX + screenX * MinimapVP_SrcW / MinimapVP_DestW
    worldY = MinimapVP_SrcY + screenY * MinimapVP_SrcH / MinimapVP_DestH
    tileX = CSng(MINIMAP_MIN_TILE_X + (worldX / MinimapVP_CellPxW - MinimapVP_MapGridX) * MINIMAP_TILE_COUNT_X - 0.5)
    tileY = CSng(MINIMAP_MIN_TILE_Y + (worldY / MinimapVP_CellPxH - MinimapVP_MapGridY) * MINIMAP_TILE_COUNT_Y - 0.5)
    If tileX < 1 Then tileX = 1
    If tileX > 100 Then tileX = 100
    If tileY < 1 Then tileY = 1
    If tileY > 100 Then tileY = 100
    MinimapScreenToTile = True
End Function

Private Sub LongToRgba4(ByVal colorValue As Long, ByRef out() As RGBA)
    Dim i As Long
    For i = 0 To 3
        Call SetRGBA(out(i), CByte(colorValue And &HFF&), CByte((colorValue \ &H100&) And &HFF&), CByte((colorValue \ &H10000) And &HFF&))
    Next i
End Sub

Private Sub DrawCenteredSquare(ByVal screenX As Long, ByVal screenY As Long, ByVal size As Long, ByRef col() As RGBA)
    Call Batch_Textured_Box_File(screenX - size \ 2, screenY - size \ 2, 1, 1, 0, 0, "white_pixel.png", col, False, 0, size, size)
End Sub

Private Sub DrawCenteredPointRgba(ByVal tileX As Long, ByVal tileY As Long, ByRef fillCol() As RGBA, ByRef borderCol() As RGBA, ByVal total As Long, ByVal fillSize As Long)
    Dim screenX As Long, screenY As Long
    If Not MinimapTileToScreen(tileX, tileY, screenX, screenY) Then Exit Sub
    Call DrawCenteredSquare(screenX, screenY, total, borderCol)
    Call DrawCenteredSquare(screenX, screenY, fillSize, fillCol)
End Sub

' Punto con color propio (quest/simbolos: el color depende de cada NPC).
Private Sub DrawCenteredPoint(ByVal tileX As Long, ByVal tileY As Long, ByVal fillColor As Long, ByVal borderColor As Long, ByVal total As Long, ByVal fillSize As Long)
    Dim fillCol(0 To 3) As RGBA
    Dim borderCol(0 To 3) As RGBA
    Call LongToRgba4(fillColor, fillCol)
    Call LongToRgba4(borderColor, borderCol)
    Call DrawCenteredPointRgba(tileX, tileY, fillCol, borderCol, total, fillSize)
End Sub

' Capa de un solo color: convierte los colores una vez para todos los puntos.
Private Sub DrawCenteredLayer(ByRef arr() As t_RadarHostile, ByVal pointCount As Integer, ByVal fillColor As Long, ByVal borderColor As Long, ByVal total As Long, ByVal fillSize As Long)
    If pointCount <= 0 Then Exit Sub
    Dim fillCol(0 To 3) As RGBA
    Dim borderCol(0 To 3) As RGBA
    Call LongToRgba4(fillColor, fillCol)
    Call LongToRgba4(borderColor, borderCol)
    Dim k As Long
    For k = 1 To pointCount
        Call DrawCenteredPointRgba(arr(k).x, arr(k).y, fillCol, borderCol, total, fillSize)
    Next k
End Sub

Public Sub RenderMinimapRadar()
    On Error GoTo RenderMinimapRadar_Err
    Dim total As Long, fillSize As Long
    Select Case RadarPointSize
        Case 1: total = 3: fillSize = 1
        Case 2: total = 4: fillSize = 2
        Case Else: total = 5: fillSize = 3
    End Select
    Dim k As Long
    Dim symColor As Long
    Call DrawCenteredLayer(RadarNeutralNpc, RadarNeutralNpcCount, RadarFill_NeutralNpc, RadarBorder_NeutralNpc, total, fillSize)
    If RadarShowQuests Then
        For k = 1 To RadarQuestNpcCount
            Select Case RadarQuestNpcState(k)
                Case 1: symColor = RadarQuestSymbol_Available
                Case 2: symColor = RadarQuestSymbol_InProgress
                Case 3: symColor = RadarQuestSymbol_Ready
                Case Else: symColor = RadarFill_QuestNpc
            End Select
            Call DrawCenteredPoint(RadarQuestNpc(k).x, RadarQuestNpc(k).y, symColor, RadarBorder_QuestNpc, total, fillSize)
        Next k
    End If
    If RadarShowSymbols Then
        For k = 1 To RadarNpcSymbolCount
            Select Case RadarNpcSymbolSubtype(k)
                Case 1: symColor = RadarNpcSymbol_Sacerdote
                Case 2: symColor = RadarNpcSymbol_Banquero
                Case Else: symColor = RadarFill_NeutralNpc
            End Select
            Call DrawCenteredPoint(RadarNpcSymbol(k).x, RadarNpcSymbol(k).y, symColor, RGB(0, 0, 0), total, fillSize)
        Next k
    End If
    Call DrawCenteredLayer(RadarGuardiaReal, RadarGuardiaRealCount, RadarFill_GuardiaReal, RadarBorder_GuardiaReal, total, fillSize)
    Call DrawCenteredLayer(RadarGuardiaCaos, RadarGuardiaCaosCount, RadarFill_GuardiaCaos, RadarBorder_GuardiaCaos, total, fillSize)
    Call DrawCenteredLayer(RadarMascota, RadarMascotaCount, RadarFill_Mascota, RadarBorder_Mascota, total, fillSize)
    Call DrawCenteredLayer(RadarUserCiudadano, RadarUserCiudadanoCount, RadarFill_UserCiudadano, RadarBorder_UserCiudadano, total, fillSize)
    Call DrawCenteredLayer(RadarUserCriminal, RadarUserCriminalCount, RadarFill_UserCriminal, RadarBorder_UserCriminal, total, fillSize)
    Call DrawCenteredLayer(RadarUserImperial, RadarUserImperialCount, RadarFill_UserImperial, RadarBorder_UserImperial, total, fillSize)
    Call DrawCenteredLayer(RadarUserCaos, RadarUserCaosCount, RadarFill_UserCaos, RadarBorder_UserCaos, total, fillSize)
    Call DrawCenteredLayer(RadarUserGm, RadarUserGmCount, RadarFill_UserGm, RadarBorder_UserGm, total, fillSize)
    Call DrawCenteredLayer(RadarHostiles, RadarHostileCount, RadarFill_Hostile, RadarBorder_Hostile, total, fillSize)
    If RadarShowBosses Then Call DrawCenteredLayer(RadarJefe, RadarJefeCount, RadarFill_Jefe, RadarBorder_Jefe, 7, 5)
    If RadarShowClan Then Call DrawCenteredLayer(RadarClanMembers, RadarClanMembersCount, RadarFill_ClanMember, RadarBorder_ClanMember, total, fillSize)
    If RadarShowParty Then Call DrawCenteredLayer(RadarPartyMembers, RadarPartyMembersCount, RadarFill_PartyMember, RadarBorder_PartyMember, total, fillSize)
    If RadarShowPings Then Call DrawCenteredPings
    Exit Sub
RenderMinimapRadar_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModMinimap.RenderMinimapRadar", Erl)
End Sub

' Plan radar-v2 E6: marcas de party en el minimapa centrado (cuadrado hueco amarillo que late).
Private Sub DrawCenteredPings()
    Dim col(0 To 3) As RGBA
    Call LongToRgba4(RGB(255, 255, 0), col)
    Dim i As Long, side As Long
    Dim px As Byte, py As Byte
    Dim screenX As Long, screenY As Long
    Dim half As Long
    For i = 1 To MinimapPingSlots()
        side = MinimapPingSize(i, px, py)
        If side > 0 Then
            If MinimapTileToScreen(px, py, screenX, screenY) Then
                half = side \ 2
                Call Batch_Textured_Box_File(screenX - half, screenY - half, 1, 1, 0, 0, "white_pixel.png", col, False, 0, side, 1)
                Call Batch_Textured_Box_File(screenX - half, screenY + half, 1, 1, 0, 0, "white_pixel.png", col, False, 0, side, 1)
                Call Batch_Textured_Box_File(screenX - half, screenY - half, 1, 1, 0, 0, "white_pixel.png", col, False, 0, 1, side)
                Call Batch_Textured_Box_File(screenX + half, screenY - half, 1, 1, 0, 0, "white_pixel.png", col, False, 0, 1, side)
            End If
        End If
    Next i
End Sub
