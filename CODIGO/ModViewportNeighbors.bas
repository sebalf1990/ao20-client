Attribute VB_Name = "ModViewportNeighbors"
Option Explicit

Private Const CSM_FIVE_LAYER_SIGNATURE_NEIGHBOR As Long = &H324C3557 ' same as Recursos.CSM_FIVE_LAYER_SIGNATURE

Private Const DIR_NORTH As Integer = 1
Private Const DIR_SOUTH As Integer = 2
Private Const DIR_WEST  As Integer = 3
Private Const DIR_EAST  As Integer = 4

Private Const MAP_MAX_TILES As Integer = 100
Private Const NEIGHBOR_MIN_TE_COUNT As Integer = 60
Private Const MAX_MAP_ID_GUESS As Integer = 1200
Private Const CSM_SIZE_BLOCKED As Long = 5
Private Const CSM_SIZE_GRH As Long = 8
Private Const CSM_SIZE_TRIGGER As Long = 6
Private Const CSM_SIZE_LIGHT As Long = 9
Private Const CSM_SIZE_PARTICLE As Long = 8
Private Const CSM_SIZE_NPC As Long = 6
Private Const CSM_SIZE_OBJ As Long = 8
Private Const CSM_SIZE_TE As Long = 10

Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(1 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    x As Integer
    y As Integer
    lados As Byte
End Type

Private Type tDatosGrh
    x As Integer
    y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    x As Integer
    y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    x As Integer
    y As Integer
    color As RGBA
    Rango As Byte
End Type

Private Type tDatosParticulas
    x As Integer
    y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    x As Integer
    y As Integer
    NpcIndex As Integer
End Type

Private Type tDatosObjs
    x As Integer
    y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    x As Integer
    y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    backup_mode As Byte
    restrict_mode As String
    music_numberHi As Long
    music_numberLow As Long
    Seguro As Byte
    zone As String
    terrain As String
    ambient As String
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String
    LLUVIA As Byte
    NIEVE As Byte
    niebla As Byte
End Type

Private Type tNeighborBlock
    Graphic(1 To 4) As Grh
    light_value(3) As RGBA
    Blocked As Integer
End Type

Private g_neighbor_map(1 To 4) As Integer
Private g_neighbor_count(1 To 4) As Integer
Private g_neighbor_source_coord(1 To 4) As Integer
Private g_neighbor_dest_coord(1 To 4) As Integer
Private g_neighbor_loaded(1 To 4) As Boolean
Private g_neighbor_data(1 To 4, 1 To MAP_MAX_TILES, 1 To MAP_MAX_TILES) As tNeighborBlock
Private g_cached_resource_map As Integer

Public Sub ViewportNeighbor_PrepareForMap(ByVal ResourceMap As Integer)
    On Error GoTo ViewportNeighbor_PrepareForMap_Err

    Dim direction As Integer
    Dim targetMap As Integer
    Dim sourceCoord As Integer
    Dim destCoord As Integer

    Call ViewportNeighbor_Clear
    g_cached_resource_map = ResourceMap

    For direction = DIR_NORTH To DIR_EAST
        sourceCoord = 0
        destCoord = 0
        targetMap = ViewportNeighbor_FindDominantNeighbor(ResourceMap, direction, g_neighbor_count(direction), sourceCoord, destCoord)
        If targetMap > 0 And g_neighbor_count(direction) >= NEIGHBOR_MIN_TE_COUNT Then
            If ViewportNeighbor_LoadMap(direction, targetMap) Then
                g_neighbor_map(direction) = targetMap
                g_neighbor_source_coord(direction) = sourceCoord
                g_neighbor_dest_coord(direction) = destCoord
                g_neighbor_loaded(direction) = True
            End If
        End If
    Next direction

    Exit Sub
ViewportNeighbor_PrepareForMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_PrepareForMap", Erl)
End Sub

Public Sub ViewportNeighbor_RenderLayer(ByVal Layer As Integer, _
                                        ByVal RawMinX As Integer, _
                                        ByVal RawMaxX As Integer, _
                                        ByVal RawMinY As Integer, _
                                        ByVal RawMaxY As Integer, _
                                        ByVal BaseScreenX As Integer, _
                                        ByVal BaseScreenY As Integer)
    On Error GoTo ViewportNeighbor_RenderLayer_Err

    If Not ViewportDebug_IsFullscreen() Then Exit Sub
    If Layer < 1 Or Layer > 4 Then Exit Sub

    Dim virtualX As Integer
    Dim virtualY As Integer
    Dim localX As Integer
    Dim localY As Integer
    Dim direction As Integer
    Dim screenX As Integer
    Dim screenY As Integer

    For virtualY = RawMinY To RawMaxY
        For virtualX = RawMinX To RawMaxX
            If virtualX < XMinMapSize Or virtualX > XMaxMapSize Or virtualY < YMinMapSize Or virtualY > YMaxMapSize Then
                If ViewportNeighbor_MapVirtualTile(virtualX, virtualY, direction, localX, localY) Then
                    screenX = BaseScreenX + virtualX * TilePixelWidth
                    screenY = BaseScreenY + virtualY * TilePixelHeight
                    With g_neighbor_data(direction, localX, localY)
                        If .Graphic(Layer).GrhIndex <> 0 Then
                            Call Draw_Grh(.Graphic(Layer), screenX, screenY, 1, 1, .light_value, False, localX, localY)
                        End If
                    End With
                End If
            End If
        Next virtualX
    Next virtualY

    Exit Sub
ViewportNeighbor_RenderLayer_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_RenderLayer", Erl)
End Sub

Private Sub ViewportNeighbor_Clear()
    On Error GoTo ViewportNeighbor_Clear_Err

    Dim direction As Integer

    For direction = DIR_NORTH To DIR_EAST
        g_neighbor_map(direction) = 0
        g_neighbor_count(direction) = 0
        g_neighbor_source_coord(direction) = 0
        g_neighbor_dest_coord(direction) = 0
        g_neighbor_loaded(direction) = False
    Next direction
    Erase g_neighbor_data

    Exit Sub
ViewportNeighbor_Clear_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_Clear", Erl)
End Sub

Private Function ViewportNeighbor_FindDominantNeighbor(ByVal ResourceMap As Integer, _
                                                       ByVal direction As Integer, _
                                                       ByRef bestCount As Integer, _
                                                       ByRef bestSourceCoord As Integer, _
                                                       ByRef bestDestCoord As Integer) As Integer
    On Error GoTo ViewportNeighbor_FindDominantNeighbor_Err

    Dim fh      As Integer
    Dim MH      As tMapHeader
    Dim MS      As tMapSize
    Dim MD      As tMapDat
    Dim MapPath As String
    Dim TEs()   As tDatosTE
    Dim counts(1 To MAX_MAP_ID_GUESS) As Integer
    Dim sourceCounts(1 To MAP_MAX_TILES) As Integer
    Dim destCounts(1 To MAP_MAX_TILES) As Integer
    Dim i       As Long
    Dim mapId   As Integer
    Dim sourceCoord As Integer
    Dim destCoord As Integer

    MapPath = ViewportNeighbor_MapPath(ResourceMap)
    If LenB(Dir$(MapPath)) = 0 Then Exit Function

    fh = FreeFile
    Open MapPath For Binary As #fh
    ' Plan 15.001: this reader only knows the legacy 4-layer CSM. A five-layer
    ' file (signature W5L2, see Recursos.ReadCsmHeader) is skipped, not misread.
    Dim csmSignature As Long
    Get #fh, , csmSignature
    If csmSignature = CSM_FIVE_LAYER_SIGNATURE_NEIGHBOR Then
        Close #fh
        Exit Function
    End If
    Seek #fh, 1
    Get #fh, , MH
    Get #fh, , MS
    Get #fh, , MD
    Call ViewportNeighbor_SkipMapPayload(fh, MH, False)

    If MH.NumeroTE > 0 Then
        ReDim TEs(1 To MH.NumeroTE)
        Get #fh, , TEs
        For i = 1 To MH.NumeroTE
            mapId = TEs(i).DestM
            If mapId > 0 And mapId <= MAX_MAP_ID_GUESS Then
                If ViewportNeighbor_IsCardinalTE(TEs(i), direction) Then
                    counts(mapId) = counts(mapId) + 1
                    If counts(mapId) > bestCount Then
                        bestCount = counts(mapId)
                        ViewportNeighbor_FindDominantNeighbor = mapId
                    End If
                End If
            End If
        Next i

        If ViewportNeighbor_FindDominantNeighbor > 0 Then
            For i = 1 To MH.NumeroTE
                If TEs(i).DestM = ViewportNeighbor_FindDominantNeighbor Then
                    If ViewportNeighbor_IsCardinalTE(TEs(i), direction) Then
                        Call ViewportNeighbor_TECoords(TEs(i), direction, sourceCoord, destCoord)
                        If sourceCoord >= XMinMapSize And sourceCoord <= XMaxMapSize Then
                            sourceCounts(sourceCoord) = sourceCounts(sourceCoord) + 1
                        End If
                        If destCoord >= XMinMapSize And destCoord <= XMaxMapSize Then
                            destCounts(destCoord) = destCounts(destCoord) + 1
                        End If
                    End If
                End If
            Next i
            bestSourceCoord = ViewportNeighbor_MostUsedCoord(sourceCounts)
            bestDestCoord = ViewportNeighbor_MostUsedCoord(destCounts)
        End If
    End If

    Close #fh
    Exit Function
ViewportNeighbor_FindDominantNeighbor_Err:
    On Error Resume Next
    Close #fh
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_FindDominantNeighbor", Erl)
End Function

Private Function ViewportNeighbor_LoadMap(ByVal direction As Integer, ByVal ResourceMap As Integer) As Boolean
    On Error GoTo ViewportNeighbor_LoadMap_Err

    Dim fh           As Integer
    Dim MH           As tMapHeader
    Dim MS           As tMapSize
    Dim MD           As tMapDat
    Dim Blqs()       As tDatosBloqueados
    Dim L()          As tDatosGrh
    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE
    Dim MapPath      As String
    Dim i            As Long
    Dim x            As Integer
    Dim y            As Integer
    Dim layer        As Integer

    MapPath = ViewportNeighbor_MapPath(ResourceMap)
    If LenB(Dir$(MapPath)) = 0 Then Exit Function

    Call ViewportNeighbor_InitDirection(direction)

    fh = FreeFile
    Open MapPath For Binary As #fh
    ' Plan 15.001: this reader only knows the legacy 4-layer CSM. A five-layer
    ' file (signature W5L2, see Recursos.ReadCsmHeader) is skipped, not misread.
    Dim csmSignature As Long
    Get #fh, , csmSignature
    If csmSignature = CSM_FIVE_LAYER_SIGNATURE_NEIGHBOR Then
        Close #fh
        Exit Function
    End If
    Seek #fh, 1
    Get #fh, , MH
    Get #fh, , MS
    Get #fh, , MD

    With MH
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs
            For i = 1 To .NumeroBloqueados
                g_neighbor_data(direction, Blqs(i).x, Blqs(i).y).Blocked = Blqs(i).lados
            Next i
        End If

        For layer = 1 To 4
            If .NumeroLayers(layer) > 0 Then
                ReDim L(1 To .NumeroLayers(layer))
                Get #fh, , L
                For i = 1 To .NumeroLayers(layer)
                    x = L(i).x
                    y = L(i).y
                    With g_neighbor_data(direction, x, y)
                        .Graphic(layer).GrhIndex = L(i).GrhIndex
                        .Graphic(layer).x = x * TilePixelWidth
                        .Graphic(layer).y = y * TilePixelHeight
                        Call InitGrh(.Graphic(layer), .Graphic(layer).GrhIndex)
                        If layer = 1 Then
                            If ViewportNeighbor_IsWaterGrh(.Graphic(1).GrhIndex) Then
                                .Blocked = .Blocked Or FLAG_AGUA
                            ElseIf ViewportNeighbor_IsLavaGrh(.Graphic(1).GrhIndex) Then
                                .Blocked = .Blocked Or FLAG_LAVA
                            End If
                        End If
                    End With
                Next i
            End If
        Next layer

        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
        End If

        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
        End If

        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
        End If

        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
        End If

        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs
        End If

        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs
        End If
    End With

    Close #fh
    ViewportNeighbor_LoadMap = True
    Exit Function
ViewportNeighbor_LoadMap_Err:
    On Error Resume Next
    Close #fh
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_LoadMap", Erl)
End Function

Private Sub ViewportNeighbor_InitDirection(ByVal direction As Integer)
    On Error GoTo ViewportNeighbor_InitDirection_Err

    Dim x As Integer
    Dim y As Integer

    For x = 1 To MAP_MAX_TILES
        For y = 1 To MAP_MAX_TILES
            g_neighbor_data(direction, x, y).light_value(0) = global_light
            g_neighbor_data(direction, x, y).light_value(1) = global_light
            g_neighbor_data(direction, x, y).light_value(2) = global_light
            g_neighbor_data(direction, x, y).light_value(3) = global_light
        Next y
    Next x

    Exit Sub
ViewportNeighbor_InitDirection_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_InitDirection", Erl)
End Sub

Private Function ViewportNeighbor_MapVirtualTile(ByVal virtualX As Integer, _
                                                 ByVal virtualY As Integer, _
                                                 ByRef direction As Integer, _
                                                 ByRef localX As Integer, _
                                                 ByRef localY As Integer) As Boolean
    If virtualX < XMinMapSize And virtualY >= YMinMapSize And virtualY <= YMaxMapSize Then
        direction = DIR_WEST
        If Not g_neighbor_loaded(direction) Then Exit Function
        localX = g_neighbor_dest_coord(direction) + (virtualX - g_neighbor_source_coord(direction))
        localY = virtualY
    ElseIf virtualX > XMaxMapSize And virtualY >= YMinMapSize And virtualY <= YMaxMapSize Then
        direction = DIR_EAST
        If Not g_neighbor_loaded(direction) Then Exit Function
        localX = g_neighbor_dest_coord(direction) + (virtualX - g_neighbor_source_coord(direction))
        localY = virtualY
    ElseIf virtualY < YMinMapSize And virtualX >= XMinMapSize And virtualX <= XMaxMapSize Then
        direction = DIR_NORTH
        If Not g_neighbor_loaded(direction) Then Exit Function
        localX = virtualX
        localY = g_neighbor_dest_coord(direction) + (virtualY - g_neighbor_source_coord(direction))
    ElseIf virtualY > YMaxMapSize And virtualX >= XMinMapSize And virtualX <= XMaxMapSize Then
        direction = DIR_SOUTH
        If Not g_neighbor_loaded(direction) Then Exit Function
        localX = virtualX
        localY = g_neighbor_dest_coord(direction) + (virtualY - g_neighbor_source_coord(direction))
    Else
        Exit Function
    End If

    If localX < XMinMapSize Or localX > XMaxMapSize Then Exit Function
    If localY < YMinMapSize Or localY > YMaxMapSize Then Exit Function

    ViewportNeighbor_MapVirtualTile = True
End Function

Private Sub ViewportNeighbor_TECoords(ByRef TE As tDatosTE, _
                                      ByVal direction As Integer, _
                                      ByRef sourceCoord As Integer, _
                                      ByRef destCoord As Integer)
    Select Case direction
        Case DIR_NORTH, DIR_SOUTH
            sourceCoord = TE.y
            destCoord = TE.DestY
        Case DIR_WEST, DIR_EAST
            sourceCoord = TE.x
            destCoord = TE.DestX
    End Select
End Sub

Private Function ViewportNeighbor_MostUsedCoord(ByRef counts() As Integer) As Integer
    Dim coord As Integer
    Dim bestCount As Integer

    For coord = XMinMapSize To XMaxMapSize
        If counts(coord) > bestCount Then
            bestCount = counts(coord)
            ViewportNeighbor_MostUsedCoord = coord
        End If
    Next coord
End Function

Private Function ViewportNeighbor_IsCardinalTE(ByRef TE As tDatosTE, ByVal direction As Integer) As Boolean
    Select Case direction
        Case DIR_NORTH
            ViewportNeighbor_IsCardinalTE = (TE.y <= 15 And TE.DestY >= 86)
        Case DIR_SOUTH
            ViewportNeighbor_IsCardinalTE = (TE.y >= 86 And TE.DestY <= 15)
        Case DIR_WEST
            ViewportNeighbor_IsCardinalTE = (TE.x <= 18 And TE.DestX >= 83)
        Case DIR_EAST
            ViewportNeighbor_IsCardinalTE = (TE.x >= 83 And TE.DestX <= 18)
    End Select
End Function

Private Sub ViewportNeighbor_SkipMapPayload(ByVal fh As Integer, ByRef MH As tMapHeader, ByVal IncludeTE As Boolean)
    If MH.NumeroBloqueados > 0 Then Seek #fh, Seek(fh) + MH.NumeroBloqueados * CSM_SIZE_BLOCKED
    If MH.NumeroLayers(1) > 0 Then Seek #fh, Seek(fh) + MH.NumeroLayers(1) * CSM_SIZE_GRH
    If MH.NumeroLayers(2) > 0 Then Seek #fh, Seek(fh) + MH.NumeroLayers(2) * CSM_SIZE_GRH
    If MH.NumeroLayers(3) > 0 Then Seek #fh, Seek(fh) + MH.NumeroLayers(3) * CSM_SIZE_GRH
    If MH.NumeroLayers(4) > 0 Then Seek #fh, Seek(fh) + MH.NumeroLayers(4) * CSM_SIZE_GRH
    If MH.NumeroTriggers > 0 Then Seek #fh, Seek(fh) + MH.NumeroTriggers * CSM_SIZE_TRIGGER
    If MH.NumeroParticulas > 0 Then Seek #fh, Seek(fh) + MH.NumeroParticulas * CSM_SIZE_PARTICLE
    If MH.NumeroLuces > 0 Then Seek #fh, Seek(fh) + MH.NumeroLuces * CSM_SIZE_LIGHT
    If MH.NumeroOBJs > 0 Then Seek #fh, Seek(fh) + MH.NumeroOBJs * CSM_SIZE_OBJ
    If MH.NumeroNPCs > 0 Then Seek #fh, Seek(fh) + MH.NumeroNPCs * CSM_SIZE_NPC
    If IncludeTE And MH.NumeroTE > 0 Then Seek #fh, Seek(fh) + MH.NumeroTE * CSM_SIZE_TE
End Sub

Private Function ViewportNeighbor_MapPath(ByVal ResourceMap As Integer) As String
#If Compresion = 1 Then
    ' Bajo Compresion, los .csm viven en el archivo de OUTPUT. El loader principal
    ' (Recursos.bas) extrae el mapa actual al temp; los vecinos hay que extraerlos aca.
    Dim fileName As String
    fileName = "mapa" & ResourceMap & ".csm"
    If LenB(Dir$(Windows_Temp_Dir & fileName)) = 0 Then
        Call Extract_File(Maps, App.path & "\..\Recursos\OUTPUT\", fileName, Windows_Temp_Dir, ResourcesPassword, False)
    End If
    ViewportNeighbor_MapPath = Windows_Temp_Dir & fileName
#Else
    ViewportNeighbor_MapPath = App.path & "\..\Recursos\Mapas\mapa" & ResourceMap & ".csm"
#End If
End Function

Private Function ViewportNeighbor_IsWaterGrh(ByVal GrhIndex As Long) As Boolean
    ViewportNeighbor_IsWaterGrh = (GrhIndex >= 1505 And GrhIndex <= 1520) Or (GrhIndex >= 124 And GrhIndex <= 139) Or (GrhIndex >= 24223 And GrhIndex <= 24238) Or ( _
            GrhIndex >= 24303 And GrhIndex <= 24318) Or (GrhIndex >= 468 And GrhIndex <= 483) Or (GrhIndex >= 44668 And GrhIndex <= 44683) Or (GrhIndex >= 24143 And _
            GrhIndex <= 24158) Or (GrhIndex >= 12628 And GrhIndex <= 12643) Or (GrhIndex >= 2948 And GrhIndex <= 2963)
End Function

Private Function ViewportNeighbor_IsLavaGrh(ByVal GrhIndex As Long) As Boolean
    ViewportNeighbor_IsLavaGrh = (GrhIndex >= 57400 And GrhIndex <= 57415) Or (GrhIndex >= 16101 And GrhIndex <= 16116) Or (GrhIndex >= 26767 And GrhIndex <= 26782)
End Function
