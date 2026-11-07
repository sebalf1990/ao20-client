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

' E4c: el bloque ya no guarda luz por tile (todas valian global_light del momento de la carga y en
' cache quedarian viejas): se dibuja con la luz global vigente (g_neighbor_light).
' Tamano: Grh = 40 bytes en memoria (LenB) -> bloque = 4 * 40 + 2 = 162, alineado a 164 bytes;
' un mapa (100x100) = 1,64 MB. Se loguea Len/LenB real al reservar la cache ("[NEIGH] pool").
Private Type tNeighborBlock
    Graphic(1 To 4) As Grh
    Blocked As Integer
End Type

' E4c: cache de vecinos por id de mapa (= archivo mapa{id}.csm realmente abierto, R2), con
' NEIGHBOR_POOL_SIZE bloques y desalojo LRU. 9 bloques * 1,64 MB = ~14,8 MB, reservados recien la
' primera vez que se carga un vecino (en HUD clasico puro no se reserva nada). 9 = 4 cardinales en uso
' + margen para que ida y vuelta y caminatas cortas no relean (y espacio para esquinas en E4e).
Private Const NEIGHBOR_POOL_SIZE As Integer = 9

' Plan del mapa actual por direccion (se recalcula en cada cambio de mapa desde la lista de TE).
Private g_neighbor_map(1 To 4) As Integer
Private g_neighbor_count(1 To 4) As Integer
Private g_neighbor_source_coord(1 To 4) As Integer
Private g_neighbor_dest_coord(1 To 4) As Integer
Private g_neighbor_loaded(1 To 4) As Boolean
Private g_neighbor_slot(1 To 4) As Integer
Private g_cached_resource_map As Integer
Private g_neighbor_light(3) As RGBA

' Cache: g_pool_data(x, y, slot); el slot va ultimo para que cada mapa quede contiguo en memoria.
Private g_pool_data() As tNeighborBlock
Private g_pool_ready As Boolean
Private g_slot_map(1 To NEIGHBOR_POOL_SIZE) As Integer
Private g_slot_last_use(1 To NEIGHBOR_POOL_SIZE) As Long
Private g_neighbor_use_clock As Long
Private g_empty_block As tNeighborBlock

Public Sub ViewportNeighbor_PrepareForMap(ByVal ResourceMap As Integer)
    On Error GoTo ViewportNeighbor_PrepareForMap_Err

    Dim direction As Integer
    Dim slot As Integer
    Dim targetMap(1 To 4) As Integer
    Dim sourceCoord(1 To 4) As Integer
    Dim destCoord(1 To 4) As Integer
    Dim wanted(1 To 4) As Boolean
    Dim hits As Integer
    Dim loads As Integer
    Dim summary As String

    Call ViewportNeighbor_ResetDirections
    g_cached_resource_map = ResourceMap
    If ResourceMap < 1 Then Exit Sub

    Call ViewportNeighbor_FindNeighbors(ResourceMap, targetMap, g_neighbor_count, sourceCoord, destCoord)

    ' E4c: primero se marcan como usados en este cruce los vecinos que ya estan en cache, asi el LRU
    ' no desaloja uno que se necesita ahora al cargar los que faltan.
    g_neighbor_use_clock = g_neighbor_use_clock + 1
    For direction = DIR_NORTH To DIR_EAST
        wanted(direction) = (targetMap(direction) > 0 And g_neighbor_count(direction) >= NEIGHBOR_MIN_TE_COUNT)
        If wanted(direction) Then
            slot = ViewportNeighbor_FindSlot(targetMap(direction))
            If slot > 0 Then g_slot_last_use(slot) = g_neighbor_use_clock
        End If
    Next direction

    For direction = DIR_NORTH To DIR_EAST
        If wanted(direction) Then
            slot = ViewportNeighbor_FindSlot(targetMap(direction))
            If slot > 0 Then
                hits = hits + 1
            Else
                slot = ViewportNeighbor_AcquireSlot()
                If slot > 0 Then
                    If ViewportNeighbor_LoadMap(slot, targetMap(direction)) Then
                        g_slot_map(slot) = targetMap(direction)
                        loads = loads + 1
                    Else
                        slot = 0
                    End If
                End If
            End If
            If slot > 0 Then
                g_slot_last_use(slot) = g_neighbor_use_clock
                g_neighbor_slot(direction) = slot
                g_neighbor_map(direction) = targetMap(direction)
                g_neighbor_source_coord(direction) = sourceCoord(direction)
                g_neighbor_dest_coord(direction) = destCoord(direction)
                g_neighbor_loaded(direction) = True
            End If
        End If
        summary = summary & " " & Mid$("NSWE", direction, 1) & "=" & g_neighbor_map(direction)
    Next direction

    Call ViewportDebug_AppendDiagLog("[NEIGH] map=" & ResourceMap & summary & " hits=" & hits & " loads=" & loads)
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
    Dim perfT As Double

    perfT = Perf_FrameStart()
    g_neighbor_light(0) = global_light
    g_neighbor_light(1) = global_light
    g_neighbor_light(2) = global_light
    g_neighbor_light(3) = global_light
    For virtualY = RawMinY To RawMaxY
        For virtualX = RawMinX To RawMaxX
            If virtualX < XMinMapSize Or virtualX > XMaxMapSize Or virtualY < YMinMapSize Or virtualY > YMaxMapSize Then
                If ViewportNeighbor_MapVirtualTile(virtualX, virtualY, direction, localX, localY) Then
                    screenX = BaseScreenX + virtualX * TilePixelWidth
                    screenY = BaseScreenY + virtualY * TilePixelHeight
                    With g_pool_data(localX, localY, g_neighbor_slot(direction))
                        If .Graphic(Layer).GrhIndex <> 0 Then
                            Call Draw_Grh(.Graphic(Layer), screenX, screenY, 1, 1, g_neighbor_light, False, localX, localY)
                        End If
                    End With
                End If
            End If
        Next virtualX
    Next virtualY
    Call Perf_FrameAdd(PERF_PIECE_NEIGHBORS, perfT)

    Exit Sub
ViewportNeighbor_RenderLayer_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_RenderLayer", Erl)
End Sub

' E4c: solo se resetea el plan por direccion; la cache de bloques se conserva (antes: Erase de ~7 MB
' por cruce y recarga de los 4 vecinos).
Private Sub ViewportNeighbor_ResetDirections()
    Dim direction As Integer

    For direction = DIR_NORTH To DIR_EAST
        g_neighbor_map(direction) = 0
        g_neighbor_count(direction) = 0
        g_neighbor_source_coord(direction) = 0
        g_neighbor_dest_coord(direction) = 0
        g_neighbor_loaded(direction) = False
        g_neighbor_slot(direction) = 0
    Next direction
End Sub

Private Function ViewportNeighbor_FindSlot(ByVal mapId As Integer) As Integer
    Dim slot As Integer

    If Not g_pool_ready Or mapId < 1 Then Exit Function
    For slot = 1 To NEIGHBOR_POOL_SIZE
        If g_slot_map(slot) = mapId Then
            ViewportNeighbor_FindSlot = slot
            Exit Function
        End If
    Next slot
End Function

' Devuelve un slot libre o, si no hay, el usado hace mas tiempo que no se necesite en este cruce.
' El slot devuelto queda marcado libre (g_slot_map = 0) hasta que la carga termine bien.
Private Function ViewportNeighbor_AcquireSlot() As Integer
    Dim slot As Integer
    Dim best As Integer

    If Not g_pool_ready Then
        ReDim g_pool_data(1 To MAP_MAX_TILES, 1 To MAP_MAX_TILES, 1 To NEIGHBOR_POOL_SIZE)
        g_pool_ready = True
        Call ViewportDebug_AppendDiagLog("[NEIGH] pool slots=" & NEIGHBOR_POOL_SIZE & " block Len=" & Len(g_pool_data(1, 1, 1)) & " LenB=" & LenB(g_pool_data(1, 1, 1)) & _
                " bytes=" & CStr(CLng(MAP_MAX_TILES) * MAP_MAX_TILES * NEIGHBOR_POOL_SIZE * LenB(g_pool_data(1, 1, 1))))
    End If

    For slot = 1 To NEIGHBOR_POOL_SIZE
        If g_slot_map(slot) = 0 Then
            best = slot
            Exit For
        End If
    Next slot

    If best = 0 Then
        For slot = 1 To NEIGHBOR_POOL_SIZE
            If g_slot_last_use(slot) < g_neighbor_use_clock Then
                If best = 0 Then
                    best = slot
                ElseIf g_slot_last_use(slot) < g_slot_last_use(best) Then
                    best = slot
                End If
            End If
        Next slot
    End If

    If best > 0 Then g_slot_map(best) = 0
    ViewportNeighbor_AcquireSlot = best
End Function

' E4b: una sola apertura del .csm actual y una sola pasada por su lista de TE para los 4 lados
' (antes se abria 4 veces, una por direccion). Mismo criterio de desempate que antes: por lado gana
' el primer mapa que alcanza el maximo de TE, recorriendo la lista en orden.
Private Sub ViewportNeighbor_FindNeighbors(ByVal ResourceMap As Integer, _
                                           ByRef bestMap() As Integer, _
                                           ByRef bestCount() As Integer, _
                                           ByRef bestSourceCoord() As Integer, _
                                           ByRef bestDestCoord() As Integer)
    On Error GoTo ViewportNeighbor_FindNeighbors_Err

    Dim fh      As Integer
    Dim MH      As tMapHeader
    Dim MS      As tMapSize
    Dim MD      As tMapDat
    Dim MapPath As String
    Dim TEs()   As tDatosTE
    Dim counts(1 To 4, 1 To MAX_MAP_ID_GUESS) As Integer
    Dim sourceCounts(1 To 4, 1 To MAP_MAX_TILES) As Integer
    Dim destCounts(1 To 4, 1 To MAP_MAX_TILES) As Integer
    Dim i       As Long
    Dim mapId   As Integer
    Dim direction As Integer
    Dim sourceCoord As Integer
    Dim destCoord As Integer
    Dim csmSignature As Long

    MapPath = ViewportNeighbor_MapPath(ResourceMap)
    If LenB(Dir$(MapPath)) = 0 Then Exit Sub

    fh = FreeFile
    Open MapPath For Binary As #fh
    ' Plan 15.001: this reader only knows the legacy 4-layer CSM. A five-layer
    ' file (signature W5L2, see Recursos.ReadCsmHeader) is skipped, not misread.
    Get #fh, , csmSignature
    If csmSignature = CSM_FIVE_LAYER_SIGNATURE_NEIGHBOR Then
        Close #fh
        fh = 0
        Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
        Exit Sub
    End If
    Seek #fh, 1
    Get #fh, , MH
    Get #fh, , MS
    Get #fh, , MD
    Call ViewportNeighbor_SkipMapPayload(fh, MH, False)

    If MH.NumeroTE > 0 Then
        ReDim TEs(1 To MH.NumeroTE)
        Get #fh, , TEs
    End If
    Close #fh
    fh = 0
    Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
    If MH.NumeroTE <= 0 Then Exit Sub

    For i = 1 To MH.NumeroTE
        mapId = TEs(i).DestM
        If mapId > 0 And mapId <= MAX_MAP_ID_GUESS Then
            For direction = DIR_NORTH To DIR_EAST
                If ViewportNeighbor_IsCardinalTE(TEs(i), direction) Then
                    counts(direction, mapId) = counts(direction, mapId) + 1
                    If counts(direction, mapId) > bestCount(direction) Then
                        bestCount(direction) = counts(direction, mapId)
                        bestMap(direction) = mapId
                    End If
                End If
            Next direction
        End If
    Next i

    For i = 1 To MH.NumeroTE
        For direction = DIR_NORTH To DIR_EAST
            If bestMap(direction) > 0 Then
                If TEs(i).DestM = bestMap(direction) Then
                    If ViewportNeighbor_IsCardinalTE(TEs(i), direction) Then
                        Call ViewportNeighbor_TECoords(TEs(i), direction, sourceCoord, destCoord)
                        If sourceCoord >= XMinMapSize And sourceCoord <= XMaxMapSize Then
                            sourceCounts(direction, sourceCoord) = sourceCounts(direction, sourceCoord) + 1
                        End If
                        If destCoord >= XMinMapSize And destCoord <= XMaxMapSize Then
                            destCounts(direction, destCoord) = destCounts(direction, destCoord) + 1
                        End If
                    End If
                End If
            End If
        Next direction
    Next i

    For direction = DIR_NORTH To DIR_EAST
        If bestMap(direction) > 0 Then
            bestSourceCoord(direction) = ViewportNeighbor_MostUsedCoord(sourceCounts, direction)
            bestDestCoord(direction) = ViewportNeighbor_MostUsedCoord(destCounts, direction)
        End If
    Next direction

    Exit Sub
ViewportNeighbor_FindNeighbors_Err:
    Dim errNumber As Long, errDescription As String, errLine As Long
    errNumber = Err.Number: errDescription = Err.Description: errLine = Erl
    On Error Resume Next
    If fh <> 0 Then Close #fh
    Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
    Call RegistrarError(errNumber, errDescription, "ModViewportNeighbors.ViewportNeighbor_FindNeighbors", errLine)
End Sub

Private Function ViewportNeighbor_LoadMap(ByVal slot As Integer, ByVal ResourceMap As Integer) As Boolean
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

    Call ViewportNeighbor_ClearSlot(slot)

    fh = FreeFile
    Open MapPath For Binary As #fh
    ' Plan 15.001: this reader only knows the legacy 4-layer CSM. A five-layer
    ' file (signature W5L2, see Recursos.ReadCsmHeader) is skipped, not misread.
    Dim csmSignature As Long
    Get #fh, , csmSignature
    If csmSignature = CSM_FIVE_LAYER_SIGNATURE_NEIGHBOR Then
        Close #fh
        fh = 0
        Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
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
                ' E1: coordenadas fuera de 1..MAP_MAX_TILES se saltean (no abortan la carga).
                If Blqs(i).x >= 1 And Blqs(i).x <= MAP_MAX_TILES And Blqs(i).y >= 1 And Blqs(i).y <= MAP_MAX_TILES Then
                    g_pool_data(Blqs(i).x, Blqs(i).y, slot).Blocked = Blqs(i).lados
                End If
            Next i
        End If

        For layer = 1 To 4
            If .NumeroLayers(layer) > 0 Then
                ReDim L(1 To .NumeroLayers(layer))
                Get #fh, , L
                For i = 1 To .NumeroLayers(layer)
                    x = L(i).x
                    y = L(i).y
                    ' E1: coordenadas fuera de 1..MAP_MAX_TILES se saltean (no abortan la carga).
                    If x >= 1 And x <= MAP_MAX_TILES And y >= 1 And y <= MAP_MAX_TILES Then
                        With g_pool_data(x, y, slot)
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
                    End If
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
    fh = 0
    Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
    ViewportNeighbor_LoadMap = True
    Exit Function
ViewportNeighbor_LoadMap_Err:
    Dim errNumber As Long, errDescription As String, errLine As Long
    errNumber = Err.Number: errDescription = Err.Description: errLine = Erl
    On Error Resume Next
    If fh <> 0 Then Close #fh
    Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
    Call RegistrarError(errNumber, errDescription, "ModViewportNeighbors.ViewportNeighbor_LoadMap", errLine)
End Function

' E4c: vacia solo el slot que se va a recargar (la carga solo escribe los tiles presentes en el .csm).
Private Sub ViewportNeighbor_ClearSlot(ByVal slot As Integer)
    On Error GoTo ViewportNeighbor_ClearSlot_Err

    Dim x As Integer
    Dim y As Integer

    For y = 1 To MAP_MAX_TILES
        For x = 1 To MAP_MAX_TILES
            g_pool_data(x, y, slot) = g_empty_block
        Next x
    Next y

    Exit Sub
ViewportNeighbor_ClearSlot_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_ClearSlot", Erl)
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

Private Function ViewportNeighbor_MostUsedCoord(ByRef counts() As Integer, ByVal direction As Integer) As Integer
    Dim coord As Integer
    Dim bestCount As Integer

    For coord = XMinMapSize To XMaxMapSize
        If counts(direction, coord) > bestCount Then
            bestCount = counts(direction, coord)
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

' E4b: con Compresion=1 el .csm extraido al temp se borra apenas se termina de leer (antes los de
' los vecinos quedaban para siempre). Sin Compresion no hay temporales: no hace nada.
Private Sub ViewportNeighbor_ReleaseMapFile(ByVal ResourceMap As Integer)
#If Compresion = 1 Then
    On Error GoTo ViewportNeighbor_ReleaseMapFile_Err
    Dim tempPath As String
    tempPath = Windows_Temp_Dir & "mapa" & ResourceMap & ".csm"
    ' Delete_File sobre un archivo inexistente lo crea vacio y falla en su handler: chequear antes.
    If LenB(Dir$(tempPath)) <> 0 Then Call Delete_File(tempPath)
    Exit Sub
ViewportNeighbor_ReleaseMapFile_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_ReleaseMapFile", Erl)
#End If
End Sub

Private Function ViewportNeighbor_IsWaterGrh(ByVal GrhIndex As Long) As Boolean
    ViewportNeighbor_IsWaterGrh = (GrhIndex >= 1505 And GrhIndex <= 1520) Or (GrhIndex >= 124 And GrhIndex <= 139) Or (GrhIndex >= 24223 And GrhIndex <= 24238) Or ( _
            GrhIndex >= 24303 And GrhIndex <= 24318) Or (GrhIndex >= 468 And GrhIndex <= 483) Or (GrhIndex >= 44668 And GrhIndex <= 44683) Or (GrhIndex >= 24143 And _
            GrhIndex <= 24158) Or (GrhIndex >= 12628 And GrhIndex <= 12643) Or (GrhIndex >= 2948 And GrhIndex <= 2963)
End Function

Private Function ViewportNeighbor_IsLavaGrh(ByVal GrhIndex As Long) As Boolean
    ViewportNeighbor_IsLavaGrh = (GrhIndex >= 57400 And GrhIndex <= 57415) Or (GrhIndex >= 16101 And GrhIndex <= 16116) Or (GrhIndex >= 26767 And GrhIndex <= 26782)
End Function
