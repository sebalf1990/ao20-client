Attribute VB_Name = "ModViewportNeighbors"
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Const CSM_FIVE_LAYER_SIGNATURE_NEIGHBOR As Long = &H324C3557 ' same as Recursos.CSM_FIVE_LAYER_SIGNATURE

Private Const DIR_NORTH As Integer = 1
Private Const DIR_SOUTH As Integer = 2
Private Const DIR_WEST  As Integer = 3
Private Const DIR_EAST  As Integer = 4

Private Const MAP_MAX_TILES As Integer = 100
Private Const NEIGHBOR_MIN_TE_COUNT As Integer = 60
Private Const MAX_MAP_ID_GUESS As Integer = 1200
Private Const CSM_MAX_LAYERS As Integer = 5
Private Const CSM_MAX_RECORDS As Long = 10000
Private Const CSM_SIZE_BLOCKED As Long = 5
Private Const CSM_SIZE_GRH As Long = 8
Private Const CSM_SIZE_TRIGGER As Long = 6
Private Const CSM_SIZE_LIGHT As Long = 9
Private Const CSM_SIZE_PARTICLE As Long = 8
Private Const CSM_SIZE_NPC As Long = 6
Private Const CSM_SIZE_OBJ As Long = 8
Private Const CSM_SIZE_TE As Long = 10

' E4d: precarga repartida. Con el jugador a PREFETCH_EDGE_TILES o menos de la linea de cruce de un lado,
' se cargan en la cache los vecinos del vecino de ese lado (lo que se necesita apenas cruce). La carga
' se hace en rebanadas desde el bucle principal con un presupuesto de ms por frame (VB6 un solo hilo).
Private Const PREFETCH_EDGE_TILES As Integer = 15
Private Const PREFETCH_BUDGET_IDLE_MS As Double = 3#
Private Const PREFETCH_BUDGET_MOVING_MS As Double = 1#
' Tras un cambio de mapa la posicion del jugador sigue siendo la vieja hasta el CharacterCreate:
' la precarga por cercania espera un poco (los cardinales faltantes no esperan).
Private Const PREFETCH_SETTLE_MS As Double = 300#
' Tamano de rebanada: registros del .csm por paso (cada Grh puede cargar una textura) y filas por paso
' al vaciar el slot. Entre paso y paso se mira el reloj.
Private Const LOAD_CHUNK_RECORDS As Long = 64
Private Const CLEAR_CHUNK_ROWS As Integer = 10
Private Const MAX_FAILED_MAPS As Integer = 8

' Fases de la carga en curso (una sola a la vez).
Private Const JOB_IDLE As Integer = 0
Private Const JOB_CLEAR As Integer = 1
Private Const JOB_BLOCKED As Integer = 2
Private Const JOB_LAYERS As Integer = 3
Private Const JOB_TAIL As Integer = 4

' Mismo layout que Recursos.tMapHeader (5 capas). En un .csm viejo de 4 capas las capas 3 y 4 del
' archivo van a 4 y 5 (como Recursos.ReadCsmHeader) y la 3 queda vacia.
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(1 To 5) As Long
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
' Capas del bloque: 1 piso, 2 encima del piso, 3 capa 3 vieja (capa 4 del mapa), 4 techos.
Private Type tNeighborBlock
    Graphic(1 To 4) As Grh
    Blocked As Integer
End Type

' E4c/E4d: cache de vecinos por id de mapa (= archivo mapa{id}.csm realmente abierto, R2), con
' NEIGHBOR_POOL_SIZE bloques y desalojo LRU. 12 bloques * 1,64 MB = ~19,7 MB, reservados recien la
' primera vez que se carga un vecino (en HUD clasico puro no se reserva nada). 12 = 4 cardinales +
' 4 esquinas + lo precargado cerca de una esquina (2 vecinos lejanos + el mapa actual) + 1 de margen.
Private Const NEIGHBOR_POOL_SIZE As Integer = 12

' Plan del mapa actual por direccion (se recalcula en cada cambio de mapa). g_neighbor_map es el mapa
' elegido para ese lado aunque todavia no este cargado; g_neighbor_loaded dice si ya se dibuja.
Private g_neighbor_map(1 To 4) As Integer
Private g_neighbor_count(1 To 4) As Integer
Private g_neighbor_source_coord(1 To 4) As Integer
Private g_neighbor_dest_coord(1 To 4) As Integer
Private g_neighbor_offset(1 To 4) As Integer
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
' E4d: un slot con g_slot_pin = g_pin_gen no se desaloja (cardinales del mapa actual y lo que se
' precarga para el lado al que el jugador esta cerca). g_pin_gen sube en cada cambio de mapa y cada
' vez que cambia el conjunto de lados cercanos.
Private g_slot_pin(1 To NEIGHBOR_POOL_SIZE) As Long
Private g_pin_gen As Long
Private g_near_mask As Integer
' E4d: vecinos ya resueltos de cada mapa en cache (la lista de TE se lee al cargarlo, en la misma
' apertura del archivo): la precarga y el cambio de mapa no necesitan otra pasada por el .csm.
Private g_slot_nb_map(1 To NEIGHBOR_POOL_SIZE, 1 To 4) As Integer
Private g_slot_nb_count(1 To NEIGHBOR_POOL_SIZE, 1 To 4) As Integer
Private g_slot_nb_source(1 To NEIGHBOR_POOL_SIZE, 1 To 4) As Integer
Private g_slot_nb_dest(1 To NEIGHBOR_POOL_SIZE, 1 To 4) As Integer

' E4d: carga en curso (en rebanadas o de una).
Private g_job_phase As Integer
Private g_job_map As Integer
Private g_job_slot As Integer
Private g_job_fh As Integer
Private g_job_header As tMapHeader
Private g_job_row As Integer
Private g_job_layer As Integer
Private g_job_done As Long
Private g_job_ms As Double
Private g_job_ticks As Long
Private g_job_ok As Boolean
Private g_job_reason As String
Private g_failed_maps(1 To MAX_FAILED_MAPS) As Integer
Private g_failed_count As Integer
Private g_deferred_count As Integer
Private g_prepare_ms As Double
Private g_qpc_freq As Double

' Cambio de mapa (SwitchMap en FULL): engancha lo que ya esta en cache y carga de una los que faltan.
Public Sub ViewportNeighbor_PrepareForMap(ByVal ResourceMap As Integer)
    Call ViewportNeighbor_Prepare(ResourceMap, False)
End Sub

' E4d: al entrar con F11 se enganchan los vecinos que ya estan en cache y los que faltan se cargan en
' rebanadas desde ViewportNeighbor_IdleTick (unos frames sin vecino en vez de un tiron).
Public Sub ViewportNeighbor_EnterFull(ByVal ResourceMap As Integer)
    On Error GoTo ViewportNeighbor_EnterFull_Err

    Dim startMs As Double

    startMs = ViewportNeighbor_Now()
    Call ViewportNeighbor_Prepare(ResourceMap, True)
    Call ViewportDebug_AppendDiagLog("[PERF] EnterFull neighbors=" & format$(ViewportNeighbor_Now() - startMs, "0.0") & "ms deferred=" & g_deferred_count)
    Exit Sub
ViewportNeighbor_EnterFull_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_EnterFull", Erl)
End Sub

' E4d: llamar antes de CargarMapa. Si la precarga en curso tiene abierto justo el mapa que se va a
' cargar, se corta (con Compresion=1 los dos usarian el mismo temporal).
Public Sub ViewportNeighbor_BeforeMapLoad(ByVal ResourceMap As Integer)
    On Error GoTo ViewportNeighbor_BeforeMapLoad_Err

    If g_job_phase = JOB_IDLE Then Exit Sub
    If g_job_map = ResourceMap Then Call ViewportNeighbor_JobAbort("map change")
    Exit Sub
ViewportNeighbor_BeforeMapLoad_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_BeforeMapLoad", Erl)
End Sub

' E4d: una vez por vuelta del bucle principal (engine.start, en juego). Solo en FULL. Avanza la carga
' en curso dentro del presupuesto del frame o elige la proxima precarga.
Public Sub ViewportNeighbor_IdleTick()
    On Error GoTo ViewportNeighbor_IdleTick_Err

    Dim budgetMs As Double

    If Not ViewportDebug_IsFullscreen() Then
        If g_job_phase <> JOB_IDLE Then Call ViewportNeighbor_JobAbort("left full")
        Exit Sub
    End If
    If g_cached_resource_map < 1 Then Exit Sub

    Call ViewportNeighbor_UpdatePins
    If g_job_phase = JOB_IDLE Then
        If Not ViewportNeighbor_StartNextPrefetch() Then Exit Sub
    End If

    If UserMoving Then
        budgetMs = PREFETCH_BUDGET_MOVING_MS
    Else
        budgetMs = PREFETCH_BUDGET_IDLE_MS
    End If
    g_job_ticks = g_job_ticks + 1
    If ViewportNeighbor_JobRun(budgetMs) Then Call ViewportNeighbor_JobFinished
    Exit Sub
ViewportNeighbor_IdleTick_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_IdleTick", Erl)
End Sub

Private Sub ViewportNeighbor_Prepare(ByVal ResourceMap As Integer, ByVal deferMisses As Boolean)
    On Error GoTo ViewportNeighbor_Prepare_Err

    Dim direction As Integer
    Dim slot As Integer
    Dim currentSlot As Integer
    Dim targetMap(1 To 4) As Integer
    Dim sourceCoord(1 To 4) As Integer
    Dim destCoord(1 To 4) As Integer
    Dim hits As Integer
    Dim loads As Integer
    Dim summary As String
    Dim scanText As String

    Call ViewportNeighbor_ResetDirections
    g_cached_resource_map = ResourceMap
    g_failed_count = 0
    g_deferred_count = 0
    g_near_mask = -1
    g_prepare_ms = ViewportNeighbor_Now()
    If ResourceMap < 1 Then
        If g_job_phase <> JOB_IDLE Then Call ViewportNeighbor_JobAbort("no map")
        Exit Sub
    End If

    ' E4d: si el mapa nuevo esta en cache (se entro a un vecino) su lista de vecinos ya esta resuelta y
    ' no se relee el .csm; si no, una pasada por su lista de TE.
    currentSlot = ViewportNeighbor_FindSlot(ResourceMap)
    If currentSlot > 0 Then
        For direction = DIR_NORTH To DIR_EAST
            targetMap(direction) = g_slot_nb_map(currentSlot, direction)
            g_neighbor_count(direction) = g_slot_nb_count(currentSlot, direction)
            sourceCoord(direction) = g_slot_nb_source(currentSlot, direction)
            destCoord(direction) = g_slot_nb_dest(currentSlot, direction)
        Next direction
        Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
        scanText = " scan=cache"
    Else
        Call ViewportNeighbor_FindNeighbors(ResourceMap, targetMap, g_neighbor_count, sourceCoord, destCoord)
        scanText = " scan=file"
    End If

    For direction = DIR_NORTH To DIR_EAST
        If targetMap(direction) > 0 And g_neighbor_count(direction) >= NEIGHBOR_MIN_TE_COUNT Then
            g_neighbor_map(direction) = targetMap(direction)
            g_neighbor_source_coord(direction) = sourceCoord(direction)
            g_neighbor_dest_coord(direction) = destCoord(direction)
            g_neighbor_offset(direction) = destCoord(direction) - sourceCoord(direction)
        End If
    Next direction

    ' Una precarga a medias: si es uno de los vecinos de ahora se termina, si no se corta.
    Call ViewportNeighbor_JobSettle

    ' Primero se fijan los vecinos que ya estan en cache, asi el LRU no desaloja uno que se necesita
    ' ahora al cargar los que faltan.
    g_pin_gen = g_pin_gen + 1
    For direction = DIR_NORTH To DIR_EAST
        If g_neighbor_map(direction) > 0 Then
            slot = ViewportNeighbor_FindSlot(g_neighbor_map(direction))
            If slot > 0 Then
                Call ViewportNeighbor_Attach(direction, slot)
                hits = hits + 1
            End If
        End If
    Next direction

    For direction = DIR_NORTH To DIR_EAST
        If g_neighbor_map(direction) > 0 And Not g_neighbor_loaded(direction) Then
            ' Un mapa vecino por dos lados comparte slot: el segundo lado lo encuentra ya cargado.
            slot = ViewportNeighbor_FindSlot(g_neighbor_map(direction))
            If slot > 0 Then
                Call ViewportNeighbor_Attach(direction, slot)
            ElseIf deferMisses Then
                g_deferred_count = g_deferred_count + 1
            Else
                slot = ViewportNeighbor_AcquireSlot()
                If slot > 0 Then
                    If ViewportNeighbor_LoadNow(slot, g_neighbor_map(direction), Mid$("NSWE", direction, 1)) Then
                        Call ViewportNeighbor_Attach(direction, slot)
                        loads = loads + 1
                    End If
                End If
            End If
        End If
        summary = summary & " " & Mid$("NSWE", direction, 1) & "=" & g_neighbor_map(direction)
    Next direction

    Call ViewportDebug_AppendDiagLog("[NEIGH] map=" & ResourceMap & summary & " hits=" & hits & " loads=" & loads & " deferred=" & g_deferred_count & scanText)
    Exit Sub
ViewportNeighbor_Prepare_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModViewportNeighbors.ViewportNeighbor_Prepare", Erl)
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
    Dim slot As Integer
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
                If ViewportNeighbor_MapVirtualTile(virtualX, virtualY, slot, localX, localY) Then
                    screenX = BaseScreenX + virtualX * TilePixelWidth
                    screenY = BaseScreenY + virtualY * TilePixelHeight
                    With g_pool_data(localX, localY, slot)
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
        g_neighbor_offset(direction) = 0
        g_neighbor_loaded(direction) = False
        g_neighbor_slot(direction) = 0
    Next direction
End Sub

Private Sub ViewportNeighbor_Attach(ByVal direction As Integer, ByVal slot As Integer)
    g_neighbor_slot(direction) = slot
    g_neighbor_loaded(direction) = True
    Call ViewportNeighbor_Touch(slot)
End Sub

' Marca el slot como usado ahora y lo fija contra el desalojo en la generacion vigente.
Private Sub ViewportNeighbor_Touch(ByVal slot As Integer)
    If slot < 1 Or slot > NEIGHBOR_POOL_SIZE Then Exit Sub
    g_neighbor_use_clock = g_neighbor_use_clock + 1
    g_slot_last_use(slot) = g_neighbor_use_clock
    g_slot_pin(slot) = g_pin_gen
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

Private Function ViewportNeighbor_IsJobSlot(ByVal slot As Integer) As Boolean
    ViewportNeighbor_IsJobSlot = (g_job_phase <> JOB_IDLE And slot = g_job_slot)
End Function

' Devuelve un slot libre o, si no hay, el usado hace mas tiempo entre los no fijados. Nunca el de la
' carga en curso. El slot devuelto queda libre (g_slot_map = 0) hasta que la carga termine bien.
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
        If g_slot_map(slot) = 0 And Not ViewportNeighbor_IsJobSlot(slot) Then
            best = slot
            Exit For
        End If
    Next slot

    If best = 0 Then
        For slot = 1 To NEIGHBOR_POOL_SIZE
            If g_slot_pin(slot) <> g_pin_gen And Not ViewportNeighbor_IsJobSlot(slot) Then
                If best = 0 Then
                    best = slot
                ElseIf g_slot_last_use(slot) < g_slot_last_use(best) Then
                    best = slot
                End If
            End If
        Next slot
    End If

    If best > 0 Then
        g_slot_map(best) = 0
        Call ViewportNeighbor_ForgetSlot(best)
    End If
    ViewportNeighbor_AcquireSlot = best
End Function

' Red de seguridad: ningun lado sigue apuntando a un slot que se va a reusar.
Private Sub ViewportNeighbor_ForgetSlot(ByVal slot As Integer)
    Dim direction As Integer

    For direction = DIR_NORTH To DIR_EAST
        If g_neighbor_slot(direction) = slot Then
            g_neighbor_slot(direction) = 0
            g_neighbor_loaded(direction) = False
        End If
    Next direction
End Sub

' E4d: al cambiar el conjunto de lados cercanos se renueva la generacion de fijados: quedan fijos los
' cardinales cargados y lo que la proxima eleccion de precarga vuelva a pedir.
Private Sub ViewportNeighbor_UpdatePins()
    Dim direction As Integer
    Dim mask As Integer

    For direction = DIR_NORTH To DIR_EAST
        mask = mask * 2
        If ViewportNeighbor_IsNear(direction) Then mask = mask + 1
    Next direction
    If mask = g_near_mask Then Exit Sub

    g_near_mask = mask
    g_pin_gen = g_pin_gen + 1
    For direction = DIR_NORTH To DIR_EAST
        If g_neighbor_loaded(direction) Then Call ViewportNeighbor_Touch(g_neighbor_slot(direction))
    Next direction
End Sub

' Tiles entre el jugador y la linea de cruce de ese lado (la coordenada de TE mas usada; si el lado no
' tiene vecino, la linea habitual 10/91).
Private Function ViewportNeighbor_DistanceToSide(ByVal direction As Integer) As Long
    Dim lineCoord As Long

    lineCoord = g_neighbor_source_coord(direction)
    If lineCoord = 0 Then
        If direction = DIR_NORTH Or direction = DIR_WEST Then
            lineCoord = 10
        Else
            lineCoord = 91
        End If
    End If
    Select Case direction
        Case DIR_NORTH
            ViewportNeighbor_DistanceToSide = UserPos.y - lineCoord
        Case DIR_SOUTH
            ViewportNeighbor_DistanceToSide = lineCoord - UserPos.y
        Case DIR_WEST
            ViewportNeighbor_DistanceToSide = UserPos.x - lineCoord
        Case Else
            ViewportNeighbor_DistanceToSide = lineCoord - UserPos.x
    End Select
End Function

Private Function ViewportNeighbor_IsNear(ByVal direction As Integer) As Boolean
    If g_neighbor_map(direction) = 0 Then Exit Function
    ViewportNeighbor_IsNear = (ViewportNeighbor_DistanceToSide(direction) <= PREFETCH_EDGE_TILES)
End Function

' True si mapId hace falta y no esta en cache. Si ya esta, lo fija (para que no se desaloje).
Private Function ViewportNeighbor_IsWantedMissing(ByVal mapId As Integer) As Boolean
    Dim slot As Integer

    If mapId < 1 Then Exit Function
    slot = ViewportNeighbor_FindSlot(mapId)
    If slot > 0 Then
        If g_slot_pin(slot) <> g_pin_gen Then Call ViewportNeighbor_Touch(slot)
        Exit Function
    End If
    If ViewportNeighbor_IsFailed(mapId) Then Exit Function
    ViewportNeighbor_IsWantedMissing = True
End Function

' Prioridad: 1) cardinales del mapa actual que faltan (F11 diferido), el mas cercano primero;
' 2) vecinos del vecino del lado al que el jugador esta cerca (incluye el mapa actual, que pasa a ser
' el vecino opuesto apenas cruce).
Private Function ViewportNeighbor_PickCandidate(ByRef mapId As Integer, ByRef reason As String) As Boolean
    Dim direction As Integer
    Dim second As Integer
    Dim slot As Integer
    Dim best As Integer
    Dim distance As Long
    Dim bestDistance As Long
    Dim candidate As Integer

    For direction = DIR_NORTH To DIR_EAST
        If g_neighbor_map(direction) > 0 And Not g_neighbor_loaded(direction) Then
            slot = ViewportNeighbor_FindSlot(g_neighbor_map(direction))
            If slot > 0 Then
                Call ViewportNeighbor_Attach(direction, slot)
            ElseIf Not ViewportNeighbor_IsFailed(g_neighbor_map(direction)) Then
                distance = ViewportNeighbor_DistanceToSide(direction)
                If best = 0 Or distance < bestDistance Then
                    best = direction
                    bestDistance = distance
                End If
            End If
        End If
    Next direction
    If best > 0 Then
        mapId = g_neighbor_map(best)
        reason = Mid$("NSWE", best, 1)
        ViewportNeighbor_PickCandidate = True
        Exit Function
    End If

    If ViewportNeighbor_Now() - g_prepare_ms < PREFETCH_SETTLE_MS Then Exit Function

    For direction = DIR_NORTH To DIR_EAST
        If g_neighbor_loaded(direction) Then
            If ViewportNeighbor_IsNear(direction) Then
                slot = g_neighbor_slot(direction)
                For second = DIR_NORTH To DIR_EAST
                    candidate = g_slot_nb_map(slot, second)
                    If candidate > 0 And g_slot_nb_count(slot, second) >= NEIGHBOR_MIN_TE_COUNT Then
                        If ViewportNeighbor_IsWantedMissing(candidate) Then
                            mapId = candidate
                            reason = Mid$("NSWE", direction, 1) & "." & Mid$("NSWE", second, 1)
                            ViewportNeighbor_PickCandidate = True
                            Exit Function
                        End If
                    End If
                Next second
            End If
        End If
    Next direction
End Function

Private Function ViewportNeighbor_StartNextPrefetch() As Boolean
    Dim mapId As Integer
    Dim reason As String
    Dim slot As Integer

    If Not ViewportNeighbor_PickCandidate(mapId, reason) Then Exit Function
    slot = ViewportNeighbor_AcquireSlot()
    If slot = 0 Then Exit Function
    g_job_ticks = 0
    If Not ViewportNeighbor_JobStart(slot, mapId, reason) Then
        Call ViewportNeighbor_MarkFailed(mapId)
        Exit Function
    End If
    ViewportNeighbor_StartNextPrefetch = True
End Function

Private Sub ViewportNeighbor_JobFinished()
    Dim direction As Integer

    If Not g_job_ok Then
        Call ViewportDebug_AppendDiagLog("[NEIGH] prefetch fail map=" & g_job_map & " for=" & g_job_reason)
        Exit Sub
    End If
    Call ViewportDebug_AppendDiagLog("[NEIGH] prefetch map=" & g_job_map & " for=" & g_job_reason & " ms=" & format$(g_job_ms, "0.0") & " ticks=" & g_job_ticks)
    Call ViewportNeighbor_Touch(g_job_slot)
    For direction = DIR_NORTH To DIR_EAST
        If Not g_neighbor_loaded(direction) And g_neighbor_map(direction) = g_job_map Then
            Call ViewportNeighbor_Attach(direction, g_job_slot)
        End If
    Next direction
End Sub

' Carga en curso al cambiar de mapa: se termina de una si el mapa nuevo la necesita, si no se corta.
Private Sub ViewportNeighbor_JobSettle()
    Dim direction As Integer
    Dim needed As Boolean

    If g_job_phase = JOB_IDLE Then Exit Sub
    For direction = DIR_NORTH To DIR_EAST
        If g_neighbor_map(direction) = g_job_map Then needed = True
    Next direction
    If needed Then
        Call ViewportNeighbor_JobRun(0#)
        g_job_reason = g_job_reason & "+cross"
        Call ViewportNeighbor_JobFinished
    Else
        Call ViewportNeighbor_JobAbort("map change")
    End If
End Sub

Private Function ViewportNeighbor_LoadNow(ByVal slot As Integer, ByVal mapId As Integer, ByVal reason As String) As Boolean
    If Not ViewportNeighbor_JobStart(slot, mapId, reason) Then
        Call ViewportNeighbor_MarkFailed(mapId)
        Exit Function
    End If
    Call ViewportNeighbor_JobRun(0#)
    ViewportNeighbor_LoadNow = g_job_ok
End Function

' Abre el .csm del mapa y lee su cabecera. La carga sigue en ViewportNeighbor_JobRun.
Private Function ViewportNeighbor_JobStart(ByVal slot As Integer, ByVal mapId As Integer, ByVal reason As String) As Boolean
    On Error GoTo ViewportNeighbor_JobStart_Err

    Dim MS      As tMapSize
    Dim MD      As tMapDat
    Dim MapPath As String
    Dim startMs As Double

    startMs = ViewportNeighbor_Now()
    g_job_ok = False
    g_job_map = mapId
    g_job_slot = slot
    g_job_reason = reason
    g_job_ms = 0
    MapPath = ViewportNeighbor_MapPath(mapId)
    If LenB(Dir$(MapPath)) = 0 Then Exit Function

    g_job_fh = FreeFile
    Open MapPath For Binary As #g_job_fh
    If Not ViewportNeighbor_ReadHeader(g_job_fh, g_job_header) Then
        Call ViewportNeighbor_JobClose
        Exit Function
    End If
    Get #g_job_fh, , MS
    Get #g_job_fh, , MD

    g_job_phase = JOB_CLEAR
    g_job_row = 1
    g_job_layer = 1
    g_job_done = 0
    g_job_ms = ViewportNeighbor_Now() - startMs
    ViewportNeighbor_JobStart = True
    Exit Function
ViewportNeighbor_JobStart_Err:
    Dim errNumber As Long, errDescription As String, errLine As Long
    errNumber = Err.Number: errDescription = Err.Description: errLine = Erl
    On Error Resume Next
    Call ViewportNeighbor_JobClose
    Call RegistrarError(errNumber, errDescription, "ModViewportNeighbors.ViewportNeighbor_JobStart", errLine)
End Function

' Avanza la carga en curso. budgetMs = 0: hasta el final. Devuelve True cuando la carga termino (bien
' o mal, ver g_job_ok).
Private Function ViewportNeighbor_JobRun(ByVal budgetMs As Double) As Boolean
    On Error GoTo ViewportNeighbor_JobRun_Err

    Dim startMs    As Double
    Dim blocks()   As tDatosBloqueados
    Dim graphics() As tDatosGrh
    Dim TEs()      As tDatosTE
    Dim teCount    As Long
    Dim chunk      As Long
    Dim i          As Long
    Dim x          As Integer
    Dim y          As Integer
    Dim lastRow    As Integer

    startMs = ViewportNeighbor_Now()
    Do While g_job_phase <> JOB_IDLE
        Select Case g_job_phase
            Case JOB_CLEAR
                ' La carga solo escribe los tiles presentes en el .csm: primero se vacia el slot.
                lastRow = g_job_row + CLEAR_CHUNK_ROWS - 1
                If lastRow > MAP_MAX_TILES Then lastRow = MAP_MAX_TILES
                For y = g_job_row To lastRow
                    For x = 1 To MAP_MAX_TILES
                        g_pool_data(x, y, g_job_slot) = g_empty_block
                    Next x
                Next y
                g_job_row = lastRow + 1
                If g_job_row > MAP_MAX_TILES Then
                    g_job_phase = JOB_BLOCKED
                    g_job_done = 0
                End If

            Case JOB_BLOCKED
                chunk = ViewportNeighbor_ChunkSize(g_job_header.NumeroBloqueados - g_job_done)
                If chunk <= 0 Then
                    g_job_phase = JOB_LAYERS
                    g_job_layer = 1
                    g_job_done = 0
                Else
                    ReDim blocks(1 To chunk)
                    Get #g_job_fh, , blocks
                    For i = 1 To chunk
                        ' E1: coordenadas fuera de 1..MAP_MAX_TILES se saltean (no abortan la carga).
                        If blocks(i).x >= 1 And blocks(i).x <= MAP_MAX_TILES And blocks(i).y >= 1 And blocks(i).y <= MAP_MAX_TILES Then
                            g_pool_data(blocks(i).x, blocks(i).y, g_job_slot).Blocked = blocks(i).lados
                        End If
                    Next i
                    g_job_done = g_job_done + chunk
                End If

            Case JOB_LAYERS
                If g_job_layer > CSM_MAX_LAYERS Then
                    g_job_phase = JOB_TAIL
                Else
                    chunk = ViewportNeighbor_ChunkSize(g_job_header.NumeroLayers(g_job_layer) - g_job_done)
                    If chunk <= 0 Then
                        g_job_layer = g_job_layer + 1
                        g_job_done = 0
                    Else
                        ReDim graphics(1 To chunk)
                        Get #g_job_fh, , graphics
                        For i = 1 To chunk
                            Call ViewportNeighbor_StoreGrh(g_job_slot, g_job_layer, graphics(i))
                        Next i
                        g_job_done = g_job_done + chunk
                    End If
                End If

            Case Else
                ' Resto del archivo: solo interesa la lista de TE (vecinos de este mapa, para la precarga).
                Call ViewportNeighbor_SkipTail(g_job_fh, g_job_header)
                teCount = g_job_header.NumeroTE
                If teCount > 0 Then
                    ReDim TEs(1 To teCount)
                    Get #g_job_fh, , TEs
                End If
                Call ViewportNeighbor_JobClose
                Call ViewportNeighbor_StoreSlotNeighbors(g_job_slot, TEs, teCount)
                g_slot_map(g_job_slot) = g_job_map
                g_job_ok = True
        End Select

        If budgetMs > 0 And g_job_phase <> JOB_IDLE Then
            If ViewportNeighbor_Now() - startMs >= budgetMs Then Exit Do
        End If
    Loop

    g_job_ms = g_job_ms + (ViewportNeighbor_Now() - startMs)
    ViewportNeighbor_JobRun = (g_job_phase = JOB_IDLE)
    Exit Function
ViewportNeighbor_JobRun_Err:
    Dim errNumber As Long, errDescription As String, errLine As Long
    errNumber = Err.Number: errDescription = Err.Description: errLine = Erl
    On Error Resume Next
    g_job_ok = False
    Call ViewportNeighbor_JobClose
    Call ViewportNeighbor_MarkFailed(g_job_map)
    Call RegistrarError(errNumber, errDescription, "ModViewportNeighbors.ViewportNeighbor_JobRun", errLine)
    ViewportNeighbor_JobRun = True
End Function

Private Function ViewportNeighbor_ChunkSize(ByVal remaining As Long) As Long
    If remaining > LOAD_CHUNK_RECORDS Then
        ViewportNeighbor_ChunkSize = LOAD_CHUNK_RECORDS
    Else
        ViewportNeighbor_ChunkSize = remaining
    End If
End Function

' Cierra el archivo de la carga en curso y borra su temporal (Compresion=1). El slot queda libre si la
' carga no habia terminado (g_slot_map se escribe recien al final).
Private Sub ViewportNeighbor_JobClose()
    On Error Resume Next
    If g_job_fh <> 0 Then Close #g_job_fh
    g_job_fh = 0
    g_job_phase = JOB_IDLE
    If g_job_map > 0 Then Call ViewportNeighbor_ReleaseMapFile(g_job_map)
End Sub

Private Sub ViewportNeighbor_JobAbort(ByVal reason As String)
    If g_job_phase = JOB_IDLE Then Exit Sub
    Call ViewportDebug_AppendDiagLog("[NEIGH] prefetch abort map=" & g_job_map & " for=" & g_job_reason & " why=" & reason)
    g_job_ok = False
    Call ViewportNeighbor_JobClose
End Sub

' Capa del archivo -> capa del bloque (ver tNeighborBlock).
Private Sub ViewportNeighbor_StoreGrh(ByVal slot As Integer, ByVal fileLayer As Integer, ByRef record As tDatosGrh)
    Dim x As Integer
    Dim y As Integer
    Dim blockLayer As Integer

    x = record.x
    y = record.y
    ' E1: coordenadas fuera de 1..MAP_MAX_TILES se saltean (no abortan la carga).
    If x < 1 Or x > MAP_MAX_TILES Or y < 1 Or y > MAP_MAX_TILES Then Exit Sub

    Select Case fileLayer
        Case 1
            blockLayer = 1
        Case 2
            blockLayer = 2
        Case 4
            blockLayer = 3
        Case 5
            blockLayer = 4
        Case Else
            Exit Sub
    End Select

    With g_pool_data(x, y, slot)
        .Graphic(blockLayer).GrhIndex = record.GrhIndex
        .Graphic(blockLayer).x = x * TilePixelWidth
        .Graphic(blockLayer).y = y * TilePixelHeight
        Call InitGrh(.Graphic(blockLayer), record.GrhIndex)
        If blockLayer = 1 Then
            If ViewportNeighbor_IsWaterGrh(.Graphic(1).GrhIndex) Then
                .Blocked = .Blocked Or FLAG_AGUA
            ElseIf ViewportNeighbor_IsLavaGrh(.Graphic(1).GrhIndex) Then
                .Blocked = .Blocked Or FLAG_LAVA
            End If
        End If
    End With
End Sub

Private Sub ViewportNeighbor_StoreSlotNeighbors(ByVal slot As Integer, ByRef TEs() As tDatosTE, ByVal teCount As Long)
    Dim direction As Integer
    Dim bestMap(1 To 4) As Integer
    Dim bestCount(1 To 4) As Integer
    Dim bestSourceCoord(1 To 4) As Integer
    Dim bestDestCoord(1 To 4) As Integer

    Call ViewportNeighbor_ResolveFromTEs(TEs, teCount, bestMap, bestCount, bestSourceCoord, bestDestCoord)
    For direction = DIR_NORTH To DIR_EAST
        g_slot_nb_map(slot, direction) = bestMap(direction)
        g_slot_nb_count(slot, direction) = bestCount(direction)
        g_slot_nb_source(slot, direction) = bestSourceCoord(direction)
        g_slot_nb_dest(slot, direction) = bestDestCoord(direction)
    Next direction
End Sub

Private Sub ViewportNeighbor_MarkFailed(ByVal mapId As Integer)
    If ViewportNeighbor_IsFailed(mapId) Then Exit Sub
    If g_failed_count >= MAX_FAILED_MAPS Then Exit Sub
    g_failed_count = g_failed_count + 1
    g_failed_maps(g_failed_count) = mapId
End Sub

Private Function ViewportNeighbor_IsFailed(ByVal mapId As Integer) As Boolean
    Dim i As Integer

    For i = 1 To g_failed_count
        If g_failed_maps(i) = mapId Then
            ViewportNeighbor_IsFailed = True
            Exit Function
        End If
    Next i
End Function

' Milisegundos de alta resolucion, independiente de PERF_ENABLED (el presupuesto de la precarga lo
' necesita siempre). 0 si el equipo no tiene contador: la carga se hace entera en un paso.
Private Function ViewportNeighbor_Now() As Double
    Dim counter As Currency
    Dim freq As Currency

    If g_qpc_freq = 0 Then
        If QueryPerformanceFrequency(freq) = 0 Then
            g_qpc_freq = -1
        Else
            g_qpc_freq = CDbl(freq)
        End If
    End If
    If g_qpc_freq <= 0 Then Exit Function
    If QueryPerformanceCounter(counter) = 0 Then Exit Function
    ViewportNeighbor_Now = CDbl(counter) / g_qpc_freq * 1000#
End Function

' Lee la cabecera del .csm (con o sin firma W5L2) al layout de 5 capas. False si no es valida.
Private Function ViewportNeighbor_ReadHeader(ByVal fh As Integer, ByRef MH As tMapHeader) As Boolean
    Dim emptyHeader As tMapHeader
    Dim signatureOrBlocked As Long
    Dim layer As Integer

    MH = emptyHeader
    Get #fh, , signatureOrBlocked
    ' Plan 15.001: this reader only knows the legacy 4-layer CSM. A five-layer
    ' file (signature W5L2, see Recursos.ReadCsmHeader) is skipped, not misread.
    If signatureOrBlocked = CSM_FIVE_LAYER_SIGNATURE_NEIGHBOR Then Exit Function
    MH.NumeroBloqueados = signatureOrBlocked
    Get #fh, , MH.NumeroLayers(1)
    Get #fh, , MH.NumeroLayers(2)
    Get #fh, , MH.NumeroLayers(4)
    Get #fh, , MH.NumeroLayers(5)
    Get #fh, , MH.NumeroTriggers
    Get #fh, , MH.NumeroLuces
    Get #fh, , MH.NumeroParticulas
    Get #fh, , MH.NumeroNPCs
    Get #fh, , MH.NumeroOBJs
    Get #fh, , MH.NumeroTE

    If MH.NumeroBloqueados < 0 Or MH.NumeroBloqueados > CSM_MAX_RECORDS Then Exit Function
    For layer = 1 To CSM_MAX_LAYERS
        If MH.NumeroLayers(layer) < 0 Or MH.NumeroLayers(layer) > CSM_MAX_RECORDS Then Exit Function
    Next layer
    If MH.NumeroTriggers < 0 Or MH.NumeroLuces < 0 Or MH.NumeroParticulas < 0 Then Exit Function
    If MH.NumeroNPCs < 0 Or MH.NumeroOBJs < 0 Or MH.NumeroTE < 0 Then Exit Function
    ViewportNeighbor_ReadHeader = True
End Function

' E4b: una sola apertura del .csm actual y una sola pasada por su lista de TE para los 4 lados
' (antes se abria 4 veces, una por direccion).
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

    MapPath = ViewportNeighbor_MapPath(ResourceMap)
    If LenB(Dir$(MapPath)) = 0 Then Exit Sub

    fh = FreeFile
    Open MapPath For Binary As #fh
    If Not ViewportNeighbor_ReadHeader(fh, MH) Then
        Close #fh
        fh = 0
        Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
        Exit Sub
    End If
    Get #fh, , MS
    Get #fh, , MD
    Call ViewportNeighbor_SkipBlocksAndLayers(fh, MH)
    Call ViewportNeighbor_SkipTail(fh, MH)

    If MH.NumeroTE > 0 Then
        ReDim TEs(1 To MH.NumeroTE)
        Get #fh, , TEs
    End If
    Close #fh
    fh = 0
    Call ViewportNeighbor_ReleaseMapFile(ResourceMap)

    Call ViewportNeighbor_ResolveFromTEs(TEs, MH.NumeroTE, bestMap, bestCount, bestSourceCoord, bestDestCoord)
    Exit Sub
ViewportNeighbor_FindNeighbors_Err:
    Dim errNumber As Long, errDescription As String, errLine As Long
    errNumber = Err.Number: errDescription = Err.Description: errLine = Erl
    On Error Resume Next
    If fh <> 0 Then Close #fh
    Call ViewportNeighbor_ReleaseMapFile(ResourceMap)
    Call RegistrarError(errNumber, errDescription, "ModViewportNeighbors.ViewportNeighbor_FindNeighbors", errLine)
End Sub

' Vecino por lado a partir de una lista de TE. Mismo criterio de desempate que antes: por lado gana el
' primer mapa que alcanza el maximo de TE, recorriendo la lista en orden. Los arreglos de salida tienen
' que venir en cero.
Private Sub ViewportNeighbor_ResolveFromTEs(ByRef TEs() As tDatosTE, _
                                            ByVal teCount As Long, _
                                            ByRef bestMap() As Integer, _
                                            ByRef bestCount() As Integer, _
                                            ByRef bestSourceCoord() As Integer, _
                                            ByRef bestDestCoord() As Integer)
    Dim counts(1 To 4, 1 To MAX_MAP_ID_GUESS) As Integer
    Dim sourceCounts(1 To 4, 1 To MAP_MAX_TILES) As Integer
    Dim destCounts(1 To 4, 1 To MAP_MAX_TILES) As Integer
    Dim i           As Long
    Dim mapId       As Integer
    Dim direction   As Integer
    Dim sourceCoord As Integer
    Dim destCoord   As Integer

    If teCount <= 0 Then Exit Sub

    For i = 1 To teCount
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

    For i = 1 To teCount
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
End Sub

' Tile fuera del mapa actual -> slot y tile local del vecino que lo cubre.
Private Function ViewportNeighbor_MapVirtualTile(ByVal virtualX As Integer, _
                                                 ByVal virtualY As Integer, _
                                                 ByRef slot As Integer, _
                                                 ByRef localX As Integer, _
                                                 ByRef localY As Integer) As Boolean
    Dim direction As Integer
    Dim insideX As Boolean
    Dim insideY As Boolean

    insideX = (virtualX >= XMinMapSize And virtualX <= XMaxMapSize)
    insideY = (virtualY >= YMinMapSize And virtualY <= YMaxMapSize)

    If insideY Then
        If virtualX < XMinMapSize Then
            direction = DIR_WEST
        Else
            direction = DIR_EAST
        End If
        If Not g_neighbor_loaded(direction) Then Exit Function
        localX = virtualX + g_neighbor_offset(direction)
        localY = virtualY
        slot = g_neighbor_slot(direction)
    ElseIf insideX Then
        If virtualY < YMinMapSize Then
            direction = DIR_NORTH
        Else
            direction = DIR_SOUTH
        End If
        If Not g_neighbor_loaded(direction) Then Exit Function
        localX = virtualX
        localY = virtualY + g_neighbor_offset(direction)
        slot = g_neighbor_slot(direction)
    Else
        Exit Function
    End If

    If slot < 1 Then Exit Function
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

Private Sub ViewportNeighbor_SkipBlocksAndLayers(ByVal fh As Integer, ByRef MH As tMapHeader)
    Dim layer As Integer

    If MH.NumeroBloqueados > 0 Then Seek #fh, Seek(fh) + MH.NumeroBloqueados * CSM_SIZE_BLOCKED
    For layer = 1 To CSM_MAX_LAYERS
        If MH.NumeroLayers(layer) > 0 Then Seek #fh, Seek(fh) + MH.NumeroLayers(layer) * CSM_SIZE_GRH
    Next layer
End Sub

' Triggers, particulas, luces, objetos y NPCs (todo lo que hay entre las capas y la lista de TE).
Private Sub ViewportNeighbor_SkipTail(ByVal fh As Integer, ByRef MH As tMapHeader)
    If MH.NumeroTriggers > 0 Then Seek #fh, Seek(fh) + MH.NumeroTriggers * CSM_SIZE_TRIGGER
    If MH.NumeroParticulas > 0 Then Seek #fh, Seek(fh) + MH.NumeroParticulas * CSM_SIZE_PARTICLE
    If MH.NumeroLuces > 0 Then Seek #fh, Seek(fh) + MH.NumeroLuces * CSM_SIZE_LIGHT
    If MH.NumeroOBJs > 0 Then Seek #fh, Seek(fh) + MH.NumeroOBJs * CSM_SIZE_OBJ
    If MH.NumeroNPCs > 0 Then Seek #fh, Seek(fh) + MH.NumeroNPCs * CSM_SIZE_NPC
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
