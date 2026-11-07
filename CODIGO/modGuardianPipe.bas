Attribute VB_Name = "modGuardianPipe"
' E4c antitrampas: cliente <-> guardian (neao-guardian.exe) por pipe local. Diseno: e4-diseno.md (4, 5).
'
' El guardian (Rust) crea \\.\pipe\ao20-guardian-<pid de este proceso>. Este modulo:
'   - recibe el desafio del server (GuardianPipeOnChallenge) y recien ahi empieza a hablar con el guardian;
'   - manda "HELLO <N hex> <char_id> <intervalo>" y luego "BEAT" cada <intervalo> segundos;
'   - cada linea completa que contesta el guardian se reenvia al server (WriteGuardianHeartbeat).
' El cliente NO calcula firmas y nunca ve el secreto. Invariante (4): no se manda ningun heartbeat si no llego antes
' un eGuardianChallenge en esta conexion (GuardianPipeReset lo apaga en cada desconexion).
' NUNCA bloquea: el pipe se abre en modo NOWAIT, se lee solo si PeekNamedPipe dice que hay bytes, y sin pipe se
' reintenta cada 30 s sin mostrar nada. Ningun error muestra cartel ni cierra el cliente.
Option Explicit

Private Declare Function GP_CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GP_SetNamedPipeHandleState Lib "kernel32" Alias "SetNamedPipeHandleState" (ByVal hNamedPipe As Long, ByRef lpMode As Long, ByVal lpMaxCollectionCount As Long, ByVal lpCollectDataTimeout As Long) As Long
Private Declare Function GP_WriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GP_ReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GP_PeekNamedPipe Lib "kernel32" Alias "PeekNamedPipe" (ByVal hNamedPipe As Long, ByVal lpBuffer As Long, ByVal nBufferSize As Long, ByVal lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByVal lpBytesLeftThisMessage As Long) As Long
Private Declare Function GP_CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function GP_GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long

Private Const GP_GENERIC_READ_WRITE As Long = &HC0000000
Private Const GP_OPEN_EXISTING As Long = 3
Private Const GP_INVALID_HANDLE As Long = -1
Private Const GP_PIPE_NOWAIT As Long = 1               ' PIPE_READMODE_BYTE (0) Or PIPE_NOWAIT (1)
Private Const GP_PIPE_NAME_PREFIX As String = "\\.\pipe\ao20-guardian-"

Private Const GP_POLL_MS As Long = 250                 ' maximo una vuelta de trabajo cada 250 ms
Private Const GP_OPEN_RETRY_MS As Long = 30000         ' sin pipe: reintento cada 30 s
Private Const GP_FIRST_BEAT_RETRY_MS As Long = 3000    ' hasta la primera linea tras un HELLO: BEAT cada 3 s (el guardian
                                                       ' no firma hasta terminar sus chequeos; un BEAT temprano es inofensivo)
Private Const GP_DEFAULT_INTERVAL_SEC As Long = 20
Private Const GP_MAX_INTERVAL_SEC As Long = 3600
Private Const GP_MAX_LINE As Long = 512                ' tope de una linea de heartbeat (el server rechaza mas)
Private Const GP_READ_CHUNK As Long = 512
Private Const GP_N_HEX_LEN As Long = 32

Private m_Pending        As Boolean   ' hubo desafio en esta conexion
Private m_NHex           As String
Private m_CharId         As Long
Private m_IntervalSec    As Long
Private m_Pipe           As Long      ' 0 = cerrado
Private m_NeedHello      As Boolean
Private m_GotLineSinceHello As Boolean
Private m_RxBuffer       As String
Private m_RxOverflow     As Boolean
Private m_LastPoll       As Long
Private m_LastOpenTry    As Long
Private m_OpenTried      As Boolean
Private m_LastBeat       As Long

' Handler de eGuardianChallenge: guarda el desafio y (re)arranca la sesion con el guardian.
Public Sub GuardianPipeOnChallenge(ByVal nHex As String, ByVal intervalSec As Long)
    On Error GoTo GuardianPipeOnChallenge_Err
    Dim charId As Long
    Call GuardianPipeReset
    nHex = LCase$(nHex)
    If Not IsValidNHex(nHex) Then Exit Sub
    If intervalSec <= 0 Or intervalSec > GP_MAX_INTERVAL_SEC Then intervalSec = GP_DEFAULT_INTERVAL_SEC
    charId = GetSelectedCharIDFromName(userName)
    If charId <= 0 Then Exit Sub
    m_NHex = nHex
    m_CharId = charId
    m_IntervalSec = intervalSec
    m_NeedHello = True
    m_Pending = True
    Exit Sub
GuardianPipeOnChallenge_Err:
    Call GuardianPipeReset
End Sub

' Apaga todo: cierra el pipe y olvida el desafio. Se llama al desconectar, al cerrar el cliente y antes de cada desafio.
Public Sub GuardianPipeReset()
    On Error Resume Next
    Call ClosePipe
    m_Pending = False
    m_NHex = vbNullString
    m_CharId = 0
    m_IntervalSec = 0
    m_NeedHello = False
    m_GotLineSinceHello = False
    m_RxBuffer = vbNullString
    m_RxOverflow = False
    m_OpenTried = False
End Sub

' Se llama desde el loop principal (junto a UpdateAntiCheat). Sin desafio pendiente es solo un If.
Public Sub GuardianPipePoll()
    On Error GoTo GuardianPipePoll_Err
    If Not m_Pending Then Exit Sub
    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    If TicksElapsed(m_LastPoll, nowTick) < GP_POLL_MS And m_LastPoll <> 0 Then Exit Sub
    m_LastPoll = nowTick
    If m_Pipe = 0 Then
        If m_OpenTried Then
            If TicksElapsed(m_LastOpenTry, nowTick) < GP_OPEN_RETRY_MS Then Exit Sub
        End If
        m_OpenTried = True
        m_LastOpenTry = nowTick
        If Not OpenPipe() Then Exit Sub
        m_NeedHello = True
    End If
    If m_NeedHello Then
        If Not SendPipeLine("HELLO " & m_NHex & " " & CStr(m_CharId) & " " & CStr(m_IntervalSec)) Then Exit Sub
        m_NeedHello = False
        m_GotLineSinceHello = False
        m_LastBeat = nowTick
    End If
    If Not ReadPipeLines() Then Exit Sub
    If m_Pending Then Call MaybeSendBeat(nowTick)
    Exit Sub
GuardianPipePoll_Err:
    ' Nunca cartel ni crash: se cierra el pipe y se reintenta mas tarde.
    On Error Resume Next
    Call ClosePipe
    m_NeedHello = True
    m_OpenTried = True
    m_LastOpenTry = GetTickCountRaw()
End Sub

Private Sub MaybeSendBeat(ByVal nowTick As Long)
    Dim gapMs As Long
    If m_GotLineSinceHello Then
        gapMs = m_IntervalSec * 1000
    Else
        gapMs = GP_FIRST_BEAT_RETRY_MS
    End If
    If TicksElapsed(m_LastBeat, nowTick) < gapMs Then Exit Sub
    If SendPipeLine("BEAT") Then
        m_LastBeat = nowTick
    End If
End Sub

Private Function IsValidNHex(ByVal nHex As String) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(nHex) <> GP_N_HEX_LEN Then Exit Function
    For i = 1 To GP_N_HEX_LEN
        ch = Mid$(nHex, i, 1)
        If InStr(1, "0123456789abcdef", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsValidNHex = True
End Function

Private Function OpenPipe() As Boolean
    On Error GoTo OpenPipe_Err
    Dim h As Long
    Dim pipeMode As Long
    h = GP_CreateFile(GP_PIPE_NAME_PREFIX & CStr(GP_GetCurrentProcessId()), GP_GENERIC_READ_WRITE, 0&, 0&, GP_OPEN_EXISTING, 0&, 0&)
    If h = GP_INVALID_HANDLE Or h = 0 Then Exit Function
    pipeMode = GP_PIPE_NOWAIT
    Call GP_SetNamedPipeHandleState(h, pipeMode, 0&, 0&)
    m_Pipe = h
    m_RxBuffer = vbNullString
    m_RxOverflow = False
    OpenPipe = True
    Exit Function
OpenPipe_Err:
    OpenPipe = False
End Function

Private Sub ClosePipe()
    On Error Resume Next
    If m_Pipe <> 0 Then
        Call GP_CloseHandle(m_Pipe)
        m_Pipe = 0
    End If
End Sub

' Pipe roto: se cierra y se reintenta en 30 s (con HELLO de nuevo al reconectar).
Private Sub OnPipeBroken()
    On Error Resume Next
    Call ClosePipe
    m_NeedHello = True
    m_OpenTried = True
    m_LastOpenTry = GetTickCountRaw()
End Sub

Private Function SendPipeLine(ByVal text As String) As Boolean
    On Error GoTo SendPipeLine_Err
    Dim data() As Byte
    Dim written As Long
    Dim ok As Long
    If m_Pipe = 0 Then Exit Function
    data = StrConv(text & vbLf, vbFromUnicode)
    ok = GP_WriteFile(m_Pipe, data(0), UBound(data) + 1, written, 0&)
    If ok = 0 Or written <> UBound(data) + 1 Then
        Call OnPipeBroken
        Exit Function
    End If
    SendPipeLine = True
    Exit Function
SendPipeLine_Err:
    Call OnPipeBroken
End Function

' Lee lo que haya (nunca espera) y reenvia cada linea completa al server. False = el pipe se rompio.
Private Function ReadPipeLines() As Boolean
    On Error GoTo ReadPipeLines_Err
    Dim avail As Long
    Dim got As Long
    Dim chunk() As Byte
    Dim ok As Long
    Dim passes As Long
    If m_Pipe = 0 Then Exit Function
    ReadPipeLines = True
    Do While passes < 8
        passes = passes + 1
        avail = 0
        ok = GP_PeekNamedPipe(m_Pipe, 0&, 0&, 0&, avail, 0&)
        If ok = 0 Then
            Call OnPipeBroken
            ReadPipeLines = False
            Exit Function
        End If
        If avail <= 0 Then Exit Do
        If avail > GP_READ_CHUNK Then avail = GP_READ_CHUNK
        ReDim chunk(0 To avail - 1) As Byte
        got = 0
        ok = GP_ReadFile(m_Pipe, chunk(0), avail, got, 0&)
        If ok = 0 Or got <= 0 Then
            Call OnPipeBroken
            ReadPipeLines = False
            Exit Function
        End If
        If got < avail Then ReDim Preserve chunk(0 To got - 1) As Byte
        Call ConsumeBytes(StrConv(chunk, vbUnicode))
    Loop
    Exit Function
ReadPipeLines_Err:
    Call OnPipeBroken
    ReadPipeLines = False
End Function

Private Sub ConsumeBytes(ByVal incoming As String)
    Dim pos As Long
    Dim lineText As String
    m_RxBuffer = m_RxBuffer & incoming
    Do
        pos = InStr(1, m_RxBuffer, vbLf, vbBinaryCompare)
        If pos = 0 Then Exit Do
        lineText = Left$(m_RxBuffer, pos - 1)
        m_RxBuffer = Mid$(m_RxBuffer, pos + 1)
        If m_RxOverflow Then
            m_RxOverflow = False
        Else
            If Right$(lineText, 1) = vbCr Then lineText = Left$(lineText, Len(lineText) - 1)
            If Len(lineText) > 0 And Len(lineText) <= GP_MAX_LINE Then
                m_GotLineSinceHello = True
                If m_Pending Then Call WriteGuardianHeartbeat(lineText)
            End If
        End If
    Loop
    ' Sin salto de linea y ya pasado el tope: es basura, se descarta hasta el proximo salto.
    If Len(m_RxBuffer) > GP_MAX_LINE Then
        m_RxBuffer = vbNullString
        m_RxOverflow = True
    End If
End Sub
