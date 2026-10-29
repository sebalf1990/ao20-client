Attribute VB_Name = "ModFullscreenLayout"
' Argentum 20 Game Client
'
'    Layout declarativo para el modo viewport fullscreen.
'    Reemplaza el ocultar-todo del SetHudControlsVisible legacy por una tabla
'    que define que controles se ocultan vs se mueven a su zona FULL.
'
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md
'
Option Explicit

Public Enum e_FullAction
    eFlaHide = 0
    eFlaMove = 1
End Enum

Private Type t_FullscreenSlot
    ctlName     As String
    action      As e_FullAction
    fullLeft    As Long
    fullTop     As Long
    origLeft    As Long
    origTop     As Long
    origVisible As Boolean
End Type

Private g_slots()       As t_FullscreenSlot
Private g_slot_count    As Integer
Private g_layout_inited As Boolean

' frmMain.ScaleMode = 3 (Pixel). Los .Left/.Top runtime estan en pixeles directo,
' aunque el .frm los guarde como twips en disco. No multiplicamos por TwipsPerPixel.
Private Sub AddSlot(ByVal ctlName As String, ByVal action As e_FullAction, Optional ByVal fullLeftPx As Long = 0, Optional ByVal fullTopPx As Long = 0)
    ReDim Preserve g_slots(g_slot_count) As t_FullscreenSlot
    g_slots(g_slot_count).ctlName = ctlName
    g_slots(g_slot_count).action = action
    g_slots(g_slot_count).fullLeft = fullLeftPx
    g_slots(g_slot_count).fullTop = fullTopPx
    g_slot_count = g_slot_count + 1
End Sub

Public Sub Init_FullscreenLayout()
    On Error GoTo Init_FullscreenLayout_Err
    If g_layout_inited Then Exit Sub
    g_slot_count = 0

    ' --- Etapa 1: solo PictureBoxes pueden moverse sobre el renderer ---
    ' VB6 Labels/Images son lightweight y NO se renderizan encima de un PictureBox
    ' grande (el renderer). NombrePJ/Coord/lblhora se reemplazaran en Etapa 2 via DX.
    Call AddSlot("MiniMap", eFlaMove, 884, 70)
    Call AddSlot("NombrePJ", eFlaHide)
    Call AddSlot("Coord", eFlaHide)
    Call AddSlot("lblhora", eFlaHide)

    ' --- Etapa 1: contenedores y botones del HUD que se ocultan ---
    ' (mismo comportamiento que el SetHudControlsVisible legacy, ahora declarativo)
    Call AddSlot("panel", eFlaHide)
    Call AddSlot("panelInf", eFlaHide)
    Call AddSlot("imgMAO", eFlaHide)
    Call AddSlot("imgManual", eFlaHide)
    Call AddSlot("imgHechizos", eFlaHide)
    Call AddSlot("imgInventario", eFlaHide)
    Call AddSlot("btnZoomIn", eFlaHide)
    Call AddSlot("btnZoomOut", eFlaHide)
    Call AddSlot("btnExp", eFlaHide)
    Call AddSlot("btnTotalExp", eFlaHide)
    Call AddSlot("btnExpRemaining", eFlaHide)

    ' Etapas siguientes agregaran:
    '  Etapa 2: party, displays vitales (HP/Mana/Sta), stats (Fuerza/Agilidad), mirror consola
    '  Etapa 3: botones de seguros + iconos de menus a bottom-right
    '  Etapa 4: panel DX de buffs/CDs

    g_layout_inited = True
    Exit Sub
Init_FullscreenLayout_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenLayout.Init_FullscreenLayout", Erl)
    Resume Next
End Sub

Public Sub Apply_FullscreenLayout(ByVal enterFull As Boolean)
    On Error GoTo Apply_FullscreenLayout_Err
    If Not g_layout_inited Then Call Init_FullscreenLayout
    Call ViewportDebug_AppendDiagLog("[Layout] Apply enterFull=" & enterFull & " slots=" & g_slot_count)

    Dim i As Integer
    For i = 0 To g_slot_count - 1
        Call ApplyOneSlot(i, enterFull)
    Next i
    Call ViewportDebug_AppendDiagLog("[Layout] Apply done")
    Exit Sub
Apply_FullscreenLayout_Err:
    Call ViewportDebug_AppendDiagLog("[Layout] FATAL err=" & Err.Number & " " & Err.Description)
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenLayout.Apply_FullscreenLayout", Erl)
    Resume Next
End Sub

Private Sub ApplyOneSlot(ByVal idx As Integer, ByVal enterFull As Boolean)
    On Error GoTo ApplyOneSlot_Err
    Dim ctl As Object
    Set ctl = SafeGetControl(g_slots(idx).ctlName)
    If ctl Is Nothing Then
        Call ViewportDebug_AppendDiagLog("[Layout] NOTFOUND name=" & g_slots(idx).ctlName)
        Exit Sub
    End If

    If enterFull Then
        g_slots(idx).origLeft = ctl.Left
        g_slots(idx).origTop = ctl.Top
        g_slots(idx).origVisible = SafeGetVisible(ctl)

        Select Case g_slots(idx).action
            Case eFlaHide
                Call SafeSetVisible(ctl, False)
                Call ViewportDebug_AppendDiagLog("[Layout] HIDE " & g_slots(idx).ctlName & " wasVisible=" & g_slots(idx).origVisible)
            Case eFlaMove
                ctl.Left = g_slots(idx).fullLeft
                ctl.Top = g_slots(idx).fullTop
                Call SafeSetVisible(ctl, True)
                ctl.ZOrder 0
                Call ViewportDebug_AppendDiagLog("[Layout] MOVE " & g_slots(idx).ctlName & " to L=" & ctl.Left & " T=" & ctl.Top & " V=" & ctl.Visible & " (orig L=" & g_slots(idx).origLeft & " T=" & g_slots(idx).origTop & " V=" & g_slots(idx).origVisible & ")")
        End Select
    Else
        ctl.Left = g_slots(idx).origLeft
        ctl.Top = g_slots(idx).origTop
        Call SafeSetVisible(ctl, g_slots(idx).origVisible)
    End If
    Exit Sub
ApplyOneSlot_Err:
    Call ViewportDebug_AppendDiagLog("[Layout] ERR slot=" & g_slots(idx).ctlName & " err=" & Err.Number & " " & Err.Description)
    Resume Next
End Sub

Private Function SafeGetControl(ByVal ctlName As String) As Object
    On Error Resume Next
    Set SafeGetControl = frmMain.Controls(ctlName)
End Function

Private Function SafeGetVisible(ByVal ctl As Object) As Boolean
    On Error Resume Next
    SafeGetVisible = ctl.Visible
End Function

Private Sub SafeSetVisible(ByVal ctl As Object, ByVal v As Boolean)
    On Error Resume Next
    ctl.Visible = v
End Sub
