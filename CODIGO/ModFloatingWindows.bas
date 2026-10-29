Attribute VB_Name = "ModFloatingWindows"
' Argentum 20 Game Client
'
'    Infraestructura spike de ventanas flotantes FULL.
'    Etapa 5: dummy de Inventario/Hechizos con abrir, cerrar, drag, pin,
'    z-order y persistencia basica. No renderiza items ni hechizos reales todavia.
'
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md (Etapa 5)
'
Option Explicit

Public Enum e_FloatingWindowId
    efwInventory = 0
    efwSpells = 1
End Enum

Private Const FLOAT_W        As Long = 250
Private Const FLOAT_H        As Long = 310
Private Const HEADER_H       As Long = 24
Private Const BUTTON_W       As Long = 20
Private Const INV_OFFSET_X   As Long = 6
Private Const INV_OFFSET_Y   As Long = 36
Private Const INV_SPACE      As Long = 3
Private Const LAYOUT_FILE    As String = "fullscreen_windows.ini"

' Layout de la ventana flotante de Hechizos (ref: picFullSpellList anidado en picFullSpellWindow)
Private Const SPELL_LIST_TOP As Long = 28
Private Const SPELL_LIST_H   As Long = 200
Private Const SPELL_FOOTER_Y As Long = 240
Private Const SPELL_BTN_H    As Long = 22
Private Const SPELL_LANZAR_X As Long = 10
Private Const SPELL_LANZAR_W As Long = 80
Private Const SPELL_BTN_W    As Long = 28
Private Const SPELL_UP_X     As Long = 100
Private Const SPELL_DOWN_X   As Long = 132
Private Const SPELL_INFO_X   As Long = 170

Private g_floating_active As Boolean
Private g_dragging        As Boolean
Private g_drag_window     As e_FloatingWindowId
Private g_drag_x          As Long
Private g_drag_y          As Long
Private g_pinned(0 To 1)  As Boolean
Private g_loaded          As Boolean

Public Sub Set_FloatingWindowsActive(ByVal active As Boolean)
    On Error GoTo Set_FloatingWindowsActive_Err
    g_floating_active = active
    If active Then
        Call LoadFloatingWindowState
        Call ApplyWindowState(efwInventory)
        Call ApplyWindowState(efwSpells)
    Else
        Call SaveFloatingWindowState(efwInventory)
        Call SaveFloatingWindowState(efwSpells)
        Call BindMainInventoryToClassic
        Call BindMainSpellsToClassic
        frmMain.picFullSpellList.Visible = False
        frmMain.picFullInvWindow.Visible = False
        frmMain.picFullSpellWindow.Visible = False
        g_dragging = False
    End If
    Exit Sub
Set_FloatingWindowsActive_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.Set_FloatingWindowsActive", Erl)
    Resume Next
End Sub

Public Sub Toggle_FloatingWindow(ByVal windowId As e_FloatingWindowId)
    On Error GoTo Toggle_FloatingWindow_Err
    If Not g_floating_active Then Exit Sub
    Dim ctl As PictureBox
    Set ctl = WindowCtl(windowId)
    ctl.Visible = Not ctl.Visible
    If windowId = efwInventory Then
        If ctl.Visible Then
            Call BindMainInventoryToFloating
        Else
            Call BindMainInventoryToClassic
        End If
    ElseIf windowId = efwSpells Then
        If ctl.Visible Then
            frmMain.picFullSpellList.Visible = True
            Call BindMainSpellsToFloating
        Else
            frmMain.picFullSpellList.Visible = False
            Call BindMainSpellsToClassic
        End If
    End If
    If ctl.Visible Then
        ctl.ZOrder 0
        Call RenderFloatingWindow(windowId)
    End If
    Call SaveFloatingWindowState(windowId)
    Exit Sub
Toggle_FloatingWindow_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.Toggle_FloatingWindow", Erl)
    Resume Next
End Sub

Public Sub FloatingWindow_MouseDown(ByVal windowId As e_FloatingWindowId, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
    On Error GoTo FloatingWindow_MouseDown_Err
    If Button <> vbLeftButton Then Exit Sub
    Dim ctl As PictureBox
    Set ctl = WindowCtl(windowId)
    ctl.ZOrder 0

    If y <= HEADER_H Then
        If x >= ctl.ScaleWidth - BUTTON_W Then
            ctl.Visible = False
            If windowId = efwInventory Then
                Call BindMainInventoryToClassic
            ElseIf windowId = efwSpells Then
                frmMain.picFullSpellList.Visible = False
                Call BindMainSpellsToClassic
            End If
            Call SaveFloatingWindowState(windowId)
            Exit Sub
        End If
        If x >= ctl.ScaleWidth - BUTTON_W * 2 Then
            g_pinned(windowId) = Not g_pinned(windowId)
            Call RenderFloatingWindow(windowId)
            Call SaveFloatingWindowState(windowId)
            Exit Sub
        End If
        If Not g_pinned(windowId) Then
            g_dragging = True
            g_drag_window = windowId
            g_drag_x = CLng(x)
            g_drag_y = CLng(y)
        End If
        Exit Sub
    End If

    ' Hit-test del footer en la ventana de Hechizos
    If windowId = efwSpells Then
        If y >= SPELL_FOOTER_Y And y < SPELL_FOOTER_Y + SPELL_BTN_H Then
            If x >= SPELL_LANZAR_X And x < SPELL_LANZAR_X + SPELL_LANZAR_W Then
                Call frmMain.Public_CastSelectedSpell
            ElseIf x >= SPELL_UP_X And x < SPELL_UP_X + SPELL_BTN_W Then
                Call frmMain.Public_MoveSpell(1)
            ElseIf x >= SPELL_DOWN_X And x < SPELL_DOWN_X + SPELL_BTN_W Then
                Call frmMain.Public_MoveSpell(0)
            ElseIf x >= SPELL_INFO_X And x < SPELL_INFO_X + SPELL_BTN_W Then
                Call frmMain.Public_RequestSpellInfo
            End If
        End If
    End If
    Exit Sub
FloatingWindow_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.FloatingWindow_MouseDown", Erl)
    Resume Next
End Sub

Public Sub FloatingWindow_MouseMove(ByVal windowId As e_FloatingWindowId, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
    On Error GoTo FloatingWindow_MouseMove_Err
    If Not g_dragging Then Exit Sub
    If windowId <> g_drag_window Then Exit Sub
    If Button <> vbLeftButton Then Exit Sub

    Dim ctl As PictureBox
    Set ctl = WindowCtl(windowId)
    ctl.Left = ctl.Left + CLng(x) - g_drag_x
    ctl.Top = ctl.Top + CLng(y) - g_drag_y
    Call ClampWindowToViewport(ctl)
    Exit Sub
FloatingWindow_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.FloatingWindow_MouseMove", Erl)
    Resume Next
End Sub

Public Sub FloatingWindow_MouseUp(ByVal windowId As e_FloatingWindowId, ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
    On Error GoTo FloatingWindow_MouseUp_Err
    If g_dragging And windowId = g_drag_window Then
        g_dragging = False
        Call SaveFloatingWindowState(windowId)
    End If
    Exit Sub
FloatingWindow_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.FloatingWindow_MouseUp", Erl)
    Resume Next
End Sub

Private Sub ApplyWindowState(ByVal windowId As e_FloatingWindowId)
    On Error GoTo ApplyWindowState_Err
    Dim ctl As PictureBox
    Set ctl = WindowCtl(windowId)
    ctl.Width = FLOAT_W
    ctl.Height = FLOAT_H
    Call ClampWindowToViewport(ctl)
    ctl.Visible = (val(GetVarOrDefault(LayoutPath(), SectionName(windowId), "Visible", "0")) <> 0)
    If windowId = efwInventory Then
        If ctl.Visible Then
            Call BindMainInventoryToFloating
        Else
            Call BindMainInventoryToClassic
        End If
    ElseIf windowId = efwSpells Then
        If ctl.Visible Then
            frmMain.picFullSpellList.Visible = True
            Call BindMainSpellsToFloating
        Else
            frmMain.picFullSpellList.Visible = False
            Call BindMainSpellsToClassic
        End If
    End If
    If ctl.Visible Then ctl.ZOrder 0
    If ctl.Visible Then Call ModMenuTransparency.MenuTransparency_ApplyToForm(ctl.hWnd)
    If windowId = efwSpells And frmMain.picFullSpellList.Visible Then
        Call ModMenuTransparency.MenuTransparency_ApplyToForm(frmMain.picFullSpellList.hWnd)
    End If
    Call RenderFloatingWindow(windowId)
    Exit Sub
ApplyWindowState_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.ApplyWindowState", Erl)
    Resume Next
End Sub

Private Sub LoadFloatingWindowState()
    On Error GoTo LoadFloatingWindowState_Err
    If g_loaded Then Exit Sub
    Call LoadOneWindow(efwInventory, 270, 180)
    Call LoadOneWindow(efwSpells, 540, 180)
    g_loaded = True
    Exit Sub
LoadFloatingWindowState_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.LoadFloatingWindowState", Erl)
    Resume Next
End Sub

Private Sub LoadOneWindow(ByVal windowId As e_FloatingWindowId, ByVal defaultLeft As Long, ByVal defaultTop As Long)
    Dim ctl As PictureBox
    Set ctl = WindowCtl(windowId)
    ctl.Width = FLOAT_W
    ctl.Height = FLOAT_H
    ctl.Left = CLng(val(GetVarOrDefault(LayoutPath(), SectionName(windowId), "Left", CStr(defaultLeft))))
    ctl.Top = CLng(val(GetVarOrDefault(LayoutPath(), SectionName(windowId), "Top", CStr(defaultTop))))
    g_pinned(windowId) = (val(GetVarOrDefault(LayoutPath(), SectionName(windowId), "Pinned", "0")) <> 0)
    Call ClampWindowToViewport(ctl)
End Sub

Private Sub SaveFloatingWindowState(ByVal windowId As e_FloatingWindowId)
    On Error GoTo SaveFloatingWindowState_Err
    Dim ctl As PictureBox
    Set ctl = WindowCtl(windowId)
    Call WriteVar(LayoutPath(), SectionName(windowId), "Left", CStr(ctl.Left))
    Call WriteVar(LayoutPath(), SectionName(windowId), "Top", CStr(ctl.Top))
    Call WriteVar(LayoutPath(), SectionName(windowId), "Pinned", IIf(g_pinned(windowId), "1", "0"))
    Call WriteVar(LayoutPath(), SectionName(windowId), "Visible", IIf(ctl.Visible, "1", "0"))
    Exit Sub
SaveFloatingWindowState_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.SaveFloatingWindowState", Erl)
    Resume Next
End Sub

Private Sub RenderFloatingWindow(ByVal windowId As e_FloatingWindowId)
    On Error GoTo RenderFloatingWindow_Err
    If windowId = efwInventory Then Exit Sub
    Dim ctl As PictureBox
    Set ctl = WindowCtl(windowId)

    ctl.BackColor = RGB(14, 14, 14)
    ctl.ForeColor = RGB(235, 235, 235)
    If windowId <> efwInventory Then
        ctl.Cls
        ctl.Line (0, HEADER_H)-(ctl.ScaleWidth, ctl.ScaleHeight), RGB(9, 9, 9), BF
    End If
    ctl.Line (0, 0)-(ctl.ScaleWidth, HEADER_H), RGB(32, 32, 32), BF
    ctl.Line (0, 0)-(ctl.ScaleWidth - 1, ctl.ScaleHeight - 1), RGB(88, 88, 88), B

    ctl.CurrentX = 8
    ctl.CurrentY = 5
    ctl.Print WindowTitle(windowId)

    ctl.Line (ctl.ScaleWidth - BUTTON_W * 2, 0)-(ctl.ScaleWidth - BUTTON_W - 1, HEADER_H), RGB(55, 55, 55), BF
    ctl.Line (ctl.ScaleWidth - BUTTON_W, 0)-(ctl.ScaleWidth - 1, HEADER_H), RGB(75, 30, 30), BF
    ctl.CurrentX = ctl.ScaleWidth - BUTTON_W * 2 + 6
    ctl.CurrentY = 5
    ctl.Print IIf(g_pinned(windowId), "P", "p")
    ctl.CurrentX = ctl.ScaleWidth - BUTTON_W + 6
    ctl.CurrentY = 5
    ctl.Print "x"

    If windowId = efwInventory Then
        ctl.Refresh
        Exit Sub
    End If

    ' Footer: botones Lanzar / ARRIBA / ABAJO / ?
    Call DrawSpellFooterButton(ctl, SPELL_LANZAR_X, SPELL_FOOTER_Y, SPELL_LANZAR_W, SPELL_BTN_H, "Lanzar")
    Call DrawSpellFooterButton(ctl, SPELL_UP_X, SPELL_FOOTER_Y, SPELL_BTN_W, SPELL_BTN_H, "^")
    Call DrawSpellFooterButton(ctl, SPELL_DOWN_X, SPELL_FOOTER_Y, SPELL_BTN_W, SPELL_BTN_H, "v")
    Call DrawSpellFooterButton(ctl, SPELL_INFO_X, SPELL_FOOTER_Y, SPELL_BTN_W, SPELL_BTN_H, "?")
    ctl.Refresh
    Exit Sub
RenderFloatingWindow_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.RenderFloatingWindow", Erl)
    Resume Next
End Sub

Public Function Is_FloatingInventoryVisible() As Boolean
    On Error Resume Next
    Is_FloatingInventoryVisible = g_floating_active And frmMain.picFullInvWindow.Visible
End Function

Public Function FloatingInventoryHwnd() As Long
    FloatingInventoryHwnd = frmMain.picFullInvWindow.hWnd
End Function

Public Function FloatingInventoryWidth() As Long
    FloatingInventoryWidth = frmMain.picFullInvWindow.ScaleWidth
End Function

Public Function FloatingInventoryHeight() As Long
    FloatingInventoryHeight = frmMain.picFullInvWindow.ScaleHeight
End Function

Public Sub Render_FloatingInventoryChrome()
    On Error GoTo Render_FloatingInventoryChrome_Err
    If Not Is_FloatingInventoryVisible() Then Exit Sub

    Dim W As Long
    Dim H As Long
    W = frmMain.picFullInvWindow.ScaleWidth
    H = frmMain.picFullInvWindow.ScaleHeight

    Dim chromeA As Byte
    chromeA = MenuTransparency_ScaleAlpha(255)

    Dim headerBg As RGBA
    headerBg = RGBA_From_Comp(32, 32, 32, chromeA)
    Call Engine_Draw_Box(0, 0, W, HEADER_H, headerBg)

    Dim borderColor As RGBA
    borderColor = RGBA_From_Comp(88, 88, 88, chromeA)
    Call Engine_Draw_Box(0, 0, W, 1, borderColor)
    Call Engine_Draw_Box(0, H - 1, W, 1, borderColor)
    Call Engine_Draw_Box(0, 0, 1, H, borderColor)
    Call Engine_Draw_Box(W - 1, 0, 1, H, borderColor)

    Dim pinBg As RGBA
    pinBg = RGBA_From_Comp(55, 55, 55, chromeA)
    Call Engine_Draw_Box(W - BUTTON_W * 2, 0, BUTTON_W, HEADER_H, pinBg)

    Dim closeBg As RGBA
    closeBg = RGBA_From_Comp(75, 30, 30, chromeA)
    Call Engine_Draw_Box(W - BUTTON_W, 0, BUTTON_W, HEADER_H, closeBg)

    Dim textColor(3) As RGBA
    Call RGBAList(textColor, 235, 235, 235, 255)
    Call Engine_Text_Render("Inventario", 8, 4, textColor, 1, False)

    Dim pinTxt As String
    If g_pinned(efwInventory) Then
        pinTxt = "P"
    Else
        pinTxt = "p"
    End If
    Call Engine_Text_Render(pinTxt, W - BUTTON_W * 2 + 6, 4, textColor, 1, False)
    Call Engine_Text_Render("x", W - BUTTON_W + 6, 4, textColor, 1, False)
    Exit Sub
Render_FloatingInventoryChrome_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.Render_FloatingInventoryChrome", Erl)
    Resume Next
End Sub

Private Sub BindMainInventoryToFloating()
    On Error GoTo BindMainInventoryToFloating_Err
    Call frmMain.Inventario.BindWindow(frmMain.picFullInvWindow, frmMain.picFullInvWindow.ScaleWidth - INV_OFFSET_X * 2, frmMain.picFullInvWindow.ScaleHeight - INV_OFFSET_Y - 8, INV_OFFSET_X, INV_OFFSET_Y, INV_SPACE, INV_SPACE, 9)
    Exit Sub
BindMainInventoryToFloating_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.BindMainInventoryToFloating", Erl)
    Resume Next
End Sub

Private Sub BindMainInventoryToClassic()
    On Error GoTo BindMainInventoryToClassic_Err
    Call frmMain.Inventario.BindWindow(frmMain.picInv, frmMain.picInv.ScaleWidth, frmMain.picInv.ScaleHeight, 0, 0, 3, 3, 9)
    Exit Sub
BindMainInventoryToClassic_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.BindMainInventoryToClassic", Erl)
    Resume Next
End Sub

Private Sub BindMainSpellsToFloating()
    On Error GoTo BindMainSpellsToFloating_Err
    If hlst Is Nothing Then Exit Sub
    Call hlst.BindPicture(frmMain.picFullSpellList)
    Exit Sub
BindMainSpellsToFloating_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.BindMainSpellsToFloating", Erl)
    Resume Next
End Sub

Private Sub BindMainSpellsToClassic()
    On Error GoTo BindMainSpellsToClassic_Err
    If hlst Is Nothing Then Exit Sub
    Call hlst.BindPicture(frmMain.picHechiz)
    Exit Sub
BindMainSpellsToClassic_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFloatingWindows.BindMainSpellsToClassic", Erl)
    Resume Next
End Sub

Public Function Is_FloatingSpellsVisible() As Boolean
    On Error Resume Next
    Is_FloatingSpellsVisible = g_floating_active And frmMain.picFullSpellWindow.Visible
End Function

Public Sub Refresh_FloatingWindowsTransparency()
    On Error Resume Next
    If Not g_floating_active Then Exit Sub
    If frmMain.picFullInvWindow.Visible Then
        Call ModMenuTransparency.MenuTransparency_ApplyToForm(frmMain.picFullInvWindow.hWnd)
        Call Render_FloatingInventoryChrome
    End If
    If frmMain.picFullSpellWindow.Visible Then
        Call ModMenuTransparency.MenuTransparency_ApplyToForm(frmMain.picFullSpellWindow.hWnd)
        If frmMain.picFullSpellList.Visible Then
            Call ModMenuTransparency.MenuTransparency_ApplyToForm(frmMain.picFullSpellList.hWnd)
        End If
        Call RenderFloatingWindow(efwSpells)
    End If
End Sub

Private Sub DrawSpellFooterButton(ByRef ctl As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal label As String)
    On Error Resume Next
    ctl.Line (x, y)-(x + w - 1, y + h - 1), RGB(55, 55, 55), BF
    ctl.Line (x, y)-(x + w - 1, y + h - 1), RGB(100, 100, 100), B
    ctl.CurrentX = x + 6
    ctl.CurrentY = y + 4
    ctl.Print label
End Sub

Private Sub ClampWindowToViewport(ByRef ctl As PictureBox)
    If ctl.Left < g_viewport_logical_left Then ctl.Left = g_viewport_logical_left
    If ctl.Top < g_viewport_logical_top Then ctl.Top = g_viewport_logical_top
    If ctl.Left + ctl.Width > g_viewport_logical_left + g_viewport_logical_width Then ctl.Left = g_viewport_logical_left + g_viewport_logical_width - ctl.Width
    If ctl.Top + ctl.Height > g_viewport_logical_top + g_viewport_logical_height Then ctl.Top = g_viewport_logical_top + g_viewport_logical_height - ctl.Height
End Sub

Private Function WindowCtl(ByVal windowId As e_FloatingWindowId) As PictureBox
    Select Case windowId
        Case efwInventory
            Set WindowCtl = frmMain.picFullInvWindow
        Case efwSpells
            Set WindowCtl = frmMain.picFullSpellWindow
    End Select
End Function

Private Function SectionName(ByVal windowId As e_FloatingWindowId) As String
    Select Case windowId
        Case efwInventory
            SectionName = "Inventory"
        Case efwSpells
            SectionName = "Spells"
    End Select
End Function

Private Function WindowTitle(ByVal windowId As e_FloatingWindowId) As String
    Select Case windowId
        Case efwInventory
            WindowTitle = "Inventario"
        Case efwSpells
            WindowTitle = "Hechizos"
    End Select
End Function

Private Function LayoutPath() As String
    LayoutPath = App.path & "\" & LAYOUT_FILE
End Function
