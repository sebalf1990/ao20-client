Attribute VB_Name = "ModGameplayUI"
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'

' ============== Viewport fullscreen toggle (PoC 2026-05-20) ==============
' Declaraciones a nivel de modulo (subs estan al final del archivo).
Private g_renderer_state_saved As Boolean
Private g_pending_restore_full As Boolean
Private Const SENDTXT_HUD_LEFT   As Long = 40
Private Const SENDTXT_HUD_TOP    As Long = 120
Private Const SENDTXT_HUD_WIDTH  As Long = 546
Private Const SENDTXT_HUD_HEIGHT As Long = 24
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Global G_LastSelectedSlot As Integer
Global G_LasSelectedObjIndex As Integer

Public Sub SetupGameplayUI()
    frmMain.shapexy.Left = 1200
    frmMain.shapexy.Top = 1200
    frmMain.shapexy.BackColor = RGB(170, 0, 0)
    frmMain.NombrePJ.Caption = userName
    ' Detect links in console
    Call EnableURLDetect(frmMain.RecTxt.hWnd, frmMain.hWnd)
    Call Make_Transparent_Richtext(frmMain.RecTxt.hWnd)
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(frmMain)
    frmMain.panel.Picture = LoadInterface("centroinventario.bmp")
    frmMain.picInv.visible = True
    frmMain.picHechiz.visible = False
    frmMain.cmdlanzar.visible = False
    frmMain.imgSpellInfo.visible = False
    frmMain.cmdMoverHechi(0).visible = False
    frmMain.cmdMoverHechi(1).visible = False
    Call frmMain.Inventario.ReDraw
    frmMain.Left = 0
    frmMain.Top = 0
    frmMain.Width = D3DWindow.BackBufferWidth * screen.TwipsPerPixelX
    frmMain.Height = D3DWindow.BackBufferHeight * screen.TwipsPerPixelY
    Call SetViewportLogicalHud
    Call ApplyRendererHudClip
    frmMain.renderer.ZOrder 1
    frmMain.visible = True
    ActiveInventoryTab = eInventory
    Call LoadHotkeys
End Sub

#If DEBUGGING = 1 Then
Public Sub CaptureDebugClickInfo(ByVal tx As Integer, ByVal ty As Integer)
    On Error GoTo CaptureDebugClickInfo_Err
    Dim info As String
    Dim charIdx As Integer
    If tx < XMinMapSize Or tx > XMaxMapSize Then Exit Sub
    If ty < YMinMapSize Or ty > YMaxMapSize Then Exit Sub
    charIdx = MapData(tx, ty).charindex
    If charIdx > 0 And charIdx <= UBound(charlist) Then
        With charlist(charIdx)
            If .EsNpc Then
                info = .nombre & Chr(13) & _
                       "Nro: " & .NpcNumber & Chr(13) & _
                       "Body: " & .iBody & "  Idle: " & .BodyIdle & Chr(13) & _
                       "HP: " & .UserMinHp & "/" & .UserMaxHp
            End If
        End With
    End If
    If LenB(info) = 0 Then
        With MapData(tx, ty).OBJInfo
            If .ObjIndex > 0 Then
                info = ObjData(.ObjIndex).Name & Chr(13) & _
                       "Nro: " & .ObjIndex & Chr(13) & _
                       "GrhIndex: " & MapData(tx, ty).ObjGrh.GrhIndex
            End If
        End With
    End If
    g_debug_click_info = info
    Exit Sub
CaptureDebugClickInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUI.CaptureDebugClickInfo", Erl)
End Sub
#End If

Public Sub OnClick(ByVal MouseButton As Long, ByVal MouseShift As Long)
    On Error GoTo OnClick_Err
    If pausa Then Exit Sub
    If IsGameDialogOpen Then Exit Sub
    If mascota.visible Then
        If Sqr((mouseX - mascota.PosX) ^ 2 + (mouseY - mascota.PosY) ^ 2) < 30 Then
            mascota.dialog = ""
        End If
    End If
    If cartel_visible Then
        If mouseX > 50 And mouseY > 478 And mouseX < 671 And mouseY < 585 Then
            If tutorial_index > 0 Then
                Call nextCartel
            Else
                Call cerrarCartel
            End If
        End If
    End If
    Dim MouseAction As e_MouseAction
    Select Case MouseButton
        Case vbLeftButton
            MouseAction = ACCION1
        Case vbRightButton
            MouseAction = ACCION2
        Case vbMiddleButton
            MouseAction = ACCION3
        Case Else
            Exit Sub
    End Select
    ' Hook PoC viewport: registrar click para marker visual del overlay debug.
    If tX > 0 And tY > 0 Then Call ViewportDebug_OnClick(tX, tY)
#If DEBUGGING = 1 Then
    If EsGM And IsSet(FeatureToggles, eShowGmDebugData) Then
        Call CaptureDebugClickInfo(tX, tY)
    End If
#End If
    If MouseAction = e_MouseAction.eThrowOrLook Then
        If Not Comerciando Then
            If MouseShift = 0 Then
                If UsingSkill = 0 Or frmMain.MacroLadder.enabled Then
                    Call WriteLeftClick(tX, tY)
                Else
                    Dim SendSkill As Boolean
                    If UsingSkill = magia Then
                        If ModoHechizos = BloqueoLanzar Then
                            SendSkill = IIf((mouseX >= g_viewport_logical_left And mouseX <= g_viewport_logical_left + g_viewport_logical_width And mouseY >= g_viewport_logical_top And _
                                    mouseY <= g_viewport_logical_top + g_viewport_logical_height), True, False)
                            If Not SendSkill Then
                                Exit Sub
                            End If
                            Call MainTimer.Restart(TimersIndex.CastAttack)
                            Call MainTimer.Restart(TimersIndex.CastSpell)
                        Else
                            If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                                If MainTimer.Check(TimersIndex.CastSpell) Then
                                    SendSkill = IIf((mouseX >= g_viewport_logical_left And mouseX <= g_viewport_logical_left + g_viewport_logical_width And mouseY >= g_viewport_logical_top _
                                            And mouseY <= g_viewport_logical_top + g_viewport_logical_height), True, False)
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                    Call MainTimer.Restart(TimersIndex.CastAttack)
                                ElseIf ModoHechizos = SinBloqueo Then
                                    SendSkill = IIf((mouseX >= g_viewport_logical_left And mouseX <= g_viewport_logical_left + g_viewport_logical_width And mouseY >= g_viewport_logical_top _
                                            And mouseY <= g_viewport_logical_top + g_viewport_logical_height), True, False)
                                    If Not SendSkill Then
                                        Exit Sub
                                    End If
                                    With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                        Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_LANZAMIENTO_RAPIDO"), .red, .green, .blue, .bold, .italic) ' MENSAJE_LANZAMIENTO_RAPIDO=No puedes lanzar hechizos tan rápido.
                                    End With
                                Else
                                    Exit Sub
                                End If
                            ElseIf ModoHechizos = SinBloqueo Then
                                SendSkill = IIf((mouseX >= g_viewport_logical_left And mouseX <= g_viewport_logical_left + g_viewport_logical_width And mouseY >= g_viewport_logical_top And _
                                        mouseY <= g_viewport_logical_top + g_viewport_logical_height), True, False)
                                If Not SendSkill Then
                                    Exit Sub
                                End If
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_ATAQUE_RAPIDO_GOLPE"), .red, .green, .blue, .bold, .italic) ' MENSAJE_ATAQUE_RAPIDO_GOLPE=No puedes lanzar tan rápido después de un golpe.
                                End With
                            Else
                                Exit Sub
                            End If
                        End If
                    End If
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If ModoHechizos = BloqueoLanzar Then
                            SendSkill = True
                            Call MainTimer.Restart(TimersIndex.Attack) ' Prevengo flecha-golpe
                            Call MainTimer.Restart(TimersIndex.CastSpell) ' flecha-hechizo
                            Call MainTimer.Restart(TimersIndex.Arrows)
                        ElseIf MainTimer.Check(TimersIndex.AttackSpell, False) Then
                            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                                If MainTimer.Check(TimersIndex.Arrows, False) Then
                                    SendSkill = True
                                    Call MainTimer.Restart(TimersIndex.Attack) ' Prevengo flecha-golpe
                                    Call MainTimer.Restart(TimersIndex.CastSpell) ' flecha-hechizo
                                    Call MainTimer.Restart(TimersIndex.Arrows)
                                ElseIf ModoHechizos = SinBloqueo Then
                                    SendSkill = True
                                    Call ShowGameplayModeMessage("MENSAJE_LANZAMIENTO_RAPIDO_FLECHAS") ' MENSAJE_LANZAMIENTO_RAPIDO_FLECHAS=No puedes lanzar flechas tan rápido.
                                Else
                                    Exit Sub
                                End If
                            ElseIf ModoHechizos = SinBloqueo Then
                                SendSkill = True
                            Else
                                Exit Sub
                            End If
                        ElseIf ModoHechizos = SinBloqueo Then
                            SendSkill = True
                        Else
                            Exit Sub
                        End If
                    End If
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Robar Or UsingSkill = Domar Or UsingSkill = Grupo Or UsingSkill = MarcaDeClan Or UsingSkill = MarcaDeGM) Then
                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            If UsingSkill = MarcaDeGM Then
                                Dim Pos As Integer
                                If MapData(tX, tY).charindex <> 0 Then
                                    Pos = InStr(charlist(MapData(tX, tY).charindex).nombre, "<")
                                    If Pos = 0 Then Pos = LenB(charlist(MapData(tX, tY).charindex).nombre) + 2
                                    frmPanelgm.cboListaUsus.text = Left$(charlist(MapData(tX, tY).charindex).nombre, Pos - 2)
                                End If
                            Else
                                SendSkill = True
                            End If
                        End If
                    End If
                    If (UsingSkill = eSkill.Alquimia Or UsingSkill = eSkill.TargetableItem) Then
                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            Call FormParser.Parse_Form(GetGameplayForm)
                            If CursoresGraficos = 0 Then
                                GetGameplayForm.MousePointer = vbDefault
                            End If
                        End If
                    Else
                        If UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = eSkill.Pescar Or UsingSkill = eSkill.Smelting Then
                            Call WriteStartAutomatedAction(tX, tY, UsingSkill)
                        End If
                    End If
                    If SendSkill Then
                        If UsingSkill = eSkill.magia Then
                            If ComprobarPosibleMacro(mouseX, mouseY) Then
                                Call WriteWorkLeftClick(tX + RandomNumber(-2, 2), tY + RandomNumber(-2, 2), UsingSkill)
                            Else
                                Call WriteWorkLeftClick(tX, tY, UsingSkill)
                            End If
                        Else
                            Call WriteWorkLeftClick(tX, tY, UsingSkill)
                        End If
                    End If
                    Call FormParser.Parse_Form(GetGameplayForm)
                    If CursoresGraficos = 0 Then
                        GetGameplayForm.MousePointer = vbDefault
                    End If
                    UsaLanzar = False
                    UsingSkill = 0
                End If
            Else
                Call WriteWarpChar("YO", UserMap, tX, tY)
            End If
            If cartel Then cartel = False
        End If
    ElseIf MouseAction = e_MouseAction.eInteract Then
        Call WriteDoubleClick(tX, tY)
    ElseIf MouseAction = e_MouseAction.eAttack Then
        If UserDescansar Then Exit Sub
        If MainTimer.Check(TimersIndex.CastAttack, False) Then
            If WriteAttack() Then
                Call MainTimer.Restart(TimersIndex.AttackSpell)
            End If
        End If
    ElseIf MouseAction = e_MouseAction.eWhisper Then
        Dim charindex As Integer
        charindex = MapData(tX, tY).charindex
        If charindex = 0 And tY < YMaxMapSize Then
            charindex = MapData(tX, tY + 1).charindex
        End If
        If charindex <> 0 Then
            If charlist(charindex).nombre <> charlist(UserCharIndex).nombre Then
                If charlist(charindex).EsNpc = False Then
                    frmMain.SendTxt.text = "\" & charlist(charindex).nombre & " "
                    If frmMain.SendTxtCmsg.visible = False Then
                        frmMain.SendTxt.visible = True
                        frmMain.SendTxt.SetFocus
                        frmMain.SendTxt.SelStart = Len(frmMain.SendTxt.text)
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
OnClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUi.OnClick", Erl)
    Resume Next
End Sub

Public Sub HandleQuestionResponse(ByVal Result As Boolean)
    If PreguntaLocal Then
        If Result Then
            Select Case PreguntaNUM
                Case 1 '¿Destruir item?
                    Call WriteDrop(DestItemSlot, DestItemCant)
                Case 2 ' Denunciar
                    Call WriteDenounce(targetName)
            End Select
        Else
            Select Case PreguntaNUM
                Case 1
                    DestItemSlot = 0
                    DestItemCant = 0
            End Select
        End If
    Else
        Call WriteResponderPregunta(Result)
    End If
    Pregunta = False
    PreguntaLocal = False
End Sub

Public Sub HandleGameplayAreaMouseUp(ByVal Button As Integer, _
                                     ByVal x As Integer, _
                                     ByVal y As Integer, _
                                     ByVal FormTop As Long, _
                                     ByVal FormLeft As Long, _
                                     ByVal FormHeight As Long, _
                                     ByRef GameplayArea As Rect)
    clicX = x
    clicY = y
    
    Dim MouseAction As e_MouseAction
    Select Case Button
        Case vbLeftButton:  MouseAction = ACCION1
        Case vbRightButton: MouseAction = ACCION2
        Case vbMiddleButton: MouseAction = ACCION3
        Case Else: Exit Sub
    End Select
    
    Select Case MouseAction
    
        Case e_MouseAction.eThrowOrLook
            If HandleMouseInput(x, y) Then
            ElseIf Pregunta Then
                If x >= 419 And x <= 433 And y >= 243 And y <= 260 Then
                    Call HandleQuestionResponse(False)
                    Exit Sub
                ElseIf x >= 443 And x <= 458 And y >= 243 And y <= 260 Then
                    Call HandleQuestionResponse(True)
                    Exit Sub
                End If
            End If
        Case e_MouseAction.eInteract
            Call ShowInteractionMenu(FormTop, FormLeft, FormHeight, x, y, GameplayArea)
    End Select
End Sub

Public Sub HandleChatMsg(ByVal InputText As String)
    Dim str2 As String
    Dim str1 As String
    If LenB(InputText) <> 0 Then
        If Left$(InputText, 1) = "/" Then
            If UCase$(Left$(InputText, 7)) = "/GRUPO " Then
                SendingType = 5
            ElseIf UCase$(Left$(InputText, 6)) = "/CMSG " Then
                SendingType = 4
            ElseIf UCase$(Left$(InputText, 6)) = "/GRMG " Then
                SendingType = 6
            ElseIf UCase$(Left$(InputText, 6)) = "/RMSG " Then
                SendingType = 8
            Else
                SendingType = 1
            End If
            If InputText <> "" Then Call ParseUserCommand(InputText)
            'Shout
        ElseIf Left$(InputText, 1) = "-" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("-" & Right$(InputText, Len(InputText) - 1))
            SendingType = 2
            'Global
        ElseIf Left$(InputText, 1) = ";" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("/CONSOLA " & Right$(InputText, Len(InputText) - 1))
            sndPrivateTo = ""
        ElseIf Left$(InputText, 1) = "/RMSG" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("/RMSG " & Right$(InputText, Len(InputText) - 1))
            SendingType = 8
            sndPrivateTo = ""
            'Faccion
        ElseIf Left$(InputText, 1) = "/FMSG" Then
            If Right$(InputText, Len(InputText) - 1) <> "" Then Call ParseUserCommand("/FMSG " & Right$(InputText, Len(InputText) - 1))
            SendingType = 9
            'Privado
        ElseIf Left$(InputText, 1) = "\" Then
            Dim mensaje As String
            str1 = Right$(InputText, Len(InputText) - 1)
            str2 = ReadField(1, str1, 32)
            mensaje = Right$(InputText, Len(str1) - Len(str2) - 1)
            sndPrivateTo = str2
            SendingType = 3
            If str1 <> "" Then Call WriteWhisper(sndPrivateTo, mensaje)
            'Say
        Else
            If InputText <> "" Then Call ParseUserCommand(InputText)
            SendingType = 1
            sndPrivateTo = ""
        End If
    Else
        SendingType = 1
        sndPrivateTo = ""
    End If
End Sub

Public Sub UseSelectInvItem()
End Sub

Public Sub SetInvItem(ByVal Slot As Byte, _
                      ByVal ObjIndex As Integer, _
                      ByVal Amount As Integer, _
                      ByVal Equipped As Byte, _
                      ByVal GrhIndex As Long, _
                      ByVal ObjType As Integer, _
                      ByVal MaxHit As Integer, _
                      ByVal MinHit As Integer, _
                      ByVal Def As Integer, _
                      ByVal value As Single, _
                      ByVal Name As String, _
                      ByVal CanUse As Byte, _
                      ByVal ElementalTags As Long, _
                      ByVal IsBindable As Byte)
    If Slot < 1 Or Slot > UBound(UserInventory.Slots) Then Exit Sub
    With UserInventory.Slots(Slot)
        .Amount = Amount
        .Def = Def
        .Equipped = Equipped
        .GrhIndex = GrhIndex
        .MaxHit = MaxHit
        .MinHit = MinHit
        .Name = Name
        .ObjIndex = ObjIndex
        .ObjType = ObjType
        .Valor = value
        .PuedeUsar = CanUse
        .ElementalTags = ElementalTags
        .IsBindable = IsBindable > 0
    End With
    Call frmMain.Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, Def, value, Name, ElementalTags, CanUse, IsBindable > 0)
End Sub

Public Sub SelectItemSlot(ByVal Slot As Integer)
    UserInventory.SelectedSlot = Slot
End Sub

Public Function GetSelectedItemSlot() As Integer
    GetSelectedItemSlot = frmMain.Inventario.SelectedItem
End Function

Public Function IsItemSelected() As Boolean
    IsItemSelected = frmMain.Inventario.IsItemSelected
End Function

Private Function IsProjectileTargetReady() As Boolean
    If Not MainTimer.Check(TimersIndex.AttackSpell, False) Then Exit Function
    If Not MainTimer.Check(TimersIndex.CastAttack, False) Then Exit Function
    If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Function
    IsProjectileTargetReady = True
End Function

Private Function IsEquippedProjectileSlot(ByVal Slot As Integer) As Boolean
    Dim objIndex As Integer
    If Slot <= 0 Then Exit Function
    If Not frmMain.Inventario.Equipped(Slot) Then Exit Function
    objIndex = frmMain.Inventario.ObjIndex(Slot)
    If objIndex <= 0 Then Exit Function
    If ObjData(objIndex).ObjType <> eObjType.otWeapon Then Exit Function
    IsEquippedProjectileSlot = (ObjData(objIndex).proyectil = 1)
End Function

Private Sub RequestProjectileTarget(ByVal Slot As Byte, ByVal UseItemWithU As Boolean)
    If UsingSkill = Proyectiles Then Exit Sub
    If ModoHechizos = BloqueoLanzar Then
        If Not IsProjectileTargetReady() Then Exit Sub
    End If
    If UseItemWithU Then
        Call WriteUseItemU(Slot, True)
    Else
        Call WriteUseItem(Slot, True)
    End If
End Sub

Private Sub ShowGameplayModeMessage(ByVal MessageKey As String)
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        Call ShowConsoleMsg(JsonLanguage.Item(MessageKey), .red, .green, .blue, .bold, .italic)
    End With
End Sub

Public Sub UseItemKey()
    Dim Slot As Byte
    If frmMain.Inventario.IsItemSelected Then
        Slot = frmMain.Inventario.SelectedItem
        If IsEquippedProjectileSlot(Slot) Then
            Call RequestProjectileTarget(Slot, True)
        Else
            Call WriteUseItemU(Slot)
        End If
    End If
End Sub

Public Sub UserItemClick(Optional ByVal DisableMacro As Boolean = True)
    If frmCarp.visible Or frmHerrero.visible Or frmComerciar.visible Or frmBancoObj.visible Then Exit Sub
    If pausa Then Exit Sub
    If UserMeditar Then Exit Sub
    If DisableMacro And frmMain.macrotrabajo.enabled Then frmMain.DesactivarMacroTrabajo
    If Not IsItemSelected Then Exit Sub
    Call UserOrEquipItem(frmMain.Inventario.SelectedItem, frmMain.Inventario.Equipped(frmMain.Inventario.SelectedItem), frmMain.Inventario.ObjIndex( _
            frmMain.Inventario.SelectedItem))
End Sub

Public Sub UserOrEquipItem(ByVal Slot As Integer, ByVal Equipped As Boolean, ByVal ObjIndex As Integer)
    Dim ObjType As Byte
    ObjType = ObjData(ObjIndex).ObjType
    G_LastSelectedSlot = Slot
    G_LasSelectedObjIndex = ObjIndex
    Select Case ObjType
        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otAnillos, eObjType.otManchas
            If Not Equipped Then
                Call WriteEquipItem(Slot)
            Else
                Call WriteUseItem(Slot)
            End If
        Case eObjType.otWeapon
            If ObjData(ObjIndex).proyectil = 1 And Equipped Then
                Call RequestProjectileTarget(Slot, False)
            Else
                If Not Equipped Then
                    Call WriteEquipItem(Slot)
                End If
            End If
        Case eObjType.OtHerramientas
            If Equipped Then
                Call WriteUseItem(Slot)
            Else
                If Not Equipped Then
                    Call WriteEquipItem(Slot)
                End If
            End If
        Case eObjType.otMinerales
            Call WriteUseItem(Slot)
        Case eObjType.OtDonador
            If Not Equipped Then
                Call WriteEquipItem(Slot)
            End If
        Case eObjType.otCollectibleCard
            frmCollectibleCard.cmdAccept.visible = True
            frmCollectibleCard.visible = True
        Case eObjType.otCastleSpawner
            Call WriteUseItem(Slot)
        Case Else
            Call WriteUseItem(Slot)
    End Select
End Sub

Public Sub HandleKeyUp(KeyCode As Integer, Shift As Integer)
    ' Toggles del viewport (consumen la tecla antes que cualquier hotkey).
    ' Activos incluso con dialogos abiertos, pero no cuando el chat tiene foco.
    ' Las teclas vienen de BindKeys (configurables via Teclas.ini).
    If Not IsInputFocus Then
        If KeyCode = BindKeys(e_KeyAction.eToggleViewportFullscreen).KeyCode And KeyCode <> 0 Then
            Call ToggleViewportFullscreen
            Exit Sub
        ElseIf KeyCode = BindKeys(e_KeyAction.eToggleViewportDebug).KeyCode And KeyCode <> 0 Then
            Call ViewportDebug_ToggleOverlay
            Exit Sub
        ElseIf ViewportDebug_IsFullscreen() Then
            If KeyCode = BindKeys(e_KeyAction.eToggleInventoryWindow).KeyCode And KeyCode <> 0 Then
                Call Toggle_FloatingWindow(efwInventory)
                Exit Sub
            ElseIf KeyCode = BindKeys(e_KeyAction.eToggleSpellsWindow).KeyCode And KeyCode <> 0 Then
                Call Toggle_FloatingWindow(efwSpells)
                Exit Sub
            End If
        End If
    End If
    If Not IsInputFocus Then
        If Not IsDialogOpen Then
            If Accionar(KeyCode) Then
                Exit Sub
            ElseIf KeyCode = BindKeys(e_KeyAction.eSendText).KeyCode Then
                Call OpenChatInput
            ElseIf KeyCode = vbKeyDelete Then
                Call OpenAndFocusClanChat
            ElseIf KeyCode = vbKeyEscape And Not UserSaliendo Then
                Call HandleEsc
            ElseIf KeyCode = 27 And UserSaliendo Then
                Call WriteCancelarExit
            ElseIf KeyCode = 80 And PescandoEspecial Then
                Call IntentarObtenerPezEspecial
            ElseIf KeyCode = vbKeyF1 Then
                Call ParseUserCommand("/SM")
                Call ParseUserCommand("/IRA " & targetName)
            End If
        End If
    Else
        Call FocusInput
    End If
End Sub

Public Sub HandleEsc()
    frmCerrar.Show , frmMain
End Sub

Public Function IsDialogOpen() As Boolean
    IsDialogOpen = pausa Or frmComerciar.visible Or frmComerciarUsu.visible Or frmBancoObj.visible Or frmGoliath.visible Or IsGameDialogOpen
End Function

Public Function IsInputFocus() As Boolean
    IsInputFocus = frmMain.SendTxt.visible Or frmMain.SendTxtCmsg.visible
End Function

Public Sub OpenAndFocusClanChat()
    If Not frmMain.SendTxt.visible Then
        If ViewportDebug_IsFullscreen() Then
            frmMain.SendTxtCmsg.Left = 5
            frmMain.SendTxtCmsg.Top = VIEWPORT_FULL_HEIGHT - 32
            frmMain.SendTxtCmsg.Width = hotkey_render_posX - 15
            frmMain.SendTxtCmsg.Height = SENDTXT_HUD_HEIGHT
        Else
            Call PositionHudChatInput
        End If
        frmMain.SendTxtCmsg.visible = True
        frmMain.SendTxtCmsg.ZOrder 0
        frmMain.SendTxtCmsg.SetFocus
    End If
    Call DialogosClanes.toggle_dialogs_visibility(True)
End Sub

Public Sub OpenChatInput()
    If Not frmCantidad.visible Then
        Call frmMain.CompletarEnvioMensajes
        StartOpenChatTime = GetTickCount
        If ViewportDebug_IsFullscreen() Then
            Call PositionFullscreenChatInput
        Else
            Call PositionHudChatInput
        End If
        frmMain.SendTxt.visible = True
        frmMain.SendTxt.ZOrder 0
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub FocusInput()
    If frmMain.SendTxt.visible Then
        frmMain.SendTxt.SetFocus
    End If
    If frmMain.SendTxtCmsg.visible Then
        frmMain.SendTxtCmsg.SetFocus
    End If
End Sub

Public Function GetGameplayForm() As Form
    Set GetGameplayForm = frmMain
End Function

Public Sub UseSpell(ByVal SpellSlot As Byte, ByVal SpellName As String)
    If pausa Then Exit Sub
    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Lanzar Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    iClickTick = TempTick
    LastMacroButton = tMacroButton.Lanzar
    If SpellName <> "(Vacío)" Then
        If UserStats.estado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_ESTAS_MUERTO"), .red, .green, .blue, .bold, .italic)
            End With
        Else
            If ModoHechizos = BloqueoLanzar Then
                If Not MainTimer.Check(TimersIndex.AttackSpell, False) Or Not MainTimer.Check(TimersIndex.CastSpell, False) Then
                    Exit Sub
                End If
            End If
            Call WriteCastSpell(SpellSlot)
            UsaMacro = True
            UsaLanzar = True
        End If
    End If
End Sub

Public Sub UpdateMapPos()
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    If frmMapaGrande.visible Then
        Call frmMapaGrande.ActualizarPosicionMapa
    End If
End Sub

Public Sub RequestSkills()
    If pausa Or tutorial_index > 0 Then Exit Sub
    If MostrarTutorial And tutorial_index <= 0 Then
        If tutorial(4).Activo = 1 Then
            tutorial_index = e_tutorialIndex.TUTORIAL_SkillPoints
            'TUTORIAL MAPA INSEGURO
            Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).Grh, -1, &H164B8A, , , False, 100, 479, 100, 535, _
                    640, 530, 64, 64)
            Exit Sub
        End If
    End If
    LlegaronSkills = True
    Call WriteRequestSkills
End Sub

Public Sub EquipSelectedItem()
    If frmMain.Inventario.IsItemSelected Then Call WriteEquipItem(frmMain.Inventario.SelectedItem)
End Sub

Public Sub OpenCreateObjectMenu()
    On Error GoTo createObj_Click_Err
    Dim i As Long
    For i = 1 To NumOBJs
        If ObjData(i).Name <> "" Then
            Dim subelemento As ListItem
            Set subelemento = FrmObjetos.ListView1.ListItems.Add(, , ObjData(i).Name)
            subelemento.SubItems(1) = i
        End If
    Next i
    GetGameplayForm().SetFocus
    FrmObjetos.Show , GetGameplayForm
    Exit Sub
createObj_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.createObj_Click", Erl)
    Resume Next
End Sub

Public Sub SelectInventoryTab()
    ActiveInventoryTab = eInventory
    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Inventario Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    iClickTick = TempTick
    LastMacroButton = tMacroButton.Inventario
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(1, hlst.ListIndex, hlst.Scroll)
    End If
End Sub

Public Sub SelectSpellTab()
    ActiveInventoryTab = eSpellList
    TempTick = GetTickCount And &H7FFFFFFF
    If TempTick - iClickTick < IntervaloEntreClicks And Not iClickTick = 0 And LastMacroButton <> tMacroButton.Hechizos Then
        Call WriteLogMacroClickHechizo(tMacro.Coordenadas)
    End If
    iClickTick = TempTick
    LastMacroButton = tMacroButton.Hechizos
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, hlst.ListIndex, hlst.Scroll)
    End If
End Sub

Public Sub GetMinimapPosition(ByRef x As Single, ByRef y As Single)
    x = x * (100 - 2 * HalfWindowTileWidth - 4) / 100 + HalfWindowTileWidth + 2
    y = y * (100 - 2 * HalfWindowTileHeight - 4) / 100 + HalfWindowTileHeight + 2
End Sub

Public Sub RequestMeditate()
    If UserStats.minman = UserStats.maxman Then Exit Sub
    If UserStats.estado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_ESTAS_MUERTO"), .red, .green, .blue, .bold, .italic) ' MENSAJE_ESTAS_MUERTO=¡Estás muerto!
        End With
        Exit Sub
    End If
    Call WriteMeditate
End Sub

Public Sub SetHotkey(ByVal Index As Integer, ByVal LastKnownSlot As Integer, ByVal HotkeyType As e_HotkeyType, ByVal HotkeySlot As Integer)
    HotkeyList(HotkeySlot).Index = Index
    HotkeyList(HotkeySlot).LastKnownSlot = LastKnownSlot
    HotkeyList(HotkeySlot).Type = HotkeyType
    Call SaveHotkey(Index, LastKnownSlot, HotkeyType, HotkeySlot)
    Call WriteSetHotkeySlot(HotkeySlot, Index, LastKnownSlot, HotkeyType)
End Sub

Public Sub ClearHotkeys()
    Dim i As Integer
    For i = 0 To HotKeyCount - 1
        HotkeyList(i).Index = -1
        HotkeyList(i).LastKnownSlot = -1
        HotkeyList(i).Type = Unknown
    Next i
End Sub

Public Sub ClearHotkeySlot(ByVal HotkeySlot As Integer)
    If HotkeySlot < 0 Or HotkeySlot >= HotKeyCount Then Exit Sub
    HotkeyList(HotkeySlot).Index = -1
    HotkeyList(HotkeySlot).LastKnownSlot = -1
    HotkeyList(HotkeySlot).Type = Unknown
    HotkeyList(HotkeySlot).CommandText = ""
    Call SaveHotkey(-1, -1, Unknown, HotkeySlot)
    Call WriteSetHotkeySlot(CByte(HotkeySlot), -1, -1, Unknown)
End Sub

Public Sub SetHotkeyCommand(ByVal HotkeySlot As Integer, ByVal CmdText As String)
    If HotkeySlot < 0 Or HotkeySlot >= HotKeyCount Then Exit Sub
    HotkeyList(HotkeySlot).Type = e_HotkeyType.Command
    HotkeyList(HotkeySlot).Index = 0
    HotkeyList(HotkeySlot).LastKnownSlot = 0
    HotkeyList(HotkeySlot).CommandText = CmdText
    Call SaveHotkey(0, 0, e_HotkeyType.Command, HotkeySlot)
End Sub

Public Sub ShowInteractionMenu(ByVal FormTop As Long, _
                                ByVal FormLeft As Long, _
                                ByVal FormHeight As Long, _
                                ByVal x As Integer, _
                                ByVal y As Integer, _
                                ByRef GameplayArea As RECT)
    On Error GoTo ShowInteractionMenu_Err
    Dim charindex As Integer
    charindex = MapData(tX, tY).charindex
    If charindex = 0 Then charindex = MapData(tX, tY + 1).charindex
    If charindex = 0 Or charindex = UserCharIndex Then Exit Sub

    Dim Frm As Form
    Call WriteLeftClick(tX, tY)
    TargetX = tX
    TargetY = tY

    If charlist(charindex).EsMascota Then
        Set Frm = MenuNPC
    ElseIf Not charlist(charindex).EsNpc Then
        targetName = charlist(charindex).nombre
        If charlist(UserCharIndex).priv > 0 And Shift = 0 Then
            Set Frm = MenuGM
        Else
            Set Frm = MenuUser
        End If
    End If

    If Not Frm Is Nothing Then
        Frm.Show
        Frm.Left = FormLeft + (GameplayArea.Left + x + 1) * screen.TwipsPerPixelX
        If (GameplayArea.Top + y) * screen.TwipsPerPixelY + Frm.Height > FormHeight Then
            Frm.Top = FormTop + (GameplayArea.Top + y) * screen.TwipsPerPixelY - Frm.Height
        Else
            Frm.Top = FormTop + (GameplayArea.Top + y) * screen.TwipsPerPixelY
        End If
        Set Frm = Nothing
    End If
    Exit Sub
ShowInteractionMenu_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUi.ShowInteractionMenu", Erl)
    Resume Next
End Sub

' ============== Viewport fullscreen toggle (PoC 2026-05-20) ==============
' Plan: ia/plans/2026/mayo/20.002.viewport-fullscreen-poc.md
' (Declaraciones a nivel de modulo movidas al principio del archivo, ver inicio.)

Public Sub ToggleViewportFullscreen()
    If g_renderer_state_saved Then
        Call ExitViewportFullscreen
    Else
        Call EnterViewportFullscreen
    End If
End Sub

Public Sub EnterViewportFullscreen()
    On Error GoTo EnterViewportFullscreen_Err
    If g_renderer_state_saved Then Exit Sub
    Call ViewportDebug_AppendDiagLog("[Enter] start")
    g_renderer_state_saved = True
    Call SetViewportLogicalFull
    Call ClearRendererClip
    frmMain.renderer.ZOrder 0
    Call Apply_FullscreenLayout(True)
    Call Set_FullscreenHudActive(True)
    Call Set_FullscreenGmBarActive(True)
    Call Set_FullscreenBuffsPanelActive(True)
    Call Set_FullscreenButtonsActive(True)
    Call Set_FloatingWindowsActive(True)
    If frmMain.SendTxt.visible Then
        Call PositionFullscreenChatInput
        frmMain.SendTxt.ZOrder 0
    End If
    If frmMain.SendTxtCmsg.visible Then
        frmMain.SendTxtCmsg.Left = 5
        frmMain.SendTxtCmsg.Top = VIEWPORT_FULL_HEIGHT - 32
        frmMain.SendTxtCmsg.Width = hotkey_render_posX - 15
        frmMain.SendTxtCmsg.Height = SENDTXT_HUD_HEIGHT
        frmMain.SendTxtCmsg.ZOrder 0
    End If
    Call ViewportDebug_AppendDiagLog("[Enter] logical viewport FULL")
    Call Recalc_TileEngine_Viewport
    Call ViewportDebug_AppendDiagLog("[Enter] recalc done")
    Call ViewportDebug_SetViewportMode(True)
    Call ViewportDebug_AppendDiagLog("[Enter] OK")
    Exit Sub
EnterViewportFullscreen_Err:
    Call ViewportDebug_AppendDiagLog("[Enter] ERROR " & Err.Number & " " & Err.Description & " line=" & Erl)
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUI.EnterViewportFullscreen", Erl)
    Resume Next
End Sub

Public Sub ExitViewportFullscreen()
    On Error GoTo ExitViewportFullscreen_Err
    If Not g_renderer_state_saved Then Exit Sub
    Call ViewportDebug_AppendDiagLog("[Exit] start")
    g_renderer_state_saved = False
    Call SetViewportLogicalHud
    Call ApplyRendererHudClip
    Call PositionHudChatInput
    Call Set_FloatingWindowsActive(False)
    Call Set_FullscreenButtonsActive(False)
    Call Set_FullscreenBuffsPanelActive(False)
    frmMain.renderer.ZOrder 1
    Call Set_FullscreenGmBarActive(False)
    Call Set_FullscreenHudActive(False)
    Call Apply_FullscreenLayout(False)
    Call Recalc_TileEngine_Viewport
    Call ViewportDebug_SetViewportMode(False)
    Call ViewportDebug_AppendDiagLog("[Exit] OK")
    Exit Sub
ExitViewportFullscreen_Err:
    Call ViewportDebug_AppendDiagLog("[Exit] ERROR " & Err.Number & " " & Err.Description & " line=" & Erl)
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUI.ExitViewportFullscreen", Erl)
    Resume Next
End Sub

Public Function Is_RendererStateFullscreen() As Boolean
    Is_RendererStateFullscreen = g_renderer_state_saved
End Function

Public Sub Mark_PendingRestoreFullscreen()
    g_pending_restore_full = True
End Sub

Public Sub Try_RestoreFullscreenAfterLogin()
    On Error Resume Next
    If g_pending_restore_full Then
        g_pending_restore_full = False
        If Not g_renderer_state_saved Then Call EnterViewportFullscreen
    End If
End Sub

Private Sub PositionHudChatInput()
    On Error Resume Next
    frmMain.SendTxt.Left = SENDTXT_HUD_LEFT
    frmMain.SendTxt.Top = SENDTXT_HUD_TOP
    frmMain.SendTxt.Width = SENDTXT_HUD_WIDTH
    frmMain.SendTxt.Height = SENDTXT_HUD_HEIGHT
    frmMain.SendTxtCmsg.Left = SENDTXT_HUD_LEFT
    frmMain.SendTxtCmsg.Top = SENDTXT_HUD_TOP
    frmMain.SendTxtCmsg.Width = SENDTXT_HUD_WIDTH
    frmMain.SendTxtCmsg.Height = SENDTXT_HUD_HEIGHT
End Sub

Private Sub PositionFullscreenChatInput()
    On Error Resume Next
    frmMain.SendTxt.Left = 5
    frmMain.SendTxt.Top = VIEWPORT_FULL_HEIGHT - 32
    frmMain.SendTxt.Width = hotkey_render_posX - 15
    frmMain.SendTxt.Height = SENDTXT_HUD_HEIGHT
End Sub

Private Sub ApplyRendererHudClip()
    On Error GoTo ApplyRendererHudClip_Err
    Dim hRegion As Long
    hRegion = CreateRectRgn(g_viewport_logical_left, g_viewport_logical_top, g_viewport_logical_left + g_viewport_logical_width, g_viewport_logical_top + g_viewport_logical_height)
    If hRegion <> 0 Then
        If SetWindowRgn(frmMain.renderer.hWnd, hRegion, 1) = 0 Then
            Call DeleteObject(hRegion)
        End If
    End If
    Exit Sub
ApplyRendererHudClip_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUI.ApplyRendererHudClip", Erl)
    Resume Next
End Sub

Private Sub ClearRendererClip()
    On Error GoTo ClearRendererClip_Err
    Call SetWindowRgn(frmMain.renderer.hWnd, 0, 1)
    Exit Sub
ClearRendererClip_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGameplayUI.ClearRendererClip", Erl)
    Resume Next
End Sub

' SetHudControlsVisible legacy fue reemplazado por ModFullscreenLayout.Apply_FullscreenLayout.
' Plan: ia/plans/2026/mayo/21.002.viewport-fullscreen-hud-flotante.md (Etapa 1).
