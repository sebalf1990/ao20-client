Attribute VB_Name = "ModFullscreenBuffsPanel"
' Argentum 20 Game Client
'
'    Panel DX de buffs, debuffs y cooldowns modernos para viewport FULL.
'    Fuente de verdad: BuffList, DeBuffList y CDList, alimentadas por
'    eSendSkillCdUpdate. Fuerza/Agilidad quedan fuera de este primer corte.
'
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md (Etapa 4)
'
Option Explicit

Private Const CELL_SIZE       As Long = 32
Private Const CELL_STEP       As Long = 38
Private Const PANEL_MAX_W     As Long = 500
Private Const PANEL_Y         As Long = 70
Private Const PANEL_PAD       As Long = 4
Private Const TEXT_OFFSET_Y   As Long = 24

Private g_fullscreen_buffs_active As Boolean
Private g_cooldown_color As RGBA

' Plan 25.003: Effect49 (Potenciado) parpadea y suena SND_DOPA en los ultimos 12s.
Private Const CLIENT_EFFECT_TYPE_ID_POTENCIADO As Integer = 49
Private Const POTENCIADO_FINAL_PHASE_MS As Long = 12000
Private g_potenciado_last_dopa_second As Long

Public Sub Set_FullscreenBuffsPanelActive(ByVal active As Boolean)
    g_fullscreen_buffs_active = active
End Sub

Public Function Is_FullscreenBuffsPanelActive() As Boolean
    Is_FullscreenBuffsPanelActive = g_fullscreen_buffs_active
End Function

Public Sub Render_FullscreenBuffsPanel()
    On Error GoTo Render_FullscreenBuffsPanel_Err
    If Not g_fullscreen_buffs_active Then Exit Sub
    If g_game_state Is Nothing Then Exit Sub
    If g_game_state.state <> e_state_gameplay_screen Then Exit Sub

    Dim CurrTime As Long
    CurrTime = GetTickCount()
    Call UpdateEffectTime(BuffList, CurrTime)
    Call UpdateEffectTime(DeBuffList, CurrTime)
    Call UpdateEffectTime(CDList, CurrTime)

    Dim totalCount As Long
    totalCount = BuffList.EffectCount + DeBuffList.EffectCount + CDList.EffectCount
    If totalCount <= 0 Then Exit Sub

    Dim maxCells As Long
    maxCells = PANEL_MAX_W \ CELL_STEP
    If maxCells < 1 Then maxCells = 1
    If totalCount > maxCells Then totalCount = maxCells

    Dim panelW As Long
    panelW = totalCount * CELL_STEP - (CELL_STEP - CELL_SIZE) + PANEL_PAD * 2

    ' V9.4: panel a la izquierda del minimapa, crece hacia la izquierda
    Const MINIMAP_LEFT_REF As Long = 884
    Dim startX As Long
    startX = MINIMAP_LEFT_REF - panelW - 8
    If startX < PANEL_PAD Then startX = PANEL_PAD

    ' V9: sin fondo, los iconos van sueltos sobre el viewport

    Dim colors(3) As RGBA
    Call RGBAList(colors, 255, 255, 255, 255)
    Call SetRGBA(g_cooldown_color, 125, 125, 125, 120)

    Dim currentX As Long
    Dim rendered As Long
    currentX = startX
    rendered = 0

    Call RenderEffectList(BuffList, currentX, PANEL_Y, CurrTime, colors, rendered, maxCells, False)
    Call RenderEffectList(DeBuffList, currentX, PANEL_Y, CurrTime, colors, rendered, maxCells, False)
    Call RenderEffectList(CDList, currentX, PANEL_Y, CurrTime, colors, rendered, maxCells, True)
    Exit Sub
Render_FullscreenBuffsPanel_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenBuffsPanel.Render_FullscreenBuffsPanel", Erl)
    Resume Next
End Sub

Private Sub RenderEffectList(ByRef EffectList As t_ActiveEffectList, ByRef currentX As Long, ByVal y As Long, ByVal CurrTime As Long, ByRef colors() As RGBA, ByRef rendered As Long, ByVal maxCells As Long, ByVal isCooldown As Boolean)
    On Error GoTo RenderEffectList_Err
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        If rendered >= maxCells Then Exit Sub
        Call DrawFullscreenEffect(currentX, y, EffectList.EffectList(i), CurrTime, colors, isCooldown)
        currentX = currentX + CELL_STEP
        rendered = rendered + 1
    Next i
    Exit Sub
RenderEffectList_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenBuffsPanel.RenderEffectList", Erl)
    Resume Next
End Sub

Private Sub DrawFullscreenEffect(ByVal x As Long, ByVal y As Long, ByRef Effect As t_ActiveEffect, ByVal CurrTime As Long, ByRef colors() As RGBA, ByVal isCooldown As Boolean)
    On Error GoTo DrawFullscreenEffect_Err
    If Effect.Grh <= 0 Then Exit Sub

    ' Plan 25.003: Potenciado parpadea en los ultimos 12s y dispara SND_DOPA cada vez que cruza un segundo.
    Dim renderColors(3) As RGBA
    Dim isPotenciadoFinalPhase As Boolean
    Dim remainingMs As Long
    Dim i As Integer
    For i = 0 To 3
        renderColors(i) = colors(i)
    Next i
    If Effect.TypeId = CLIENT_EFFECT_TYPE_ID_POTENCIADO And Effect.duration > 0 Then
        remainingMs = Effect.duration - (CurrTime - Effect.startTime)
        If remainingMs > 0 And remainingMs <= POTENCIADO_FINAL_PHASE_MS Then
            isPotenciadoFinalPhase = True
            ' Blink ~4Hz alternando alpha 255/80.
            If (CurrTime \ 250) Mod 2 = 0 Then
                Dim k As Integer
                For k = 0 To 3
                    renderColors(k).a = 80
                Next k
            End If
            Dim currentSecond As Long
            currentSecond = (remainingMs + 999) \ 1000
            If currentSecond <> g_potenciado_last_dopa_second Then
                g_potenciado_last_dopa_second = currentSecond
                Call ao20audio.StopWav(SND_DOPA)
                Call ao20audio.PlayWav(SND_DOPA)
            End If
        Else
            g_potenciado_last_dopa_second = 0
        End If
    End If

    Dim Grh As Grh
    Call InitGrh(Grh, Effect.Grh)
    Call Grh_Render_Advance(Grh, x, y, CELL_SIZE, CELL_SIZE, renderColors)

    Dim angle As Single
    If Effect.duration > 0 Then
        angle = (CurrTime - Effect.startTime) * 360 / Effect.duration
        If angle < 0 Then angle = 0
        If angle > 360 Then angle = 360
        Call Engine_Draw_Load(x + CELL_SIZE \ 2, y + CELL_SIZE \ 2, CELL_SIZE, CELL_SIZE, g_cooldown_color, angle)
    End If

    Dim labelTxt As String
    labelTxt = EffectLabel(Effect, CurrTime, isCooldown)
    If LenB(labelTxt) > 0 Then
        Call Engine_Text_Render(labelTxt, x + 2, y + TEXT_OFFSET_Y, COLOR_WHITE, 1, False)
    End If
    Exit Sub
DrawFullscreenEffect_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenBuffsPanel.DrawFullscreenEffect", Erl)
    Resume Next
End Sub

Private Function EffectLabel(ByRef Effect As t_ActiveEffect, ByVal CurrTime As Long, ByVal isCooldown As Boolean) As String
    On Error GoTo EffectLabel_Err
    ' Plan 20.002: regla global del motor para EOT (buffs/debuffs): stacks solo si hay 2 o mas,
    ' nunca segundos (el barrido radial es el timer). Cooldowns conservan segundos (plan 25.003).
    If Effect.StackCount > 1 Then
        EffectLabel = CStr(Effect.StackCount)
        Exit Function
    End If
    If isCooldown And Effect.duration > 0 Then
        Dim remaining As Long
        remaining = (Effect.duration - (CurrTime - Effect.startTime) + 999) \ 1000
        If remaining < 0 Then remaining = 0
        EffectLabel = CStr(remaining)
    Else
        EffectLabel = ""
    End If
    Exit Function
EffectLabel_Err:
    EffectLabel = ""
End Function
