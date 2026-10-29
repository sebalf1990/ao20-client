Attribute VB_Name = "ModFullscreenGmBar"
' Argentum 20 Game Client
'
'    Barra superior FULL (GM toolbar). DX overlay.
'    Para todos: Argentum20 + Online + FPS.
'    Solo GM: PanelGM, Crear Obj, Spawn NPC, Invisible.
'    Hit-test desde renderer_MouseDown.
'
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md (Etapa 9)
'
Option Explicit

Public Const GMBAR_HEIGHT      As Long = 22

Private Const GMBAR_LOGO_X     As Long = 8
Private Const GMBAR_ONLINE_X   As Long = 110
Private Const GMBAR_TEXT_Y     As Long = 4

' Coordenadas relativas centrales de los items GM
Private Const GM_BTN_PANELGM_X   As Long = -210
Private Const GM_BTN_PANELGM_W   As Long = 80
Private Const GM_BTN_CREAR_X     As Long = -120
Private Const GM_BTN_CREAR_W     As Long = 90
Private Const GM_BTN_SPAWN_X     As Long = -20
Private Const GM_BTN_SPAWN_W     As Long = 90
Private Const GM_BTN_INVIS_X     As Long = 80
Private Const GM_BTN_INVIS_W     As Long = 80

Private g_gmbar_active As Boolean

Public Sub Set_FullscreenGmBarActive(ByVal active As Boolean)
    g_gmbar_active = active
End Sub

Public Function Is_FullscreenGmBarActive() As Boolean
    Is_FullscreenGmBarActive = g_gmbar_active
End Function

Public Sub Render_FullscreenGmBar()
    On Error GoTo Err_h
    If Not g_gmbar_active Then Exit Sub
    If Not EsGM Then Exit Sub
    Dim sw As Long
    sw = frmMain.renderer.ScaleWidth

    Dim bg As RGBA
    bg = RGBA_From_Comp(15, 70, 65, 230)
    Call Engine_Draw_Box(0, 0, sw, GMBAR_HEIGHT, bg)

    Dim borderC As RGBA
    borderC = RGBA_From_Comp(60, 130, 120, 240)
    Call Engine_Draw_Box(0, GMBAR_HEIGHT - 1, sw, 1, borderC)

    Dim white(3) As RGBA
    Call RGBAList(white, 240, 240, 240, 255)

    Call Engine_Text_Render("Argentum20", GMBAR_LOGO_X, GMBAR_TEXT_Y, white, 1, True)
    Call Engine_Text_Render("Online: " & usersOnline, GMBAR_ONLINE_X, GMBAR_TEXT_Y, white, 1, False)

    If EsGM Then
        Dim cx As Long
        cx = sw \ 2
        Call Engine_Text_Render("PanelGM", cx + GM_BTN_PANELGM_X, GMBAR_TEXT_Y, white, 1, True)
        Call Engine_Text_Render("Crear Obj", cx + GM_BTN_CREAR_X, GMBAR_TEXT_Y, white, 1, True)
        Call Engine_Text_Render("Spawn NPC", cx + GM_BTN_SPAWN_X, GMBAR_TEXT_Y, white, 1, True)
        Call Engine_Text_Render("Invisible", cx + GM_BTN_INVIS_X, GMBAR_TEXT_Y, white, 1, True)
    End If

    Dim fpsTxt As String
    fpsTxt = "FPS: " & engine.fps
    Dim fpsW As Long
    fpsW = Engine_Text_Width(fpsTxt, False)
    Call Engine_Text_Render(fpsTxt, sw - fpsW - 8, GMBAR_TEXT_Y, white, 1, False)
    Exit Sub
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenGmBar.Render_FullscreenGmBar", Erl)
    Resume Next
End Sub

' Devuelve True si el click fue consumido por la GM bar
Public Function Handle_GmBarClick(ByVal x As Long, ByVal y As Long) As Boolean
    On Error GoTo Err_h
    Handle_GmBarClick = False
    If Not g_gmbar_active Then Exit Function
    If Not EsGM Then Exit Function
    If y < 0 Or y > GMBAR_HEIGHT Then Exit Function

    ' Cualquier click sobre la barra (zona reservada) se consume
    Handle_GmBarClick = True

    If Not EsGM Then Exit Function

    Dim sw As Long, cx As Long
    sw = frmMain.renderer.ScaleWidth
    cx = sw \ 2

    If x >= cx + GM_BTN_PANELGM_X And x < cx + GM_BTN_PANELGM_X + GM_BTN_PANELGM_W Then
        frmPanelgm.Width = 4860
        Call WriteSOSShowList
        Call WriteGMPanel
        Exit Function
    End If
    If x >= cx + GM_BTN_CREAR_X And x < cx + GM_BTN_CREAR_X + GM_BTN_CREAR_W Then
        Call OpenCreateObjectMenu
        Exit Function
    End If
    If x >= cx + GM_BTN_SPAWN_X And x < cx + GM_BTN_SPAWN_X + GM_BTN_SPAWN_W Then
        Call WriteSpawnListRequest
        Exit Function
    End If
    If x >= cx + GM_BTN_INVIS_X And x < cx + GM_BTN_INVIS_X + GM_BTN_INVIS_W Then
        Call ParseUserCommand("/INVISIBLE")
        Exit Function
    End If
    Exit Function
Err_h:
    Call RegistrarError(Err.Number, Err.Description, "ModFullscreenGmBar.Handle_GmBarClick", Erl)
    Resume Next
End Function
