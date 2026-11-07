Attribute VB_Name = "ModMenuTransparency"
' Argentum 20 Game Client
'
'    Transparencia configurable para menus/ventanas auxiliares.
'    Plan: ia/archivo/plans/2026/2026-05-21-viewport-fullscreen-hud-flotante.md (Etapa 9 V9.12)
'
Option Explicit

Private Const APP_NAME    As String = "UI"
Private Const KEY_NAME    As String = "MenuAlpha"
Private Const ALPHA_MIN   As Integer = 80
Private Const ALPHA_DEF   As Integer = 230

Private Const GWL_EXSTYLE     As Long = (-20)
Private Const WS_EX_LAYERED   As Long = &H80000
Private Const LWA_ALPHA       As Long = &H2
Private Const RDW_INVALIDATE  As Long = &H1
Private Const RDW_ERASE       As Long = &H4
Private Const RDW_FRAME       As Long = &H400
Private Const RDW_ALLCHILDREN As Long = &H80

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public g_menu_alpha As Integer

Public Sub MenuTransparency_EnsureLoaded()
    If g_menu_alpha > 0 Then Exit Sub
    Dim v As String
    Dim parsedAlpha As Double
    v = GetSetting(APP_NAME, KEY_NAME)
    ' E1: Val a Double + clamp 80..255 (CInt desbordaba con valores enormes y dejaba alpha 0).
    If LenB(v) = 0 Or Not IsNumeric(v) Then
        parsedAlpha = ALPHA_DEF
    Else
        parsedAlpha = val(v)
    End If
    If parsedAlpha < ALPHA_MIN Then parsedAlpha = ALPHA_MIN
    If parsedAlpha > 255 Then parsedAlpha = 255
    g_menu_alpha = CInt(parsedAlpha)
End Sub

Public Sub MenuTransparency_Set(ByVal value As Integer)
    If value < ALPHA_MIN Then value = ALPHA_MIN
    If value > 255 Then value = 255
    g_menu_alpha = value
    Call SaveSetting(APP_NAME, KEY_NAME, CStr(value))
End Sub

' E3: igual que MenuTransparency_Set pero sin escribir el ini (para el arrastre del slider).
Public Sub MenuTransparency_Preview(ByVal value As Integer)
    If value < ALPHA_MIN Then value = ALPHA_MIN
    If value > 255 Then value = 255
    g_menu_alpha = value
End Sub

Public Function MenuTransparency_Min() As Integer
    MenuTransparency_Min = ALPHA_MIN
End Function

Public Function MenuTransparency_Default() As Integer
    MenuTransparency_Default = ALPHA_DEF
End Function

' Aplica el alpha actual al form pasado (WS_EX_LAYERED + SetLayeredWindowAttributes).
' E3: con alpha 255 (opaco) se QUITA WS_EX_LAYERED en vez de dejarlo puesto: una ventana
' layered cuesta de mas al compositor aunque sea opaca.
' Llamar desde Form_Load del form que se quiere transparentar.
Public Sub MenuTransparency_ApplyToForm(ByVal hWnd As Long)
    On Error GoTo MenuTransparency_ApplyToForm_Err
    Call MenuTransparency_EnsureLoaded
    Dim style As Long
    style = GetWindowLong(hWnd, GWL_EXSTYLE)
    If g_menu_alpha >= 255 Then
        If (style And WS_EX_LAYERED) <> 0 Then
            Call SetWindowLong(hWnd, GWL_EXSTYLE, style And Not WS_EX_LAYERED)
            Call RedrawWindow(hWnd, 0&, 0&, RDW_INVALIDATE Or RDW_ERASE Or RDW_FRAME Or RDW_ALLCHILDREN)
        End If
        Exit Sub
    End If
    style = style Or WS_EX_LAYERED
    Call SetWindowLong(hWnd, GWL_EXSTYLE, style)
    Call SetLayeredWindowAttributes(hWnd, 0, CByte(g_menu_alpha), LWA_ALPHA)
    Exit Sub
MenuTransparency_ApplyToForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModMenuTransparency.MenuTransparency_ApplyToForm", Erl)
    Resume Next
End Sub

' Devuelve un alpha (0..255) escalado por la configuracion del usuario.
' Usar para ajustar las alphas que se pasan a Engine_Draw_Box en ventanas flotantes.
Public Function MenuTransparency_ScaleAlpha(ByVal baseAlpha As Long) As Byte
    Call MenuTransparency_EnsureLoaded
    Dim r As Long
    r = (baseAlpha * g_menu_alpha) \ 255
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    MenuTransparency_ScaleAlpha = CByte(r)
End Function
