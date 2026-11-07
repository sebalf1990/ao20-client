VERSION 5.00
Begin VB.Form frmRadarOpciones
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones del radar"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   4440
      Width           =   1200
   End
   Begin VB.CommandButton cmdRadarGm
      Caption         =   "Radar de GM (/SHOWNPC)"
      Height          =   375
      Left            =   2460
      TabIndex        =   9
      Top             =   3900
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdRestaurar
      Caption         =   "Restaurar colores"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3900
      Width           =   2220
   End
   Begin VB.CheckBox chkJefes
      Caption         =   "Jefes (con detectar NPCs)"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   4300
   End
   Begin VB.CheckBox chkSimbolos
      Caption         =   "Sacerdotes y banqueros"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3180
      Width           =   4300
   End
   Begin VB.CheckBox chkQuests
      Caption         =   "Quests (! y ?)"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   4300
   End
   Begin VB.CheckBox chkClan
      Caption         =   "Compañeros de clan"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2580
      Width           =   4300
   End
   Begin VB.CheckBox chkParty
      Caption         =   "Compañeros de party"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   4300
   End
   Begin VB.ComboBox cmbPaleta
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   2520
   End
   Begin VB.ComboBox cmbPunto
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   990
      Width           =   2520
   End
   Begin VB.ComboBox cmbMinimapa
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   2520
   End
   Begin VB.Label lblCapas
      Caption         =   "Mostrar en el minimapa:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1980
      Width           =   4560
   End
   Begin VB.Label lblPaleta
      Caption         =   "Colores"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1485
      Width           =   1980
   End
   Begin VB.Label lblPunto
      Caption         =   "Tamaño del punto"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1035
      Width           =   1980
   End
   Begin VB.Label lblMinimapa
      Caption         =   "Minimapa"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   585
      Width           =   1980
   End
   Begin VB.Label lblTitulo
      Caption         =   "Los cambios se aplican al instante."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   150
      Width           =   4560
   End
End
Attribute VB_Name = "frmRadarOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
' Plan 2026-11-07-radar-v2, E2b: player radar options. Every change is saved in
' Configuracion.ini (section OPCIONES) and applied at once (ModUtils.ReloadRadarColors).
Option Explicit

Private loadingOptions As Boolean

Private Sub Form_Load()
    On Error GoTo Form_Load_Err
    loadingOptions = True
    Call cmbMinimapa.AddItem("Clásico")
    Call cmbMinimapa.AddItem("Centrado")
    Call cmbPunto.AddItem("Como radar_colors.ini")
    Call cmbPunto.AddItem("Chico")
    Call cmbPunto.AddItem("Mediano")
    Call cmbPunto.AddItem("Grande")
    Call cmbPaleta.AddItem("Normal")
    Call cmbPaleta.AddItem("Alto contraste")
    Call cmbPaleta.AddItem("Daltónico")
    Call LoadControlsFromOptions
    cmdRadarGm.visible = IsLocalUserGm()
    loadingOptions = False
    Exit Sub
Form_Load_Err:
    loadingOptions = False
    Call RegistrarError(Err.Number, Err.Description, "frmRadarOpciones.Form_Load", Erl)
End Sub

Private Sub LoadControlsFromOptions()
    If CenteredMinimap <> 0 Then
        cmbMinimapa.ListIndex = 1
    Else
        cmbMinimapa.ListIndex = 0
    End If
    If RadarPointSizeOption <= 3 Then
        cmbPunto.ListIndex = RadarPointSizeOption
    Else
        cmbPunto.ListIndex = 0
    End If
    If RadarPalette <= 2 Then
        cmbPaleta.ListIndex = RadarPalette
    Else
        cmbPaleta.ListIndex = 0
    End If
    chkParty.value = BoolToCheck(RadarShowParty)
    chkClan.value = BoolToCheck(RadarShowClan)
    chkQuests.value = BoolToCheck(RadarShowQuests)
    chkSimbolos.value = BoolToCheck(RadarShowSymbols)
    chkJefes.value = BoolToCheck(RadarShowBosses)
End Sub

Private Function BoolToCheck(ByVal flag As Boolean) As Integer
    If flag Then
        BoolToCheck = vbChecked
    Else
        BoolToCheck = vbUnchecked
    End If
End Function

Private Function IsLocalUserGm() As Boolean
    On Error GoTo IsLocalUserGm_Err
    If UserCharIndex > 0 Then IsLocalUserGm = (charlist(UserCharIndex).priv > 0)
    Exit Function
IsLocalUserGm_Err:
    IsLocalUserGm = False
End Function

Private Sub SaveCheck(ByVal settingName As String, ByVal box As CheckBox)
    On Error GoTo SaveCheck_Err
    If loadingOptions Then Exit Sub
    Call SaveSetting("OPCIONES", settingName, IIf(box.value = vbChecked, "1", "0"))
    Call ReloadRadarColors
    Exit Sub
SaveCheck_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmRadarOpciones.SaveCheck", Erl)
End Sub

Private Sub cmbMinimapa_Click()
    On Error GoTo cmbMinimapa_Click_Err
    If loadingOptions Then Exit Sub
    CenteredMinimap = cmbMinimapa.ListIndex
    Call SaveSetting("OPCIONES", "CenteredMinimap", CStr(CenteredMinimap))
    frmMain.btnZoomIn.visible = (CenteredMinimap = 1)
    frmMain.btnZoomOut.visible = (CenteredMinimap = 1)
    Call DibujarMiniMapa
    Exit Sub
cmbMinimapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmRadarOpciones.cmbMinimapa_Click", Erl)
End Sub

Private Sub cmbPunto_Click()
    On Error GoTo cmbPunto_Click_Err
    If loadingOptions Then Exit Sub
    Call SaveSetting("OPCIONES", "RadarPointSize", CStr(cmbPunto.ListIndex))
    Call ReloadRadarColors
    Exit Sub
cmbPunto_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmRadarOpciones.cmbPunto_Click", Erl)
End Sub

Private Sub cmbPaleta_Click()
    On Error GoTo cmbPaleta_Click_Err
    If loadingOptions Then Exit Sub
    Call SaveSetting("OPCIONES", "RadarPalette", CStr(cmbPaleta.ListIndex))
    Call ReloadRadarColors
    Exit Sub
cmbPaleta_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmRadarOpciones.cmbPaleta_Click", Erl)
End Sub

Private Sub chkParty_Click()
    Call SaveCheck("RadarShowParty", chkParty)
End Sub

Private Sub chkClan_Click()
    Call SaveCheck("RadarShowClan", chkClan)
End Sub

Private Sub chkQuests_Click()
    Call SaveCheck("RadarShowQuests", chkQuests)
End Sub

Private Sub chkSimbolos_Click()
    Call SaveCheck("RadarShowSymbols", chkSimbolos)
End Sub

Private Sub chkJefes_Click()
    Call SaveCheck("RadarShowBosses", chkJefes)
End Sub

Private Sub cmdRestaurar_Click()
    On Error GoTo cmdRestaurar_Click_Err
    Call SaveSetting("OPCIONES", "RadarPalette", "0")
    Call SaveSetting("OPCIONES", "RadarPointSize", "0")
    Call ReloadRadarColors
    loadingOptions = True
    Call LoadControlsFromOptions
    loadingOptions = False
    Exit Sub
cmdRestaurar_Click_Err:
    loadingOptions = False
    Call RegistrarError(Err.Number, Err.Description, "frmRadarOpciones.cmdRestaurar_Click", Erl)
End Sub

Private Sub cmdRadarGm_Click()
    On Error GoTo cmdRadarGm_Click_Err
    Call WriteNpcRadarToggle
    Exit Sub
cmdRadarGm_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmRadarOpciones.cmdRadarGm_Click", Erl)
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
