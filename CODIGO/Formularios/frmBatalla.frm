VERSION 5.00
Begin VB.Form frmBatalla 
   BorderStyle     =   0  'None
   ClientHeight    =   7620
   ClientLeft      =   -60
   ClientTop       =   -60
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image cmdEntrarAl 
      Height          =   390
      Left            =   4740
      Top             =   4530
      Width           =   1470
   End
   Begin VB.Image cmdIrA 
      Height          =   390
      Left            =   1200
      Top             =   4530
      Width           =   1470
   End
   Begin VB.Image cmdOrganizar 
      Height          =   390
      Left            =   4770
      Top             =   2400
      Width           =   1470
   End
   Begin VB.Image cmdIrAl 
      Height          =   390
      Left            =   1200
      Top             =   2370
      Width           =   1470
   End
   Begin VB.Image cmdVolver 
      Height          =   540
      Left            =   6960
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lblClasificacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Madera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   5430
      TabIndex        =   6
      Top             =   660
      Width           =   1050
   End
   Begin VB.Label lblELOUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000 Puntos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   1290
      TabIndex        =   5
      Top             =   660
      Width           =   1050
   End
   Begin VB.Label TopELO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Nadie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   0
      Left            =   4290
      TabIndex        =   4
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label TopELO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Nadie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   4290
      TabIndex        =   3
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label TopELO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Nadie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   4290
      TabIndex        =   2
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label TopELO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Nadie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   3
      Left            =   4290
      TabIndex        =   1
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label TopELO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Nadie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   4
      Left            =   4290
      TabIndex        =   0
      Top             =   6840
      Width           =   2415
   End
End
Attribute VB_Name = "frmBatalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cBotonCerrar    As clsGraphicalButton
Private cBotonRanked    As clsGraphicalButton
Private cBotonArena    As clsGraphicalButton
Private cBotonPlantes    As clsGraphicalButton
Private cBotonRetos    As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Private Sub cmdEntrarAl_Click()
    Call WriteBatallaPVP(2)
    Unload Me
End Sub

Private Sub cmdIrA_Click()
    If MsgBox("¿Seguro que quieres entrar a la Arena de Rinkel?", vbYesNo, "Atencion!") = vbNo Then Exit Sub
    Call WriteBatallaPVP(1)
    Unload Me
End Sub

Private Sub cmdIrDuelo_Click()
    Call WriteBatallaPVP(0)
    Unload Me
End Sub

Private Sub cmdRetos_Click()
    Unload Me
    Call FrmRetos.Show(vbModeless, frmMain)
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Public Sub Iniciar_Labels()
    Dim UserClasificacion As String
    
    For i = 1 To 5
        TopELO(i - 1).Caption = Ranking(i).name & " - ELO: " & Ranking(i).ELO
    Next i
        
    UserClasificacion = AsignarClasificacion(UserELO)
        
    lblELOUser.Caption = UserELO & " Puntos"
    lblClasificacion.Caption = UserClasificacion
End Sub

Private Function AsignarClasificacion(ByVal UserELO As Double) As String

    If UserELO < 1100 Then
        AsignarClasificacion = "Madera"
    ElseIf UserELO <= 1300 Then
        AsignarClasificacion = "Bronce"
    ElseIf UserELO <= 1500 Then
        AsignarClasificacion = "Oro"
    ElseIf UserELO <= 1700 Then
        AsignarClasificacion = "Platino"
    ElseIf UserELO > 1900 Then
        AsignarClasificacion = "Diamante"

    End If

End Function

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("batalla.bmp")
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set LastButtonPressed = New clsGraphicalButton
    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonRanked = New clsGraphicalButton
    Set cBotonArena = New clsGraphicalButton
    Set cBotonPlantes = New clsGraphicalButton
    Set cBotonRetos = New clsGraphicalButton
    
                                 
    Call cBotonCerrar.Initialize(cmdVolver, "27.gif", _
                                    "28.gif", _
                                    "29.gif", Me)
                                    
    Call cBotonRanked.Initialize(cmdIrAl, "45.gif", _
                                    "46.gif", _
                                    "47.gif", Me)
                                    
    Call cBotonRetos.Initialize(cmdOrganizar, "45.gif", _
                                    "46.gif", _
                                    "47.gif", Me)
                                    
    Call cBotonArena.Initialize(cmdIrA, "45.gif", _
                                    "46.gif", _
                                    "47.gif", Me)
                                    
    Call cBotonPlantes.Initialize(cmdEntrarAl, "45.gif", _
                                    "46.gif", _
                                    "47.gif", Me)
                                 

End Sub
