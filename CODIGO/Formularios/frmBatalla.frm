VERSION 5.00
Begin VB.Form frmBatalla 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9330
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   7200
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
   ScaleHeight     =   9330
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraDagaRusa 
      Caption         =   "Duelo de Plantes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   3480
      TabIndex        =   26
      Top             =   1890
      Width           =   3615
      Begin VB.CommandButton cmdEntrarAl 
         Caption         =   "Entrar al juego"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   900
         TabIndex        =   28
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label lblLabel2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enfrenta a tu adversario cara a cara, sin moverte, solo a golpes, y gana recompensas! (No se puede usar hechizos)"
         Height          =   615
         Left            =   210
         TabIndex        =   27
         Top             =   270
         Width           =   3165
      End
   End
   Begin VB.Frame FraRetos 
      Caption         =   "Retos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   3480
      TabIndex        =   23
      Top             =   180
      Width           =   3615
      Begin VB.CommandButton cmdRetos 
         Caption         =   "Organizar"
         Height          =   360
         Left            =   840
         TabIndex        =   25
         Top             =   870
         Width           =   1875
      End
      Begin VB.Label lblDuelosDe 
         BackStyle       =   0  'Transparent
         Caption         =   "Pelea solo o con compañeros en duelos de 1v1, 2v2 y 3v3 por oro."
         Height          =   405
         Left            =   150
         TabIndex        =   24
         Top             =   330
         Width           =   3240
      End
   End
   Begin VB.Frame FraArenaDe 
      Caption         =   "Arena de la Muerte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   3255
      Begin VB.CommandButton cmdIrA 
         Caption         =   "Ir a la Arena"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   570
         TabIndex        =   22
         Top             =   1050
         Width           =   1875
      End
      Begin VB.Label lblYtyhyhr 
         BackStyle       =   0  'Transparent
         Caption         =   "Tendrás que sobrevivir a oleadas de criaturas para conseguir el tesoro del desierto."
         Height          =   555
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2985
      End
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   18
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ranking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   6975
      Begin VB.Frame Frame5 
         Caption         =   "Tabla de Clasificaciones         "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   6495
         Begin VB.Label Label4 
            Caption         =   "< 1900 Diamante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1560
            Width           =   5775
         End
         Begin VB.Label Label9 
            Caption         =   "<= 1700 > 1900 Platino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   5775
         End
         Begin VB.Label Label8 
            Caption         =   "<= 1500 > 1700 Oro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   6015
         End
         Begin VB.Label Label7 
            Caption         =   "<= 1300 > 1500 Plata"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   5775
         End
         Begin VB.Label Label6 
            Caption         =   "<= 1100 > 1300 Bronce"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   5895
         End
         Begin VB.Label Label5 
            Caption         =   "1100 < Madera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   5655
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Top 5 ELO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   6495
         Begin VB.Label TopELO 
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
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label TopELO 
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
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label TopELO 
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
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label TopELO 
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
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label TopELO 
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Label lblELOUser 
         Caption         =   "Tu ELO es de 1000 puntos, estas en la clasificación Madera."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label Label3 
         Caption         =   $"frmBatalla.frx":0000
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Torneos Ranked"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdIrDuelo 
         Caption         =   "Ir al Duelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   510
         TabIndex        =   19
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Duela contra otro usuario al mejor de 3, mejora tu ELO y gana recompensas!"
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2925
      End
   End
End
Attribute VB_Name = "frmBatalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
        
    lblELOUser.Caption = "Tu ELO es de " & UserELO & " estas en la clasificacion " & UserClasificacion
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

