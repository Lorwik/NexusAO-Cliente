VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox headview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1695
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   34
      Top             =   4545
      Width           =   375
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D961
      Left            =   870
      List            =   "frmCrearPersonaje.frx":5D97A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3810
      Width           =   2055
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D9AD
      Left            =   840
      List            =   "frmCrearPersonaje.frx":5D9B7
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3150
      Width           =   2055
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D9D0
      Left            =   870
      List            =   "frmCrearPersonaje.frx":5D9D2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2490
      Width           =   2055
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1050
      Width           =   5865
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D9D4
      Left            =   8550
      List            =   "frmCrearPersonaje.frx":5D9D6
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3585
      Width           =   2745
   End
   Begin VB.Image imgNoDisp 
      Height          =   2145
      Left            =   8415
      Top             =   780
      Width           =   3045
   End
   Begin VB.Image Head 
      Height          =   600
      Index           =   0
      Left            =   1320
      Top             =   4440
      Width           =   390
   End
   Begin VB.Image Head 
      Height          =   600
      Index           =   1
      Left            =   2160
      Top             =   4440
      Width           =   390
   End
   Begin VB.Image imgClase 
      Height          =   3570
      Left            =   8490
      Stretch         =   -1  'True
      Top             =   4230
      Width           =   2835
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5310
      TabIndex        =   33
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   5310
      TabIndex        =   32
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5310
      TabIndex        =   31
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5310
      TabIndex        =   30
      Top             =   3450
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5310
      TabIndex        =   29
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5310
      TabIndex        =   28
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   5310
      TabIndex        =   27
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   5310
      TabIndex        =   26
      Top             =   4950
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   5310
      TabIndex        =   25
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   5310
      TabIndex        =   24
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   5310
      TabIndex        =   23
      Top             =   6090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   5310
      TabIndex        =   22
      Top             =   6450
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   5310
      TabIndex        =   21
      Top             =   6840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   5310
      TabIndex        =   20
      Top             =   7200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   7365
      TabIndex        =   19
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   7365
      TabIndex        =   18
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   7365
      TabIndex        =   17
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   7365
      TabIndex        =   16
      Top             =   3450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5D9D8
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DB2A
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DC7C
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DDCE
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DF20
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E072
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E1C4
      Top             =   3570
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E316
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E468
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E5BA
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E70C
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E85E
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E9B0
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5EB02
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5EC54
      Top             =   7170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5EDA6
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5EEF8
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F04A
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F19C
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F2EE
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F440
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F592
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F6E4
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F836
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F988
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FADA
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FC2C
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FD7E
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FED0
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60022
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60174
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":602C6
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60418
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":6056A
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":606BC
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":6080E
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60960
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60AB2
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60C04
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60D56
      Top             =   3540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60EA8
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60FFA
      Top             =   2790
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   7365
      TabIndex        =   15
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   7365
      TabIndex        =   14
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   7365
      TabIndex        =   13
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   7365
      TabIndex        =   12
      Top             =   4950
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":6114C
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":6129E
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   7365
      TabIndex        =   11
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":613F0
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   45
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61542
      Top             =   5430
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   7365
      TabIndex        =   10
      Top             =   5700
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61694
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":617E6
      Top             =   5790
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   7365
      TabIndex        =   9
      Top             =   6090
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61938
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   49
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61A8A
      Top             =   6180
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   7365
      TabIndex        =   8
      Top             =   6450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61BDC
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   51
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61D2E
      Top             =   6540
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   7365
      TabIndex        =   7
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61E80
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   53
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61FD2
      Top             =   6930
      Width           =   195
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2610
      TabIndex        =   6
      Top             =   8220
      Width           =   6795
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6795
      TabIndex        =   1
      Top             =   7260
      Width           =   270
   End
   Begin VB.Image boton 
      Height          =   615
      Index           =   1
      Left            =   720
      MouseIcon       =   "frmCrearPersonaje.frx":62124
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   1605
   End
   Begin VB.Image boton 
      Height          =   570
      Index           =   0
      Left            =   9600
      MouseIcon       =   "frmCrearPersonaje.frx":62276
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   1680
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.13.3
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private Type tModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DanoArmas As Double
    DanoProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double
End Type

Private ModRaza()  As tModRaza
Private ModClase() As tModClase

Public Actual As Integer
Public SkillPoints As Byte
Private MaxEleccion As Integer, MinEleccion As Integer

Private botonCrear As Boolean

Private Function CheckData() As Boolean
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Comprobacion antes de crear el PJ
'**************************************
    
    '¿Puso un nombre?
    If LenB(txtNombre.Text) = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_NOMBRE_PJ").item("TEXTO")
        txtNombre.SetFocus
        Exit Function
    End If

    '¿Selecciono una raza?
    If UserRaza = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_RAZA").item("TEXTO")
        Exit Function
    End If
    
    '¿Selecciono el Sexo?
    If UserSexo = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_SEXO").item("TEXTO")
        Exit Function
    End If
    
    '¿Seleciono la clase?
    If UserClase = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_CLASE").item("TEXTO")
        Exit Function
    End If

    '¿Estamos intentando crear sin tener el AccountName?
    If Len(AccountName) = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_HASH").item("TEXTO")
        Exit Function
    End If
    
    '¿El nombre de usuario supera los 30 caracteres?
    If LenB(UserName) > 30 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(1)
        Exit Function
    End If
    
    If UserHogar = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_HOGAR").item("TEXTO")
        Exit Function
    End If
    
    If SkillPoints > 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_SKILLS").item("TEXTO")
        Exit Function
    End If
    
    CheckData = True

End Function

Private Sub Boton_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    Select Case Index
    
        Case 0
            
            Dim Count   As Byte
            Dim i       As Integer
            Dim k       As Object
            
            i = 1
            For Each k In Skill
                UserSkills(i) = k.Caption
                i = i + 1
            Next
            
            'Nombre de usuario
            UserName = LTrim(txtNombre.Text)
                    
            '¿El nombre esta vacio y es correcto?
            If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(2))
                Exit Sub
            End If
            
            'Solo permitimos 1 espacio en los nombres
            For i = 1 To Len(UserName)
                If mid(UserName, i, 1) = Chr(32) Then Count = Count + 1
            Next i
            
            If Count > 1 Then
                Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(3))
                Exit Sub
            End If
            
            UserHogar = lstHogar.ListIndex + 1
            
            'Comprobamos que todo este OK
            If Not CheckData Then Exit Sub
            
            EstadoLogin = E_MODO.CrearNuevoPJ
            
            'Limpio la lista de hechizos
            frmMain.hlst.Clear
                
            'Conexion!!!
            If Not frmMain.Client.State = sckConnected Then
                Call MostrarMensaje(JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO"))
                Unload Me
            Else
                'Si ya mandamos el paquete, evitamos que se pueda volver a mandar
                botonCrear = True
                Call Login
                botonCrear = False
            End If
            
            'Mandamos el tutorial de inicio
            'bShowTutorial = True
            
        Case 1
            If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                    Sound.NextMusic = MUS_VolverInicio
                    Sound.Fading = 500
                End If
            End If
            
            Unload Me
            
            frmCharList.Visible = True

    End Select
End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

    Randomize Timer
    
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
    If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Sub Command1_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    Dim indice
    If (Index And &H1) = 0 Then
        If SkillPoints > 0 Then
            indice = Index \ 2
            Skill(indice).Caption = Val(Skill(indice).Caption) + 1
            SkillPoints = SkillPoints - 1
        End If
    Else
        If SkillPoints < 10 Then
            
            indice = Index \ 2
            If Val(Skill(indice).Caption) > 0 Then
                Skill(indice).Caption = Val(Skill(indice).Caption) - 1
                SkillPoints = SkillPoints + 1
            End If
        End If
    End If
    
    Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("cp-interface.bmp")
    
    Call LoadCharInfo
    
    SkillPoints = 10
    Puntos.Caption = SkillPoints
    
    Dim i As Integer
    
    lstProfesion.Clear
    For i = LBound(ListaClases) To UBound(ListaClases)
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstHogar.Clear
    
    For i = LBound(Ciudades()) To UBound(Ciudades())
        lstHogar.AddItem Ciudades(i)
    Next i
    
    lstRaza.Clear
    
    For i = LBound(ListaRazas()) To UBound(ListaRazas())
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    lstProfesion.Clear
    
    For i = LBound(ListaClases()) To UBound(ListaClases())
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstProfesion.ListIndex = 1
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserHead = 0
    
End Sub

Private Sub Head_Click(Index As Integer)
    
    Call Sound.Sound_Play(SND_CLICK)
    
    Select Case Index
    
        Case 0
            UserHead = CheckCabeza(UserHead - 1)

        Case 1
            UserHead = CheckCabeza(UserHead + 1)
    
    End Select
    
    If UserHead > 0 Then Call DrawHead(UserHead)

End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    imgClase.Picture = General_Load_Picture_From_Resource(LCase(lstProfesion.Text & ".bmp"))
    
    UserClase = lstProfesion.ListIndex + 1
    
End Sub

Private Sub lstGenero_Click()
    UserSexo = lstGenero.ListIndex + 1
    Call DameCabezas
    
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.ListIndex + 1
    Call DameCabezas
    
End Sub

Sub DameCabezas()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Asignamos un cuerpo y unac abeza segun la raza y el sexo
'**************************************

    Select Case UserSexo
    
        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_H_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_H_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_H_CUERPO_DESNUDO
                    
                Case eRaza.Orco
                    UserHead = eCabezas.ORCO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ORCO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_M_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_M_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_M_CUERPO_DESNUDO
                    
                Case eRaza.Orco
                    UserHead = eCabezas.ORCO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ORCO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case Else
            UserHead = 0
            UserBody = 0
            
    End Select
    
    If UserHead > 0 Then Call DrawHead(UserHead)
    
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

On Error GoTo errhandler

    Select Case UserSexo

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano

                    If Head > eCabezas.HUMANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_H_PRIMER_CABEZA + (Head - eCabezas.HUMANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.HUMANO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_H_ULTIMA_CABEZA - (eCabezas.HUMANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo

                    If Head > eCabezas.ELFO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_H_PRIMER_CABEZA + (Head - eCabezas.ELFO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ELFO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_H_ULTIMA_CABEZA - (eCabezas.ELFO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > eCabezas.DROW_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.DROW_H_PRIMER_CABEZA + (Head - eCabezas.DROW_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.DROW_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.DROW_H_ULTIMA_CABEZA - (eCabezas.DROW_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano

                    If Head > eCabezas.ENANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_H_PRIMER_CABEZA + (Head - eCabezas.ENANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ENANO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_H_ULTIMA_CABEZA - (eCabezas.ENANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo

                    If Head > eCabezas.GNOMO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_H_PRIMER_CABEZA + (Head - eCabezas.GNOMO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.GNOMO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_H_ULTIMA_CABEZA - (eCabezas.GNOMO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Orco

                    If Head > eCabezas.ORCO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_H_PRIMER_CABEZA + (Head - eCabezas.ORCO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ORCO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_H_ULTIMA_CABEZA - (eCabezas.ORCO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    CheckCabeza = CheckCabeza(Head)
                    
            End Select
        
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano

                    If Head > eCabezas.HUMANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_M_PRIMER_CABEZA + (Head - eCabezas.HUMANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.HUMANO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_M_ULTIMA_CABEZA - (eCabezas.HUMANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo

                    If Head > eCabezas.ELFO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_M_PRIMER_CABEZA + (Head - eCabezas.ELFO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ELFO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_M_ULTIMA_CABEZA - (eCabezas.ELFO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > eCabezas.DROW_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.DROW_M_PRIMER_CABEZA + (Head - eCabezas.DROW_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.DROW_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.DROW_M_ULTIMA_CABEZA - (eCabezas.DROW_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano

                    If Head > eCabezas.ENANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_M_PRIMER_CABEZA + (Head - eCabezas.ENANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ENANO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_M_ULTIMA_CABEZA - (eCabezas.ENANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo

                    If Head > eCabezas.GNOMO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_M_PRIMER_CABEZA + (Head - eCabezas.GNOMO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.GNOMO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_M_ULTIMA_CABEZA - (eCabezas.GNOMO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Orco

                    If Head > eCabezas.ORCO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_M_PRIMER_CABEZA + (Head - eCabezas.ORCO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ORCO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_M_ULTIMA_CABEZA - (eCabezas.ORCO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    CheckCabeza = Head
                    
            End Select

        Case Else
            CheckCabeza = Head
            
    End Select
    
errhandler:

    If Err.number Then
        Call LogError(Err.number, Err.Description, "frmCrearPersonaje.CheckCabeza")
    End If
    
    Exit Function
    
End Function

Private Sub LoadCharInfo()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Carga los modificadores de cada raza
'**************************************

    Dim SearchVar As String
    Dim i         As Integer

    ReDim ModRaza(1 To NUMRAZAS)

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(Carga.Path(Lenguajes) & "CharInfo_" & Language & ".dat")
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
    
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", vbNullString)
        
            .Fuerza = CSng(Lector.GetValue("MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = CSng(Lector.GetValue("MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = CSng(Lector.GetValue("MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = CSng(Lector.GetValue("MODRAZA", SearchVar + "Carisma"))
            .Constitucion = CSng(Lector.GetValue("MODRAZA", SearchVar + "Constitucion"))
        End With
        
    Next i

End Sub

Private Sub DrawHead(ByVal Head As Integer)

    Dim DR  As RECT
    Dim Grh As Long

    Grh = HeadData(Head).Head(3).GrhIndex
    
    With headview
        DR.Right = .Width - 5
        DR.Bottom = .Height - 3
        DR.Left = -5
        DR.Top = -3
    End With
        
    Call DrawGrhtoHdc(headview, Grh, DR)

End Sub

Private Sub txtNombre_GotFocus()
    lblInfo.Caption = "Sea cuidadoso al seleccionar el nombre de su personaje, NexusAO es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."

End Sub
