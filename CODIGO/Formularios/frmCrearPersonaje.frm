VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "NexusAO"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox headview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   3225
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   7740
      Width           =   555
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H80000012&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   3360
      List            =   "frmCrearPersonaje.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5475
      Width           =   2055
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":004C
      Left            =   3360
      List            =   "frmCrearPersonaje.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4485
      Width           =   2055
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H80000012&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":006F
      Left            =   3360
      List            =   "frmCrearPersonaje.frx":0071
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3180
      MaxLength       =   25
      TabIndex        =   2
      Top             =   2460
      Width           =   3585
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0073
      Left            =   3360
      List            =   "frmCrearPersonaje.frx":0075
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   6615
      Width           =   2055
   End
   Begin VB.Image Headmenos 
      Height          =   420
      Left            =   2760
      Top             =   7800
      Width           =   315
   End
   Begin VB.Image Headmas 
      Height          =   420
      Left            =   3900
      Top             =   7800
      Width           =   315
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
      Left            =   11250
      TabIndex        =   33
      Top             =   3240
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
      Left            =   11250
      TabIndex        =   32
      Top             =   3030
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
      Left            =   11250
      TabIndex        =   31
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
      Index           =   3
      Left            =   11250
      TabIndex        =   30
      Top             =   3660
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
      Left            =   11250
      TabIndex        =   29
      Top             =   3870
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
      Left            =   11250
      TabIndex        =   28
      Top             =   4080
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
      Left            =   11250
      TabIndex        =   27
      Top             =   4290
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
      Left            =   11250
      TabIndex        =   26
      Top             =   4500
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
      Left            =   11250
      TabIndex        =   25
      Top             =   4710
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
      Left            =   11250
      TabIndex        =   24
      Top             =   4920
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
      Left            =   11250
      TabIndex        =   23
      Top             =   5160
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
      Left            =   11250
      TabIndex        =   22
      Top             =   5400
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
      Left            =   11250
      TabIndex        =   21
      Top             =   5610
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
      Left            =   11250
      TabIndex        =   20
      Top             =   5820
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
      Left            =   11250
      TabIndex        =   19
      Top             =   6060
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
      Left            =   11250
      TabIndex        =   18
      Top             =   6270
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
      Left            =   11250
      TabIndex        =   17
      Top             =   6480
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
      Left            =   11250
      TabIndex        =   16
      Top             =   6690
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":0077
      Top             =   7320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":01C9
      Top             =   7320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":031B
      Top             =   7110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":046D
      Top             =   7110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":05BF
      Top             =   6900
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":0711
      Top             =   6900
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":0863
      Top             =   6690
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":09B5
      Top             =   6690
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":0B07
      Top             =   6480
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":0C59
      Top             =   6480
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":0DAB
      Top             =   6270
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":0EFD
      Top             =   6270
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":104F
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":11A1
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":12F3
      Top             =   5850
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":1445
      Top             =   5610
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":1597
      Top             =   5400
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":16E9
      Top             =   5160
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":183B
      Top             =   4950
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":198D
      Top             =   4740
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":1ADF
      Top             =   4530
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":1C31
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":1D83
      Top             =   4110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":1ED5
      Top             =   3900
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":2027
      Top             =   3690
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":2179
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":22CB
      Top             =   3240
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":241D
      Top             =   3030
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":256F
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":26C1
      Top             =   5850
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":2813
      Top             =   5640
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":2965
      Top             =   5400
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":2AB7
      Top             =   5160
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":2C09
      Top             =   4950
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":2D5B
      Top             =   4710
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":2EAD
      Top             =   4530
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":2FFF
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":3151
      Top             =   4110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":32A3
      Top             =   3900
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":33F5
      Top             =   3690
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":3547
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":3699
      Top             =   3240
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
      Left            =   11250
      TabIndex        =   15
      Top             =   7320
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
      Left            =   11250
      TabIndex        =   14
      Top             =   7110
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
      Left            =   11250
      TabIndex        =   13
      Top             =   6900
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
      Left            =   11250
      TabIndex        =   12
      Top             =   7530
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":37EB
      Top             =   7530
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":393D
      Top             =   7530
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
      Left            =   11250
      TabIndex        =   11
      Top             =   7755
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":3A8F
      Top             =   7770
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   45
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":3BE1
      Top             =   7770
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
      Left            =   11250
      TabIndex        =   10
      Top             =   7965
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":3D33
      Top             =   7980
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":3E85
      Top             =   7980
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
      Left            =   11250
      TabIndex        =   9
      Top             =   8175
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":3FD7
      Top             =   8190
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   49
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":4129
      Top             =   8190
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
      Left            =   11250
      TabIndex        =   8
      Top             =   8370
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":427B
      Top             =   8400
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   51
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":43CD
      Top             =   8385
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
      Left            =   11250
      TabIndex        =   7
      Top             =   8580
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   11580
      MouseIcon       =   "frmCrearPersonaje.frx":451F
      Top             =   8610
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   53
      Left            =   11010
      MouseIcon       =   "frmCrearPersonaje.frx":4671
      Top             =   8580
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
      Left            =   4320
      TabIndex        =   6
      Top             =   10830
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
      Left            =   12630
      TabIndex        =   1
      Top             =   2490
      Width           =   270
   End
   Begin VB.Image btnVolver 
      Height          =   525
      Left            =   810
      MouseIcon       =   "frmCrearPersonaje.frx":47C3
      MousePointer    =   99  'Custom
      Top             =   9780
      Width           =   2250
   End
   Begin VB.Image btnCrear 
      Height          =   525
      Left            =   12300
      MouseIcon       =   "frmCrearPersonaje.frx":4915
      MousePointer    =   99  'Custom
      Top             =   9810
      Width           =   2250
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

Private ModRaza()         As tModRaza

Private ModClase()        As tModClase

Public Actual             As Integer

Public SkillPoints        As Byte

Private MaxEleccion       As Integer, MinEleccion As Integer

Private botonCrear        As Boolean

Private cBotonVolver      As clsGraphicalButton

Private cBotonCrear       As clsGraphicalButton

Private cBotonCabezaMas   As clsGraphicalButton

Private cBotonCabezaMenos As clsGraphicalButton

Public LastButtonPressed  As clsGraphicalButton

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

Private Sub btnCrear_Click()
    Call Sound.Sound_Play(SND_CLICK)
            
    Dim Count As Byte

    Dim i     As Integer

    Dim k     As Object
            
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
End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

    Randomize Timer
    
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

    If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Sub btnVolver_Click()
    Call Sound.Sound_Play(SND_CLICK)

    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_VolverInicio
            Sound.Fading = 500

        End If

    End If
            
    Unload Me
            
    frmCharList.Visible = True

End Sub

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
    
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()

    Dim i As Byte
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Set cBotonVolver = New clsGraphicalButton
    Set cBotonCrear = New clsGraphicalButton
    Set cBotonCabezaMas = New clsGraphicalButton
    Set cBotonCabezaMenos = New clsGraphicalButton
    
    btnVolver.MouseIcon = picMouseIcon
    btnCrear.MouseIcon = picMouseIcon
    
    'Numero de command1
    For i = 0 To 53
        Command1(i).MouseIcon = picMouseIcon
    Next i
    
    Headmas.MouseIcon = picMouseIcon
    Headmenos.MouseIcon = picMouseIcon
    
    Call cBotonVolver.Initialize(btnVolver, "btnvolver_n.gif", _
                                 "btnvolver_h.gif", _
                                 "btnvolver_p.gif", Me)
                                 
    Call cBotonCrear.Initialize(btnCrear, "1.gif", _
                                 "3.gif", _
                                 "2.gif", Me)
                                 
    Call cBotonCabezaMas.Initialize(Headmas, "4.gif", _
                                 "6.gif", _
                                 "5.gif", Me)
                                 
    Call cBotonCabezaMenos.Initialize(Headmenos, "7.gif", _
                                 "9.gif", _
                                 "8.gif", Me)
                                 
                                 
                                 
End Sub

Private Sub Headmas_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    UserHead = CheckCabeza(UserHead + 1)
    
    If UserHead > 0 Then Call DrawHead(UserHead)
    
End Sub

Private Sub Headmenos_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    UserHead = CheckCabeza(UserHead - 1)
    
    If UserHead > 0 Then Call DrawHead(UserHead)
    
End Sub

Private Sub lstProfesion_Click()

    On Error Resume Next

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
