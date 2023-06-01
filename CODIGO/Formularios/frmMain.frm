VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   15
   ClientTop       =   -3300
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H00008080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004080&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9165
      Left            =   105
      ScaleHeight     =   611
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2250
      Width           =   10800
      Begin VB.Frame FrameMenu 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   9330
         TabIndex        =   19
         Top             =   6210
         Visible         =   0   'False
         Width           =   1425
         Begin NexusAO_Client.uAOButton cmdGrupo 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Grupo"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":10CA
            PICF            =   "frmMain.frx":10E6
            PICH            =   "frmMain.frx":1102
            PICV            =   "frmMain.frx":111E
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAO_Client.uAOButton cmdEstadisticas 
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   450
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Estadisticas"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":113A
            PICF            =   "frmMain.frx":1156
            PICH            =   "frmMain.frx":1172
            PICV            =   "frmMain.frx":118E
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAO_Client.uAOButton cmdClanes 
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   780
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Clanes"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":11AA
            PICF            =   "frmMain.frx":11C6
            PICH            =   "frmMain.frx":11E2
            PICV            =   "frmMain.frx":11FE
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAO_Client.uAOButton cmdQuest 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   1110
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Quest"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":121A
            PICF            =   "frmMain.frx":1236
            PICH            =   "frmMain.frx":1252
            PICV            =   "frmMain.frx":126E
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAO_Client.uAOButton cmdTorneos 
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   1440
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Torneos"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":128A
            PICF            =   "frmMain.frx":12A6
            PICH            =   "frmMain.frx":12C2
            PICV            =   "frmMain.frx":12DE
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAO_Client.uAOButton cmdOpciones 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   1770
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Opciones"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":12FA
            PICF            =   "frmMain.frx":1316
            PICH            =   "frmMain.frx":1332
            PICV            =   "frmMain.frx":134E
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAO_Client.uAOButton cmdCerrar 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   2100
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Desconectar"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":136A
            PICF            =   "frmMain.frx":1386
            PICH            =   "frmMain.frx":13A2
            PICV            =   "frmMain.frx":13BE
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAO_Client.uAOButton UAOCerrarMenú 
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Top             =   2520
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            TX              =   "Cerrar Menú"
            ENAB            =   -1  'True
            FCOL            =   7314354
            OCOL            =   16777215
            PICE            =   "frmMain.frx":13DA
            PICF            =   "frmMain.frx":13F6
            PICH            =   "frmMain.frx":1412
            PICV            =   "frmMain.frx":142E
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   150
      MaxLength       =   500
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1860
      Visible         =   0   'False
      Width           =   8580
   End
   Begin VB.PictureBox MiniMapa 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   9360
      MouseIcon       =   "frmMain.frx":144A
      ScaleHeight     =   97
      ScaleMode       =   0  'User
      ScaleWidth      =   97
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   180
      Width           =   1455
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   45
         Left            =   750
         Top             =   750
         Width           =   45
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   11670
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   2535
      Width           =   3000
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3540
      Left            =   11850
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   2685
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1740
      Left            =   150
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   60
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   3069
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":159C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image cmdPVP 
      Height          =   480
      Left            =   13200
      Top             =   10830
      Width           =   1725
   End
   Begin VB.Image imgSafe 
      Height          =   375
      Index           =   0
      Left            =   12030
      Top             =   10140
      Width           =   375
   End
   Begin VB.Image imgSafe 
      Height          =   375
      Index           =   1
      Left            =   12585
      Top             =   10155
      Width           =   375
   End
   Begin VB.Image cmdMenu 
      Height          =   480
      Left            =   11400
      Top             =   10830
      Width           =   1725
   End
   Begin VB.Label MapExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
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
      Height          =   135
      Index           =   1
      Left            =   11970
      TabIndex        =   18
      Top             =   960
      Width           =   2805
   End
   Begin VB.Shape ExpShp 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   11940
      Top             =   990
      Width           =   2835
   End
   Begin VB.Label MapExp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   9210
      TabIndex        =   17
      Top             =   1890
      Width           =   1725
   End
   Begin VB.Label lblFU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   13560
      TabIndex        =   16
      Top             =   10260
      Width           =   240
   End
   Begin VB.Label lblAG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14220
      TabIndex        =   15
      Top             =   10260
      Width           =   240
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   12630
      TabIndex        =   14
      Top             =   8115
      Width           =   1095
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   11880
      TabIndex        =   13
      Top             =   8640
      Width           =   2505
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   12480
      TabIndex        =   12
      Top             =   9165
      Width           =   1365
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   11940
      TabIndex        =   11
      Top             =   9675
      Width           =   1095
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   13350
      TabIndex        =   10
      Top             =   9690
      Width           =   1095
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   12210
      MousePointer    =   99  'Custom
      Top             =   6690
      Width           =   300
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   13290
      Top             =   9690
      Width           =   1155
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   11955
      Top             =   9690
      Width           =   1095
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   11910
      Top             =   9180
      Width           =   2535
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   11910
      Top             =   8655
      Width           =   2535
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   11910
      Top             =   8145
      Width           =   2535
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11535
      TabIndex        =   9
      Top             =   870
      Width           =   345
   End
   Begin VB.Image lblChat 
      Height          =   270
      Left            =   8805
      Top             =   1845
      Width           =   300
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11910
      TabIndex        =   8
      Top             =   150
      Width           =   2565
   End
   Begin VB.Image btnInfo 
      Height          =   465
      Left            =   13620
      MouseIcon       =   "frmMain.frx":1619
      MousePointer    =   99  'Custom
      Top             =   6480
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image btnLanzar 
      Height          =   465
      Left            =   11730
      MouseIcon       =   "frmMain.frx":176B
      MousePointer    =   99  'Custom
      Top             =   6480
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image imgSolapaHech 
      Height          =   420
      Left            =   13290
      MouseIcon       =   "frmMain.frx":18BD
      Top             =   1860
      Width           =   1545
   End
   Begin VB.Image imgSolapaInv 
      Height          =   420
      Left            =   11550
      MouseIcon       =   "frmMain.frx":1A0F
      Top             =   1860
      Width           =   1545
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   14565
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Image cmdMoverHechiDown 
      Height          =   420
      Left            =   14580
      MouseIcon       =   "frmMain.frx":1B61
      MousePointer    =   99  'Custom
      Top             =   4470
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cmdMoverHechiTop 
      Height          =   420
      Left            =   14580
      MouseIcon       =   "frmMain.frx":1CB3
      MousePointer    =   99  'Custom
      Top             =   4050
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label GldLbl 
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
      Height          =   165
      Left            =   12840
      TabIndex        =   0
      Top             =   6750
      Width           =   1440
   End
   Begin VB.Image InvEqu 
      Height          =   4770
      Left            =   11520
      Top             =   2400
      Width           =   3360
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmMain
'    Project    : ARGENTUM
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Public tX                  As Byte
Public tY                  As Byte
Public MouseX              As Long
Public MouseY              As Long
Public MouseBoton          As Long
Public MouseShift          As Long
Private clicX              As Long
Private clicY              As Long

Public UltPos As Integer

Private clsFormulario           As clsFormMovementManager
Private cBotonInventario        As clsGraphicalButton
Private cBotonHechizos          As clsGraphicalButton
Private cBotonLanzar            As clsGraphicalButton
Private cBotonInfo              As clsGraphicalButton
Private cBotonBubble            As clsGraphicalButton
Private cBotonMenu              As clsGraphicalButton
Private cBotonPVP               As clsGraphicalButton
Private cBotonMoverHechiTop     As clsGraphicalButton
Private cBotonMoverHechiDown    As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Public WithEvents Client   As clsSocket
Attribute Client.VB_VarHelpID = -1

Private FirstTimeChat      As Boolean

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn             As Boolean

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Sub btnGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub cmdCerrar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    If UserParalizado Then 'Inmo

        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_SALIR").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
        End With
        
        Exit Sub
        
    End If
    
    ' Nos desconectamos y lo mando al Panel de la Cuenta
    Call WriteQuit
End Sub

Private Sub cmdClanes_Click()
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub cmdMoverHechiDown_Click()

    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub

        Dim sTemp As String

        If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

        Call WriteMoveSpell(False, hlst.ListIndex + 1)

        sTemp = hlst.List(hlst.ListIndex + 1)
        hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
        hlst.List(hlst.ListIndex) = sTemp
        hlst.ListIndex = hlst.ListIndex + 1

    End If

End Sub

Private Sub cmdMoverHechiTop_Click()

    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub

        Dim sTemp As String

        If hlst.ListIndex = 0 Then Exit Sub
    
        Call WriteMoveSpell(True, hlst.ListIndex + 1)
        
        sTemp = hlst.List(hlst.ListIndex - 1)
        hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
        hlst.List(hlst.ListIndex) = sTemp
        hlst.ListIndex = hlst.ListIndex - 1

    End If

End Sub

Private Sub cmdPVP_Click()
    LlegoRank = False
    Call WriteSolicitarRank
    Call FlushBuffer
            
    Do While Not LlegoRank
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmBatalla.Iniciar_Labels
    frmBatalla.Show , frmMain
    LlegoRank = False
End Sub

Private Sub imgSolapaInv_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    InvEqu.Picture = General_Load_Picture_From_Resource("18.bmp", True)
            
    ' Activo controles de inventario
    picInv.Visible = True
    cmdDropGold.Visible = True
    GldLbl.Visible = True
        
    ' Desactivo controles de hechizo
    hlst.Visible = False
    btnInfo.Visible = False
    btnLanzar.Visible = False
            
    cmdMoverHechiTop.Visible = False
    cmdMoverHechiDown.Visible = False
            
    DoEvents
    Call Inventario.DrawInventory
End Sub

Private Sub imgSolapaHech_Click()
    
    Call Sound.Sound_Play(SND_CLICK)
    InvEqu.Picture = General_Load_Picture_From_Resource("17.bmp", True)
            
    ' Activo controles de hechizos
    hlst.Visible = True
    btnInfo.Visible = True
    btnLanzar.Visible = True
            
    cmdMoverHechiTop.Visible = True
    cmdMoverHechiDown.Visible = True
            
    ' Desactivo controles de inventario
    picInv.Visible = False
    cmdDropGold.Visible = False
    GldLbl.Visible = False
            
End Sub

Private Sub cmdDropGold_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        frmCantidad.Show , frmMain
    End If
End Sub

Private Sub cmdGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub cmdMenu_Click()
    FrameMenu.Visible = Not FrameMenu.Visible
End Sub

Private Sub Form_Activate()

    Call Inventario.DrawInventory

End Sub

Private Sub Form_Load()
    ClientSetup.SkinSeleccionado = GetVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "SkinSelected")
    
    Me.Picture = General_Load_Picture_From_Resource("1.bmp", True)
    InvEqu.Picture = General_Load_Picture_From_Resource("18.bmp", True)
    
    If Not ResolucionCambiada Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        Call clsFormulario.Initialize(Me, 120)
    End If
        
    Call LoadButtons
    
    ' Seteamos el caption
    Me.Caption = Form_Caption
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)

    ' Reseteamos el tamanio de la ventana para que no queden bordes blancos
    Me.Width = 15360
    Me.Height = 11520
        
    ' Detect links in console
    Call EnableURLDetect(RecTxt.hWnd, Me.hWnd)
    
    ' Make the console transparent
    Call SetWindowLong(RecTxt.hWnd, -20, &H20&)
    RecTxt.BackColor = RGB(24, 23, 21)
    
    CtrlMaskOn = False
    
    FirstTimeChat = True
    SendingType = 1
    
    UltPos = -1
    
End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    Dim i As Integer

    Set LastButtonPressed = New clsGraphicalButton
    
    lblMinimizar.MouseIcon = picMouseIcon
    
    Set cBotonInventario = New clsGraphicalButton
    Set cBotonHechizos = New clsGraphicalButton
    Set cBotonLanzar = New clsGraphicalButton
    Set cBotonInfo = New clsGraphicalButton
    Set cBotonBubble = New clsGraphicalButton
    Set cBotonMenu = New clsGraphicalButton
    Set cBotonPVP = New clsGraphicalButton
    Set cBotonMoverHechiTop = New clsGraphicalButton
    Set cBotonMoverHechiDown = New clsGraphicalButton
    
                                 
    Call cBotonLanzar.Initialize(imgSolapaInv, "2.bmp", _
                                 "3.bmp", _
                                 "4.bmp", Me, , , , , True)
                                 
    Call cBotonLanzar.Initialize(imgSolapaHech, "5.bmp", _
                                 "6.bmp", _
                                 "7.bmp", Me, , , , , True)
    
    Call cBotonLanzar.Initialize(btnLanzar, "8.bmp", _
                                 "9.bmp", _
                                 "10.bmp", Me, , , , , True)
                                     
    Call cBotonInfo.Initialize(btnInfo, "11.bmp", _
                               "12.bmp", _
                               "13.bmp", Me, , , , , True)
                                                           
    Call cBotonBubble.Initialize(lblChat, "23.bmp", _
                               "24.bmp", _
                               "25.bmp", Me, , , , , True)
                               
    Call cBotonMenu.Initialize(cmdMenu, "14.gif", _
                               "15.gif", _
                               "16.gif", Me, , , , , True)
                               
    Call cBotonPVP.Initialize(cmdPVP, "17.gif", _
                               "18.gif", _
                               "19.gif", Me, , , , , True)
                               
    Call cBotonMoverHechiTop.Initialize(cmdMoverHechiTop, "flechatop_n.gif", _
                                        "flechatop_h.gif", _
                                        "flechatop_p.gif", Me, , , , , True)
                                        
    Call cBotonMoverHechiDown.Initialize(cmdMoverHechiDown, "flechadown_n.gif", _
                                        "flechadown_h.gif", _
                                        "flechadown_p.gif", Me, , , , , True)
                                
    For i = 0 To 1
        imgSafe(i).MouseIcon = picMouseIcon
    Next i

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2010
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '18/11/2010: Amraphen - Agregue el handle correspondiente para las nuevas configuraciones de teclas (CTRL+0..9).
    '***************************************************
    If (Not SendTxt.Visible) Then
        
        If KeyCode = vbKeyControl Then

            'Chequeo que no se haya usado un CTRL + tecla antes de disparar las bindings.
            If CtrlMaskOn Then
                CtrlMaskOn = False
                Exit Sub
            End If
        End If
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If ClientSetup.bMusic = CONST_MP3 Then
                        Sound.Music_Stop
                        ClientSetup.bMusic = CONST_DESHABILITADA
                    Else
                        ClientSetup.bMusic = CONST_MP3
                    End If
                        
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    'Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
                    ClientSetup.FPSShow = Not ClientSetup.FPSShow
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatSafe)
                    Call WriteCombatToggle
                    
            End Select
            
        End If
        
    
        Select Case KeyCode
            
            Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                Call Mod_General.Client_Screenshot(frmMain.hDC, 1024, 768)
                    
            Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
                Call frmOpciones.Show(vbModeless, frmMain)
                
            Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
    
                Call WriteQuit
                
            Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
    
                If Shift <> 0 Then Exit Sub
                
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
    
                    If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
                End If
                
                If frmCustomKeys.Visible Then Exit Sub 'Chequeo si esta visible la ventana de configuracion de teclas.
                
                Call WriteAttack
                
            Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                
                If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                    Call CompletarEnvioMensajes
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                Else
                    Call Enviar_SendTxt
                End If
            
        End Select
     End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DisableURLDetect
    
End Sub

Private Sub btnClanes_Click()
    
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub cmdEstadisticas_Click()

    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    LlegoFamily = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call WriteRequestFamily
    Call FlushBuffer

    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    
    Alocados = SkillPoints
    frmEstadisticas.lblLibres.Caption = SkillPoints
    
    Call frmEstadisticas.MostrarAsignacion
    
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    
End Sub

Private Sub btnMapa_Click()
    
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub imgsafe_Click(Index As Integer)

    Select Case Index

        Case eSMType.sSafemode
            Call WriteSafeToggle
            
        Case eSMType.sCombatmode
            Call WriteCombatToggle
            
    End Select
    
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub LbLChat_Click()
    frmMensaje.PopupMenuMensaje
End Sub

Private Sub lblMana_Click()

   Call ParseUserCommand("/MEDITAR")
End Sub

Private Sub cmdOpciones_Click()
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub lblMP_Click()
    Call ParseUserCommand("/MEDITAR")
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(3), _
                        False, False, True)
End Sub

Private Sub MapExp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

     
     
    With charlist(UserCharIndex)
    
        If UltPos <> Index Then
        
            If UltPos >= 0 Then
                If Index = 1 Then
                    MapExp(Index).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel)) & "%"
                    
                Else
                
                    If ClientSetup.VerLugar = 1 Then
                        MapExp(Index).Caption = mapInfo.name
                        
                    Else
                        MapExp(Index).Caption = "Posición: " & UserMap & ", " & .Pos.X & "  " & .Pos.Y
                    
                    End If

                End If
            End If
            
    
            If Index = 1 Then
                MapExp(Index).Caption = UserExp & "/" & UserPasarNivel
                
            Else

                If ClientSetup.VerLugar = 1 Then
                    MapExp(Index).Caption = mapInfo.name
                        
                Else
                    MapExp(Index).Caption = "Posición: " & UserMap & ", " & .Pos.X & "  " & .Pos.Y
                    
                End If
            End If
            
            If UserPasarNivel = 0 Then
                MapExp(Index).Caption = "¡Nivel máximo!"
            End If
                
            UltPos = Index
        End If
        
    End With
    
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    StartCheckingLinks
End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Para borrar el mensaje de fondo
    If FirstTimeChat Then
        SendTxt.Text = vbNullString
        FirstTimeChat = False
        ' Cambiamos el color de texto al original
        SendTxt.ForeColor = &HE0E0E0
    End If
    
errhandler:
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else

        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else

                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()

    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call WriteUseItem(Inventario.SelectedItem)
    End If
    
End Sub

Private Sub EquiparItem()

    If UserEstado = 1 Then
    
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
        
    Else
    
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
            Call WriteEquipItem(Inventario.SelectedItem)
        End If
        
    End If
End Sub

Private Sub btnLanzar_Click()
    
    If hlst.List(hlst.ListIndex) <> JsonLanguage.item("NADA").item("TEXTO") And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
    
End Sub

Private Sub btnLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub btnInfo_Click()
    
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
    
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
    '¿Hizo click derecho?
    If Button = 2 Then
        If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
            Call WriteAccionClick(tX, tY)
        End If
    End If
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub MainViewPic_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteAccionClick(tX, tY)
    End If
    
End Sub

Private Sub MainViewPic_Click()

    If Cartel Then Cartel = False
    
    Dim MENSAJE_ADVERTENCIA As String
    Dim VAR_LANZANDO        As String
    
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If Not InGameArea() Then Exit Sub
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then

                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1

                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If

                'Invitando party
                If InvitandoParty = True Then
                    frmMain.MousePointer = vbDefault
                    Call WriteInvitarPartyClick(tX, tY)
                    InvitandoParty = False
                    Exit Sub
                End If
    
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        'frmMain.MousePointer = vbDefault
                        'UsingSkill = 0

                        'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            'VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                            'MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                            'MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                            
                            'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                        'End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            'frmMain.MousePointer = vbDefault
                            'UsingSkill = 0

                            'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                'VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                                'MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                'MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                
                                'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                            'End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0

                                'LwK: ¿Poner aqui el bloqueo del cursor si no paso el intervalo del hechizo?
                                Exit Sub
                            End If
                        Else

                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0

                                'LwK: ¿Poner aqui el bloqueo del cursor si no paso el intervalo del hechizo?
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                'Call WriteRightClick(tx, tY) 'Proximamnete lo implementaremos..
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then

            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteAccionClick(tX, tY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewPic.Left
    MouseY = Y - MainViewPic.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewPic.Width Then
        MouseX = MainViewPic.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewPic.Height Then
        MouseY = MainViewPic.Height
    End If
    
    LastButtonPressed.ToggleToNormal
    
    ' Disable links checking (not over consola)
    StopCheckingLinks
    
    If UltPos >= 0 Then
        If UserPasarNivel = 0 Then
            frmMain.MapExp(1).Caption = "¡Nivel máximo!"
            
        Else
            If UltPos = 1 Then
                MapExp(UltPos).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel)) & "%"
                
            Else
    
                If ClientSetup.VerLugar = 1 Then
                    MapExp(UltPos).Caption = mapInfo.name
                    
                Else
                    MapExp(0).Caption = "Posición: " & UserMap & ", " & charlist(UserCharIndex).Pos.X & "  " & charlist(UserCharIndex).Pos.Y
                
                End If
    
            End If
            
            UltPos = -1
        End If
    End If
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

    Inventario.SelectGold

    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub picInv_DblClick()
'**********************************************
'Autor: Lorwik
'Fecha: 14/07/2020
'Descripcion: DobleClick sobre el inventario
'**********************************************
    'Esta validacion es para que el juego no rompa si hacemos doble click
    If MirandoTrabajo > 0 Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    '¿Es un slot valido?
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call WriteAccionInventario(Inventario.SelectedItem)
    End If
    
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Sound.Sound_Play(SND_CLICK)
End Sub

Private Sub RecTxt_Change()

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    
    ElseIf (Not Comerciando) And _
           (Not MirandoAsignarSkills) And _
           (Not frmMSG.Visible) And _
           (Not MirandoForo) And _
           (Not frmEstadisticas.Visible) And _
           (Not frmCantidad.Visible) And _
           (Not MirandoParty) Then

        If picInv.Visible Then
            picInv.SetFocus
                        
        ElseIf hlst.Visible Then
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedi se inserten caracteres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = JsonLanguage.item("MENSAJE_SOY_CHEATER").item("TEXTO")
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long
        Dim tempstr   As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub CompletarEnvioMensajes()

    Select Case SendingType
        Case 1
            SendTxt.Text = vbNullString
        Case 2
            SendTxt.Text = "-"
        Case 3
            SendTxt.Text = ("\" & sndPrivateTo & " ")
        Case 4
            SendTxt.Text = "/CMSG "
        Case 5
            SendTxt.Text = "/PMSG "
        Case 6
            SendTxt.Text = "; "
    End Select
    
    stxtbuffer = SendTxt.Text
    SendTxt.SelStart = Len(SendTxt.Text)

End Sub

Private Sub Enviar_SendTxt()
    
    Dim str1 As String
    Dim str2 As String
    
    If Len(stxtbuffer) > 255 Then stxtbuffer = mid$(stxtbuffer, 1, 255)
    
    'Send text
    If Left$(stxtbuffer, 1) = "/" Then
        Call ParseUserCommand(stxtbuffer)

    'Shout
    ElseIf Left$(stxtbuffer, 1) = "-" Then
        If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> vbNullString Then Call ParseUserCommand(stxtbuffer)
        SendingType = 2
        
    'Global
    ElseIf Left$(stxtbuffer, 1) = ";" Then
        If LenB(Right$(stxtbuffer, Len(stxtbuffer) - 1)) > 0 And InStr(stxtbuffer, ">") = 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 6

    'Privado
    ElseIf Left$(stxtbuffer, 1) = "\" Then
        str1 = Right$(stxtbuffer, Len(stxtbuffer) - 1)
        str2 = ReadField(1, str1, 32)
        If LenB(str1) > 0 And InStr(str1, ">") = 0 Then Call ParseUserCommand("\" & str1)
        sndPrivateTo = str2
        SendingType = 3
                
    'Say
    Else
        If LenB(stxtbuffer) > 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 1
    End If

    stxtbuffer = vbNullString
    SendTxt.Text = vbNullString
    SendTxt.Visible = False
    
End Sub

Private Sub AbrirMenuViewPort()
    'TODO: No usar variable de compilacion y acceder a esto desde el config.ini
    #If (ConMenuseConextuales = 1) Then

        If tX >= MinXBorder And tY >= MinYBorder And tY <= MaxYBorder And tX <= MaxXBorder Then

            If MapData(tX, tY).CharIndex > 0 Then
                If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
                    Dim m As frmMenuseFashion
                    Set m = New frmMenuseFashion
            
                    Load m
                    m.SetCallback Me
                    m.SetMenuId 1
                    m.ListaInit 2, False
            
                    If LenB(charlist(MapData(tX, tY).CharIndex).Nombre) <> 0 Then
                        m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
                    Else
                        m.ListaSetItem 0, "<NPC>", True
                    End If
                    m.ListaSetItem 1, JsonLanguage.item("COMERCIAR").item("TEXTO")
            
                    m.ListaFin
                    m.Show , Me

                End If
            End If
        End If

    #End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)

    Select Case MenuId

        Case 0 'Inventario

            Select Case Sel

                Case 0

                Case 1

                Case 2 'Tirar
                    Call TirarItem

                Case 3 'Usar

                    If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
                        Call UsarItem
                    End If

                Case 3 'equipar
                    Call EquiparItem
            End Select
    
        Case 1 'FrameMenu del ViewPort del engine

            Select Case Sel

                Case 0 'Nombre
                    Call WriteLeftClick(tX, tY)
        
                Case 1 'Comerciar
                    Call WriteLeftClick(tX, tY)
                    Call WriteCommerceStart
            End Select
    End Select
End Sub
 
''''''''''''''''''''''''''''''''''''''
'     WINDOWS API                    '
''''''''''''''''''''''''''''''''''''''
Private Sub Client_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    Select Case EstadoLogin

        Case E_MODO.CrearNuevoPJ, E_MODO.Normal
            Call Login

        Case E_MODO.Dados
            frmCrearPersonaje.Show
        
    End Select
 
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim RD     As String
    Dim Data() As Byte
    
    Client.GetData RD, vbByte, bytesTotal
    Data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
    
End Sub

Private Sub Client_CloseSck()
    
    Debug.Print "Cerrando la conexion via API de Windows..."

    If frmMain.Visible = True Then frmMain.Visible = False
    Call ResetAllInfo
    Mod_Declaraciones.Conectando = True
    frmConnect.Show
End Sub

Private Sub Client_Error(ByVal number As Integer, _
                         Description As String, _
                         ByVal sCode As Long, _
                         ByVal Source As String, _
                         ByVal HelpFile As String, _
                         ByVal HelpContext As Long, _
                         CancelDisplay As Boolean)
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    
    frmConnect.MousePointer = 1
 
    If Client.State <> sckClosed Then Client.CloseSck
    
    Mod_Declaraciones.Conectando = True
    frmConnect.Show
 
End Sub

Private Function InGameArea() As Boolean
'********************************************************************
'Author: NicoNZ
'Last Modification: 29/09/2019
'Checks if last click was performed within or outside the game area.
'********************************************************************
    If clicX < 0 Or clicX > frmMain.MainViewPic.Width Then Exit Function
    If clicY < 0 Or clicY > frmMain.MainViewPic.Height Then Exit Function
    
    InGameArea = True
End Function

Private Sub hlst_Click()
    
    With hlst

        .BackColor = vbBlack

    End With

End Sub

Private Sub Minimapa_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If Button = vbRightButton Then
        Call WriteWarpChar("YO", UserMap, CByte(X - 1), CByte(Y - 1))
        Call ActualizarMiniMapa
        
    End If
    
End Sub
    'fin Incorporado ReyarB

Public Sub ActualizarMiniMapa()
    '***************************************************
    'Author: Martin Gomez (Samke)
    'Last Modify Date: 21/03/2020 (ReyarB)
    'Integrado por Reyarb
    'Se agrego campo de vision del render (Recox)
    'Ajustadas las coordenadas para centrarlo (WyroX)
    'Ajuste de coordenadas y tamaÃ±o del visor (ReyarB)
    '***************************************************
    Me.UserM.Left = UserPos.X - 2
    Me.UserM.Top = UserPos.Y - 2
    Me.MiniMapa.Refresh
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
    
    Select Case Index
    
        Case eSMType.sCombatmode
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(3), _
                                     True, False, True)
                                        
                imgSafe(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO")
                imgSafe(Index).Picture = General_Load_Picture_From_Resource("19.bmp", True)
                
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(3), _
                                     True, False, True)
                                        
                imgSafe(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO")
                imgSafe(Index).Picture = General_Load_Picture_From_Resource("20.bmp", True)
                
            End If
            
            
            
        Case eSMType.sSafemode
            
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(3), _
                                     True, False, True)
                                        
                imgSafe(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO")
                imgSafe(Index).Picture = General_Load_Picture_From_Resource("21.bmp", True)
                
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(3), _
                                     True, False, True)
                                        
                imgSafe(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO")
                imgSafe(Index).Picture = General_Load_Picture_From_Resource("22.bmp", True)

            End If
        
    End Select
    
End Sub

Private Sub UAOCerrarMenú_Click()
    FrameMenu.Visible = False
End Sub
