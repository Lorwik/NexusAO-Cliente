VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   11505
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   6  'Inside Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   767
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6750
      Top             =   2490
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   11640
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   11
      Top             =   2460
      Width           =   3150
   End
   Begin VB.TextBox SendTxt 
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
      Height          =   240
      Left            =   150
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1830
      Visible         =   0   'False
      Width           =   9090
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6300
      Top             =   2490
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer SonidosMapas 
      Interval        =   20000
      Left            =   5340
      Top             =   2490
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1725
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   60
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   3043
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":24034E
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
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   11850
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   9150
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   610
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   718
      TabIndex        =   20
      Top             =   2250
      Width           =   10770
      Begin VB.Frame fMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2865
         Left            =   9150
         TabIndex        =   21
         Top             =   6240
         Visible         =   0   'False
         Width           =   1575
         Begin NexusAOClient.uAOButton uAOMenu 
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   22
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            TX              =   "Estadisticas"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":2403CB
            PICF            =   "frmMain.frx":2403E7
            PICH            =   "frmMain.frx":240403
            PICV            =   "frmMain.frx":24041F
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAOClient.uAOButton uAOMenu 
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   23
            Top             =   540
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            TX              =   "Clanes"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":24043B
            PICF            =   "frmMain.frx":240457
            PICH            =   "frmMain.frx":240473
            PICV            =   "frmMain.frx":24048F
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAOClient.uAOButton uAOMenu 
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   24
            Top             =   900
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            TX              =   "Grupo"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":2404AB
            PICF            =   "frmMain.frx":2404C7
            PICH            =   "frmMain.frx":2404E3
            PICV            =   "frmMain.frx":2404FF
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAOClient.uAOButton uAOMenu 
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   25
            Top             =   1260
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            TX              =   "Mapa"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":24051B
            PICF            =   "frmMain.frx":240537
            PICH            =   "frmMain.frx":240553
            PICV            =   "frmMain.frx":24056F
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAOClient.uAOButton uAOMenu 
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   26
            Top             =   1620
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            TX              =   "Opciones"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":24058B
            PICF            =   "frmMain.frx":2405A7
            PICH            =   "frmMain.frx":2405C3
            PICV            =   "frmMain.frx":2405DF
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAOClient.uAOButton uAOMenu 
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   27
            Top             =   1980
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            TX              =   "Desconectar"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":2405FB
            PICF            =   "frmMain.frx":240617
            PICH            =   "frmMain.frx":240633
            PICV            =   "frmMain.frx":24064F
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin NexusAOClient.uAOButton uAOMenu 
            Height          =   255
            Index           =   6
            Left            =   90
            TabIndex        =   28
            Top             =   2460
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            TX              =   "Cerrar Menú"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":24066B
            PICF            =   "frmMain.frx":240687
            PICH            =   "frmMain.frx":2406A3
            PICV            =   "frmMain.frx":2406BF
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Image imgSafe 
      Height          =   375
      Index           =   1
      Left            =   12570
      Top             =   10125
      Width           =   375
   End
   Begin VB.Image imgSafe 
      Height          =   375
      Index           =   0
      Left            =   12015
      Top             =   10110
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   480
      Left            =   12180
      Top             =   10860
      Width           =   2175
   End
   Begin VB.Image cmdInfo 
      Height          =   465
      Left            =   13740
      MouseIcon       =   "frmMain.frx":2406DB
      MousePointer    =   99  'Custom
      Top             =   6510
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image btnHechizos 
      Height          =   420
      Left            =   13290
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Image btnInventario 
      Height          =   420
      Left            =   11610
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   12240
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   6630
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   14400
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   10800
      Width           =   255
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9300
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   315
      Index           =   0
      Left            =   14610
      MouseIcon       =   "frmMain.frx":24082D
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":24097F
      Top             =   4200
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   315
      Index           =   1
      Left            =   14610
      MouseIcon       =   "frmMain.frx":240DCF
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":240F21
      Top             =   3855
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   11400
      Top             =   0
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   195
      Index           =   1
      Left            =   11445
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11910
      TabIndex        =   16
      Top             =   165
      Width           =   2505
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   11640
      TabIndex        =   15
      Top             =   870
      Width           =   90
   End
   Begin VB.Label lblPorcLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   12930
      TabIndex        =   14
      Top             =   1140
      Width           =   660
   End
   Begin VB.Label lblExp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp: 999999999/99999999"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   12330
      TabIndex        =   13
      Top             =   690
      Width           =   2265
   End
   Begin VB.Image CmdLanzar 
      Height          =   465
      Left            =   11730
      MouseIcon       =   "frmMain.frx":241371
      MousePointer    =   99  'Custom
      Top             =   6510
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12900
      TabIndex        =   10
      Top             =   6660
      Width           =   1425
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   14160
      TabIndex        =   4
      Top             =   10260
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   13530
      TabIndex        =   3
      Top             =   10260
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9450
      TabIndex        =   2
      Top             =   1890
      Width           =   1335
   End
   Begin VB.Image InvEqu 
      Height          =   4770
      Left            =   11520
      Top             =   2340
      Width           =   3360
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   12630
      TabIndex        =   6
      Top             =   8595
      Width           =   1095
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   12630
      TabIndex        =   5
      Top             =   9150
      Width           =   1095
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   12630
      TabIndex        =   7
      Top             =   8130
      Width           =   1095
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   11940
      TabIndex        =   8
      Top             =   9660
      Width           =   1095
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   13350
      TabIndex        =   9
      Top             =   9660
      Width           =   1095
   End
   Begin VB.Shape shpEnergia 
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   11910
      Top             =   9150
      Width           =   2535
   End
   Begin VB.Shape shpMana 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   11910
      Top             =   8610
      Width           =   2535
   End
   Begin VB.Shape shpVida 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   11910
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Shape shpHambre 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   11910
      Top             =   9660
      Width           =   1155
   End
   Begin VB.Shape shpSed 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   13290
      Top             =   9660
      Width           =   1155
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
'Nexus AO mod Argentum Online 0.13
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
'Nexus AO mod Argentum Online is based on Baronsoft's VB6 Online RPG
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

Public tX                       As Byte

Public tY                       As Byte

Public MouseX                   As Long

Public MouseY                   As Long

Public MouseBoton               As Long

Public MouseShift               As Long

Private clicX                   As Long

Private clicY                   As Long

Public IsPlaying                As Byte

Private clsFormulario           As clsFormMovementManager

Private cBotonDiamArriba        As clsGraphicalButton

Private cBotonDiamAbajo         As clsGraphicalButton

Private cBotonLanzar            As clsGraphicalButton

Private cBotonInfo              As clsGraphicalButton

Private cBotonInventario        As clsGraphicalButton

Private cBotonHechizos          As clsGraphicalButton

Private cBotonMenu              As clsGraphicalButton

Public LastButtonPressed        As clsGraphicalButton

Public picSkillStar             As Picture

Public WithEvents dragInventory As clsGrapchicalInventory
Attribute dragInventory.VB_VarHelpID = -1

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn                  As Boolean

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)

End Sub

Private Sub Form_Load()
    
    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120

    End If

    Me.Picture = LoadPicture(DirInterfaces & "Main.bmp")
    
    InvEqu.Picture = LoadPicture(DirInterfaces & "CentroInventario.jpg")
    
    Call LoadButtons
    
    Set dragInventory = Inventario
    
    ' Detect links in console
    EnableURLDetect RecTxt.hwnd, Me.hwnd
    
    ' Seteamos el caption
    Me.Caption = "Nexus AO"
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)
    
    ' Reseteamos el tamanio de la ventana para que no queden bordes blancos
    Me.Width = 15360
    Me.Height = 11520
    
    CtrlMaskOn = False

End Sub

Private Sub LoadButtons()

    Dim GrhPath As String

    Dim i       As Integer
    
    GrhPath = DirInterfaces

    Set cBotonDiamArriba = New clsGraphicalButton
    Set cBotonDiamAbajo = New clsGraphicalButton
    Set cBotonLanzar = New clsGraphicalButton
    Set cBotonInfo = New clsGraphicalButton
    Set cBotonInventario = New clsGraphicalButton
    Set cBotonHechizos = New clsGraphicalButton
    Set cBotonMenu = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonLanzar.Initialize(CmdLanzar, GrhPath & "btnLanzar.jpg", GrhPath & "btnLanzar_Hov.jpg", GrhPath & "btnLanzar_press.jpg", Me)
    
    Call cBotonLanzar.Initialize(cmdInfo, GrhPath & "btnInfo.jpg", GrhPath & "btnInfo_Hov.jpg", GrhPath & "btnInfo_press.jpg", Me)
    
    Call cBotonInventario.Initialize(btnInventario, GrhPath & "btnInventario.jpg", GrhPath & "btnInventario_Hov.jpg", GrhPath & "btnInventario_press.jpg", Me)
    
    Call cBotonHechizos.Initialize(btnHechizos, GrhPath & "btnHechizos.jpg", GrhPath & "btnHechizos_Hov.jpg", GrhPath & "btnHechizos_press.jpg", Me)
    
    Call cBotonMenu.Initialize(imgMenu, GrhPath & "btnMenu.jpg", GrhPath & "btnMenu_Hov.jpg", GrhPath & "btnMenu_press.jpg", Me)
    
    lblDropGold.MouseIcon = picMouseIcon
    lblCerrar.MouseIcon = picMouseIcon
    
    For i = 0 To 1
        imgSafe(i).MouseIcon = picMouseIcon
    Next i
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)

    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub

        Dim sTemp As String
    
        Select Case Index

            Case 1 'subir

                If hlst.ListIndex = 0 Then Exit Sub

            Case 0 'bajar

                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

        End Select
    
        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
        
        Select Case Index

            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1

            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1

        End Select

    End If

End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)

    Select Case Index
        
        Case eSMType.sSafemode

            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
                imgSafe(Index).ToolTipText = "Seguro activado."
                imgSafe(Index).Picture = LoadPicture(DirInterfaces & "segurooff.bmp")
                
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True)
                imgSafe(Index).ToolTipText = "Seguro desactivado."
                imgSafe(Index).Picture = LoadPicture(DirInterfaces & "seguroon.bmp")

            End If
            
        Case eSMType.sCombatmode

            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_COMBATE_ACTIVADO, 0, 255, 0, True, False, True)
                imgSafe(Index).ToolTipText = "Modo Combate activado."
                imgSafe(Index).Picture = LoadPicture(DirInterfaces & "combateoff.bmp")
                
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_COMBATE_DESACTIVADO, 255, 0, 0, True, False, True)
                imgSafe(Index).ToolTipText = "Modo Combate desactivado."
                imgSafe(Index).Picture = LoadPicture(DirInterfaces & "combateon.bmp")

            End If

    End Select

    SMStatus(Index) = Mostrar

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2010
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '18/11/2010: Amraphen - Agregué el handle correspondiente para las nuevas configuraciones de teclas (CTRL+0..9).
    '***************************************************
    
    If (Not SendTxt.Visible) Then
    
        'Verificamos si se está presionando la tecla CTRL.
        If Shift = 2 Then
            If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
                If KeyCode = vbKey0 Then
                    'Si es CTRL+0 muestro la ventana de configuración de teclas.
                    Call frmCustomKeys.Show(vbModal, Me)
                    
                ElseIf KeyCode >= vbKey1 And KeyCode <= vbKey9 Then

                    'Si es CTRL+1..9 cambio la configuración.
                    If KeyCode - vbKey0 = CustomKeys.CurrentConfig Then Exit Sub
                    
                    CustomKeys.CurrentConfig = KeyCode - vbKey0
                    
                    Dim sMsg As String
                    
                    sMsg = "¡Se ha cargado la configuración "

                    If CustomKeys.CurrentConfig = 0 Then
                        sMsg = sMsg & "default"
                    Else
                        sMsg = sMsg & "perzonalizada número " & CStr(CustomKeys.CurrentConfig)

                    End If

                    sMsg = sMsg & "!"

                    Call ShowConsoleMsg(sMsg, 255, 255, 255, True)

                End If
                
                CtrlMaskOn = True
                Exit Sub

            End If

        End If
        
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
                    Audio.MusicActivated = Not Audio.MusicActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

                        End With

                    Else
                        Call WriteWork(eSkill.Domar)

                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

                        End With

                    Else
                        Call WriteWork(eSkill.Robar)

                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

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

        Else
            
            'Evito que se muestren los mensajes personalizados cuando se cambie una configuración de teclas.
            If Shift = 2 Then Exit Sub
            
            Select Case KeyCode

                    'Custom messages!
                Case vbKey0 To vbKey9

                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)

                    If LenB(CustomMessage) <> 0 Then

                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase(Left(CustomMessage, 5)) <> "/CMSG" And Left(CustomMessage, 1) <> "\" Then
                            
                            Call ParseUserCommand(CustomMessage)

                        End If

                    End If

            End Select

        End If

    End If
    
    Select Case KeyCode

        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)

            If SendTxt.Visible Then Exit Sub
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call Mod_General.Client_Screenshot(frmMain.hDC, 1024, 768)
                
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyVerFPS)
            FPSFLAG = Not FPSFLAG
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)

            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

                End With

                Exit Sub

            End If
        
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
            
            If frmCustomKeys.Visible Then Exit Sub 'Chequeo si está visible la ventana de configuración de teclas.
            
            Call WriteAttack
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus

            End If
            
    End Select

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
    DisableURLDetect

End Sub

Private Sub imgMenu_Click()

    Call Audio.PlayWave(SND_CLICK)
    fMenu.Visible = Not fMenu.Visible
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

Private Sub lblCerrar_Click()
    prgRun = False

End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = 1

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
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)

End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)

End Sub

Private Sub RecTxt_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    StartCheckingLinks

End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' Control + Shift
    If Shift = 3 Then

        On Error GoTo ErrHandler
        
        ' Only allow numeric keys
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
            
            ' Get Msg Number
            Dim NroMsg As Integer

            NroMsg = KeyCode - vbKey0 - 1
            
            ' Pressed "0", so Msg Number is 9
            If NroMsg = -1 Then NroMsg = 9
            
            'Como es KeyDown, si mantenes _
             apretado el mensaje llena la consola

            If CustomMessages.Message(NroMsg) = SendTxt.Text Then
                Exit Sub

            End If
            
            CustomMessages.Message(NroMsg) = SendTxt.Text
            
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡""" & SendTxt.Text & """ fue guardado como mensaje personalizado " & NroMsg + 1 & "!!", .Red, .Green, .Blue, .bold, .italic)

            End With
            
        End If
        
    End If
    
    Exit Sub
    
ErrHandler:

    'Did detected an invalid message??
    If Err.number = CustomMessages.InvalidMessageErrCode Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("El Mensaje es inválido. Modifiquelo por favor.", .Red, .Green, .Blue, .bold, .italic)

        End With

    End If
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
        
        If PicInv.Visible Then
            PicInv.SetFocus
        Else
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
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

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
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

        End With

    Else
        Call WritePickUp

    End If

End Sub

Private Sub UsarItem()

    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call WriteUseItem(Inventario.SelectedItem)

End Sub

Private Sub EquiparItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

        End With

    Else

        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call WriteEquipItem(Inventario.SelectedItem)

    End If

End Sub

Private Sub cmdLanzar_Click()

    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)

            End With

        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True

        End If

    End If

End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    UsaMacro = False
    CnTd = 0

End Sub

Private Sub cmdINFO_Click()

    If hlst.ListIndex <> -1 Then

        Dim Index As Integer

        Index = DevolverIndexHechizo(hlst.List(hlst.ListIndex))

        Dim Msj As String
     
        If Index <> 0 Then Msj = "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf & "Nombre:" & Hechizos(Index).Nombre & vbCrLf & "Descripción:" & Hechizos(Index).Desc & vbCrLf & "Skill requerido: " & Hechizos(Index).SkillRequerido & " de magia." & vbCrLf & "Maná necesario: " & Hechizos(Index).ManaRequerida & vbCrLf & "Energía necesaria: " & Hechizos(Index).EnergiaRequerida & vbCrLf & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
                                             
        Call ShowConsoleMsg(Msj, 210, 220, 220)
        
    End If

End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
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
        Call WriteDoubleClick(tX, tY)

    End If

End Sub

Private Sub SendTxt_Click()
    SendTxt.Tag = 0 ' GSZAO

End Sub

Private Sub MainViewPic_Click()

    If Cartel Then Cartel = False

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

                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0

                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .Red, .Green, .Blue, .bold, .italic)

                        End With

                        Exit Sub

                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0

                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .Red, .Green, .Blue, .bold, .italic)

                            End With

                            Exit Sub

                        End If

                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0

                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .Red, .Green, .Blue, .bold, .italic)

                                End With

                                Exit Sub

                            End If

                        Else

                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0

                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .Red, .Green, .Blue, .bold, .italic)

                                End With

                                Exit Sub

                            End If

                        End If

                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
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
    If Not MirandoForo And Not Comerciando Then _
        Call WriteDoubleClick(tX, tY)

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

Private Sub btnInventario_Click()

    Call Audio.PlayWave(SND_CLICK)
    
    If PicInv.Visible Then Exit Sub

    InvEqu.Picture = LoadPicture(DirInterfaces & "Centroinventario.jpg")

    ' Activo controles de inventario
    PicInv.Visible = True

    ' Desactivo controles de hechizo
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
End Sub

Private Sub btnHechizos_Click()
    
    Call Audio.PlayWave(SND_CLICK)
    
    If hlst.Visible Then Exit Sub

    InvEqu.Picture = LoadPicture(DirInterfaces & "Centrohechizos.jpg")
    
    ' Activo controles de hechizos
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    ' Desactivo controles de inventario
    PicInv.Visible = False

End Sub

Private Sub picInv_DblClick()

    If MirandoCarpinteria Or MirandoHerreria Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    Call UsarItem

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)

End Sub

Private Sub RecTxt_Change()

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (Not MirandoParty) Then
         
        If PicInv.Visible Then
            PicInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        hlst.SetFocus

    End If

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
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

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

    Private Sub Socket1_Connect()
    
        'Clean input and output buffers
        Call incomingData.ReadASCIIStringFixed(incomingData.length)
        Call outgoingData.ReadASCIIStringFixed(outgoingData.length)

        Select Case EstadoLogin

            Case E_MODO.CrearNuevoPj
                Call Login
        
            Case E_MODO.Normal
                Call Login
        
            Case E_MODO.Dados
                Call Audio.PlayMIDI("7.mid")
                frmCrearPersonaje.Show vbModal

        End Select

    End Sub

Private Sub Socket1_Disconnect()
    ResetAllInfo
    Socket1.Cleanup

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, _
                              ErrorString As String, _
                              Response As Integer)

    '*********************************************
    'Handle socket errors
    '*********************************************
    Select Case ErrorCode

        Case TOO_FAST 'jajasAJ CUALQUEIRA AJJAJA
            Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
            Exit Sub

        Case REFUSED 'Vivan las negradas
            Call MsgBox("El servidor se encuentra cerrado o no te has podido conectar correctamente.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

        Case TIME_OUT
            Call MsgBox("El tiempo de espera se ha agotado, intenta nuevamente.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

        Case Else
            Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    End Select
    
    frmConnect.MousePointer = 1
    Response = 0

    frmMain.Socket1.Disconnect

End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)

    Dim RD     As String

    Dim Data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    Data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
    'Put data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData

End Sub

#End If

Private Sub AbrirMenuViewPort()
    #If (ConMenuseConextuales = 1) Then

        If tX >= MinXBorder And tY >= MinYBorder And tY <= MaxYBorder And tX <= MaxXBorder Then

            If MapData(tX, tY).CharIndex > 0 Then
                If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
                    Dim i As Long

                    Dim m As frmMenuseFashion

                    Set m = New frmMenuseFashion
            
                    Load m
                    m.SetCallback Me
                    m.SetMenuId 1
                    m.ListaInit 2, False
            
                    If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                        m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
                    Else
                        m.ListaSetItem 0, "<NPC>", True

                    End If

                    m.ListaSetItem 1, "Comerciar"
            
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
    
        Case 1 'Menu del ViewPort del engine

            Select Case Sel

                Case 0 'Nombre
                    Call WriteLeftClick(tX, tY)
        
                Case 1 'Comerciar
                    Call WriteLeftClick(tX, tY)
                    Call WriteCommerceStart

            End Select

    End Select

End Sub

Private Sub SonidosMapas_Timer()
    Sonidos.ReproducirSonidosDeMapas

End Sub

Private Sub uAOMenu_Click(Index As Integer)

    On Error GoTo uAOMenu_Click_Error
    
    Call Audio.PlayWave(SND_CLICK)
    
    Select Case Index
    
        Case 0 'Estadisticas
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer
        
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            
        Case 1 'Clanes
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            Call WriteRequestGuildLeaderInfo
            
        Case 2 'Grupo
            Call WriteRequestPartyForm
            
        Case 3 'Mapa
            Call frmMapa.Show(vbModeless, frmMain)
            
        Case 4 'Opciones
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 5 'Desconectar
            fMenu.Visible = Not fMenu.Visible
            Call ParseUserCommand("/SALIR")
            
        Case 6 'Cerrar Menú
            fMenu.Visible = Not fMenu.Visible
    End Select
    
    
    On Error GoTo 0
    Exit Sub

uAOMenu_Click_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure uAOMenu_Click, line " & Erl & "."

End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    Second.Enabled = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name Then
            Unload Forms(i)
        End If
    Loop
    On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
    Second.Enabled = True
    
    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
            Call Login

        Case E_MODO.Normal
            Call Login

        Case E_MODO.Dados
            Call Audio.PlayMIDI("7.mid")
            frmCrearPersonaje.Show vbModal
    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    Data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If

Private Function InGameArea() As Boolean

    '***************************************************
    'Author: NicoNZ
    'Last Modification: 04/07/08
    'Checks if last click was performed within or outside the game area.
    '***************************************************
    If clicX < 0 Or clicX > (32 * (Round(frmMain.MainViewPic.Width / 32))) Then Exit Function
    If clicY < 0 Or clicY > (32 * (Round(frmMain.MainViewPic.Height / 32))) Then Exit Function
    
    InGameArea = True

End Function
