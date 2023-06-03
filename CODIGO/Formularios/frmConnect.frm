VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   983.04
   ScaleMode       =   0  'User
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox lst_servers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1395
      ItemData        =   "frmConnect.frx":240050
      Left            =   150
      List            =   "frmConnect.frx":240057
      TabIndex        =   2
      Top             =   9120
      Width           =   3255
   End
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "123456"
      Top             =   7710
      Width           =   2595
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   315
      Left            =   6360
      MaxLength       =   25
      TabIndex        =   0
      Text            =   "Lorwik"
      Top             =   6210
      Width           =   2565
   End
   Begin VB.Image btnRecordar 
      Height          =   390
      Left            =   6360
      Picture         =   "frmConnect.frx":240068
      Top             =   8310
      Width           =   390
   End
   Begin VB.Image btnOpciones 
      Height          =   525
      Left            =   150
      Top             =   8460
      Width           =   3210
   End
   Begin VB.Image imgRecargar 
      Height          =   375
      Left            =   3480
      Top             =   9150
      Width           =   255
   End
   Begin VB.Image imgCerrar 
      Height          =   525
      Left            =   180
      Top             =   10650
      Width           =   3210
   End
   Begin VB.Image imgRecuperar 
      Height          =   525
      Left            =   150
      MousePointer    =   99  'Custom
      Top             =   7650
      Width           =   3210
   End
   Begin VB.Image imgRegistrar 
      Height          =   525
      Left            =   150
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   3210
   End
   Begin VB.Image imgConectar 
      Height          =   525
      Left            =   6090
      MousePointer    =   99  'Custom
      Top             =   9060
      Width           =   3210
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Lector As clsIniManager

Private cBotonConectar     As clsGraphicalButton
Private cBotonRegistrar    As clsGraphicalButton
Private cBotonRecuperar    As clsGraphicalButton
Private cBotonCerrar       As clsGraphicalButton
Private cBotonOpciones     As clsGraphicalButton
Private cBotonRecordar     As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Private Sub btnOpciones_Click()
    Call frmOpciones.Show(vbModeless, frmConnect)
End Sub

Private Sub imgCerrar_Click()
    Call CloseClient
End Sub

Private Sub imgConectar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "SERVERSELECT", ServIndSel)
        
    'update user info
    AccountName = frmConnect.txtNombre.Text
    AccountPassword = frmConnect.txtPasswd.Text
        
    'Clear spell list
    frmMain.hlst.Clear
        
    If CheckUserData() = True Then
        Call Protocol.Connect(E_MODO.Normal)
    End If
    
End Sub

Private Sub imgRecargar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ListarServidores
End Sub

Private Sub imgRecuperar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "http://NexusAO.com.ar/", "", App.Path, SW_SHOWNORMAL)

End Sub

Private Sub imgRegistrar_Click()
    Call ShellExecute(0, "Open", "http://NexusAO.com.ar/", "", App.Path, SW_SHOWNORMAL)
    Call Sound.Sound_Play(SND_CLICK)
End Sub

Private Sub lst_servers_Click()
    ServIndSel = lst_servers.ListIndex
    CurServerIp = Servidor(ServIndSel).Ip
    CurServerPort = Servidor(ServIndSel).Puerto
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
  '  If KeyAscii = vbKeyReturn Then btnConectarse_Click
End Sub

Private Sub Form_Load()
    ' Seteamos el caption
    Me.Caption = Form_Caption
    
    Me.Picture = General_Load_Picture_From_Resource("conectar.bmp")
    
    ServIndSel = GetVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "SERVERSELECT")
        
    Call LoadButtons
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call CloseClient
    End If

End Sub

Private Sub LoadButtons()

    Set LastButtonPressed = New clsGraphicalButton
    
    Set cBotonConectar = New clsGraphicalButton
    Set cBotonRegistrar = New clsGraphicalButton
    Set cBotonRecuperar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonOpciones = New clsGraphicalButton
    Set cBotonRecordar = New clsGraphicalButton
    
    imgCerrar.MouseIcon = picMouseIcon
    
    Call cBotonConectar.Initialize(imgConectar, "botconectar.bmp", _
                                 "botconectarover.bmp", _
                                 "botconectardown.bmp", Me)
                                 
    Call cBotonRegistrar.Initialize(imgRegistrar, "botcrear.bmp", _
                                 "botcrearover.bmp", _
                                 "botcreardown.bmp", Me)
                                 
    Call cBotonRecuperar.Initialize(imgRecuperar, "botrecuperar.bmp", _
                                 "botrecuperarover.bmp", _
                                 "botrecuperardown.bmp", Me)
                                 
    Call cBotonCerrar.Initialize(imgCerrar, "btnsalir.bmp", _
                                 "btnsalirover.bmp", _
                                 "btnsalirdown.bmp", Me)
                                 
    Call cBotonOpciones.Initialize(btnOpciones, "63.gif", _
                                 "64.gif", _
                                 "65.gif", Me)
                                 
    Call cBotonRecordar.Initialize(btnRecordar, "btnrecordar.bmp", _
                                 "btnrecordarover.bmp", _
                                 "btnrecordardown.bmp", Me)
                                 
                                 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
End Sub
