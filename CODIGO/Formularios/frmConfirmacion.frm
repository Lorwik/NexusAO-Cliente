VERSION 5.00
Begin VB.Form frmConfirmacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5730
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCancelar 
      Height          =   525
      Left            =   300
      Top             =   2790
      Width           =   1800
   End
   Begin VB.Image imgAceptar 
      Height          =   540
      Left            =   2130
      Top             =   2790
      Width           =   3255
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
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
      Height          =   1995
      Left            =   330
      TabIndex        =   0
      Top             =   390
      Width           =   5085
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ' TODO: Traducir los textos de las imagenes via labels en visual basic, para que en el futuro si se quiere se pueda traducir a mas idiomas
    ' No ando con mas ganas/tiempo para hacer eso asi que se traducen las imagenes asi tenemos el juego en ingles.
    ' Tambien usar los controles uAObuttons para los botones, usar de ejemplo frmCambiaMotd.frm
    Me.Picture = General_Load_Picture_From_Resource("info.bmp")
    
    Call LoadButtons
End Sub

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub LoadButtons()
    Dim boton As String
    
   ' GrhPath = Carga.path(Interfaces)

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, "btnaceptar_n.gif", _
                                          "btnaceptar_h.gif", _
                                          "btnaceptar_d.gif", Me)
                                     
                                     
    Call cBotonCancelar.Initialize(imgCancelar, "13.gif", "14.gif", "15.gif", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    msg.Caption = "" 'Limpiamos el caption
    Unload Me
End Sub

Private Sub imgAceptar_Click()
    Call WriteRespuestaInstruccion(True)
    Unload Me
End Sub

Private Sub imgCancelar_Click()
    Call WriteRespuestaInstruccion(False)
    Unload Me
End Sub

Private Sub msg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

