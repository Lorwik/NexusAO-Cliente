VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   11490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   766
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2310
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3420
      Width           =   3030
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   495
      Left            =   2310
      TabIndex        =   0
      Top             =   2310
      Width           =   5460
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   8340
      TabIndex        =   2
      Text            =   "7666"
      Top             =   2460
      Width           =   825
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   9390
      TabIndex        =   4
      Text            =   "localhost"
      Top             =   2460
      Width           =   1575
   End
   Begin VB.Image imgConectarse 
      Height          =   795
      Left            =   5580
      Top             =   3150
      Width           =   2175
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   9960
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgBorrarPj 
      Height          =   375
      Left            =   8400
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgRecuperar 
      Height          =   1035
      Left            =   5070
      Top             =   4230
      Width           =   2685
   End
   Begin VB.Image imgCrearPj 
      Height          =   1005
      Left            =   2370
      Top             =   4200
      Width           =   2565
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private cBotonCrearPj       As clsGraphicalButton

Private cBotonRecuperarPass As clsGraphicalButton

Private cBotonManual        As clsGraphicalButton

Private cBotonReglamento    As clsGraphicalButton

Private cBotonCodigoFuente  As clsGraphicalButton

Private cBotonBorrarPj      As clsGraphicalButton

Private cBotonSalir         As clsGraphicalButton

Private cBotonLeerMas       As clsGraphicalButton

Private cBotonForo          As clsGraphicalButton

Private cBotonConectarse    As clsGraphicalButton

Private cBotonTeclas        As clsGraphicalButton

Public LastButtonPressed    As clsGraphicalButton

Private Sub Form_Activate()
    'On Error Resume Next

    If ServersRecibidos Then
        If CurServer <> 0 Then
            IPTxt = ServersLst(1).Ip
            PortTxt = ServersLst(1).Puerto
        Else
            IPTxt = IPdelServidor
            PortTxt = PuertoDelServidor

        End If

    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        prgRun = False

    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'Make Server IP and Port box visible
    If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
        'Port
        PortTxt.Visible = True
        'Label4.Visible = True
    
        'Server IP
        PortTxt.Text = "7666"
        IPTxt.Text = "192.168.0.2"
        IPTxt.Visible = True
        'Label5.Visible = True
    
        KeyCode = 0
        Exit Sub

    End If

End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]

    PortTxt.Text = Config_Inicio.Puerto
 
    '[CODE]:MatuX
    '
    '  El código para mostrar la versión se genera acá para
    ' evitar que por X razones luego desaparezca, como suele
    ' pasar a veces :)
    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    '[END]'
    
    Me.Picture = LoadPicture(App.path & "\Interfaces\VentanaConectar.jpg")
    
    Call LoadButtons

    Call CheckLicenseAgreement
        
End Sub

Private Sub CheckLicenseAgreement()

    'Recordatorio para cumplir la licencia, por si borrás el Boton sin leer el code...
    Dim i As Long
    
    For i = 0 To Me.Controls.Count - 1

        If Me.Controls(i).name = "imgCodigoFuente" Then
            Exit For

        End If

    Next i
    
    If i = Me.Controls.Count Then
        MsgBox "No debe eliminarse la posibilidad de bajar el código de sus servidor. Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor, incurriendo de esta forma en un delito punible por ley." & vbCrLf & vbCrLf & vbCrLf & "Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo, no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.", vbCritical Or vbApplicationModal

    End If

End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set cBotonCrearPj = New clsGraphicalButton
    Set cBotonRecuperarPass = New clsGraphicalButton
    Set cBotonBorrarPj = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonLeerMas = New clsGraphicalButton
    Set cBotonConectarse = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
        
    Call cBotonCrearPj.Initialize(imgCrearPj, GrhPath & "BotonCrearPersonajeConectar.jpg", GrhPath & "BotonCrearPersonajeRolloverConectar.jpg", GrhPath & "BotonCrearPersonajeClickConectar.jpg", Me)
                                    
    Call cBotonRecuperarPass.Initialize(imgRecuperar, GrhPath & "BotonRecuperarPass.jpg", GrhPath & "BotonRecuperarPassRollover.jpg", GrhPath & "BotonRecuperarPassClick.jpg", Me)
                                    
    Call cBotonBorrarPj.Initialize(imgBorrarPj, GrhPath & "BotonBorrarPersonaje.jpg", GrhPath & "BotonBorrarPersonajeRollover.jpg", GrhPath & "BotonBorrarPersonajeClick.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirConnect.jpg", GrhPath & "BotonBotonSalirRolloverConnect.jpg", GrhPath & "BotonSalirClickConnect.jpg", Me)
                                    
    Call cBotonConectarse.Initialize(imgConectarse, GrhPath & "BotonConectarse.jpg", GrhPath & "BotonConectarseRollover.jpg", GrhPath & "BotonConectarseClick.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal

End Sub

Private Sub CheckServers()

    If ServersRecibidos Then
        If Not IsIp(IPTxt) And CurServer <> 0 Then
            If MsgBox("Atencion, está intentando conectarse a un servidor no oficial, NoLand Studios no se hace responsable de los posibles problemas que estos servidores presenten. ¿Desea continuar?", vbYesNo) = vbNo Then
                If CurServer <> 0 Then
                    IPTxt = ServersLst(CurServer).Ip
                    PortTxt = ServersLst(CurServer).Puerto
                Else
                    IPTxt = IPdelServidor
                    PortTxt = PuertoDelServidor

                End If

                Exit Sub

            End If

        End If

    End If

    CurServer = 0
    IPdelServidor = IPTxt
    PuertoDelServidor = PortTxt

End Sub

Private Sub imgBorrarPj_Click()

    On Error GoTo errH

    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)

    Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")

End Sub

Private Sub imgConectarse_Click()
    Call CheckServers
    
    #If UsarWrench = 1 Then

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents

        End If

    #Else

        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents

        End If

    #End If
    
    'update user info
    UserName = txtNombre.Text
    
    Dim aux As String

    aux = txtPasswd.Text
    UserPassword = aux

    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
        #If UsarWrench = 1 Then
            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
        #Else
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
        #End If

    End If
    
End Sub

Private Sub imgCrearPj_Click()
    
    Call CheckServers
    
    EstadoLogin = E_MODO.Dados
    #If UsarWrench = 1 Then

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents

        End If

        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
    #Else

        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents

        End If

        frmMain.Winsock1.Connect CurServerIp, CurServerPort
    #End If

End Sub

Private Sub imgRecuperar_Click()

    On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")

End Sub

Private Sub imgSalir_Click()
    prgRun = False

End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then imgConectarse_Click

End Sub

Private Sub WebAuxiliar_BeforeNavigate2(ByVal pDisp As Object, _
                                        URL As Variant, _
                                        flags As Variant, _
                                        TargetFrameName As Variant, _
                                        PostData As Variant, _
                                        Headers As Variant, _
                                        Cancel As Boolean)
    
    If InStr(1, URL, "alkon") <> 0 Then
        Call ShellExecute(hwnd, "open", URL, vbNullString, vbNullString, SW_SHOWNORMAL)
        Cancel = True

    End If
    
End Sub
