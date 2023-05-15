VERSION 5.00
Begin VB.Form frmCharList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "ImperiumClasico"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   7
      Left            =   10575
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   7260
      Width           =   1500
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   6
      Left            =   8130
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   6
      Top             =   7290
      Width           =   1500
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   5
      Left            =   5700
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   5
      Top             =   7290
      Width           =   1500
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   4
      Left            =   3255
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   7320
      Width           =   1500
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   3
      Left            =   10575
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   4500
      Width           =   1500
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   2
      Left            =   8130
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   4500
      Width           =   1500
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   1
      Left            =   5700
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   4530
      Width           =   1500
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Index           =   0
      Left            =   3255
      ScaleHeight     =   1830
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   4500
      Width           =   1500
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   8
      Left            =   10605
      TabIndex        =   16
      Top             =   9360
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   7
      Left            =   8130
      TabIndex        =   15
      Top             =   9360
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   6
      Left            =   5640
      TabIndex        =   14
      Top             =   9360
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   5
      Left            =   3240
      TabIndex        =   13
      Top             =   9360
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   4
      Left            =   10605
      TabIndex        =   12
      Top             =   6600
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   3
      Left            =   8130
      TabIndex        =   11
      Top             =   6600
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   2
      Left            =   5640
      TabIndex        =   10
      Top             =   6600
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
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
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   6600
      Width           =   1515
   End
   Begin VB.Label lblAccData 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   8
      Top             =   3450
      Width           =   3705
   End
   Begin VB.Image imgConectar 
      Height          =   525
      Left            =   11250
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   10050
      Width           =   1800
   End
   Begin VB.Image imgBorrarPj 
      Height          =   525
      Left            =   4230
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   10020
      Width           =   1800
   End
   Begin VB.Image imgSalir 
      Height          =   525
      Left            =   2310
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   10050
      Width           =   1800
   End
   Begin VB.Image imgCrearPJ 
      Height          =   525
      Left            =   9330
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   10050
      Width           =   1800
   End
   Begin VB.Image imgCambiarPass 
      Height          =   525
      Left            =   6810
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   10050
      Width           =   1800
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   7
      Left            =   10425
      Top             =   7155
      Width           =   1785
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   6
      Left            =   7980
      Top             =   7155
      Width           =   1785
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   5
      Left            =   5550
      Top             =   7155
      Width           =   1785
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   4
      Left            =   3090
      Top             =   7155
      Width           =   1785
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   3
      Left            =   10440
      Top             =   4395
      Width           =   1785
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   2
      Left            =   7980
      Top             =   4395
      Width           =   1785
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   1
      Left            =   5550
      Top             =   4395
      Width           =   1785
   End
   Begin VB.Image imgAcc 
      Height          =   2070
      Index           =   0
      Left            =   3090
      Top             =   4395
      Width           =   1785
   End
End
Attribute VB_Name = "frmCharList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonConectar     As clsGraphicalButton
Private cBotonCrearPJ      As clsGraphicalButton
Private cBotonSalir        As clsGraphicalButton
Private cBotonBorrar       As clsGraphicalButton
Private cBotonCambiarPass  As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call imgSalir_Click
    End If

End Sub

Private Sub Form_Load()

    Dim i As Long

    Unload frmConnect
    
    Me.Picture = General_Load_Picture_From_Resource("cuenta.bmp")
    Me.Icon = frmMain.Icon
    ' Seteamos el caption
    Me.Caption = Form_Caption

    For i = 1 To 8
        lblAccData(i).Caption = vbNullString
    Next i

    Me.lblAccData(0).Caption = AccountName
    
    Call LoadButtons
    
End Sub

Private Sub imgBorrarPj_Click()
    If PJAccSelected < 1 Then
        Call MostrarMensaje(JsonLanguage.item("ERROR_PERSONAJE_NO_SELECCIONADO").item("TEXTO"))
        Exit Sub
    End If
                            
    frmBorrarPJ.Show
End Sub

Private Sub imgCambiarPass_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "http://NexusAO.com.ar/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub imgConectar_Click()
    If Not frmMain.Client.State = sckConnected Then
        MsgBox JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO")
        frmConnect.Show
                
    Else
        If Mod_Declaraciones.Conectando Then
            Mod_Declaraciones.Conectando = False
            Call WriteLoginExistingChar
                    
            DoEvents
            
            Call FlushBuffer
        End If
    End If
End Sub

Private Sub imgCrearPJ_Click()
    If NumberOfCharacters > 9 Then
        MsgBox JsonLanguage.item("ERROR_DEMASIADOS_PJS").item("TEXTO")
        Exit Sub
    End If
            
    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_CrearPersonaje
            Sound.Fading = 500
        End If
    End If
            
    Dim LoopC As Long
        
    For LoopC = 1 To 10
        If LenB(lblAccData(LoopC).Caption) = 0 Then
            frmCrearPersonaje.Show
            Exit Sub
        End If
    Next LoopC
End Sub

Private Sub imgSalir_Click()
    frmMain.Client.CloseSck
    Call ResetAllInfoAccounts
    Call ListarServidores
    frmConnect.Visible = True
    
    Unload Me
End Sub

Private Sub picChar_Click(Index As Integer)
    On Error Resume Next
    
    If LenB(cPJ(Index + 1).Nombre) <> 0 Then
        'El PJ seleccionado queda guardado
        UserName = cPJ(Index + 1).Nombre
        PJAccSelected = Index + 1
    End If

End Sub

Private Sub picChar_DblClick(Index As Integer)

    Call imgConectar_Click

End Sub

Private Sub LoadButtons()

    Set LastButtonPressed = New clsGraphicalButton
    
    Set cBotonConectar = New clsGraphicalButton
    Set cBotonCrearPJ = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonBorrar = New clsGraphicalButton
    Set cBotonCambiarPass = New clsGraphicalButton
    
    Call cBotonCambiarPass.Initialize(imgCambiarPass, "22.gif", _
                                 "26.gif", _
                                 "23.gif", Me)
                                 
    Call cBotonCrearPJ.Initialize(imgCrearPJ, "10.gif", _
                                 "12.gif", _
                                 "11.gif", Me)
                                 
    Call cBotonSalir.Initialize(imgSalir, "13.gif", _
                                 "15.gif", _
                                 "14.gif", Me)
                                 
    Call cBotonBorrar.Initialize(imgBorrarPj, "16.gif", _
                                 "18.gif", _
                                 "17.gif", Me)
                                 
    Call cBotonConectar.Initialize(imgConectar, "19.gif", _
                                 "21.gif", _
                                 "20.gif", Me)
                                 
                                 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
End Sub

