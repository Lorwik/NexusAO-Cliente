VERSION 5.00
Begin VB.Form FrmRetos 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9570
   ClipControls    =   0   'False
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
   Picture         =   "FrmRetos.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCompa 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   1
      Left            =   6270
      TabIndex        =   10
      Top             =   3765
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCompa 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   0
      Left            =   6270
      TabIndex        =   9
      Top             =   3405
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtOponente 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   2
      Left            =   2790
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtOponente 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   1
      Left            =   2790
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtGld 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
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
      Left            =   3090
      TabIndex        =   6
      Text            =   "0"
      Top             =   4980
      Width           =   2025
   End
   Begin VB.TextBox txtOponente 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   0
      Left            =   2790
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image tresvstres 
      Height          =   525
      Left            =   6420
      Top             =   2310
      Width           =   1725
   End
   Begin VB.Image dosvsdos 
      Height          =   495
      Left            =   3960
      Top             =   2250
      Width           =   1665
   End
   Begin VB.Image unovsuno 
      Height          =   495
      Left            =   1500
      Top             =   2250
      Width           =   1605
   End
   Begin VB.Image Salir 
      Height          =   465
      Left            =   9090
      Top             =   60
      Width           =   525
   End
   Begin VB.Image Comenzar 
      Height          =   525
      Left            =   6090
      Top             =   4920
      Width           =   1605
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona el tipo de reto que quieres jugar."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   1170
      Width           =   5400
   End
   Begin VB.Label lblCompa2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliado 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4950
      TabIndex        =   4
      Top             =   3765
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblOponente3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   1470
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblCompa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliado 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4950
      TabIndex        =   2
      Top             =   3405
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblOponente2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   1470
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblOponente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   1470
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "FrmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RetoModo As Byte

Private Sub dosvsdos_Click()
    txtOponente(1).Visible = True
    txtOponente(1).Visible = True
    txtOponente(2).Visible = False
    lblOponente.Visible = True
    lblOponente2.Visible = True
    lblOponente3.Visible = False
    txtCompa(0).Visible = True
    txtCompa(1).Visible = False
    lblCompa.Visible = True
    lblCompa2.Visible = False
    RetoModo = 2
    
End Sub

Private Sub Form_Load()

    RetoModo = 2
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    Call clsFormulario.Initialize(Me)
    
    Me.Picture = General_Load_Picture_From_Resource("retos.bmp", False)
    
    Call LoadTextsForm

    Set picNegrita = General_Load_Picture_From_Resource("129.bmp", False)
    Set picCursiva = General_Load_Picture_From_Resource("130.bmp", False)

End Sub

Private Sub LoadTextsForm()
    Me.lblDesc.Caption = JsonLanguage.item("FRM_RETOS_DESC").item("TEXTO")
    Me.lblOponente.Caption = JsonLanguage.item("FRM_RETOS_LBLOP").item("TEXTO")
    Me.lblOponente2.Caption = JsonLanguage.item("FRM_RETOS_LBLOP2").item("TEXTO")
    Me.lblOponente3.Caption = JsonLanguage.item("FRM_RETOS_LBLOP3").item("TEXTO")
    Me.lblCompa.Caption = JsonLanguage.item("FRM_RETOS_COMPA").item("TEXTO")
    Me.lblCompa2.Caption = JsonLanguage.item("FRM_RETOS_COMPA2").item("TEXTO")
    
End Sub

Private Sub Comenzar_Click()

    Dim ErrorMsg As String
    Dim ListUser As String
    
        If Not CheckDataReto(RetoModo, ListUser, ErrorMsg) Then
                MsgBox ErrorMsg
                Exit Sub
        End If
            
        Call Protocol.WriteFightSend(ListUser, Val(txtGld.Text))
        Unload Me
        
End Sub

Private Function CheckDataReto(ByVal Selected As Byte, _
                                ByRef ListUser As String, _
                                ByRef ErrorMsg As String) As Boolean
    CheckDataReto = False
    
    Dim a As Long
    
    If Val(txtGld.Text) < 0 Then
        ErrorMsg = "La apuesta minima es por 0 monedas de oro"
        Exit Function
    End If
    
    If Len(txtOponente(0).Text) <= 0 Then
        ErrorMsg = "Debes seleccionar al oponente nro 1"
        Exit Function
    End If
    
    ListUser = txtOponente(0).Text
    
    Select Case Selected
        Case 2
            If Len(txtOponente(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente nro 2"
                Exit Function
            End If
            
            If Len(txtCompa(0).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu aliado"
                Exit Function
            End If
            
            ListUser = txtOponente(0).Text & "-" & txtOponente(1).Text & "-" & txtCompa(0).Text
        Case 3
            If Len(txtOponente(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente nro 2"
                Exit Function
            End If
            
            If Len(txtOponente(2).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente nro 3"
                Exit Function
            End If
            
            If Len(txtCompa(0).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu aliado nro 2"
                Exit Function
            End If
            
            If Len(txtCompa(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu aliado nro 3"
                Exit Function
            End If
            
            ListUser = txtOponente(0).Text & "-" & txtOponente(1) & "-" & txtOponente(2) & "-" & txtCompa(0).Text & "-" & txtCompa(1).Text
    End Select
    
    
    CheckDataReto = True
End Function

Private Sub Salir_Click()
    Unload Me
End Sub

Private Sub tresvstres_Click()
    txtOponente(0).Visible = True
    txtOponente(1).Visible = True
    txtOponente(2).Visible = True
    lblOponente.Visible = True
    lblOponente2.Visible = True
    lblOponente3.Visible = True
    txtCompa(0).Visible = True
    txtCompa(1).Visible = True
    lblCompa.Visible = True
    lblCompa2.Visible = True
    RetoModo = 3
    
End Sub

Private Sub unovsuno_Click()
    txtOponente(0).Visible = True
    txtOponente(1).Visible = False
    txtOponente(2).Visible = False
    lblOponente.Visible = True
    lblOponente2.Visible = False
    lblOponente3.Visible = False
    txtCompa(0).Visible = False
    txtCompa(1).Visible = False
    lblCompa.Visible = False
    lblCompa2.Visible = False
    RetoModo = 1
    
End Sub
