VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6840
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
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   5925
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   3540
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   3855
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   4170
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   4500
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   4800
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   6
      Top             =   5130
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   7
      Top             =   5460
      Width           =   5835
   End
   Begin VB.TextBox txtCodex1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   8
      Top             =   5820
      Width           =   5835
   End
   Begin VB.Image imgConfirmar 
      Height          =   390
      Left            =   2550
      Tag             =   "1"
      Top             =   6270
      Width           =   1470
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   6450
      Tag             =   "1"
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private clsFormulario As clsFormMovementManager

Private cBotonConfirmar As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Const MAX_DESC_LENGTH As Integer = 520
Private Const MAX_CODEX_LENGTH As Integer = 100

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("ventanacodex.bmp", False)
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonConfirmar = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonConfirmar.Initialize(imgConfirmar, "42.gif", _
                                    "43.gif", _
                                    "44.gif", Me)

    Call cBotonSalir.Initialize(imgSalir, "36.gif", _
                                    "37.gif", _
                                    "38.gif", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgConfirmar_Click()
    Dim fdesc As String
    Dim Codex() As String
    Dim k As Byte
    Dim Cont As Byte

    fdesc = Replace(txtDesc, vbNewLine, "�", , , vbBinaryCompare)


    Cont = 0
    For k = 0 To txtCodex1.UBound
        If LenB(txtCodex1(k).Text) <> 0 Then Cont = Cont + 1
    Next k
    
    If Cont < 4 Then
        MsgBox "Debes definir al menos cuatro mandamientos."
        Exit Sub
    End If
                
    ReDim Codex(txtCodex1.UBound) As String
    For k = 0 To txtCodex1.UBound
        Codex(k) = txtCodex1(k)
    Next k

    If CreandoClan Then
        Call WriteCreateNewGuild(fdesc, ClanName, Site, Codex)
    Else
        Call WriteClanCodexUpdate(fdesc, Codex)
    End If

    CreandoClan = False
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub

Private Sub txtCodex1_Change(Index As Integer)
    If Len(txtCodex1.item(Index).Text) > MAX_CODEX_LENGTH Then _
        txtCodex1.item(Index).Text = Left$(txtCodex1.item(Index).Text, MAX_CODEX_LENGTH)
End Sub

Private Sub txtCodex1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub txtDesc_Change()
    If Len(txtDesc.Text) > MAX_DESC_LENGTH Then _
        txtDesc.Text = Left$(txtDesc.Text, MAX_DESC_LENGTH)
End Sub
