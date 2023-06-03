VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creacion de un Clan"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildFoundation.frx":0000
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClanName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1815
      Width           =   3345
   End
   Begin VB.TextBox txtWeb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   3345
   End
   Begin VB.Image imgSiguiente 
      Height          =   390
      Left            =   2280
      Tag             =   "1"
      Top             =   3240
      Width           =   1470
   End
   Begin VB.Image imgCancelar 
      Height          =   390
      Left            =   300
      Tag             =   "1"
      Top             =   3240
      Width           =   1470
   End
End
Attribute VB_Name = "frmGuildFoundation"
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

Private cBotonSiguiente As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    ' TODO: Traducir los textos de las imagenes via labels en visual basic, para que en el futuro si se quiere se pueda traducir a mas idiomas
    ' No ando con mas ganas/tiempo para hacer eso asi que se traducen las imagenes asi tenemos el juego en ingles.
    ' Tambien usar los controles uAObuttons para los botones, usar de ejemplo frmCambiaMotd.frm
    Me.Picture = General_Load_Picture_From_Resource("89.bmp", False)
        
    Call LoadButtons
    
    If Len(txtClanName.Text) <= 30 Then
        If Not AsciiValidos(txtClanName) Then
            MsgBox JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(2)
            Exit Sub
        End If
    Else
        MsgBox JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(3)
        Exit Sub
    End If

End Sub

Private Sub LoadButtons()

    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonSiguiente.Initialize(imgSiguiente, "66.gif", _
                                    "67.gif", _
                                    "68.gif", Me)

    Call cBotonCancelar.Initialize(imgCancelar, "69.gif", _
                                    "70.gif", _
                                    "71.gif", Me)

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgSiguiente_Click()
    ClanName = txtClanName.Text
    Site = txtWeb.Text
    Unload Me
    frmGuildDetails.Show , frmMain
End Sub

Private Sub txtWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
