VERSION 5.00
Begin VB.Form frmGuildURL 
   BorderStyle     =   0  'None
   Caption         =   "Oficial Web Site"
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6225
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
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUrl 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Left            =   420
      TabIndex        =   0
      Top             =   600
      Width           =   5355
   End
   Begin VB.Image imgAceptar 
      Height          =   540
      Left            =   1350
      Tag             =   "1"
      Top             =   1020
      Width           =   3255
   End
End
Attribute VB_Name = "frmGuildURL"
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

Private cBotonAceptar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("ventanaurlclan.bmp", False)
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonAceptar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonAceptar.Initialize(imgAceptar, "btnaceptar_n.gif", _
                                    "btnaceptar_h.gif", _
                                    "btnaceptar_d.gif", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()
    If Len(txtUrl.Text) <> 0 Then _
        Call WriteGuildNewWebsite(txtUrl.Text)
    
    Unload Me
End Sub

Private Sub txtUrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
