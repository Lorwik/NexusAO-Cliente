VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   0  'None
   Caption         =   "Ingreso"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4680
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
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   1035
      Left            =   300
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Image imgEnviar 
      Height          =   525
      Left            =   3360
      Tag             =   "1"
      Top             =   2760
      Width           =   945
   End
   Begin VB.Image imgCerrar 
      Height          =   525
      Left            =   240
      Tag             =   "1"
      Top             =   2760
      Width           =   945
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private clsFormulario    As clsFormMovementManager

Private cBotonCerrar     As clsGraphicalButton

Private cBotonEnviar     As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Dim CName                As String

Public Sub RecieveSolicitud(ByVal GuildName As String)

    CName = GuildName

End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirInterfaces & "VentanaIngreso.jpg")
    
    Call LoadButtons

End Sub

Private Sub LoadButtons()

    Dim GrhPath As String
    
    GrhPath = DirInterfaces

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonEnviar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarIngreso.jpg", GrhPath & "BotonCerrarRolloverIngreso.jpg", GrhPath & "BotonCerrarClickIngreso.jpg", Me)

    Call cBotonEnviar.Initialize(imgEnviar, GrhPath & "BotonEnviarIngreso.jpg", GrhPath & "BotonEnviarRolloverIngreso.jpg", GrhPath & "BotonEnviarClickIngreso.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgCerrar_Click()
    Unload Me

End Sub

Private Sub imgEnviar_Click()
    Call WriteGuildRequestMembership(CName, Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "º"))

    Unload Me

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal

End Sub
