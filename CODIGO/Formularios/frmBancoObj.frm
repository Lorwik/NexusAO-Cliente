VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   0  'User
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3510
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "1"
      Top             =   6900
      Width           =   510
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   450
      ScaleHeight     =   3630
      ScaleWidth      =   3030
      TabIndex        =   3
      Top             =   2400
      Width           =   3030
   End
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   3930
      ScaleHeight     =   16.021
      ScaleMode       =   0  'User
      ScaleWidth      =   1042.579
      TabIndex        =   2
      Top             =   2400
      Width           =   3030
   End
   Begin VB.Image imgCerrar 
      Height          =   540
      Left            =   6870
      Tag             =   "0"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image imgDepositar 
      Height          =   525
      Left            =   4530
      MousePointer    =   99  'Custom
      Top             =   6750
      Width           =   2175
   End
   Begin VB.Image imgRetirar 
      Height          =   525
      Left            =   780
      MousePointer    =   99  'Custom
      Top             =   6750
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1290
      Width           =   5475
   End
End
Attribute VB_Name = "frmBancoObj"
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

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonRetirar As clsGraphicalButton
Private cBotonDepositar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public NoPuedeMover As Boolean
Private Shifteando As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 1 Then Shifteando = True
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift <> 1 Then Shifteando = False
    
End Sub

Private Sub cantidad_Change()

    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    'Cargamos la interfase
    Me.Picture = General_Load_Picture_From_Resource("boveda.bmp", False)
        
    Call LoadButtons
End Sub

Private Sub Form_Activate()
On Error Resume Next

    InvBanco(0).DrawInventory
    InvBanco(1).DrawInventory

End Sub

Private Sub Form_GotFocus()
On Error Resume Next

    InvBanco(0).DrawInventory
    InvBanco(1).DrawInventory

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    InvBanco(0).DrawInventory
    InvBanco(1).DrawInventory

End Sub

Private Sub LoadButtons()
    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonRetirar = New clsGraphicalButton
    Set cBotonDepositar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCerrar.Initialize(imgCerrar, "27.gif", _
                                "28.gif", _
                                "29.gif", Me)
                                
    Call cBotonRetirar.Initialize(imgRetirar, "btnretirar_n.gif", _
                                    "btnretirar_h.gif", _
                                    "btnretirar_p.gif", Me)
                                    
    Call cBotonDepositar.Initialize(imgDepositar, "30.gif", _
                                    "31.gif", _
                                    "32.gif", Me)
    
    imgRetirar.MouseIcon = picMouseIcon
    imgDepositar.MouseIcon = picMouseIcon
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
    
End Sub

Private Sub imgDepositar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    If InvBanco(1).SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    
    LastIndex2 = InvBanco(1).SelectedItem
    LasActionBuy = False
    Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
    
End Sub

Private Sub imgRetirar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    If InvBanco(0).SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    
    LastIndex1 = InvBanco(0).SelectedItem
    LasActionBuy = True
    Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
    
End Sub

Private Sub PicBancoInv_Click()

    If InvBanco(0).SelectedItem <> 0 Then

        If Shifteando Then
            LasActionBuy = True
            Call WriteBankExtractItem(InvBanco(0).SelectedItem, 10000)
            Exit Sub
        End If
        
        With UserBancoInventory(InvBanco(0).SelectedItem)
            Label1(0).Caption = .name
            
            Select Case .OBJType
                Case 2, 32
                    Label1(1).Caption = "Golpe: " & .MinHIT & "/" & .MaxHIT
                    Label1(1).Visible = True
                    
                Case 3, 16, 17
                    Label1(1).Caption = "Defensa: " & .MinDef & "/" & .MaxDef
                    Label1(1).Visible = True

                Case Else
                    Label1(1).Visible = False
                    
            End Select
            
        End With
        
    Else
        Label1(0).Caption = vbNullString
        Label1(1).Visible = False
    End If

End Sub

Private Sub PicBancoInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
    
End Sub

Private Sub PicInv_Click()
    
    If InvBanco(1).SelectedItem <> 0 Then

        If Shifteando Then
            LasActionBuy = False
            Call WriteBankDeposit(InvBanco(1).SelectedItem, 10000)
            Exit Sub
        End If
        
        With Inventario
            Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
            
            Select Case .OBJType(InvBanco(1).SelectedItem)
                Case eObjType.otWeapon, eObjType.otFlechas
                    Label1(1).Caption = "Golpe: " & .MaxHIT(InvBanco(1).SelectedItem) & "/" & .MinHIT(InvBanco(1).SelectedItem)
                    Label1(1).Visible = True
                    
                Case eObjType.otcasco, eObjType.otArmadura, eObjType.otescudo ' 3, 16, 17
                    Label1(1).Caption = "Defensa: " & .MaxDef(InvBanco(1).SelectedItem) & "/" & .MinDef(InvBanco(1).SelectedItem)
                    Label1(1).Visible = True
                    
                Case Else
                    Label1(1).Visible = False
                    
            End Select
            
        End With
    Else
        Label1(0).Caption = vbNullString
        Label1(1).Visible = False
    End If
    
End Sub

Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
    
End Sub

Private Sub imgCerrar_Click()
    Call WriteBankEnd
    NoPuedeMover = False
    
End Sub
