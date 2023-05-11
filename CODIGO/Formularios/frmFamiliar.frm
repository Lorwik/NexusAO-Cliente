VERSION 5.00
Begin VB.Form frmFamiliar 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
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
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   206
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   2400
   End
   Begin VB.PictureBox picFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2010
      ScaleHeight     =   1185
      ScaleWidth      =   840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   735
      Width           =   870
   End
   Begin VB.TextBox txtFamiliar 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.ComboBox lstFamiliar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      ItemData        =   "frmFamiliar.frx":0000
      Left            =   150
      List            =   "frmFamiliar.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblFamiInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descropcion del familiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   180
      TabIndex        =   3
      Top             =   1455
      Width           =   1635
   End
End
Attribute VB_Name = "frmFamiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    
    If UserPet.tipo = 0 Then
        Call MostrarMensaje("Seleccione su familiar o mascota.")
        Exit Sub
            
    ElseIf UserPet.Nombre = "" Then
        Call MostrarMensaje("Asigne un nombre a su familiar o mascota.")
        Exit Sub
            
    ElseIf Len(UserPet.Nombre) > 30 Then
        Call MostrarMensaje("El nombre de tu familiar o mascota debe tener menos de 30 letras.")
        Exit Sub
            
    End If
    
    UserPet.tipo = lstFamiliar.ListIndex + 1
    UserPet.Nombre = txtFamiliar.Text
    
    Call WriteAdoptarFamiliar

End Sub

Private Sub Form_Load()

    If UserClase = eClass.Mage Then
        Call CambioFamiliar(5)
        
    ElseIf UserClase = eClass.Hunter Or UserClase = eClass.Druid Then
        Call CambioFamiliar(4)
        
    End If
End Sub

Private Sub lstFamiliar_Click()

    If lstFamiliar.ListIndex > 0 Then
        lblFamiInfo.Caption = ListaFamiliares(lstFamiliar.ListIndex).Desc
        picFamiliar.Picture = General_Load_Picture_From_Resource(ListaFamiliares(lstFamiliar.ListIndex).Imagen)
    Else
        lblFamiInfo.Caption = "Selecciona tu familiar o mascota para saber más de él"
        picFamiliar.Picture = Nothing
    End If

End Sub

Private Sub CambioFamiliar(ByVal NumFamiliares As Integer)

    If NumFamiliares = 5 Then
    
        ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
        ListaFamiliares(1).name = "Elemental De Fuego"
        ListaFamiliares(1).Desc = "Hecho de puro fuego, lanzará tormentas sobre tus contrincantes."
        ListaFamiliares(1).Imagen = "elefuego.bmp"
        
        ListaFamiliares(2).name = "Elemental De Agua"
        ListaFamiliares(2).Desc = "Con su cuerpo acuoso paralizará a tus enemigos."
        ListaFamiliares(2).Imagen = "eleagua.bmp"
        
        ListaFamiliares(3).name = "Elemental De Tierra"
        ListaFamiliares(3).Desc = "Sus fuertes brazos inmovilizarán cualquier criatura viviente."
        ListaFamiliares(3).Imagen = "eletierra.bmp"
        
        ListaFamiliares(4).name = "Ely"
        ListaFamiliares(4).Desc = "Te protegerá constantemente con sus conjuros defensivos."
        ListaFamiliares(4).Imagen = "ely.bmp"
        
        ListaFamiliares(5).name = "Fuego Fatuo"
        ListaFamiliares(5).Desc = "Débil pero con gran poder mágico, siempre estará a tu lado."
        ListaFamiliares(5).Imagen = "fatuo.bmp"
        
    Else
    
        ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
        ListaFamiliares(1).name = "Tigre"
        ListaFamiliares(1).Desc = "Poseen grandes y filosas garras para atacar a tus oponentes."
        ListaFamiliares(1).Imagen = "tigre.bmp"
        
        ListaFamiliares(2).name = "Lobo"
        ListaFamiliares(2).Desc = "Astutos y arrogantes, su mordedura causa estragos en sus víctimas."
        ListaFamiliares(2).Imagen = "lobo.bmp"
        
        ListaFamiliares(3).name = "Oso Pardo"
        ListaFamiliares(3).Desc = "Se caracterizan por ser territoriales y muy resistentes."
        ListaFamiliares(3).Imagen = "oso.bmp"
        
        ListaFamiliares(4).name = "Ent"
        ListaFamiliares(4).Desc = "¡Esta robusta criatura te defenderá cual muro de piedra!"
        ListaFamiliares(4).Imagen = "ent.bmp"
    
    End If
    
    Dim i As Integer
    lstFamiliar.Clear
    lstFamiliar.AddItem ""
    For i = 1 To UBound(ListaFamiliares)
        lstFamiliar.AddItem ListaFamiliares(i).name
    Next i
    
    lstFamiliar.ListIndex = 0

End Sub

Private Sub txtfamiliar_GotFocus()
    Call MostrarMensaje("Mucho cuidado al colocarle nombre a su familiar, no puede ponerle el mismo o parecido nombre de su personaje, recuerde que es su companía. En caso de que el familiar o mascota tenga nombre inapropiado, podrá ser retirado.")
    
End Sub

