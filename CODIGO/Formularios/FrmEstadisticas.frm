VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   9795
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   9510
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
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   653
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgStatus 
      Height          =   450
      Left            =   870
      Top             =   8520
      Width           =   1950
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   3
      Left            =   8010
      TabIndex        =   51
      Top             =   7590
      Width           =   645
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   7020
      TabIndex        =   50
      Top             =   6630
      Width           =   1260
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   6600
      TabIndex        =   49
      Top             =   6510
      Width           =   225
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6750
      TabIndex        =   48
      Top             =   6660
      Width           =   2220
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   6780
      TabIndex        =   47
      Top             =   7530
      Width           =   255
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ac� van las habilidades especiales del familiar"
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
      Height          =   645
      Index           =   5
      Left            =   6600
      TabIndex        =   46
      Top             =   8310
      Width           =   2160
   End
   Begin VB.Image imgFami 
      Height          =   3990
      Left            =   6180
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Shape fExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   7005
      Top             =   6660
      Width           =   1275
   End
   Begin VB.Shape fHPShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   8010
      Top             =   7620
      Width           =   645
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas matadas"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4245
      TabIndex        =   45
      Top             =   7170
      Width           =   1335
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadanos"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4455
      TabIndex        =   44
      Top             =   7710
      Width           =   855
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criminales"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4515
      TabIndex        =   43
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Veces muerto"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4350
      TabIndex        =   42
      Top             =   6660
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   6
      Left            =   1755
      TabIndex        =   41
      Top             =   7875
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   5
      Left            =   1755
      TabIndex        =   40
      Top             =   7320
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   4
      Left            =   1755
      TabIndex        =   39
      Top             =   6720
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   3
      Left            =   1755
      TabIndex        =   38
      Top             =   7620
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   2
      Left            =   1755
      TabIndex        =   37
      Top             =   6450
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   1
      Left            =   1755
      TabIndex        =   36
      Top             =   6180
      Width           =   1020
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
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
      Height          =   135
      Index           =   0
      Left            =   1650
      TabIndex        =   35
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G�nero"
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
      Height          =   135
      Index           =   5
      Left            =   1650
      TabIndex        =   34
      Top             =   4170
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raza"
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
      Height          =   135
      Index           =   6
      Left            =   1650
      TabIndex        =   33
      Top             =   4440
      Width           =   975
   End
   Begin VB.Image cmdGuardar 
      Height          =   390
      Left            =   3720
      Tag             =   "1"
      Top             =   4500
      Width           =   1470
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   8385
      TabIndex        =   32
      Top             =   4110
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   8385
      TabIndex        =   31
      Top             =   3870
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   8385
      TabIndex        =   30
      Top             =   3630
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   8385
      TabIndex        =   29
      Top             =   3375
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   8385
      TabIndex        =   28
      Top             =   3090
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   8385
      TabIndex        =   27
      Top             =   2820
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   8385
      TabIndex        =   26
      Top             =   2100
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   8385
      TabIndex        =   25
      Top             =   2340
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   8385
      TabIndex        =   24
      Top             =   2580
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   8385
      TabIndex        =   23
      Top             =   1845
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   8385
      TabIndex        =   22
      Top             =   1605
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   8385
      TabIndex        =   21
      Top             =   1350
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   8385
      TabIndex        =   20
      Top             =   1110
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   6150
      TabIndex        =   19
      Top             =   4350
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   6150
      TabIndex        =   18
      Top             =   4125
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   6150
      TabIndex        =   17
      Top             =   3870
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   6150
      TabIndex        =   16
      Top             =   3600
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   6150
      TabIndex        =   15
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   6150
      TabIndex        =   14
      Top             =   3135
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   6150
      TabIndex        =   13
      Top             =   2880
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   6150
      TabIndex        =   12
      Top             =   2625
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   6150
      TabIndex        =   11
      Top             =   2355
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   6150
      TabIndex        =   10
      Top             =   2100
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   6150
      TabIndex        =   9
      Top             =   1860
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   6150
      TabIndex        =   8
      Top             =   1620
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6150
      TabIndex        =   7
      Top             =   1140
      Width           =   105
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6150
      TabIndex        =   6
      Top             =   1410
      Width           =   105
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   21
      Left            =   8550
      Top             =   2550
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   21
      Left            =   8070
      Top             =   2550
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   20
      Left            =   8550
      Top             =   2295
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   20
      Left            =   8070
      Top             =   2310
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   19
      Left            =   8550
      Top             =   2025
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   19
      Left            =   8070
      Top             =   2040
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   18
      Left            =   8550
      Top             =   1770
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   18
      Left            =   8070
      Top             =   1800
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   17
      Left            =   8550
      Top             =   1530
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   17
      Left            =   8070
      Top             =   1515
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   16
      Left            =   8550
      Top             =   1290
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   330
      Index           =   16
      Left            =   8070
      Top             =   1290
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   15
      Left            =   8520
      Top             =   1050
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   15
      Left            =   8070
      Top             =   1050
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   22
      Left            =   8070
      Top             =   2790
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   22
      Left            =   8550
      Top             =   2790
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   23
      Left            =   8070
      Top             =   3060
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   23
      Left            =   8550
      Top             =   3030
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   24
      Left            =   8070
      Top             =   3330
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   24
      Left            =   8550
      Top             =   3300
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   25
      Left            =   8070
      Top             =   3540
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   25
      Left            =   8550
      Top             =   3540
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   26
      Left            =   8070
      Top             =   3810
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   26
      Left            =   8550
      Top             =   3780
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   27
      Left            =   8070
      Top             =   4050
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   27
      Left            =   8550
      Top             =   4050
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   14
      Left            =   5850
      Top             =   4320
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   13
      Left            =   5850
      Top             =   4080
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   12
      Left            =   5850
      Top             =   3825
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   11
      Left            =   5850
      Top             =   3570
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   10
      Left            =   5850
      Top             =   3300
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   9
      Left            =   5850
      Top             =   3090
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   8
      Left            =   5850
      Top             =   2835
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   7
      Left            =   5850
      Top             =   2580
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   6
      Left            =   5850
      Top             =   2310
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   5
      Left            =   5850
      Top             =   2055
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   4
      Left            =   5850
      Top             =   1830
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   3
      Left            =   5850
      Top             =   1590
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   2
      Left            =   5850
      Top             =   1350
      Width           =   225
   End
   Begin VB.Image menoskill 
      Height          =   300
      Index           =   1
      Left            =   5850
      Top             =   1080
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   1
      Left            =   6300
      Top             =   1080
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   14
      Left            =   6300
      Top             =   4320
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   13
      Left            =   6300
      Top             =   4080
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   12
      Left            =   6300
      Top             =   3825
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   11
      Left            =   6300
      Top             =   3570
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   10
      Left            =   6300
      Top             =   3300
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   9
      Left            =   6300
      Top             =   3090
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   8
      Left            =   6300
      Top             =   2835
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   7
      Left            =   6300
      Top             =   2580
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   6
      Left            =   6300
      Top             =   2310
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   5
      Left            =   6300
      Top             =   2055
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   4
      Left            =   6300
      Top             =   1830
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   3
      Left            =   6300
      Top             =   1590
      Width           =   225
   End
   Begin VB.Image masskill 
      Height          =   300
      Index           =   2
      Left            =   6300
      Top             =   1350
      Width           =   225
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   23
      Left            =   10440
      Top             =   4980
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   22
      Left            =   10440
      Top             =   4590
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   24
      Left            =   10440
      Top             =   4785
      Width           =   1095
   End
   Begin VB.Label lblLibres 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8100
      TabIndex        =   5
      Top             =   4560
      Width           =   675
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   16
      Left            =   10500
      Top             =   6390
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   15
      Left            =   10500
      Top             =   6195
      Width           =   1095
   End
   Begin VB.Image imgCerrar 
      Height          =   540
      Left            =   8940
      Tag             =   "1"
      Top             =   0
      Width           =   540
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   11
      Left            =   10710
      Top             =   3660
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   10
      Left            =   10470
      Top             =   3390
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   13
      Left            =   10590
      Top             =   4110
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   14
      Left            =   10500
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   21
      Left            =   10500
      Top             =   5805
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   18
      Left            =   10500
      Top             =   6585
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   19
      Left            =   10500
      Top             =   5415
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   20
      Left            =   10500
      Top             =   5610
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   9
      Left            =   10830
      Top             =   3270
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   12
      Left            =   10590
      Top             =   3870
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   17
      Left            =   10500
      Top             =   5220
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   8
      Left            =   10830
      Top             =   3390
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   7
      Left            =   10470
      Top             =   2430
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   6
      Left            =   10710
      Top             =   3510
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   5
      Left            =   10710
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   4
      Left            =   10710
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   3
      Left            =   10470
      Top             =   2190
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   2
      Left            =   10830
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   1
      Left            =   10590
      Top             =   1950
      Width           =   1095
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2340
      TabIndex        =   4
      Top             =   2295
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2340
      TabIndex        =   3
      Top             =   2010
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2340
      TabIndex        =   2
      Top             =   1710
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2340
      TabIndex        =   1
      Top             =   1425
      Width           =   90
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2340
      TabIndex        =   0
      Top             =   1140
      Width           =   90
   End
End
Attribute VB_Name = "frmEstadisticas"
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

Private cBotonCerrar As clsGraphicalButton
Private cBotonGuardar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Const ANCHO_BARRA As Byte = 73 'pixeles
Private Const BAR_LEFT_POS As Integer = 365 'pixeles

Public Sub Iniciar_Labels()
    'Iniciamos los labels con los valores de los atributos y los skills
    Dim i As Integer
    Dim Ancho As Integer
    Dim PetExpPerc As Long
    
    For i = 1 To NUMATRIBUTOS
        Atri(i).Caption = UserAtributos(i)
    Next
    
    For i = 1 To NUMSKILLS
        Skill(i).Caption = UserSkills(i)
    Next
    
    With UserReputacion
    
        Label4(1).Caption = .AsesinoRep
        Label4(2).Caption = .BandidoRep
        Label4(4).Caption = .LadronesRep
        Label4(5).Caption = .NobleRep
        Label4(3).Caption = .BurguesRep
        Label4(6).Caption = .PlebeRep
        
        If .Promedio < 0 Then
            'Label4(7).ForeColor = vbRed
            'Label4(7).Caption = "Criminal"
        Else
            'Label4(7).ForeColor = vbBlue
            'Label4(7).Caption = "Ciudadano"
        End If
    
    End With
    
    With UserEstadisticas
        est(1).Caption = .NpcsMatados
        est(2).Caption = .CiudadanosMatados
        est(3).Caption = .CriminalesMatados
        est(4).Caption = .Muertes
        
        est(0).Caption = .Clase
        If .Genero = 1 Then
            est(5).Caption = "Hombre"
        Else
            est(5).Caption = "Mujer"
        End If
        est(6).Caption = .Raza
    End With
    
    'Ponemos las estadisticas del familiar en pantalla
    If UserPet.tipo <> 0 Then
        imgFami.Picture = Nothing
        Fami(0).Visible = True
        Fami(1).Visible = True
        Fami(2).Visible = True
        Fami(3).Visible = True
        Fami(4).Visible = True
        Fami(5).Visible = True
        fHPShp.Visible = True
        fExpShp.Visible = True
        
        Fami(0).Caption = UserPet.Nombre
        Fami(1).Caption = UserPet.ELV
        
        PetExpPerc = CLng((UserPet.EXP * 100) / UserPet.ELU)
        
        If PetExpPerc <> 0 Then
            fExpShp.Width = (((UserPet.EXP / 100) / (UserPet.ELU / 100)) * 189)
        Else
            fExpShp.Width = 0
        End If
        
        Fami(2).Caption = PetExpPerc & "%"
        
        If UserPet.MinHP = 0 Then
            Fami(3).Caption = "Muerto"
            Fami(3).ForeColor = vbWhite
            fHPShp.Width = 0
            
        Else
            fExpShp.Width = (((UserPet.MinHP / 100) / (UserPet.MaxHP / 100)) * 43)
            Fami(3).Caption = UserPet.MinHP & "/" & UserPet.MaxHP
            
        End If
        
        Fami(4).Caption = UserPet.MinHIT & "/" & UserPet.MaxHIT
        Fami(5).Caption = IIf(UserPet.Habilidad = "", "Ninguna", UserPet.Habilidad)
        
    Else
        imgFami.Picture = General_Load_Picture_From_Resource("nofami.bmp", False)
        Fami(0).Visible = False
        Fami(1).Visible = False
        Fami(2).Visible = False
        Fami(3).Visible = False
        Fami(4).Visible = False
        Fami(5).Visible = False
        fHPShp.Visible = False
        fExpShp.Visible = False
        
    End If
    
    'Flags para saber que skills se modificaron
    ReDim flags(1 To NUMSKILLS)
    
End Sub

Private Sub cmdGuardar_Click()

    Dim skillChanges(NUMSKILLS) As Byte
    Dim i As Long

    For i = 1 To NUMSKILLS
        skillChanges(i) = CByte(Skill(i).Caption) - UserSkills(i)
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Skill(i).Caption)
    Next i
    
    Call WriteModifySkills(skillChanges())
    
    SkillPoints = Alocados
    
End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("stats.bmp", False)
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()
    
    Dim i As Byte
    Dim GrhPath As String
    
    GrhPath = Carga.Path(Interfaces)
    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonGuardar = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    For i = 1 To 27
        menoskill(i).MouseIcon = picMouseIcon
        masskill(i).MouseIcon = picMouseIcon
    Next i
    
    Call cBotonCerrar.Initialize(imgCerrar, "27.gif", _
                                    "28.gif", _
                                    "29.gif", Me)
                                    
    Call cBotonGuardar.Initialize(cmdGuardar, "42.gif", _
                                    "43.gif", _
                                    "44.gif", Me)

End Sub

Public Sub MostrarAsignacion()
    Dim i As Integer

    If SkillPoints > 0 Then
        For i = 1 To 27
            masskill(i).Visible = True
            menoskill(i).Visible = True
        Next i
        
        For i = 1 To 27
            menoskill(i).Picture = General_Load_Picture_From_Resource("miniprev_n.gif", False)
            masskill(i).Picture = General_Load_Picture_From_Resource("mininext_n.gif", False)
        Next i
        
    Else
    
        For i = 1 To 27
            masskill(i).Visible = False
            menoskill(i).Visible = False
        Next i

    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me

End Sub

Private Sub imgCerrar_Click()
    Unload Me
    
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If imgCerrar.Tag = 1 Then
        imgCerrar.Picture = General_Load_Picture_From_Resource("", False)
        imgCerrar.Tag = 0
    End If

End Sub

Private Sub masskill_click(Index As Integer)

    Call SumarSkillPoint(Index)
    
End Sub

Private Sub menoskill_click(Index As Integer)

    Call RestarSkillPoint(Index)
    
End Sub

Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)
'************************************
'Autor: ????
'Fecha: ????
'Descripci�n: Suma Skills
'************************************

    If Alocados > 0 Then

        If Val(Skill(SkillIndex).Caption) < MAXSKILLPOINTS Then
            Skill(SkillIndex).Caption = Val(Skill(SkillIndex).Caption) + 1
            flags(SkillIndex) = flags(SkillIndex) + 1
            Alocados = Alocados - 1
        End If
            
    End If
    
    lblLibres.Caption = Alocados
End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)
'************************************
'Autor: ????
'Fecha: ????
'Descripci�n: Resta Skills
'************************************

    If Alocados < SkillPoints Then
        
        If Val(Skill(SkillIndex).Caption) > 0 And flags(SkillIndex) > 0 Then
            Skill(SkillIndex).Caption = Val(Skill(SkillIndex).Caption) - 1
            flags(SkillIndex) = flags(SkillIndex) - 1
            Alocados = Alocados + 1
        End If
    End If
    
    lblLibres.Caption = Alocados
End Sub
