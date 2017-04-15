VERSION 5.00
Begin VB.Form frmCaraPenggunaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: Form Cara Penggunaan ::."
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8145
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9551
      Orientation     =   0
      BackColor       =   14737632
      ColorGradient1  =   12632256
      ColorGradient2  =   0
      BorderColor     =   12632256
      ShowIcon        =   0   'False
      Caption         =   ""
      Icon            =   "frmCaraPenggunaan.frx":0000
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtKeluar 
         Height          =   4335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Text            =   "frmCaraPenggunaan.frx":039A
         Top             =   960
         Width           =   7815
      End
      Begin VB.TextBox txtMasuk 
         Height          =   4335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Text            =   "frmCaraPenggunaan.frx":0571
         Top             =   960
         Width           =   7815
      End
      Begin VB.OptionButton OptKeluar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Parkir Keluar"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OptMasuk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Parkir Masuk"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin Jasa_Parkir.jcbutton cmdKeluar 
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Keluar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
End
Attribute VB_Name = "frmCaraPenggunaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        OptMasuk.Value = True
        txtMasuk.Visible = True
        txtKeluar.Visible = False
End Sub

Private Sub OptKeluar_Click()
    If OptKeluar.Value = True Then
        txtMasuk.Visible = False
        txtKeluar.Visible = True
    End If
End Sub

Private Sub OptMasuk_Click()
    If OptMasuk.Value = True Then
        txtMasuk.Visible = True
        txtKeluar.Visible = False
    End If
End Sub
