VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".:: About Aplikasi Jasa Parkir ::."
   ClientHeight    =   5355
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8865
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3696.116
   ScaleMode       =   0  'User
   ScaleWidth      =   8324.693
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   4215
      Left            =   -120
      TabIndex        =   0
      Top             =   1680
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
      Orientation     =   0
      BackColor       =   14737632
      ColorGradient1  =   12632256
      ColorGradient2  =   0
      BorderColor     =   16777215
      ShowIcon        =   0   'False
      Caption         =   ""
      Icon            =   "frmAbout.frx":0000
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
      Begin Jasa_Parkir.jcbutton cmdOK 
         Height          =   495
         Left            =   7200
         TabIndex        =   5
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "OK"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Klik Untuk Keluar Dari Form ini"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   540
         Left            =   240
         Picture         =   "frmAbout.frx":039A
         ScaleHeight     =   337.12
         ScaleMode       =   0  'User
         ScaleWidth      =   337.12
         TabIndex        =   2
         Top             =   600
         Width           =   540
      End
      Begin VB.ListBox listProgrammer 
         Height          =   1620
         ItemData        =   "frmAbout.frx":06A4
         Left            =   960
         List            =   "frmAbout.frx":06BA
         TabIndex        =   1
         ToolTipText     =   "KLIK UNTUK MELIHAT NIM"
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label lblDisclaimer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan : Aplikasi ini dibuat oleh Mahasiswa BSI semester 3 untuk tugas UAS.  Kelas 12.3F.01"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   6660
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aplikasi Jasa Parkir ini dibuat oleh :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   4260
      End
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   0
      Picture         =   "frmAbout.frx":073A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub listProgrammer_Click()
If listProgrammer.Text = "- Yudha Tri Putra" Then
   MsgBox "NIM : 12142123 ", vbInformation, "Informasi"
    ElseIf listProgrammer.Text = "- Asti Aprilliyanti" Then
            MsgBox "NIM : 12142348 ", vbInformation, "Informasi"
        ElseIf listProgrammer.Text = "- Devi Octaviani" Then
                MsgBox "NIM : 12142094 ", vbInformation, "Informasi"
            ElseIf listProgrammer.Text = "- Bangun Subkhi Ismawanto" Then
                   MsgBox "NIM : 12142688 ", vbInformation, "Informasi"
                ElseIf listProgrammer.Text = "- Manan Sabili" Then
                       MsgBox "NIM : 12142265 ", vbInformation, "Informasi"
                    ElseIf listProgrammer.Text = "- Dwi Hardianto Putra" Then
                           MsgBox "NIM : 12142392 ", vbInformation, "Informasi"
End If
End Sub
