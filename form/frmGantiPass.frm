VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGantiPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".:: Form Ganti Password ::. "
   ClientHeight    =   3540
   ClientLeft      =   9390
   ClientTop       =   4470
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   2775
      Left            =   1680
      TabIndex        =   4
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4895
      Orientation     =   0
      BackColor       =   14737632
      ColorGradient1  =   12632256
      ColorGradient2  =   0
      BorderColor     =   12632256
      ShowIcon        =   0   'False
      Caption         =   ""
      Icon            =   "frmGantiPass.frx":0000
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
      Begin VB.TextBox txtPassLama 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1080
         Width           =   2685
      End
      Begin VB.TextBox txtNewPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   2685
      End
      Begin VB.TextBox txtPassVerify 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2280
         Width           =   2685
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input Password lama anda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2370
      End
      Begin VB.Label lblnama 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input Password baru anda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verifikasi Password baru anda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   2730
      End
   End
   Begin MSAdodcLib.Adodc adologin 
      Height          =   375
      Left            =   120
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "T_PETUGAS"
      Caption         =   "adologin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Jasa_Parkir.jcbutton cmdGanti 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Ganti Password"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ToolTip         =   "Klik Untuk Mengubah Paasword Anda"
      TooltipType     =   1
      TooltipTitle    =   "Info"
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Jasa_Parkir.jcbutton cmdBatal 
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Batal"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ToolTip         =   "Klik Untuk Membatalkan"
      TooltipType     =   1
      TooltipTitle    =   "Info"
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Jasa_Parkir.jcbutton cmdKeluar 
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Keluar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ToolTip         =   "Klik Untuk Membatalkan"
      TooltipType     =   1
      TooltipTitle    =   "Info"
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frmGantiPass.frx":039A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3975
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   1695
   End
End
Attribute VB_Name = "frmGantiPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private passlama As String
Sub baca_user(kode As String)
    adologin.Refresh
    adologin.Recordset.MoveFirst
    adologin.Recordset.Find "ID_PETUGAS='" & kode & "'"

    If adologin.Recordset.EOF Then
        MsgBox "Data Petugas tidak ditemukan", vbExclamation, "Kesalahan"
    Else
        passlama = adologin.Recordset.Fields("PASSWORD").Value
        Exit Sub
    End If
End Sub
Private Sub cmdBatal_Click()
    txtPassLama.Text = ""
    txtNewPass.Text = ""
    txtPassVerify.Text = ""
End Sub

Private Sub cmdGanti_Click()
If txtNewPass.Text = "" Or txtPassVerify.Text = "" Or txtPassLama.Text = "" Then
    MsgBox "Masih ada field yang kosong,silahkan isi semua field", vbExclamation, "Kesalahan"
ElseIf passlama <> txtPassLama.Text Then
        MsgBox "Password lama anda salah,silahkan ulangi", vbExclamation, "Kesalahan"
            ElseIf txtNewPass.Text <> txtPassVerify.Text Then
                MsgBox "Password tidak sesuai..pastikan inputan anda benar", vbExclamation, "Verifikasi Password gagal"
Else
    'set ke user yang terkait
    adologin.Recordset.Fields("PASSWORD").Value = txtNewPass.Text
    adologin.Recordset.update
    MsgBox "Password berhasil diupdate", vbInformation, "Sukses"
End If
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'txtPassLama.SetFocus
    baca_user (id_petugas)
    lblnama.Caption = nama_petugas
End Sub
