VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".::  Form Login ::."
   ClientHeight    =   6870
   ClientLeft      =   8205
   ClientTop       =   4665
   ClientWidth     =   10455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4059.022
   ScaleMode       =   0  'User
   ScaleWidth      =   9816.681
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adologin 
      Height          =   495
      Left            =   120
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   6255
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11033
      Orientation     =   0
      BackColor       =   14737632
      ColorGradient1  =   12632256
      ColorGradient2  =   0
      BorderColor     =   12632256
      ShowIcon        =   0   'False
      Caption         =   ""
      Icon            =   "frmLogin.frx":0000
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
      Begin Jasa_Parkir.jcbutton cmdBatal 
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   5520
         Width           =   1575
         _ExtentX        =   2778
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
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdLogin 
         Height          =   495
         Left            =   6600
         TabIndex        =   5
         Top             =   5520
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
         Caption         =   "Login"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Klik Untuk Login"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2400
         Width           =   2325
      End
      Begin VB.TextBox txtNama 
         Height          =   345
         Left            =   1815
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1920
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aplikasi Jasa Parkir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   3450
      End
      Begin VB.Label lbltgl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1080
      End
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright(c)2015 . Bina Sarana Informatika (12.3F.01 )"
      Height          =   240
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Top             =   6480
      Width           =   4710
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   240
      Picture         =   "frmLogin.frx":039A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7455
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   -240
      Width           =   1935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub buka_menu_master()
    frmMaster.mn_ganti_pass.Enabled = True
    frmMaster.mn_input_kasus.Enabled = True
    frmMaster.mn_input_parkir.Enabled = True
    frmMaster.mn_parkir_out.Enabled = True
    frmMaster.mn_lap_parkir.Enabled = True
    frmMaster.mn_logout.Enabled = True
    frmMaster.mn_parkir_out.Enabled = True
    frmMaster.mn_data_parkir.Enabled = True
    frmMaster.mn_data_kasus.Enabled = True
    frmMaster.mn_lap_kasus.Enabled = True
End Sub

Sub cek_user(admin As String)
If admin = "Y" Then
    frmMaster.mn_petugas.Enabled = True
    frmMaster.mn_tarif.Enabled = True
    frmMaster.mn_data_pendapatan.Enabled = True
Else
    frmMaster.mn_petugas.Enabled = False
    frmMaster.mn_tarif.Enabled = False
    frmMaster.mn_data_pendapatan.Enabled = False
End If
End Sub

Private Sub cmdBatal_Click()
    txtNama.Text = ""
    txtPassword = ""
End Sub

Private Sub cmdLogin_Click()
Dim user, pass, pesan As String
Dim isadmin As String

user = txtNama.Text
pass = txtPassword.Text

If user = "" Then
    MsgBox "Username belum di-isi", vbExclamation, "Kesalahan"
    ElseIf pass = "" Then
        MsgBox "Password belum di-isi", vbExclamation, "Kesalahan"
Else

adologin.Refresh
Do Until adologin.Recordset.EOF
    If adologin.Recordset.Fields("UserName").Value = user And adologin.Recordset.Fields("Password").Value = pass Then
        
        id_petugas = adologin.Recordset.Fields("ID_PETUGAS")
        nama_petugas = adologin.Recordset.Fields("NAMA_PETUGAS")
        isadmin = adologin.Recordset.Fields("IS_ADMIN")
        
        adologin.Recordset.Fields("LAST_LOGIN").Value = Now
        adologin.Recordset.Update
        
        frmMaster.Caption = ".:: Aplikasi Jasa Parkir ::. - " + nama_petugas
        pesan = "Selamat Datang " + nama_petugas + ", anda berhasil login"
        Msg = MsgBox(pesan, vbInformation, "Pemberitahuan")
        
        Call buka_menu_master
        Call cek_user(isadmin)
        Unload Me
        Exit Sub
    Else
        adologin.Recordset.MoveNext
    End If
Loop

MsgBox "Password atau Username Anda salah...Silahkan coba lagi!", vbExclamation, "Kesalahan"
End If
End Sub

Private Sub Form_Load()
    'frmLogin.txtNama.SetFocus 'error (?)
    lbltgl.Caption = Format$(Now, "dddd, mmmm dd, yyyy")
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdLogin_Click
    End If
End Sub
