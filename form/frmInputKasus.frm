VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInputKasus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".:: Form Input Kasus ::."
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   2400
      TabIndex        =   26
      Text            =   "-- PILIH --"
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   6840
      Width           =   7935
      Begin Jasa_Parkir.jcbutton cmdSimpan 
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "Simpan"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Klik Untuk Menyimpan Data"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdBatal 
         Height          =   495
         Left            =   3360
         TabIndex        =   20
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         ToolTip         =   "Klik Untuk Membatalkan Input"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdTambah 
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "&Input"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Klik Untuk Input Parkir Masuk"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdKeluar 
         Height          =   495
         Left            =   6480
         TabIndex        =   22
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         ToolTip         =   "Klik Untuk Keluar Form"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdedit 
         Height          =   495
         Left            =   4920
         TabIndex        =   25
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "Edit"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Klik Untuk Mengedit Status"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
   End
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3836
      Orientation     =   0
      BackColor       =   14737632
      ColorGradient1  =   12632256
      ColorGradient2  =   0
      BorderColor     =   12632256
      ShowIcon        =   0   'False
      Caption         =   ""
      Icon            =   "frmInputKasus.frx":0000
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Kasus"
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
         Left            =   4440
         TabIndex        =   24
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lblnokasus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
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
         Left            =   5760
         TabIndex        =   23
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kendaraan"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label lblkendaraan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
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
         Left            =   2280
         TabIndex        =   13
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lbljammasuk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         DataField       =   "JAM_MASUK"
         DataSource      =   "adoparkir"
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
         Left            =   2280
         TabIndex        =   12
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lbltglmasuk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         DataField       =   "TGL_MASUK"
         DataSource      =   "adoparkir"
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
         Left            =   2280
         TabIndex        =   11
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblnopol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         DataField       =   "NO_POLISI"
         DataSource      =   "adoparkir"
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
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
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
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Masuk"
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
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Masuk"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1035
      End
   End
   Begin MSAdodcLib.Adodc adoparkir 
      Height          =   330
      Left            =   5640
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      RecordSource    =   "T_PARKIR"
      Caption         =   "adoparkir"
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
   Begin MSAdodcLib.Adodc adokasus 
      Height          =   330
      Left            =   6720
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
      RecordSource    =   "T_KASUS"
      Caption         =   "adokasus"
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
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   360
   End
   Begin VB.TextBox txtKet 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2400
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5400
      Width           =   4695
   End
   Begin VB.TextBox txtNoTiket 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   1
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Label lblnama 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Petugas"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   1200
      Width           =   2565
   End
   Begin VB.Label lbltanggal 
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
      Left            =   2160
      TabIndex        =   16
      Top             =   1680
      Width           =   660
   End
   Begin VB.Label lbljam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   240
      Picture         =   "frmInputKasus.frx":039A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Status"
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
      Left            =   360
      TabIndex        =   4
      Top             =   6360
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Keterangan"
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
      Left            =   360
      TabIndex        =   3
      Top             =   5400
      Width           =   1035
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "No Tiket"
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
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT KASUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   2565
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   2655
      Left            =   -360
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmInputKasus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private update As Boolean

Sub autogen_kasus()
adokasus.Refresh

With adokasus.Recordset
If .EOF = False Then
    .MoveFirst
    Do While Not .EOF
        lblnokasus.Caption = .Fields("ID_KASUS")
        .MoveNext
    Loop
        lblnokasus.Caption = Val(Right(lblnokasus.Caption, 8)) + 1
        lblnokasus.Caption = "IK" + lblnokasus.Caption
Else
    lblnokasus.Caption = "IK10000001"
End If
End With
End Sub
Sub kunci_form()
    cmdSimpan.Enabled = False
    cmdBatal.Enabled = False
    txtNoTiket.Enabled = False
    txtKet.Enabled = False
    cmbStatus.Enabled = False
    cmdedit.Enabled = False
End Sub
Sub buka_form()
    cmdSimpan.Enabled = False
    cmdBatal.Enabled = True
    cmdedit.Enabled = True
    cmdTambah.Enabled = False
    txtNoTiket.Enabled = True
    txtKet.Enabled = False
    cmbStatus.Enabled = False
    txtNoTiket.SetFocus
End Sub
Private Sub cmdBatal_Click()
    cmdedit.Enabled = False
    cmdBatal.Enabled = False
    cmdTambah.Enabled = True
    lblnopol.Caption = ""
    lbltglmasuk.Caption = ""
    lbljammasuk.Caption = ""
    lblkendaraan.Caption = ""
    txtNoTiket.Text = ""
    txtKet.Text = ""
    cmbStatus.Text = ""
End Sub

Private Sub cmdedit_Click()
cmbStatus.Enabled = True
update = True
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    Dim cari As String
    cari = lblnokasus.Caption
adokasus.Recordset.MoveFirst
adokasus.Recordset.Find "ID_KASUS='" & cari & "'"

If txtKet.Text = "" Then
    MsgBox "Field Keterangan belum di-isi", vbExclamation, "kesalahan"
        ElseIf cmbStatus.Text = "-- PILIH --" Then
            MsgBox "Field Status belum di-pilih", vbExclamation, "kesalahan"
Else
If adokasus.Recordset.EOF Then
    adokasus.Recordset.AddNew
    adokasus.Recordset.Fields("ID_KASUS").Value = lblnokasus.Caption
    adokasus.Recordset.Fields("NO_TIKET").Value = txtNoTiket.Text
    adokasus.Recordset.Fields("KETERANGAN").Value = txtKet.Text
    adokasus.Recordset.Fields("STATUS").Value = cmbStatus.Text
    adokasus.Recordset.update
    MsgBox "Data Kasus berhasil disimpan", vbInformation, "Pemberitahuan"
    lblnopol.Caption = ""
    lbltglmasuk.Caption = ""
    lbljammasuk.Caption = ""
    lblkendaraan.Caption = ""
    lblnokasus.Caption = ""
    txtNoTiket.Text = ""
    txtKet.Text = ""
    cmbStatus.Text = ""
    cmdTambah.Enabled = True
    Call kunci_form
End If

If update = True Then
    adokasus.Recordset.Fields("ID_KASUS").Value = lblnokasus.Caption
    adokasus.Recordset.Fields("NO_TIKET").Value = txtNoTiket.Text
    adokasus.Recordset.Fields("KETERANGAN").Value = txtKet.Text
    adokasus.Recordset.Fields("STATUS").Value = cmbStatus.Text
    adokasus.Recordset.update
    MsgBox "Data Kasus berhasil diupdate", vbInformation, "Pemberitahuan"
    lblnopol.Caption = ""
    lbltglmasuk.Caption = ""
    lbljammasuk.Caption = ""
    lblkendaraan.Caption = ""
    txtNoTiket.Text = ""
    txtKet.Text = ""
    cmbStatus.Text = ""
    cmdTambah.Enabled = True
    Call kunci_form
End If
End If
adokasus.Refresh
update = True
End Sub

Private Sub cmdTambah_Click()
    Call buka_form
    Call autogen_kasus
    'cmdedit.Enabled = True
End Sub

Private Sub Form_Activate()
    lblnama.Caption = nama_petugas
End Sub

Private Sub Form_Load()
    lblnama.Caption = nama_petugas
    lblnopol.Caption = ""
    lbltglmasuk.Caption = ""
    lbljammasuk.Caption = ""
    lblkendaraan.Caption = ""
    lbltanggal.Caption = Format$(Now, "dddd, mmmm dd, yyyy")
    update = False
    Call kunci_form
    cmbStatus.AddItem "ON PROSES"
    cmbStatus.AddItem "SELESAI"
End Sub

Private Sub Timer1_Timer()
    lbljam.Caption = Time
End Sub

Private Sub txtNoTiket_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    Dim cari As String
    cari = txtNoTiket.Text
    
    If cari = "" Then
        MsgBox "Nomor Tiket belum di input", vbExclamation, "Kesalahan"
    Else
    
    adoparkir.Refresh
    adoparkir.Recordset.MoveFirst
    adoparkir.Recordset.Find "NO_TIKET='" & cari & "'"

    If adoparkir.Recordset.EOF Then
        MsgBox "Data tidak ditemukan"
    Else
            
        If adoparkir.Recordset.Fields("SUDAH_KELUAR").Value = "Y" Then
            MsgBox "Data ini sudah tidak valid", vbExclamation, "Peringatan"
        Else
        
            If adoparkir.Recordset.Fields("KODE_JENIS").Value = "P01" Then
                lblkendaraan.Caption = "MOTOR"
            Else
                lblkendaraan.Caption = "MOBIL"
            End If
                txtKet.Enabled = True
                cmbStatus.Enabled = True
                cmdSimpan.Enabled = True
        End If
    End If
    End If
adokasus.Recordset.MoveFirst
adokasus.Recordset.Find "NO_TIKET='" & cari & "'"

If adokasus.Recordset.EOF Then
'adokasus.Recordset.AddNew
    'MsgBox "Data Petugas tidak ditemukan", vbInformation, "Informasi"
Else
txtKet.Text = adokasus.Recordset!KETERANGAN
    cmbStatus.Text = adokasus.Recordset!Status
    lblnokasus.Caption = adokasus.Recordset!ID_KASUS
    cmdedit.Enabled = True
    txtKet.Enabled = False
    'txtStatus.Enabled = False
    cmbStatus.Enabled = False
End If
End If
End Sub
