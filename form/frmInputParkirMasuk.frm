VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInputParkirMasuk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".:: Form Input Parkir Masuk ::."
   ClientHeight    =   7560
   ClientLeft      =   4425
   ClientTop       =   2715
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adovalidasi 
      Height          =   375
      Left            =   6120
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   ""
      Caption         =   "adovalidasi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   6360
      Width           =   8175
      Begin Jasa_Parkir.jcbutton cmdTambah 
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputParkirMasuk.frx":0000
         backcolor       =   14935011
         caption         =   "&Input"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Input Parkir Masuk"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdSimpan 
         Height          =   495
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputParkirMasuk.frx":0028
         backcolor       =   14935011
         caption         =   "Simpan"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Menyimpan Data"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdBatal 
         Height          =   495
         Left            =   3600
         TabIndex        =   23
         Top             =   360
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputParkirMasuk.frx":0050
         backcolor       =   14935011
         caption         =   "Batal"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Membatalkan Input"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdKeluar 
         Height          =   495
         Left            =   6360
         TabIndex        =   24
         Top             =   360
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputParkirMasuk.frx":0078
         backcolor       =   14935011
         caption         =   "Keluar"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Kleuar Dari Form"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
   End
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   8175
      _extentx        =   14420
      _extenty        =   3625
      orientation     =   0
      backcolor       =   14737632
      colorgradient1  =   12632256
      colorgradient2  =   0
      bordercolor     =   12632256
      showicon        =   0   'False
      caption         =   ""
      icon            =   "frmInputParkirMasuk.frx":00A0
      forecolor       =   16777215
      font            =   "frmInputParkirMasuk.frx":043C
      Begin VB.TextBox txtTglMasuk 
         BackColor       =   &H00C0FFFF&
         DataSource      =   "adotransaksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtJamMasuk 
         BackColor       =   &H00C0FFFF&
         DataSource      =   "adotransaksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtNoTiket 
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
         Height          =   405
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Tiket"
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
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1035
      End
   End
   Begin MSAdodcLib.Adodc adotarif 
      Height          =   375
      Left            =   6120
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "T_JENIS_KENDARAAN"
      Caption         =   "adotarif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Left            =   5520
      Top             =   1800
   End
   Begin VB.OptionButton OptMobil 
      Caption         =   "Mobil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
   End
   Begin VB.OptionButton OptMotor 
      Caption         =   "Motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtNoPol 
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
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   0
      Top             =   4800
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc adoparkir 
      Height          =   375
      Left            =   6120
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "T_PARKIR"
      Caption         =   "adoparkir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   1680
      TabIndex        =   19
      Top             =   2040
      Width           =   345
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
      Left            =   1680
      TabIndex        =   18
      Top             =   1560
      Width           =   660
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
      Left            =   1680
      TabIndex        =   17
      Top             =   1080
      Width           =   2565
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Tarif Jam Pertama:"
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
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label lbltarifpertama 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      TabIndex        =   6
      Top             =   5640
      Width           =   105
   End
   Begin VB.Label lbltarifjam 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      TabIndex        =   5
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tarif Per-jam :"
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
      TabIndex        =   4
      Top             =   6000
      Width           =   1245
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Kendaraan"
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
      TabIndex        =   3
      Top             =   5280
      Width           =   1755
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARKIR MASUK"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   2850
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "No Polisi"
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
      TabIndex        =   1
      Top             =   4800
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "frmInputParkirMasuk.frx":0468
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   2535
      Left            =   -240
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmInputParkirMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub autogen_tiket()
'pola = XXXDDMMYY000
'contoh = MTR021116001

Dim tgl As String, bln As String, thn As String, kodetgl As String, thn_1 As String
tgl = Format(Date, "dd")
bln = Month(Date)
thn_1 = Year(Date)
thn = Right(thn_1, 2)

kodetgl = tgl + bln + thn
'MsgBox kodetgl

adoparkir.Refresh

With adoparkir.Recordset
If .EOF = False Then
    .MoveFirst
    Do While Not .EOF
        txtNoTiket.Text = .Fields("NO_TIKET")
        .MoveNext
    Loop
        txtNoTiket.Text = Val(Right(txtNoTiket.Text, 3)) + 1
        txtNoTiket.Text = kodetgl + txtNoTiket.Text
Else
    txtNoTiket.Text = kodetgl + "101"
End If
End With
End Sub

Sub kunci_form()
    txtNoTiket.Enabled = False
    txtTglMasuk.Enabled = False
    txtJamMasuk.Enabled = False
    txtNoPol.Enabled = False
    cmdSimpan.Enabled = False
    cmdTambah.Enabled = True
    cmdBatal.Enabled = False
    OptMotor.Enabled = False
    OptMobil.Enabled = False
End Sub

Sub bersih()
    txtNoPol.Text = ""
    lbltarifpertama.Caption = 0
    lbltarifjam.Caption = 0
    OptMotor.Value = False
    OptMobil.Value = False
End Sub

Private Sub cmdBatal_Click()
    Call bersih
    txtNoPol.SetFocus
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim kd_jenis As String
Dim jum As Integer

If OptMotor.Value = True Then
    kd_jenis = "P01"
ElseIf OptMobil.Value = True Then
        kd_jenis = "P02"
End If

'cek dulu didatabase
adovalidasi.CommandType = adCmdUnknown
adovalidasi.RecordSource = "SELECT COUNT(NO_TIKET) AS JUM FROM T_PARKIR WHERE SUDAH_KELUAR LIKE 'T' AND NO_POLISI LIKE '" & txtNoPol.Text & "'"
adovalidasi.Refresh
jum = adovalidasi.Recordset.Fields("JUM").Value

'validasi
If txtNoPol.Text = "" Then
    MsgBox "Nomor Polisi belum di-input", vbExclamation, "kesalahan"
        ElseIf kd_jenis = "" Then
            MsgBox "Jenis kendaraan belum dipilih", vbExclamation, "kesalahan"
                ElseIf jum > 0 Then
                    MsgBox "No Polisi ini sudah masuk dan belum keluar", vbExclamation, "kesalahan"
Else
    adoparkir.Recordset.AddNew
    adoparkir.Recordset.Fields("NO_TIKET").Value = txtNoTiket.Text
    adoparkir.Recordset.Fields("TGL_MASUK").Value = txtTglMasuk.Text
    adoparkir.Recordset.Fields("ID_PETUGAS").Value = id_petugas
    adoparkir.Recordset.Fields("NO_POLISI").Value = txtNoPol.Text
    adoparkir.Recordset.Fields("KODE_JENIS").Value = kd_jenis
    adoparkir.Recordset.Fields("JAM_MASUK").Value = txtJamMasuk.Text
    adoparkir.Recordset.Fields("SUDAH_KELUAR").Value = "T"
    adoparkir.Recordset.update
    MsgBox "Data Parkir berhasil disimpan", vbInformation, "Pemberitahuan"
    
    Call kunci_form
    
     If DataEnvParkir.rsCommStrukMasuk.State = adStateOpen Then
        DataEnvParkir.rsCommStrukMasuk.Close
        DataEnvParkir.CommStrukMasuk nama_petugas, txtNoTiket.Text
        CetakStrukMasuk.Show
    Else
        DataEnvParkir.CommStrukMasuk nama_petugas, txtNoTiket.Text
        CetakStrukMasuk.Show
    End If
End If
End Sub

Private Sub cmdTambah_Click()
    Call autogen_tiket
    Call bersih
    
    txtTglMasuk.Text = Date
    txtJamMasuk.Text = Time
    txtNoPol.Enabled = True
    OptMotor.Enabled = True
    OptMobil.Enabled = True
    cmdTambah.Enabled = False
    cmdSimpan.Enabled = True
    cmdBatal.Enabled = True
    txtNoPol.SetFocus
End Sub

Private Sub Form_Activate()
    lblnama.Caption = nama_petugas
End Sub

Private Sub Form_Load()
    lblnama.Caption = nama_petugas
    lbltanggal.Caption = Format$(Now, "dddd, mmmm dd, yyyy")
    Call kunci_form
End Sub

Private Sub OptMobil_Click()
adotarif.Refresh
Do Until adotarif.Recordset.EOF
    If adotarif.Recordset.Fields("KODE_JENIS").Value = "P02" Then
        lbltarifpertama.Caption = adotarif.Recordset.Fields("TARIF_JAM_P")
        lbltarifjam.Caption = adotarif.Recordset.Fields("TARIF_PERJAM")
        txtNoTiket.Text = "MBL" + Right(txtNoTiket.Text, 7)
        Exit Sub
    Else
        adotarif.Recordset.MoveNext
    End If
Loop

End Sub

Private Sub OptMotor_Click()
adotarif.Refresh
Do Until adotarif.Recordset.EOF
    If adotarif.Recordset.Fields("KODE_JENIS").Value = "P01" Then
        lbltarifpertama.Caption = adotarif.Recordset.Fields("TARIF_JAM_P")
        lbltarifjam.Caption = adotarif.Recordset.Fields("TARIF_PERJAM")
        txtNoTiket.Text = "MTR" + Right(txtNoTiket.Text, 7)
        Exit Sub
    Else
        adotarif.Recordset.MoveNext
    End If
Loop

End Sub

Private Sub Timer1_Timer()
    lbljam.Caption = Time
End Sub
