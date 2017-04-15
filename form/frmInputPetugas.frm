VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInputPetugas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: Form Input Petugas ::."
   ClientHeight    =   9495
   ClientLeft      =   6405
   ClientTop       =   1155
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   8760
   Begin VB.TextBox txtCari 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   8415
      Begin Jasa_Parkir.jcbutton cmdTambah 
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
         _extentx        =   1931
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputPetugas.frx":0000
         backcolor       =   14935011
         caption         =   "Tambah"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Menambah Data Petugas Baru"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdEdit 
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1095
         _extentx        =   1931
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputPetugas.frx":0028
         backcolor       =   14935011
         caption         =   "Edit"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Mengedit Data Petugas Baru"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdSimpan 
         Height          =   495
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputPetugas.frx":0050
         backcolor       =   14935011
         caption         =   "Simpan"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Menyimpan Data Petugas Baru"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdBatal 
         Height          =   495
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputPetugas.frx":0078
         backcolor       =   14935011
         caption         =   "Batal"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Membatalkan Input Data Petugas Baru"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdHapus 
         Height          =   495
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputPetugas.frx":00A0
         backcolor       =   14935011
         caption         =   "Hapus"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Menghapus Data Petugas Baru"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdBersih 
         Height          =   495
         Left            =   2640
         TabIndex        =   27
         Top             =   240
         Width           =   1215
         _extentx        =   2143
         _extenty        =   873
         buttonstyle     =   3
         font            =   "frmInputPetugas.frx":00C8
         backcolor       =   14935011
         caption         =   "Bersih"
         pictureeffectonover=   0
         pictureeffectondown=   0
         captioneffects  =   0
         tooltiptitle    =   "Info"
         tooltip         =   "Klik Untuk Membatalkan Input Data Petugas Baru"
         tooltiptype     =   1
         tooltipbackcolor=   0
         colorscheme     =   2
      End
   End
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   8415
      _extentx        =   14843
      _extenty        =   7223
      orientation     =   0
      backcolor       =   14737632
      colorgradient1  =   12632256
      colorgradient2  =   0
      bordercolor     =   12632256
      showicon        =   0   'False
      caption         =   "Create User"
      icon            =   "frmInputPetugas.frx":00F0
      forecolor       =   16777215
      font            =   "frmInputPetugas.frx":048C
      Begin VB.TextBox txtPass 
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
         Height          =   375
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   20
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txtUsername 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtNoTelp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtAlamat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox txtNama 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   11
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Opsional,jika tidak di-isi maka password otomatis adalah 1234"
         Height          =   195
         Left            =   1800
         TabIndex        =   24
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Left            =   1440
         TabIndex        =   23
         Top             =   3120
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   22
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
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
         TabIndex        =   21
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label lblidpetugas 
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
         Left            =   1800
         TabIndex        =   18
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         TabIndex        =   17
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Telp"
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
         TabIndex        =   16
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         TabIndex        =   15
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Top             =   960
         Width           =   555
      End
   End
   Begin MSDataGridLib.DataGrid GridPetugas 
      Bindings        =   "frmInputPetugas.frx":04B8
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   7440
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adopetugas 
      Height          =   375
      Left            =   6480
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "T_PETUGAS"
      Caption         =   "adopetugas"
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
   Begin Jasa_Parkir.jcbutton cmdCari 
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      buttonstyle     =   3
      font            =   "frmInputPetugas.frx":04D1
      backcolor       =   14935011
      caption         =   "Cari"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltiptitle    =   "Info"
      tooltip         =   "Klik Untuk Mencari Data Petugas Baru"
      tooltiptype     =   1
      tooltipbackcolor=   0
      colorscheme     =   2
   End
   Begin Jasa_Parkir.jcbutton cmdKeluar 
      Height          =   495
      Left            =   6960
      TabIndex        =   25
      Top             =   6840
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      buttonstyle     =   3
      font            =   "frmInputPetugas.frx":04F9
      backcolor       =   14935011
      caption         =   "Keluar"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltiptitle    =   "Info"
      tooltip         =   "Klik Untuk Keluar Form"
      tooltiptype     =   1
      tooltipbackcolor=   0
      colorscheme     =   2
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ketik Nama Petugas"
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
      TabIndex        =   26
      Top             =   6840
      Width           =   1830
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT PETUGAS"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "frmInputPetugas.frx":0521
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1335
      Left            =   -240
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "frmInputPetugas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    txtNama.Text = ""
    txtAlamat.Text = ""
    txtNoTelp.Text = ""
    txtUsername.Text = ""
    txtPass.Text = ""
End Sub
Sub kunci_form_petugas()
    cmdTambah.Enabled = True
    cmdedit.Enabled = False
    cmdBersih.Enabled = False
    
    txtNama.Enabled = False
    txtAlamat.Enabled = False
    txtNoTelp.Enabled = False
    txtUsername.Enabled = False
    txtPass.Enabled = False
    
    cmdBatal.Enabled = False
    cmdSimpan.Enabled = False
    cmdHapus.Enabled = True
End Sub

Public Sub buka_form_petugas()
    txtNama.Enabled = True
    txtAlamat.Enabled = True
    txtNoTelp.Enabled = True
    txtUsername.Enabled = True
    txtPass.Enabled = True
    
    cmdBatal.Enabled = True
    cmdSimpan.Enabled = True
    cmdBersih.Enabled = True
    
    cmdTambah.Enabled = False
    cmdedit.Enabled = False
    cmdHapus.Enabled = False
End Sub

Private Sub cmdBatal_Click()
    Call bersih
    Call kunci_form_petugas
End Sub

Private Sub cmdBersih_Click()
    Call bersih
End Sub

Private Sub cmdCari_Click()
Dim cari As String
cari = txtCari.Text

adopetugas.Recordset.MoveFirst
adopetugas.Recordset.Find "NAMA_PETUGAS='" & cari & "'"

If adopetugas.Recordset.EOF Then
    MsgBox "Data Petugas tidak ditemukan", vbInformation, "Informasi"
    GridPetugas.Refresh
    lblidpetugas.Caption = ""
    txtNama.Text = ""
    txtAlamat.Text = ""
    txtNoTelp.Text = ""
    txtUsername.Text = ""
    txtPass.Text = ""
Else
    GridPetugas.Refresh
    lblidpetugas.Caption = adopetugas.Recordset!id_petugas
    txtNama.Text = adopetugas.Recordset!nama_petugas
    txtAlamat.Text = adopetugas.Recordset!ALAMAT
    txtNoTelp.Text = adopetugas.Recordset!NO_TELP
    txtUsername.Text = adopetugas.Recordset!UserName
    txtPass.Text = adopetugas.Recordset!Password
    cmdedit.Enabled = True
    cmdHapus.Enabled = True
    cmdTambah.Enabled = False
End If
End Sub

Private Sub cmdedit_Click()
    Call buka_form_petugas
    cmdBersih.Enabled = True
    cmdBatal.Enabled = True
End Sub

Private Sub cmdHapus_Click()
Dim pass, isadmin As String
Dim tanya

isadmin = adopetugas.Recordset.Fields("IS_ADMIN").Value

If isadmin = "Y" Then
    MsgBox "Data Admin tidak bisa dihapus", vbExclamation, "Kesalahan"
Else
    tanya = MsgBox("Anda yakin ingin menghapus data Petugas ?", vbYesNo, "Konfirmasi")

    If tanya = vbYes Then
        adopetugas.Recordset.Delete
        adopetugas.Refresh
        GridPetugas.Refresh
    End If
End If
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim pass, isadmin As String

If txtPass.Text = "" Then
    pass = "1234"
Else
    pass = txtPass.Text
End If

If IsNumeric(txtNoTelp.Text) = False Then
            MsgBox "Nomor Telepon harus di-isi dengan angka", vbExclamation, "Kesalahan"
                ElseIf Len(txtUsername.Text) < 3 Then
                     MsgBox "Username minimal 3 karakter", vbExclamation, "Kesalahan"
                        ElseIf Len(pass) < 4 Then
                            MsgBox "Password minimal 4 karakter", vbExclamation, "Kesalahan"
Else
    adopetugas.Refresh
    adopetugas.Recordset.Find "USERNAME='" & txtUsername.Text & "'"

    If adopetugas.Recordset.EOF Then
        adopetugas.Recordset.AddNew
        adopetugas.Recordset.Fields("NAMA_PETUGAS").Value = txtNama.Text
        adopetugas.Recordset.Fields("ALAMAT").Value = txtAlamat.Text
        adopetugas.Recordset.Fields("NO_TELP").Value = txtNoTelp.Text
        adopetugas.Recordset.Fields("USERNAME").Value = txtUsername.Text
        adopetugas.Recordset.Fields("PASSWORD").Value = pass
        adopetugas.Recordset.Fields("IS_ADMIN").Value = "T"
        adopetugas.Recordset.update
        MsgBox "Data Petugas berhasil disimpan", vbInformation, "Pemberitahuan"
    Else
        isadmin = adopetugas.Recordset.Fields("IS_ADMIN").Value
        
        If isadmin = "Y" Then
            MsgBox "Data Admin tidak bisa diubah via Aplikasi", vbExclamation, "Kesalahan"
        Else
        adopetugas.Recordset.Fields("NAMA_PETUGAS").Value = txtNama.Text
        adopetugas.Recordset.Fields("ALAMAT").Value = txtAlamat.Text
        adopetugas.Recordset.Fields("NO_TELP").Value = txtNoTelp.Text
        adopetugas.Recordset.Fields("USERNAME").Value = txtUsername.Text
        adopetugas.Recordset.Fields("PASSWORD").Value = pass
        adopetugas.Recordset.Fields("LAST_UPDATE").Value = Now
        adopetugas.Recordset.Fields("IS_ADMIN").Value = "T"
        adopetugas.Recordset.update
        MsgBox "Data Petugas berhasil diupdate", vbInformation, "Pemberitahuan"
        End If
    End If
End If

adopetugas.Refresh
Call kunci_form_petugas
Call bersih
End Sub

Private Sub cmdTambah_Click()
    Call buka_form_petugas
    Call bersih
    txtNama.SetFocus
End Sub

Private Sub Form_Load()
    Call kunci_form_petugas
    cmdedit.Enabled = False
    cmdHapus.Enabled = False
End Sub
Private Sub txtCari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdCari_Click
    End If
End Sub
