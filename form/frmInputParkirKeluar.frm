VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInputParkirKeluar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".::  Form Input Parkir Keluar ::."
   ClientHeight    =   8790
   ClientLeft      =   2235
   ClientTop       =   1350
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptTidak 
      Caption         =   "Tidak"
      Height          =   255
      Left            =   3000
      TabIndex        =   36
      Top             =   6720
      Width           =   975
   End
   Begin VB.OptionButton OptYa 
      Caption         =   "Ya"
      Height          =   255
      Left            =   2160
      TabIndex        =   35
      Top             =   6720
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   7680
      Width           =   8175
      Begin Jasa_Parkir.jcbutton cmdSimpanCetak 
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "Simpan && Cetak"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Klik Untuk Mneyimpan Data dan Mencetak Struk"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdKeluar 
         Height          =   495
         Left            =   6480
         TabIndex        =   33
         Top             =   240
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
      TabIndex        =   17
      Top             =   6240
      Width           =   2295
   End
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5953
      Orientation     =   0
      BackColor       =   14737632
      ColorGradient1  =   12632256
      ColorGradient2  =   0
      BorderColor     =   12632319
      ShowIcon        =   0   'False
      Caption         =   ""
      Icon            =   "frmInputParkirKeluar.frx":0000
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
      Begin VB.TextBox txtTglKeluar 
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
         Height          =   300
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtJamKeluar 
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
         Height          =   360
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblbiayaparkir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   40
         Top             =   3000
         Width           =   105
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Biaya Parkir:"
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
         TabIndex        =   39
         Top             =   3000
         Width           =   1140
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbltarifpertama 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   29
         Top             =   2280
         Width           =   105
      End
      Begin VB.Label lbltarifjam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   28
         Top             =   2640
         Width           =   105
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label Label5 
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
         Left            =   4320
         TabIndex        =   26
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Parkir"
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
         TabIndex        =   25
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbljeniskendaraan 
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
         Left            =   6000
         TabIndex        =   24
         Top             =   480
         Width           =   345
      End
      Begin VB.Label lbllamaparkir 
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
         TabIndex        =   23
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
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
         Left            =   3240
         TabIndex        =   22
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label lblketjam 
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
         Left            =   3840
         TabIndex        =   21
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label lblkethari 
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
         TabIndex        =   20
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblketharijam 
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
         Left            =   3840
         TabIndex        =   19
         Top             =   1920
         Width           =   345
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
         TabIndex        =   16
         Top             =   1200
         Width           =   1035
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
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   900
      End
      Begin VB.Label lblnotiket 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         DataField       =   "NO_TIKET"
         DataSource      =   "adoparkir"
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
         Left            =   2160
         TabIndex        =   13
         Top             =   480
         Width           =   405
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
         Left            =   2160
         TabIndex        =   12
         Top             =   840
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
         Left            =   2160
         TabIndex        =   11
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Keluar"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Keluar"
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
         Left            =   4320
         TabIndex        =   5
         Top             =   1200
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc adotarif 
      Height          =   330
      Left            =   7800
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   6720
      Top             =   1680
   End
   Begin VB.TextBox txtTotal 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "0"
      Top             =   7080
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc adoparkir 
      Height          =   375
      Left            =   5760
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM T_PARKIR WHERE SUDAH_KELUAR LIKE 'T'"
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
   Begin VB.Label lbldenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   5880
      TabIndex        =   38
      Top             =   6720
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Denda"
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
      Left            =   4560
      TabIndex        =   37
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ada Struk ?"
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
      TabIndex        =   34
      Top             =   6720
      Width           =   1020
   End
   Begin VB.Label Label7 
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
      TabIndex        =   18
      Top             =   6240
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "frmInputParkirKeluar.frx":039A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARKIR KELUAR"
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
      TabIndex        =   10
      Top             =   360
      Width           =   3015
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
      TabIndex        =   9
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
      Left            =   1680
      TabIndex        =   8
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
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Total Biaya"
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
      Left            =   240
      TabIndex        =   1
      Top             =   7080
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   2535
      Left            =   -480
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmInputParkirKeluar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private denda As Integer
Private adastruk As String
Sub awal()
    txtNoPol.Text = ""
    txtTotal.Text = ""
    txtTglKeluar.Text = ""
    txtJamKeluar.Text = ""
    lbllamaparkir.Caption = 0
    lbljammasuk.Caption = ""
    lblnotiket.Caption = ""
    lbljeniskendaraan.Caption = ""
    lbltglmasuk.Caption = ""
    lbltarifjam.Caption = 0
    lbltarifpertama.Caption = 0
    lblbiayaparkir.Caption = 0
    lblketjam.Caption = 0
    lbldenda.Caption = 0
    lbljam.Caption = 0
    lblkethari.Caption = 0
    OptYa.Value = False
    OptTidak.Value = False
    lblnama.Caption = nama_petugas
    lbltanggal.Caption = Format$(Now, "dddd, mmmm dd, yyyy")
    txtTotal.Enabled = False
    cmdSimpanCetak.Enabled = False
End Sub
Sub baca_tarif(kode As String)
    adotarif.Refresh
    adotarif.Recordset.MoveFirst
    adotarif.Recordset.Find "KODE_JENIS='" & kode & "'"

    If adotarif.Recordset.EOF Then
        MsgBox "Data Tarif tidak ditemukan"
    Else
        lbltarifpertama.Caption = adotarif.Recordset.Fields("TARIF_JAM_P").Value
        lbltarifjam.Caption = adotarif.Recordset.Fields("TARIF_PERJAM").Value
        lbldenda.Caption = adotarif.Recordset.Fields("DENDA").Value
        denda = adotarif.Recordset.Fields("DENDA").Value
        Exit Sub
    End If
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpanCetak_Click()
'validasi
If txtNoPol.Text = "" Then
    MsgBox "No Polisi belum di-input", vbExclamation, "Kesalahan"
        ElseIf adastruk = "X" Then
            MsgBox "Pilihan ada struk belum dipilih", vbExclamation, "kesalahan"
Else
    adoparkir.Recordset.Fields("JAM_KELUAR").Value = txtJamKeluar.Text
    adoparkir.Recordset.Fields("TGL_KELUAR").Value = txtTglKeluar.Text
    adoparkir.Recordset.Fields("LAMA_PARKIR").Value = lbllamaparkir.Caption
    adoparkir.Recordset.Fields("BIAYA_PARKIR").Value = lblbiayaparkir.Caption
    adoparkir.Recordset.Fields("TOTAL_BAYAR").Value = txtTotal.Text
    adoparkir.Recordset.Fields("SUDAH_KELUAR").Value = "Y"
    adoparkir.Recordset.Fields("ADA_STRUK").Value = adastruk
    adoparkir.Recordset.update
    MsgBox "Data Parkir berhasil diupdate", vbInformation, "Pemberitahuan"
    
    If DataEnvParkir.rsCommLapParkir.State = adStateOpen Then
        DataEnvParkir.rsCommLapParkir.Close
        DataEnvParkir.CommLapParkir (lblnotiket.Caption)
        CetakStruk.Show
    Else
        DataEnvParkir.CommLapParkir (lblnotiket.Caption)
        CetakStruk.Show
    End If
        Call awal
        txtNoPol.SetFocus
End If
End Sub

Private Sub Form_Activate()
    lblnama.Caption = nama_petugas
    adastruk = "X"
End Sub

Private Sub Form_Load()
Call awal
End Sub

Private Sub OptTidak_Click()
If OptTidak.Value = True Then
    lbldenda.Caption = denda
    txtTotal.Text = Val(lbldenda.Caption) + Val(lblbiayaparkir.Caption)
    adastruk = "T"
End If
End Sub

Private Sub OptYa_Click()
If OptYa.Value = True Then
    lbldenda.Caption = 0
    txtTotal.Text = lblbiayaparkir.Caption
    adastruk = "Y"
End If
End Sub

Private Sub Timer1_Timer()
    lbljam.Caption = Time
End Sub

Private Sub txtNoPol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    Dim cari As String
    Dim lama_parkir, per_jam, menit As Integer
    Dim lama_hari, lama_jam, total_jamhari
    cari = txtNoPol.Text
    
    If cari = "" Then
        MsgBox "Nomor Tiket belum di input", vbExclamation, "Kesalahan"
    Else
    txtTglKeluar.Text = Date
    txtJamKeluar.Text = Time
    
    adoparkir.Refresh
    adoparkir.Recordset.MoveFirst
    adoparkir.Recordset.Find "NO_POLISI='" & cari & "'"

    If adoparkir.Recordset.EOF Then
        MsgBox "Data tidak ditemukan", vbExclamation, "kesalahan"
    Else
        If adoparkir.Recordset.Fields("SUDAH_KELUAR").Value = "Y" Then
            MsgBox "Data ini sudah tidak valid", vbExclamation, "Peringatan"
        Else
        
            If adoparkir.Recordset.Fields("KODE_JENIS").Value = "P01" Then
                lbljeniskendaraan.Caption = "MOTOR"
                baca_tarif ("P01")
            Else
                lbljeniskendaraan.Caption = "MOBIL"
                baca_tarif ("P02")
            End If
        
                lama_hari = CDate(txtTglKeluar.Text) - CDate(lbltglmasuk.Caption)
                lama_jam = CDate(CDate(txtJamKeluar.Text) - CDate(lbljammasuk.Caption))
                jam = Format(lama_jam, "H")
                menit = Format(lama_jam, "nn")
                
                If menit > 0 And jam > 0 Then
                    jam = jam + 1
                End If
                
                total_jamhari = lama_hari * 24
        
                lama_parkir = total_jamhari + jam
        
            If lama_hari > 0 Then
                lblkethari.Caption = "(" + CStr(lama_hari) + " Hari  - "
                lblketjam.Caption = ""
                lblketharijam.Caption = CStr(jam) + " Jam )"
            Else
                lblkethari.Caption = "-"
                lblketjam.Caption = "(" + CStr(lama_jam) + ")"
                lblketharijam.Caption = ""
            End If
        
                cmdSimpanCetak.Enabled = True
                lbllamaparkir.Caption = lama_parkir
                
                per_jam = lama_parkir - 1
                
                If lama_parkir > 1 Then
                    lblbiayaparkir.Caption = per_jam * Val(lbltarifjam.Caption) + Val(lbltarifpertama.Caption)
                Else
                    lblbiayaparkir.Caption = lbltarifpertama.Caption
                    lbllamaparkir.Caption = 1
                End If
        End If
    End If
    End If
End If
End Sub
