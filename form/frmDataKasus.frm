VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDataKasus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".::  Form Data Kasus ::."
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin Jasa_Parkir.YudhaFrame YudhaFrame1 
      Height          =   8295
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14631
      Orientation     =   0
      BackColor       =   14737632
      ColorGradient1  =   12632256
      ColorGradient2  =   0
      BorderColor     =   12632256
      ShowIcon        =   0   'False
      Caption         =   ""
      Icon            =   "frmDataKasus.frx":0000
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
      Begin Jasa_Parkir.jcbutton cmdKeluar 
         Height          =   495
         Left            =   10320
         TabIndex        =   8
         Top             =   7680
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
      Begin Jasa_Parkir.jcbutton cmdRefresh 
         Height          =   495
         Left            =   7080
         TabIndex        =   6
         Top             =   1080
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
         Caption         =   "Refresh Data"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
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
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   1080
         Width           =   4455
      End
      Begin Jasa_Parkir.jcbutton cmdCetak 
         Height          =   495
         Left            =   7080
         TabIndex        =   4
         Top             =   7680
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Cetak Data"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Klik Untuk Cetak"
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin MSDataGridLib.DataGrid dgKasus 
         Bindings        =   "frmDataKasus.frx":039A
         Height          =   5775
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10186
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin Jasa_Parkir.jcbutton cmdCari 
         Height          =   495
         Left            =   5640
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "Cari Data"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATA KASUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc adokasus 
      Height          =   375
      Left            =   360
      Top             =   5280
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
      RecordSource    =   "T_KASUS"
      Caption         =   "adokasus"
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
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   120
      Picture         =   "frmDataKasus.frx":03B1
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9015
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   -360
      Width           =   3135
   End
End
Attribute VB_Name = "frmDataKasus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub load_kasus()
    adokasus.CommandType = adCmdUnknown
    adokasus.RecordSource = "SELECT T_PARKIR.NO_POLISI, T_KASUS.NO_TIKET, T_KASUS.ID_KASUS, T_KASUS.KETERANGAN, T_KASUS.STATUS FROM T_PARKIR INNER JOIN T_KASUS ON T_PARKIR.NO_TIKET = T_KASUS.NO_TIKET"
    adokasus.Refresh
    dgKasus.Refresh
End Sub
Private Sub cmdCari_Click()
    Dim cari As String
    cari = txtCari.Text
    
    If cari = "" Then
        MsgBox "Nomor Tiket belum di input", vbExclamation, "Kesalahan"
    Else
        adokasus.Refresh
        adokasus.Recordset.MoveFirst
        adokasus.Recordset.Find "NO_TIKET='" & cari & "'"

        If adokasus.Recordset.EOF Then
            MsgBox "Data kasus tidak ditemukan", vbExclamation, "Kesalahan"
        Else
            dgKasus.Refresh
        End If
    End If
End Sub

Private Sub cmdCetak_Click()
    CetakLapKasus.Show
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call load_kasus
End Sub

Private Sub Form_Active()
    Call load_kasus
End Sub

Private Sub Form_Load()
    Call load_kasus
End Sub

Private Sub txtCari_Change()

End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdCari_Click
End If
End Sub
