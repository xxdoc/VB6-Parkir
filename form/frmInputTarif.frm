VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInputTarif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".:: Form Input Tarif dan Denda ::."
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Navigasi Data"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   5535
      Begin Jasa_Parkir.jcbutton cmdPrev 
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Caption         =   "< Previous"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Jasa_Parkir.jcbutton cmdNext 
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Caption         =   "Next >"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Info"
         TooltipBackColor=   0
         ColorScheme     =   2
      End
   End
   Begin VB.TextBox txtDenda 
      DataField       =   "DENDA"
      DataSource      =   "adotarif"
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
   End
   Begin Jasa_Parkir.jcbutton cmdSimpan 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      ToolTip         =   "Klik Untuk Menyimpan Tarif"
      TooltipType     =   1
      TooltipTitle    =   "Info"
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin MSAdodcLib.Adodc adotarif 
      Height          =   375
      Left            =   3480
      Top             =   1320
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db_parkir.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "T_JENIS_KENDARAAN"
      Caption         =   "Data"
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
   Begin VB.TextBox txtTarif 
      DataField       =   "TARIF_PERJAM"
      DataSource      =   "adotarif"
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtTarifPertama 
      DataField       =   "TARIF_JAM_P"
      DataSource      =   "adotarif"
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin Jasa_Parkir.jcbutton cmdEdit 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      ToolTip         =   "Klik Untuk Mengubah Tarif"
      TooltipType     =   1
      TooltipTitle    =   "Info"
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Jasa_Parkir.jcbutton cmdKeluar 
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      ToolTip         =   "Klik Untuk Kleuar Dari Form"
      TooltipType     =   1
      TooltipTitle    =   "Info"
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "frmInputTarif.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbljenis 
      AutoSize        =   -1  'True
      Caption         =   "N/A"
      DataField       =   "JENIS"
      DataSource      =   "adotarif"
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
      TabIndex        =   6
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tarif Per-Jam"
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
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tarif Jam Pertama"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT TARIF && DENDA"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4185
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1215
      Left            =   -240
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmInputTarif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub kunci()
    cmdSimpan.Enabled = False
    cmdedit.Enabled = True
    txtTarifPertama.Enabled = False
    txtTarif.Enabled = False
    txtDenda.Enabled = False
End Sub
Private Sub cmdedit_Click()
    cmdSimpan.Enabled = True
    cmdedit.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    txtTarifPertama.Enabled = True
    txtTarif.Enabled = True
    txtDenda.Enabled = True
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
If adotarif.Recordset.EOF Then
    adotarif.Recordset.MoveLast
Else
    adotarif.Recordset.MoveNext
End If
End Sub

Private Sub cmdPrev_Click()
If adotarif.Recordset.BOF Then
    adotarif.Recordset.MoveFirst
Else
    adotarif.Recordset.MovePrevious
End If
End Sub

Private Sub cmdSimpan_Click()
'validasi
If txtTarifPertama.Text = "" Then
    MsgBox "Tarif Jam Pertama belum di-isi", vbExclamation, "kesalahan"
    ElseIf txtTarif.Text = "" Then
        MsgBox "Tarif Per-Jam belum di-isi", vbExclamation, "kesalahan"
        ElseIf txtDenda.Text = "" Then
            MsgBox "Denda belum di-isi", vbExclamation, "kesalahan"
                ElseIf IsNumeric(txtTarifPertama.Text) = False Then
                    MsgBox "Input Harus Angka", vbExclamation, "kesalahan"
                        ElseIf IsNumeric(txtTarif.Text) = False Then
                            MsgBox "Input Harus Angka", vbExclamation, "kesalahan"
                                ElseIf IsNumeric(txtDenda.Text) = False Then
                                    MsgBox "Input Harus Angka", vbExclamation, "kesalahan"
Else
    adotarif.Recordset.Fields("TARIF_JAM_P").Value = txtTarifPertama.Text
    adotarif.Recordset.Fields("TARIF_PERJAM").Value = txtTarif.Text
    adotarif.Recordset.Fields("DENDA").Value = txtDenda.Text
    adotarif.Recordset.update
    Call kunci
    MsgBox "Data tarif dan Denda berhasil disimpan", vbInformation, "Pemberitahuan"
    
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
End Sub

Private Sub Form_Load()
    Call kunci
End Sub
