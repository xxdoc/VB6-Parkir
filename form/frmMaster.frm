VERSION 5.00
Begin VB.MDIForm frmMaster 
   BackColor       =   &H8000000C&
   Caption         =   ".:: Aplikasi Jasa Parkir ::."
   ClientHeight    =   4485
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9900
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mn_beranda 
      Caption         =   "Beranda"
      Begin VB.Menu mn_ganti_pass 
         Caption         =   "Ganti Password"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnu_pemisah_1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_logout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mn_keluar 
         Caption         =   "Keluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mn_input 
      Caption         =   "Input"
      Begin VB.Menu mn_input_parkir 
         Caption         =   "Input Parkir Masuk"
         Shortcut        =   ^I
      End
      Begin VB.Menu mn_parkir_out 
         Caption         =   "Input Parkir Keluar"
         Shortcut        =   ^O
      End
      Begin VB.Menu mn_pemisah_3 
         Caption         =   "-"
      End
      Begin VB.Menu mn_input_kasus 
         Caption         =   "Input Kasus"
         Shortcut        =   ^K
      End
      Begin VB.Menu mn_pemisah_2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_petugas 
         Caption         =   "Input Petugas"
         Shortcut        =   ^P
      End
      Begin VB.Menu mn_tarif 
         Caption         =   "Input Tarif &&  Denda Parkir"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mn_data 
      Caption         =   "Data"
      Begin VB.Menu mn_data_parkir 
         Caption         =   "Data Parkir"
      End
      Begin VB.Menu mn_data_kasus 
         Caption         =   "Data Kasus"
      End
      Begin VB.Menu mn_pemisah_4 
         Caption         =   "-"
      End
      Begin VB.Menu mn_data_pendapatan 
         Caption         =   "Data Pendapatan"
      End
   End
   Begin VB.Menu mn_laporan 
      Caption         =   "Laporan"
      Begin VB.Menu mn_lap_kasus 
         Caption         =   "Laporan Kasus"
      End
      Begin VB.Menu mn_lap_parkir 
         Caption         =   "Laporan Parkir (All)"
      End
   End
   Begin VB.Menu mn_bantuan 
      Caption         =   "Bantuan"
      Begin VB.Menu mn_cara 
         Caption         =   "Cara Penggunaan"
         Shortcut        =   ^H
      End
      Begin VB.Menu mn_tentang 
         Caption         =   "Tentang"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub kunci_menu_master()
    mn_ganti_pass.Enabled = False
    mn_input_kasus.Enabled = False
    mn_input_parkir.Enabled = False
    mn_lap_parkir.Enabled = False
    mn_petugas.Enabled = False
    mn_tarif.Enabled = False
    mn_logout.Enabled = False
    mn_parkir_out.Enabled = False
    mn_data_parkir.Enabled = False
    mn_data_kasus.Enabled = False
    mn_lap_kasus.Enabled = False
    mn_data_pendapatan.Enabled = False
End Sub
Private Sub MDIForm_Load()
    Me.Picture = LoadPicture(App.Path & "\gambar\Background.jpg")
    frmLogin.Show
    Call kunci_menu_master
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim tanya
tanya = MsgBox("Anda yakin akan keluar dari Aplikasi ?", vbInformation + vbYesNo, "Konfirmasi")
If tanya = vbYes Then
    Unload Me
Else
    Cancel = 1
    Me.Show
        If nama_petugas = "" Then
            frmLogin.Show
        End If
End If
End Sub

Private Sub mn_cara_Click()
    frmCaraPenggunaan.Show
End Sub

Private Sub mn_data_kasus_Click()
    frmDataKasus.Show
End Sub

Private Sub mn_data_parkir_Click()
    frmDataParkir.Show
End Sub

Private Sub mn_data_pendapatan_Click()
    frmDataPendapatan.Show
End Sub

Private Sub mn_ganti_pass_Click()
    frmGantiPass.Show
End Sub

Private Sub mn_input_kasus_Click()
    frmInputKasus.Show
End Sub

Private Sub mn_input_parkir_Click()
    frmInputParkirMasuk.Show
End Sub

Private Sub mn_keluar_Click()
    Unload Me
End Sub

Private Sub mn_lap_kasus_Click()
    CetakLapKasus.Show
End Sub

Private Sub mn_lap_parkir_Click()
    CetakLapParkir.Show
End Sub

Private Sub mn_login_Click()
    frmLogin.Show vbModal
End Sub

Private Sub mn_logout_Click()
    frmInputParkirMasuk.Hide
    frmInputParkirKeluar.Hide
    frmInputPetugas.Hide
    frmInputTarif.Hide
    frmCaraPenggunaan.Hide
    frmLogin.Show
    Call kunci_menu_master
    Me.Caption = ".:: Aplikasi Jasa Parkir ::."
End Sub

Private Sub mn_parkir_out_Click()
    frmInputParkirKeluar.Show
End Sub

Private Sub mn_petugas_Click()
    frmInputPetugas.Show
End Sub

Private Sub mn_tarif_Click()
    frmInputTarif.Show
End Sub

Private Sub mn_tentang_Click()
    frmAbout.Show vbModal
End Sub
