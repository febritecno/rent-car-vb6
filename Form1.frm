VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form1_home 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rental Mobil"
   ClientHeight    =   10245
   ClientLeft      =   2490
   ClientTop       =   690
   ClientWidth     =   15525
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":1A162
   ScaleHeight     =   10245
   ScaleWidth      =   15525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   2400
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu fl 
      Caption         =   "FILE"
      Begin VB.Menu cmd_ganti 
         Caption         =   "Ganti Pengguna"
         Checked         =   -1  'True
      End
      Begin VB.Menu cmd_program 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu cmd_mobil 
      Caption         =   "MOBIL"
      Begin VB.Menu cmd_merk 
         Caption         =   "Merk Mobil"
      End
      Begin VB.Menu cmd_daftarmobil 
         Caption         =   "Daftar Mobil"
      End
   End
   Begin VB.Menu cmd_anggota 
      Caption         =   "ANGGOTA"
   End
   Begin VB.Menu cmd_transaksi 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu cmd_penyewaan 
         Caption         =   "Penyewaan"
      End
      Begin VB.Menu cmd_Pengembalian 
         Caption         =   "Pengembalian"
      End
   End
   Begin VB.Menu lp 
      Caption         =   "LAPORAN"
      Begin VB.Menu pjm 
         Caption         =   "Cetak Laporan"
      End
      Begin VB.Menu cmobil 
         Caption         =   "Cetak Data Mobil"
      End
   End
   Begin VB.Menu cmd_pengaturan 
      Caption         =   "SETTING"
      Begin VB.Menu cmd_pengguna 
         Caption         =   "Pengguna"
      End
      Begin VB.Menu abt 
         Caption         =   "About"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1_home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abt_Click()
Dialog.Show
End Sub

Private Sub cmd_anggota_Click()
Form4_anggota.Show
End Sub

Private Sub cmd_daftarmobil_Click()
Form3_mobil.Show
End Sub

Private Sub cmd_ganti_Click()
If MsgBox("Anda Akan Ganti Pengguna?", vbYesNo + vbInformation, "Informasi") = vbYes Then
    Form5_login.Show
    Form1_home.Hide
End If
End Sub

Private Sub cmd_merk_Click()
Form2_merk.Show
End Sub

Private Sub cmd_Pengembalian_Click()
Form7_pengembalian.Show
End Sub

Private Sub cmd_pengguna_Click()
Form0_user.Show
End Sub

Private Sub cmd_penyewaan_Click()
Form6_penyewaan.Show
End Sub

Private Sub cmd_program_Click()
If MsgBox("Keluar Program?", vbYesNo + vbInformation, "Informasi") = vbYes Then
    End
End If
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Time
End Sub


Private Sub cmobil_Click()
CR.ReportFileName = App.Path & "\mobil.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub pjm_Click()

CR.ReportFileName = App.Path & "\trs.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub
