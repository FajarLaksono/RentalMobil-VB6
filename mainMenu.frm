VERSION 5.00
Begin VB.MDIForm mainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Rental Mobil Purwokerto V0.5"
   ClientHeight    =   9195
   ClientLeft      =   2235
   ClientTop       =   900
   ClientWidth     =   15255
   Icon            =   "mainMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mainMenu.frx":038A
   WindowState     =   2  'Maximized
   Begin VB.Menu menuPeminjaman 
      Caption         =   "Peminjaman"
      Begin VB.Menu menuPeminjamanMobil 
         Caption         =   "Peminjaman Mobil"
      End
      Begin VB.Menu menuPengembalianMobil 
         Caption         =   "Pengembalian Mobil"
      End
      Begin VB.Menu menuTableRental 
         Caption         =   "Table Rental"
      End
   End
   Begin VB.Menu menuAnggota 
      Caption         =   "Anggota"
      Begin VB.Menu menuPendaftaranAnggota 
         Caption         =   "Pendaftaran Anggota"
      End
      Begin VB.Menu menuTabelAnggota 
         Caption         =   "Tabel Anggota"
      End
   End
   Begin VB.Menu menuMobil 
      Caption         =   "Mobil"
      Begin VB.Menu menuMobilBaru 
         Caption         =   "Mobil Baru"
      End
      Begin VB.Menu menuTableMobil 
         Caption         =   "Tabel Mobil"
      End
   End
   Begin VB.Menu menuAkun 
      Caption         =   "Akun"
      Begin VB.Menu menuSetting 
         Caption         =   "Setting"
      End
      Begin VB.Menu menuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu menuAbout 
         Caption         =   "About Rental Mobil"
      End
   End
End
Attribute VB_Name = "mainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    navigation.Show
    loadingBarForMainMenu.Show
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
End Sub

Private Sub menuAbout_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    about.Show
End Sub

Private Sub menuLogout_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    login.Show
    loadingBarForMainMenu.loadingBarInisialisasi
    Unload Me
End Sub

Private Sub menuMobilBaru_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    mobilBaru.Show
End Sub

Private Sub menuPeminjamanMobil_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    peminjamanMobil.Show
End Sub

Private Sub menuPendaftaranAnggota_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    pendaftaranAnggota.Show
End Sub

Private Sub menuPengembalianMobil_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    pengembalian.Show
End Sub

Private Sub menuSetting_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    setting.Show
End Sub

Private Sub menuTabelAnggota_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    tableAnggota.Show
End Sub

Private Sub menuTableMobil_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    tableMobil.Show
End Sub

Private Sub menuTableRental_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    tableRental.Show
End Sub
