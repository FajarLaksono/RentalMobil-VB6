VERSION 5.00
Begin VB.Form navigation 
   BorderStyle     =   0  'None
   Caption         =   "navigation"
   ClientHeight    =   10530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   Icon            =   "navigation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameNavigation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   10575
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton btnPendaftaran 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Pendaftaran Anggota"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton btnPeminjaman 
         BackColor       =   &H80000004&
         Caption         =   "Peminjaman"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton btnPengembalian 
         BackColor       =   &H80000004&
         Caption         =   "Pengembalian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4080
         Width           =   2295
      End
      Begin VB.CommandButton btnTabelPenyewa 
         BackColor       =   &H80000004&
         Caption         =   "Tabel Rental"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CommandButton btnTabelMobil 
         BackColor       =   &H80000004&
         Caption         =   "Tabel Mobil"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton btnLogout 
         BackColor       =   &H80000004&
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   9240
         Width           =   2295
      End
      Begin VB.Image photoProfil 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label labelNIKpetugas 
         BackStyle       =   0  'Transparent
         Caption         =   "NIK Petugas Rental"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label labelNamaPetugasRental 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Petugas Rental"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
   End
End
Attribute VB_Name = "navigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogout_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    login.Show
    loadingBarForMainMenu.loadingBarInisialisasi
    mainMenu.Hide
End Sub

Private Sub btnPeminjaman_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    peminjamanMobil.Show
End Sub

Private Sub btnPendaftaran_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    pendaftaranAnggota.Show
End Sub

Private Sub btnPengembalian_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    pengembalian.Show
End Sub

Private Sub btnTabelMobil_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    tableMobil.Show
End Sub

Private Sub btnTabelPenyewa_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    tableRental.Show
End Sub

Private Sub Form_Load()
    'inisialisasi pojok kiri atas
    Me.Left = 0
    Me.Top = 0

    'setelah login, data dimasukan ke variabel yang berada pada sistem.
    'setelah itu kita ambil valuenya dan menampilkanya
    labelNIKpetugas.Caption = sistem.userNIK
    labelNamaPetugasRental.Caption = sistem.getName
    photoProfil.Picture = LoadPicture(App.Path + sistem.getPhoto)
    
    'set ToolTip
    btnLogout.ToolTipText = "Keluar"
    btnPeminjaman.ToolTipText = "Tampilkan Form Peminjaman"
    btnPendaftaran.ToolTipText = "Tampilkan Form Pendafaran"
    btnPengembalian.ToolTipText = "Tampilkan Form Pengembalian"
    btnTabelMobil.ToolTipText = "Tampilkan Tabel Mobil"
    btnTabelPenyewa.ToolTipText = "Tampilkan Tabel Penyewa"
    labelNamaPetugasRental.ToolTipText = "Nama : " + sistem.getName
    labelNIKpetugas.ToolTipText = "NIK : " + sistem.userNIK
    photoProfil.ToolTipText = "Klik untuk Pengaturan Akun"
End Sub

Private Sub photoProfil_Click()
    loadingBarForMainMenu.LoadingBarMainMenuTime_Timer
    loadingBarForMainMenu.loadingBarInisialisasi
    setting.Show
End Sub

