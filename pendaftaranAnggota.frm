VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form pendaftaranAnggota 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendaftaran Anggota"
   ClientHeight    =   7155
   ClientLeft      =   6690
   ClientTop       =   2340
   ClientWidth     =   8685
   Icon            =   "pendaftaranAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8685
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00C0C0C0&
      Height          =   7335
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   8655
      Begin MSAdodcLib.Adodc adodcPendaftaranAnggota 
         Height          =   330
         Left            =   120
         Top             =   6600
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcPendaftaranAnggota"
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
      Begin MSComDlg.CommonDialog commonDialog 
         Left            =   7080
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton buttonPilihFoto 
         Caption         =   "Pilih Foto"
         Height          =   435
         Left            =   6360
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox textNomerKTP 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox textNama 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox textTempatTanggalLahir 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   2
         Top             =   1200
         Width           =   3495
      End
      Begin VB.OptionButton optionLakiLaki 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Laki - laki"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optionPerempuan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Perempuan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox textNoTelepon 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   5
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox textAlamat 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   7
         Top             =   2625
         Width           =   5685
      End
      Begin VB.TextBox textRtRw 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   8
         Top             =   3120
         Width           =   5685
      End
      Begin VB.TextBox textKelDesa 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   9
         Top             =   3600
         Width           =   5685
      End
      Begin VB.TextBox textKecamatan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   10
         Top             =   4080
         Width           =   5685
      End
      Begin VB.TextBox textKabupaten 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   11
         Top             =   4560
         Width           =   5685
      End
      Begin VB.TextBox textNoSim 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   14
         Top             =   6000
         Width           =   5685
      End
      Begin VB.TextBox textKodePos 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   12
         Top             =   5040
         Width           =   5685
      End
      Begin VB.TextBox textPekerjaan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   13
         Top             =   5520
         Width           =   5685
      End
      Begin VB.CommandButton buttonSimpan 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         Height          =   375
         Left            =   7200
         TabIndex        =   15
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Image imageFotoAnggota 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   6360
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label labelNomerKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nomor KTP "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   285
         Width           =   2175
      End
      Begin VB.Label labelNama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   765
         Width           =   2175
      End
      Begin VB.Label labelTempatTanggalLahir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tempat / Tanggal Lahir"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label labelJenisKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jenis Kelamin"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label labelAlamat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alamat"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label labekNoTelepon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No Telepon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label labelNoSIM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No SIM"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6030
         Width           =   2175
      End
      Begin VB.Label labelPekerjaan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pekerjaan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   5550
         Width           =   2175
      End
      Begin VB.Label labelRtRw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "RT / RW"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label labelKelDesa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kel / Desa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label labelKecamatan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kecamatan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label labelKabupaten 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kabupaten"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label labelKodePos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kode Pos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5055
         Width           =   2175
      End
   End
End
Attribute VB_Name = "pendaftaranAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variabels
Dim defaultFotoAnggota As String
Dim alamatFotoAnggota As String
Dim tanggal As String
Dim waktu As String

Function setConnection()
    adodcPendaftaranAnggota.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcPendaftaranAnggota.RecordSource = "select * from anggota" 'SQL
End Function

Private Sub buttonPilihFoto_Click()
    'dibawah adalah cara untuk mendapatkan file, kami buat hanya bisa memilih file JPEG untuk gambar
    commonDialog.FileName = "" 'inisialisasi
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set format pengambilan file
    commonDialog.ShowOpen 'buka dialog pemilihan dari windows
    alamatFotoAnggota = commonDialog.FileName 'ambil alamat gambar ke variabel alamatFotoAnggota
    
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatFotoAnggota)) < 1 Then
        Exit Sub
    End If
    'set imageFotoAnggota sesuai gambar
    imageFotoAnggota.Picture = LoadPicture(alamatFotoAnggota)
End Sub

Private Sub buttonSimpan_Click()
    On Error GoTo errHandler

    'cek semua pengisian
    If alamatFotoAnggota = defaultFotoAnggota Or commonDialog.FileName = "" Or commonDialog.FileTitle = "" Then
        MsgBox "Foto Calon Anggota Belum Dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textNomerKTP.Text = "" Then
        MsgBox "Informasi Nomer KTP Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textNama.Text = "" Then
        MsgBox "Informasi Nama Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textTempatTanggalLahir.Text = "" Then
        MsgBox "Informasi Tanggal Lahir Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textNoTelepon.Text = "" Then
        MsgBox "Informasi No Telepon Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textAlamat.Text = "" Then
        MsgBox "Informasi Alamat Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textRtRw.Text = "" Then
        MsgBox "Informasi RT/RW Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textKelDesa.Text = "" Then
        MsgBox "Informasi Kel/Desa Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textKecamatan.Text = "" Then
        MsgBox "Informasi Kecamatan Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textKabupaten.Text = "" Then
        MsgBox "Informasi Kabupaten Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
        
    If textKodePos.Text = "" Then
        MsgBox "Informasi Kode Pos Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textPekerjaan.Text = "" Then
        MsgBox "Informasi Pekerjaan Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textNoSim.Text = "" Then
        MsgBox "Informasi Nomer SIM Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If optionLakiLaki.Value = False And optionPerempuan.Value = False Then
        MsgBox "Informasi Jenis Kelamin Calon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
        
    'Start of Pengolahan gambar
    'dibawah ini adalah cara untuk mengelolah gambar.
    'bukan hanya kita memilih melalui commondialog diatas. tapi kita juga mengcopy dan menaruhnya dalam jangkauan program
    'mengubah nama, berguna untuk mengantisipasi duplikasi gambar dengan file yang sudah disatukan dalam forder jangkauan program
    'untuk mengantisipasi duplikasi kami menambahkan tanggal+waktu pada nama file.
    'karena tanggal dan waktu tidak akan pernah sama dan terus maju.
    tanggal = Format(Date, "d-mmmm-yyyy")
    waktu = Format(Time, "h-m-s")
    'jika kita hanya menggunakan tanggal = date(), date() akan menghasilkan format/bentuk text seperti ini d/mmmm/yyyy.
    'dan pada windows kita tidak boleh menggunakan karakter / pada nama file.
    'maka dibuat "d-mmmm-yyyy".
    
    'menyatukan string
    namaFile = sistem.userNIK + tanggal + waktu

    'mencopy, memindahkan ke folder jangkauan program dan mengubah nama.
    FileCopy commonDialog.FileName, App.Path + "/images/photoAnggota/" + namaFile + commonDialog.FileTitle
    'menyimpannya dalam variabel
    alamatFotoAnggota = "/images/photoAnggota/" + namaFile + commonDialog.FileTitle
    'End of pengolahanGambar
    
    'menyimpan dalam database
    adodcPendaftaranAnggota.Recordset.AddNew
        adodcPendaftaranAnggota.Recordset!KTP = textNomerKTP.Text
        adodcPendaftaranAnggota.Recordset!nama_anggota = textNama.Text
        adodcPendaftaranAnggota.Recordset!tempat_tanggal_lahir = textTempatTanggalLahir.Text
        adodcPendaftaranAnggota.Recordset!no_telp = textNoTelepon.Text
        adodcPendaftaranAnggota.Recordset!alamat = textAlamat.Text
        adodcPendaftaranAnggota.Recordset!rt_rw = textRtRw.Text
        adodcPendaftaranAnggota.Recordset!kelDesa = textKelDesa.Text
        adodcPendaftaranAnggota.Recordset!kec = textKecamatan.Text
        adodcPendaftaranAnggota.Recordset!kab = textKabupaten.Text
        adodcPendaftaranAnggota.Recordset!kode_pos = textKodePos.Text
        adodcPendaftaranAnggota.Recordset!pekerjaan = textPekerjaan.Text
        adodcPendaftaranAnggota.Recordset!no_sim = textNoSim.Text
        adodcPendaftaranAnggota.Recordset!foto_anggota = alamatFotoAnggota
        adodcPendaftaranAnggota.Recordset!status_meminjam = "1" '1= diperbolehkan untuk meminjam | 0 = tidak boleh meminjam
        
        If optionLakiLaki.Value = True Then
            adodcPendaftaranAnggota.Recordset!jenis_kelamin = "Laki-Laki"
        End If
        
        If optionPerempuan.Value = True Then
            adodcPendaftaranAnggota.Recordset!jenis_kelamin = "Perempuan"
        End If
        
    adodcPendaftaranAnggota.Recordset.Update
    MsgBox "Terdaftar.", vbInformasi, sistem.msgTitle
    Unload Me
    
    Exit Sub
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Sub

Private Sub Form_Load()
    setConnection
    
    'set lokasi form
    Me.Top = 1500
    Me.Left = 6000
    
    'inisialisasi
    defaultFotoAnggota = App.Path + "/images/photoAnggota/defaultFoto.jpg"
    alamatFotoAnggota = defaultFotoAnggota
    imageFotoAnggota.Picture = LoadPicture(defaultFotoAnggota)
    
    'inisialisasi ToolTip
    buttonPilihFoto.ToolTipText = "Pilih foto dari komputer anda."
    buttonSimpan.ToolTipText = "Simpan"
    imageFotoAnggota.ToolTipText = "Foto Calon Anggota"
    optionLakiLaki.ToolTipText = "Laki-Laki"
    optionPerempuan.ToolTipText = "Perempuan"
    textAlamat.ToolTipText = "Alamat Calon Anggota"
    textKabupaten.ToolTipText = "Kabupaten Calon Anggota"
    textKecamatan.ToolTipText = "Kecematan Calon Anggota"
    textKelDesa.ToolTipText = "Kelurahan Calon Anggota"
    textKodePos.ToolTipText = "Kode Pos Calon Anggota"
    textNama.ToolTipText = "Nama Calon Anggota"
    textNomerKTP.ToolTipText = "Nomer KTP Calon Anggota"
    textNoSim.ToolTipText = "Nomer Surat Ijin Mengemudi Calon Anggota"
    textNoTelepon.ToolTipText = "Nomer Telpone Calon Anggota"
    textPekerjaan.ToolTipText = "Pekerjaan Calon Anggota"
    textRtRw.ToolTipText = "RT/RW Calon Anggota"
    textTempatTanggalLahir.ToolTipText = "Tempat Tanggal Lahir Calon Anggota, Format : Tempat, d mmmm yyyy"
    
    adodcPendaftaranAnggota.Refresh
End Sub
