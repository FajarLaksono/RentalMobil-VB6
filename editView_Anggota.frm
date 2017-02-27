VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form editView_Anggota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peninjauan / Editor"
   ClientHeight    =   7155
   ClientLeft      =   7155
   ClientTop       =   2190
   ClientWidth     =   8550
   Icon            =   "editView_Anggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8550
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H00C0C0C0&
      Height          =   7335
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton buttonClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   7200
         TabIndex        =   17
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton buttonView 
         Caption         =   "View"
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   6600
         Width           =   1095
      End
      Begin VB.TextBox textStatusMeminjam 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   19
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox textViewJenisKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton buttonEditSimpan 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         Top             =   6600
         Width           =   1095
      End
      Begin VB.TextBox textPekerjaan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   13
         Top             =   5520
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
      Begin VB.TextBox textNoSim 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   14
         Top             =   6000
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
      Begin VB.TextBox textKecamatan 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   10
         Top             =   4080
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
      Begin VB.TextBox textRtRw 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   8
         Top             =   3120
         Width           =   5685
      End
      Begin VB.TextBox textAlamat 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   7
         Top             =   2640
         Width           =   5685
      End
      Begin VB.TextBox textNoTelepon 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   5
         Top             =   2160
         Width           =   3495
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
      Begin VB.TextBox textTempatTanggalLahir 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   2
         Top             =   1200
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
      Begin VB.TextBox textNomerKTP 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton buttonPilihFoto 
         Caption         =   "Pilih Foto"
         Height          =   435
         Left            =   6360
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
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
      Begin VB.Label labelKodePos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kode Pos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   5055
         Width           =   2175
      End
      Begin VB.Label labelKabupaten 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kabupaten"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label labelKecamatan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kecamatan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label labelKelDesa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kel / Desa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label labelRtRw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "RT / RW"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label labelPekerjaan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pekerjaan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   5550
         Width           =   2175
      End
      Begin VB.Label labelNoSIM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No SIM"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   6030
         Width           =   2175
      End
      Begin VB.Label labekNoTelepon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No Telepon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2160
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
      Begin VB.Label labelJenisKelamin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jenis Kelamin"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label labelTempatTanggalLahir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tempat / Tanggal Lahir"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label labelNama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   765
         Width           =   2175
      End
      Begin VB.Label labelNomerKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nomor KTP "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   285
         Width           =   2175
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
   End
End
Attribute VB_Name = "editView_Anggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables
Dim judulForm As String
Dim defaultFotoAnggota As String 'akan digunakan untuk pembanding pada penyeleksian, diisi dengan defaultFoto.jpg
Dim alamatFotoAnggota As String 'idem, diisi dengan gambar yg akan kita pilih atau gunakan pada commentDialog
Dim currFotoAnggota As String 'idem, diisi dengan gambar yang sedang digunakan

Function setConnection()
    adodcPendaftaranAnggota.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcPendaftaranAnggota.RecordSource = "SELECT * FROM anggota" 'SQL
End Function

Private Sub buttonClose_Click()
    Unload Me
    tableAnggota.Show
End Sub

Private Sub buttonEditSimpan_Click()
    If buttonView.Visible = True And sistem.isEditing = True Then
        'simpan
        simpanData
    Else
        'edit mode
        persiapanEditing
        setJudul
    End If
End Sub

Private Sub buttonPilihFoto_Click()
    'dibawah adalah cara untuk mendapatkan file, kami buat hanya bisa memilih file JPEG untuk gambar
    commonDialog.FileName = "" 'inisialisasi
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set format pengambilan file
    commonDialog.ShowOpen 'buka dialog pemilihan file yang disediakan oleh windows
    alamatFotoAnggota = commonDialog.FileName 'ambil alamat gambar ke variabel alamatFotoAnggota
    
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatFotoAnggota)) < 1 Then
        Exit Sub
    End If
    'set imageFotoAnggota sesuai gambar
    imageFotoAnggota.Picture = LoadPicture(alamatFotoAnggota)
End Sub

Private Sub buttonView_Click()
    persiapanPeninjauan
    setJudul
End Sub

Private Sub Form_Load()
    If sistem.isEditing = True Then
        persiapanEditing
    Else
        persiapanPeninjauan
    End If
    
    'set lokasi
    Me.Top = 1815
    Me.Left = 7110
    
    setConnection 'set koneksi
    muatInfo 'inisialisasi
    setJudul 'set judul
    
    defaultFotoAnggota = App.Path + "/images/photoAnggota/defaultFoto.jpg" 'inisialisasi
End Sub

Function persiapanEditing()
    sistem.isEditing = True
    buttonView.Visible = True
    buttonEditSimpan.Caption = "Simpan"

    textViewJenisKelamin.Enabled = False
    textViewJenisKelamin.Visible = False
    
    optionLakiLaki.Visible = True
    optionPerempuan.Visible = True
    
    textStatusMeminjam.Visible = False
    
    buttonPilihFoto.Visible = True
    
    'Set Enable
    textNomerKTP.Enabled = True
    textNama.Enabled = True
    textTempatTanggalLahir.Enabled = True
    textNoTelepon.Enabled = True
    textAlamat.Enabled = True
    textRtRw.Enabled = True
    textKelDesa.Enabled = True
    textKecamatan.Enabled = True
    textKabupaten.Enabled = True
    textKodePos.Enabled = True
    textPekerjaan.Enabled = True
    textNoSim.Enabled = True
    
    'Set Warna
    textNomerKTP.BackColor = &H80000005
    textNama.BackColor = &H80000005
    textTempatTanggalLahir.BackColor = &H80000005
    textNoTelepon.BackColor = &H80000005
    textAlamat.BackColor = &H80000005
    textRtRw.BackColor = &H80000005
    textKelDesa.BackColor = &H80000005
    textKecamatan.BackColor = &H80000005
    textKabupaten.BackColor = &H80000005
    textKodePos.BackColor = &H80000005
    textPekerjaan.BackColor = &H80000005
    textNoSim.BackColor = &H80000005
End Function

Function persiapanPeninjauan()
    sistem.isEditing = False
    buttonView.Visible = False
    buttonEditSimpan.Caption = "Edit"

    textViewJenisKelamin.Enabled = False
    textViewJenisKelamin.Visible = True
    
    optionLakiLaki.Visible = False
    optionPerempuan.Visible = False
    
    textStatusMeminjam.Visible = True
    
    buttonPilihFoto.Visible = False
    
    'Set Enable
    textNomerKTP.Enabled = False
    textNama.Enabled = False
    textTempatTanggalLahir.Enabled = False
    textNoTelepon.Enabled = False
    textAlamat.Enabled = False
    textRtRw.Enabled = False
    textKelDesa.Enabled = False
    textKecamatan.Enabled = False
    textKabupaten.Enabled = False
    textKodePos.Enabled = False
    textPekerjaan.Enabled = False
    textNoSim.Enabled = False
    
    'Set Warna
    textNomerKTP.BackColor = &H80000004
    textNama.BackColor = &H80000004
    textTempatTanggalLahir.BackColor = &H80000004
    textNoTelepon.BackColor = &H80000004
    textAlamat.BackColor = &H80000004
    textRtRw.BackColor = &H80000004
    textKelDesa.BackColor = &H80000004
    textKecamatan.BackColor = &H80000004
    textKabupaten.BackColor = &H80000004
    textKodePos.BackColor = &H80000004
    textPekerjaan.BackColor = &H80000004
    textNoSim.BackColor = &H80000004
End Function

Function muatInfo()
    adodcPendaftaranAnggota.Refresh
    adodcPendaftaranAnggota.Recordset.Find "KTP='" & sistem.currRecord & "'"
    If Not adodcPendaftaranAnggota.Recordset.EOF Then
        textNomerKTP.Text = adodcPendaftaranAnggota.Recordset!KTP
        textNama.Text = adodcPendaftaranAnggota.Recordset!nama_anggota
        textTempatTanggalLahir.Text = adodcPendaftaranAnggota.Recordset!tempat_tanggal_lahir
        textNoTelepon.Text = adodcPendaftaranAnggota.Recordset!no_telp
        textAlamat.Text = adodcPendaftaranAnggota.Recordset!alamat
        textRtRw.Text = adodcPendaftaranAnggota.Recordset!rt_rw
        textKelDesa.Text = adodcPendaftaranAnggota.Recordset!kelDesa
        textKecamatan.Text = adodcPendaftaranAnggota.Recordset!kec
        textKabupaten.Text = adodcPendaftaranAnggota.Recordset!kab
        textKodePos.Text = adodcPendaftaranAnggota.Recordset!kode_pos
        textPekerjaan.Text = adodcPendaftaranAnggota.Recordset!pekerjaan
        textNoSim.Text = adodcPendaftaranAnggota.Recordset!no_sim
        
        textViewJenisKelamin = adodcPendaftaranAnggota.Recordset!jenis_kelamin
        If adodcPendaftaranAnggota.Recordset!jenis_kelamin = "Perempuan" Then
            optionPerempuan.Value = True
            optionLakiLaki.Value = False
        Else
            optionLakiLaki.Value = True
            optionPerempuan.Value = False
        End If
        
        If adodcPendaftaranAnggota.Recordset!status_meminjam < 1 Then
            textStatusMeminjam.Text = "Sedang Meminjam"
        Else
            textStatusMeminjam.Text = "Tidak Sedang Meminjam"
        End If
        
        
        If Not adodcPendaftaranAnggota.Recordset!foto_anggota = "" Then
            'Cek apakah file yang dimaksud ada dalam alamat.
            If Dir(App.Path + adodcPendaftaranAnggota.Recordset!foto_anggota) <> "" Then
                'jika ada maka akan merekam sesuai rekaman di database
                imageFotoAnggota.Picture = LoadPicture(App.Path + adodcPendaftaranAnggota.Recordset!foto_anggota)
                alamatFotoAnggota = App.Path + adodcPendaftaranAnggota.Recordset!foto_anggota 'inisialisasi
                currFotoAnggota = App.Path + adodcPendaftaranAnggota.Recordset!foto_anggota 'inisialisasi
            Else
                MsgBox "Terjadi kesalahan dalam pencarian file gambar anggota !", vbCritical, sistem.msgTitle
                imageFotoAnggota.Picture = LoadPicture(App.Path + "\images\photoAnggota\defaultFoto.JPG")
                alamatFotoAnggota = App.Path + "\images\photoAnggota\defaultFoto.JPG" 'inisialisasi
                currFotoAnggota = App.Path + "\images\photoAnggota\defaultFoto.JPG" 'inisialsisasi
            End If
        Else
            MsgBox "Terjadi kesalahan dalam pencarian file gambar anggota !", vbCritical, sistem.msgTitle
            imageFotoAnggota.Picture = LoadPicture(App.Path + "\images\photoAnggota\defaultFoto.JPG")
            alamatFotoAnggota = App.Path + "\images\photoAnggota\defaultFoto.JPG" 'inisialisasi
            currFotoAnggota = App.Path + "\images\photoAnggota\defaultFoto.JPG" 'inisialisasi
        End If
        
        judulForm = adodcPendaftaranAnggota.Recordset!KTP
    End If
End Function

Function setJudul()
    If sistem.isEditing = True Then
        Me.Caption = "KTP : " + judulForm + " - Editor"
    Else
        Me.Caption = "KTP : " + judulForm + " - Peninjauan"
    End If
End Function

Function simpanData()
    On Error GoTo errHandler 'sama seperti expection pada bahasa pemrograman lain, untuk menangani error pada program
    'jika error akan langsung menuju ke errHendler yang ada dibawah
    
    'cek semua pengisian
    If alamatFotoAnggota = defaultFotoAnggota Then
        MsgBox "Foto Anggota Belum Dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textNomerKTP.Text = "" Then
        MsgBox "Informasi Nomer KTP Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textNama.Text = "" Then
        MsgBox "Informasi Nama Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textTempatTanggalLahir.Text = "" Then
        MsgBox "Informasi Tanggal Lahir Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textNoTelepon.Text = "" Then
        MsgBox "Informasi No Telepon Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textAlamat.Text = "" Then
        MsgBox "Informasi Alamat Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textRtRw.Text = "" Then
        MsgBox "Informasi RT/RW Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textKelDesa.Text = "" Then
        MsgBox "Informasi Kel/Desa Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textKecamatan.Text = "" Then
        MsgBox "Informasi Kecamatan Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textKabupaten.Text = "" Then
        MsgBox "Informasi Kabupaten Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
        
    If textKodePos.Text = "" Then
        MsgBox "Informasi Kode Pos Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textPekerjaan.Text = "" Then
        MsgBox "Informasi Pekerjaan Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textNoSim.Text = "" Then
        MsgBox "Informasi Nomer SIM Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If optionLakiLaki.Value = False And optionPerempuan.Value = False Then
        MsgBox "Informasi Jenis Kelamin Anggota belum dimasukan !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
        
    'Start of Pengolahan gambar
    If Not commonDialog.FileName = "" And Not commonDialog.FileTitle = "" And Not alamatFotoAnggota = currFotoAnggota Then
        'dibawah ini adalah cara untuk mengelolah gambar.
        'bukan hanya kita memilih melalui commondialog diatas. tapi kita juga mengcopy dan menaruhnya dalam jangkauan program
        'mengubah nama, berguna untuk mengantisipasi duplikasi gambar dengan file yang sudah disatukan dalam satu forder jangkauan program
        'untuk mengantisipasi duplikasi kami menambahkan tanggal+waktu pada nama file.
        'karena tanggal dan waktu tidak akan pernah sama dan terus maju.
        tanggal = Format(Date, "d-mmmm-yyyy")
        waktu = Format(Time, "h-m-s")
        'jika kita hanya menggunakan tanggal = date(), date() akan menghasilkan format/bentuk text seperti ini d/mmmm/yyyy.
        'dan pada windows kita tidak boleh menggunakan karakter / untuk nama file.
        'maka dibuat "d-mmmm-yyyy".
        
        namaFile = sistem.userNIK + tanggal + waktu 'menyatukan string
    
        FileCopy commonDialog.FileName, App.Path + "/images/photoAnggota/" + namaFile + commonDialog.FileTitle 'mencopy, memindahkan ke folder jangkauan program dan mengubah nama.
        
        alamatFotoAnggota = "/images/photoAnggota/" + namaFile + commonDialog.FileTitle 'menyimpan alamat dalam variabel
        'End of pengolahanGambar
        adodcPendaftaranAnggota.Recordset!foto_anggota = alamatFotoAnggota 'masukan ke database
    End If
    
    'menyimpan dalam database
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
    If optionLakiLaki.Value = True Then
        adodcPendaftaranAnggota.Recordset!jenis_kelamin = "Laki-Laki"
    End If
    
    If optionPerempuan.Value = True Then
        adodcPendaftaranAnggota.Recordset!jenis_kelamin = "Perempuan"
    End If
        
    adodcPendaftaranAnggota.Recordset.Update
    MsgBox "Updated", vbInformasi, sistem.msgTitle
    Unload Me
    tableAnggota.Show
Exit Function 'keluar
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
    'digunakan untuk mengatasi error saat input data pada database, contoh : biasanya user memasukan huruf pada textbox angka.
End Function
