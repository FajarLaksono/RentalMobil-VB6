VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form editView_mobilBaru 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peninjauan / Editing"
   ClientHeight    =   5055
   ClientLeft      =   6540
   ClientTop       =   3360
   ClientWidth     =   6585
   Icon            =   "editView_mobilBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6585
   Begin VB.TextBox textKetersediaan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "Tersedia / Tidak Tersedia"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton buttonView 
      Caption         =   "View"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox textNoPlat 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.TextBox textNamaMobil 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox textHargaMobil 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Rp""#.##0;(""Rp""#.##0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CommandButton buttonEditSimpan 
      Caption         =   "EditSimpan"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton buttonClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox comboTipeMobil 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Text            =   "-- Pilih --"
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton buttonPilihGambar 
      Caption         =   "Pilih Gambar"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   3840
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adodcMobilBaru 
      Height          =   330
      Left            =   360
      Top             =   4920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc1"
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
   Begin VB.Label labelNomerPlat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No. Plat"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label labelNamaMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nama Mobil"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label labelTipeMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipe Mobil"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label labelHarga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Harga / Hari"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label labelGambar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gambar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image imageMobil 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   4455
   End
End
Attribute VB_Name = "editView_mobilBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables
Dim judulForm As String
Dim defaultFoto As String 'akan digunakan untuk pembanding pada penyeleksian, diisi dengan defaultmobil.jpg
Dim alamatFoto As String 'idem, diisi dengan gambar yg akan kita pilih atau gunakan pada commentDialog
Dim currFoto As String 'idem, diisi dengan gambar yang sedang digunakan

Function setConnection()
    adodcMobilBaru.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcMobilBaru.RecordSource = "SELECT * FROM mobil" 'SQL
    adodcMobilBaru.Refresh
End Function

Private Sub buttonClose_Click()
    Unload Me
    tableMobil.Show
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

Private Sub buttonPilihGambar_Click()
    'dibawah adalah cara untuk mendapatkan file, kami buat hanya bisa memilih file JPEG untuk gambar
    commonDialog.FileName = "" 'inisialisasi
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set format pengambilan file
    commonDialog.ShowOpen 'buka dialog pemilihan file yang disediakan oleh windows
    alamatFoto = commonDialog.FileName 'ambil alamat gambar ke variabel alamatFoto
    
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatFoto)) < 1 Then
        Exit Sub
    End If
    'set imageFotoAnggota sesuai gambar
    imageMobil.Picture = LoadPicture(alamatFoto)
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
    
    'inisialisasi ComboBox
    comboTipeMobil.AddItem "Sport"
    comboTipeMobil.AddItem "Mobil Angkut"
    comboTipeMobil.AddItem "Mobil Keluarga"
    
    'Set lokasi
    Me.Top = 1815
    Me.Left = 7110
    
    setConnection 'set koneksi
    muatInfo 'inisialisasi
    setJudul 'set judul
    
    defaultFoto = App.Path + "/images/photoMobil/defaultmobil.jpg" 'inisialisasi
End Sub
Function setJudul()
    If sistem.isEditing = True Then
        Me.Caption = "Plat Mobil : " + judulForm + " - Editor"
    Else
        Me.Caption = "Plat Mobil : " + judulForm + " - Peninjauan"
    End If
End Function

Function persiapanEditing()
    sistem.isEditing = True
    buttonView.Visible = True
    buttonEditSimpan.Caption = "Simpan"
    buttonPilihGambar.Enabled = True
    textKetersediaan.Visible = False
    
    'Set Enable
    textNoPlat.Enabled = True
    textNamaMobil.Enabled = True
    comboTipeMobil.Enabled = True
    textHargaMobil.Enabled = True
    
    'Set Warna
    textNoPlat.BackColor = &H80000005
    textNamaMobil.BackColor = &H80000005
    comboTipeMobil.BackColor = &H80000005
    textHargaMobil.BackColor = &H80000005
End Function

Function persiapanPeninjauan()
    sistem.isEditing = False
    buttonView.Visible = False
    buttonEditSimpan.Caption = "Edit"
    buttonPilihGambar.Enabled = False
    textKetersediaan.Visible = True
    
    'Set Enable
    textNoPlat.Enabled = False
    textNamaMobil.Enabled = False
    comboTipeMobil.Enabled = False
    textHargaMobil.Enabled = False
    
    'Set Warna
    textNoPlat.BackColor = &H80000004
    textNamaMobil.BackColor = &H80000004
    comboTipeMobil.BackColor = &H80000004
    textHargaMobil.BackColor = &H80000004
End Function

Function muatInfo()
    adodcMobilBaru.Refresh
    adodcMobilBaru.Recordset.Find "plat_mobil='" & sistem.currRecord & "'"
    If Not adodcMobilBaru.Recordset.EOF Then
        textNoPlat.Text = sistem.currRecord
        textNamaMobil.Text = adodcMobilBaru.Recordset!nama_mobil
        comboTipeMobil.Text = adodcMobilBaru.Recordset!tipe_mobil
        textHargaMobil.Text = adodcMobilBaru.Recordset!harga_hari
        If adodcMobilBaru.Recordset!tersedia > 0 Then
            textKetersediaan.Text = "Tersedia"
        Else
            textKetersediaan.Text = "Tidak Tersedia"
        End If
        
        If Not adodcMobilBaru.Recordset!alamat_gambar = "" Then
            'Cek apakah file yang dimaksud ada dalam alamat.
            If Dir(App.Path + adodcMobilBaru.Recordset!alamat_gambar) <> "" Then
                'jika ada maka akan merekam sesuai rekaman di database
                imageMobil.Picture = LoadPicture(App.Path + adodcMobilBaru.Recordset!alamat_gambar)
                alamatFoto = App.Path + adodcMobilBaru.Recordset!alamat_gambar 'inisialisasi
                currFoto = App.Path + adodcMobilBaru.Recordset!alamat_gambar 'inisialisasi
            Else
                MsgBox "Terjadi kesalahan dalam pencarian file gambar mobil !", vbCritical, sistem.msgTitle
                imageMobil.Picture = LoadPicture(App.Path + "\images\photoMobil\defaultmobil.JPG")
                alamatFoto = App.Path + "\images\photoMobil\defaultmobil.JPG" 'inisialisasi
                currFoto = App.Path + "\images\photoMobil\defaultmobil.JPG" 'inisialisasi
            End If
        Else
            MsgBox "Terjadi kesalahan dalam pencarian file gambar mobil !", vbCritical, sistem.msgTitle
            imageMobil.Picture = LoadPicture(App.Path + "\images\photoMobil\defaultmobil.JPG")
            alamatFoto = App.Path + "\images\photoMobil\defaultmobil.JPG" 'inisialisasi
            currFoto = App.Path + "\images\photoMobil\defaultmobil.JPG" 'inisialisasi
        End If
    Else
        MsgBox "Ada masalah dalam pencarian data.", vbCritical, sistem.msgTitle
    End If
    judulForm = adodcMobilBaru.Recordset!plat_mobil
End Function

Function simpanData()
    On Error GoTo errHandler 'sama seperti expection pada bahasa pemrograman lain, untuk menangani error pada program
    'jika error akan langsung menuju ke errHendler yang ada dibawah
    
    'save ke database
    'cek gambar apakah sama dengan nilai awal
    If alamatFoto = defaultFoto Then 'check
        MsgBox "Gambar Belum Dipilih !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    'cek jika masih kosong
    If textNoPlat.Text = "" Then
        MsgBox "No Plat Mobil belum di isi", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textNamaMobil.Text = "" Then
        MsgBox "No Plat Mobil belum di isi", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If comboTipeMobil.Text = "--Pilih--" Or comboTipeMobil.Text = "" Then
        MsgBox "Tipe Mobil belum di pilih", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textHargaMobil.Text = "" Then
        MsgBox "No Plat Mobil Belum Di isi", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    'Start of Pengolahan gambar
    If Not commonDialog.FileName = "" And Not commonDialog.FileTitle = "" And Not alamatImage = currFoto Then
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
        namaFile = textNoPlat.Text + tanggal + waktu
        'mencopy, memindahkan ke folder jangkauan program dan mengubah nama.
        FileCopy commonDialog.FileName, App.Path + "/images/photoMobil/" + namaFile + commonDialog.FileTitle
        'menyimpannya dalam variabel
        alamatGambar = "/images/photoMobil/" + namaFile + commonDialog.FileTitle
        'End of pengolahanGambar
        adodcMobilBaru.Recordset!alamat_gambar = alamatGambar
    End If
    'menyimpan dalam database
        adodcMobilBaru.Recordset!plat_mobil = textNoPlat.Text
        adodcMobilBaru.Recordset!nama_mobil = textNamaMobil.Text
        adodcMobilBaru.Recordset!tipe_mobil = comboTipeMobil.Text
        adodcMobilBaru.Recordset!harga_hari = textHargaMobil.Text
        adodcMobilBaru.Recordset!tersedia = 1 'Boolean = true, Kami tidak tau cara merubah value boolean pada access, jadi kita membuat Boolean sendiri dengan integer
    adodcMobilBaru.Recordset.Update
    MsgBox "Data mobil sudah diperbaharui dalam database.", vbInformasi, sistem.msgTitle
    Unload Me
    tableMobil.Show
    
Exit Function
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
    'digunakan untuk mengatasi error saat input data pada database, contoh : biasanya user memasukan huruf pada textbox angka.
End Function
