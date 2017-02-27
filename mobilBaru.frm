VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mobilBaru 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pendaftaran Mobil Baru"
   ClientHeight    =   5205
   ClientLeft      =   7110
   ClientTop       =   3180
   ClientWidth     =   6705
   Icon            =   "mobilBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton buttonPilihGambar 
      Caption         =   "Pilih Gambar"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   3840
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adodcMobilBaru 
      Height          =   330
      Left            =   360
      Top             =   3480
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
   Begin VB.ComboBox comboTipeMobil 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Text            =   "-- Pilih --"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton buttonCancle 
      Caption         =   "Cancle"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton buttonTambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
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
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox textNamaMobil 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox textNoPlat 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image imageMobil 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label labelGambar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gambar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2280
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
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label labelTipeMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipe Mobil"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label labelNamaMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nama Mobil"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label labelNomerPlat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No. Plat"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "mobilBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables
Dim defaultGambar As String
Dim alamatGambar As String

Dim tanggal As String
Dim waktu As String
Dim namaFile As String

Function setConnection()
    adodcMobilBaru.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcMobilBaru.RecordSource = "select * from mobil" 'SQL
End Function

Private Sub buttonCancle_Click()
    Unload Me
End Sub

Private Sub buttonPilihGambar_Click()
    'dibawah adalah cara untuk mendapatkan file, kami buat hanya bisa memilih file JPEG untuk gambar
    commonDialog.FileName = "" 'inisialisasi
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set format pengambilan file
    commonDialog.ShowOpen 'buka dialog pemilihan dari windows
    alamatGambar = commonDialog.FileName 'ambil alamat gambar ke variabel alamatGambar
    
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatGambar)) < 1 Then
        Exit Sub
    End If
    
    'set imageMobil sesuai gambar
    imageMobil.Picture = LoadPicture(alamatGambar)
End Sub

Private Sub buttonTambah_Click()
    On Error GoTo errHandler
    'save ke database
    
    'cek gambar apakah sama dengan nilai awal
    If alamatGambar = defaultGambar Or commonDialog.FileName = "" Or commonDialog.FileTitle = "" Then  'check
        MsgBox "Gambar Belum Dipilih !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    'cek jika masih kosong
    If textNoPlat.Text = "" Then
        MsgBox "No Plat Mobil belum di isi", vbInformasi, sistem.msgTitle
        Exit Sub
    Else
        'cek primary key
        adodcMobilBaru.Refresh
        adodcMobilBaru.Recordset.Find "plat_mobil='" & textNoPlat.Text & "'" 'mencari berdasarkan plat_mobil yang dimaksud
        If Not adodcMobilBaru.Recordset.EOF Then
            MsgBox "Sudah ada plat mobil dengan nomer yang sama.", vbInformasi, sistem.msgTitle
            Exit Sub
        End If
    End If
    
    If textNamaMobil.Text = "" Then
        MsgBox "No Plat Mobil belum di isi", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If comboTipeMobil.Text = "--Pilih--" Or comboTipeMobil.Text = "" Then
        MsgBox "Tipe Mobil belum di pilih", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textHargaMobil.Text = "" Then
        MsgBox "No Plat Mobil Belum Di isi", vbInformasi, sistem.msgTitle
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
    namaFile = textNoPlat.Text + tanggal + waktu
    
    'mencopy, memindahkan ke folder jangkauan program dan mengubah nama.
    FileCopy commonDialog.FileName, App.Path + "/images/photoMobil/" + namaFile + commonDialog.FileTitle
    'menyimpannya dalam variabel
    alamatGambar = "/images/photoMobil/" + namaFile + commonDialog.FileTitle
    'End of pengolahanGambar
    
    'menyimpan dalam database
    adodcMobilBaru.Recordset.AddNew
        adodcMobilBaru.Recordset!plat_mobil = textNoPlat.Text
        adodcMobilBaru.Recordset!nama_mobil = textNamaMobil.Text
        adodcMobilBaru.Recordset!tipe_mobil = comboTipeMobil.Text
        adodcMobilBaru.Recordset!harga_hari = textHargaMobil.Text
        adodcMobilBaru.Recordset!tersedia = 1 'Boolean = true, Kami tidak tau cara merubah value boolean pada access, jadi kita membuat Boolean sendiri dengan integer
        adodcMobilBaru.Recordset!alamat_gambar = alamatGambar
    adodcMobilBaru.Recordset.Update
    MsgBox "Data mobil baru sudah ditambahkan dalam database.", vbInformasi, sistem.msgTitle
    Unload Me
    
Exit Sub
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Sub

Private Sub Form_Load()
    'set koneksi
    setConnection

    'inisialisasi lokasi form
    Me.Top = 2500
    Me.Left = 7065

    'inisialisasi ComboBox
    comboTipeMobil.AddItem "Sport"
    comboTipeMobil.AddItem "Mobil Angkut"
    comboTipeMobil.AddItem "Mobil Keluarga"
    
    'inisialisasi ToolTip
    buttonCancle.ToolTipText = "Cancle"
    buttonPilihGambar.ToolTipText = "Pilih Gambar Dari Komputer"
    buttonTambah.ToolTipText = "Tambahkan ke database"
    comboTipeMobil.ToolTipText = "Pilih Tipe Mobil"
    imageMobil.ToolTipText = "Gambar Mobil"
    textHargaMobil.ToolTipText = "Harga Mobil Per hari"
    textNamaMobil.ToolTipText = "Nama Mobil"
    textNoPlat.ToolTipText = "No Plat Mobil"
    
    'inisiaslisasi gambar
    defaultGambar = App.Path + "/images/photoMobil/defaultmobil.jpg"
    alamatGambar = defaultGambar
    imageMobil.Picture = LoadPicture(defaultGambar)
    
    adodcMobilBaru.Refresh
End Sub
