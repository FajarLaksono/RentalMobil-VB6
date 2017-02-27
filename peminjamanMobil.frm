VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form peminjamanMobil 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peminjaman Mobil"
   ClientHeight    =   7440
   ClientLeft      =   7530
   ClientTop       =   2475
   ClientWidth     =   5910
   Icon            =   "peminjamanMobil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   5910
   Begin VB.ComboBox comboNoKTP 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Text            =   "-- Pilih --"
      Top             =   4920
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc adodcRental 
      Height          =   330
      Left            =   960
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Rental"
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
   Begin VB.TextBox textNamaMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox textTipeMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox textHargaHari 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Frame framePeminjamanMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Peminjaman"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5895
      Begin MSAdodcLib.Adodc adodcTableAnggota 
         Height          =   330
         Left            =   1920
         Top             =   6720
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc adodcTableMobil 
         Height          =   330
         Left            =   120
         Top             =   6720
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
      Begin VB.Timer timer 
         Left            =   3840
         Top             =   6840
      End
      Begin VB.TextBox textTotalHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   6000
         Width           =   3495
      End
      Begin VB.TextBox textJaminan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   6360
         Width           =   3495
      End
      Begin VB.TextBox textTglPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   5640
         Width           =   3495
      End
      Begin VB.TextBox textLamaPeminjaman 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   5280
         Width           =   2895
      End
      Begin VB.CommandButton buttonSimpan 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   6840
         Width           =   1215
      End
      Begin VB.ComboBox comboPlatMobil 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Text            =   "-- Pilih --"
         Top             =   480
         Width           =   3495
      End
      Begin VB.Image imageMobil 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2415
         Left            =   360
         Stretch         =   -1  'True
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label labelHari 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label labelJaminan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jaminan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label labelTotalHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Harga"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label labelTglPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tgl Peminjaman"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label labelLamaPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lama Peminjaman"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label labelNoKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No KTP Peminjam"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label labelHargaHari 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Harga / Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label labelPlatMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Plat Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label labelTipeMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipe Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label labelNamaMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "peminjamanMobil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables
Dim tanggal As String
Dim bulan As String
Dim tahun As String
Dim IdPeminjaman As Integer
Dim dapatkahIdPeminjaman As Boolean

Function setConnection()
    adodcTableMobil.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcTableMobil.RecordSource = "select * from mobil" 'SQL
    
    adodcTableAnggota.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcTableAnggota.RecordSource = "select * from anggota" 'SQL
        
    adodcRental.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcRental.RecordSource = "select * from rental" 'SQL
End Function

Function setIdPeminjaman()
    'function ini akan mengatur ID peminjaman dengan menggunakan pengulangan setiap rekaman pada database
    'menyeleksinya apakah ada ID peminjaman tersebut. jika tidak ditemukan maka akan digunakan untuk Peminjaman ini
    Do While Not dapatkahIdPeminjaman
        adodcRental.Refresh
        adodcRental.Recordset.Find "id_peminjaman='" & IdPeminjaman & "'"
        If Not adodcRental.Recordset.EOF Then
            IdPeminjaman = Val(IdPeminjaman) + 1
        Else
            dapatkahIdPeminjaman = True
        End If
    Loop
    textJaminan.Text = IdPeminjaman
End Function

Function savingToDatabase()
    On Error GoTo errHandler
    'cek adakah anggota rental mobil yang dimaksud
    adodcTableAnggota.Refresh
    adodcTableAnggota.Recordset.Find "KTP='" & comboNoKTP.Text & "'"
    If Not adodcTableAnggota.Recordset.EOF Then
        'cek jika anggota rental mobil tidak sedang meminjam mobil
        If adodcTableAnggota.Recordset!status_meminjam > 0 Then '1=boleh meminjam | 0=sedang meminjam/tidak boleh meminjam
            'cek adakah plat mobil tang dimaksud
            adodcTableMobil.Refresh
            adodcTableMobil.Recordset.Find "plat_mobil='" & comboPlatMobil.Text & "'"
            If Not adodcTableMobil.Recordset.EOF Then
                'cek ketersediaan mobil
                If adodcTableMobil.Recordset!tersedia > 0 Then '1=tersedia | 0=tidak tersedia
                    'mulai perekaman ke database
                    adodcRental.Refresh
                    adodcRental.Recordset.AddNew
                        adodcRental.Recordset!id_peminjaman = IdPeminjaman
                        adodcRental.Recordset!KTP = comboNoKTP.Text
                        adodcRental.Recordset!plat_mobil = comboPlatMobil.Text
                        adodcRental.Recordset!NIK = sistem.userNIK
                        adodcRental.Recordset!status_peminjaman = 1 'Status peminjaman, 1 = Sedang di Pinjam / 0=Sudah kembali
                        'status_peminjaman di database akan digunakan sebagai tanda bahwa 1 = sedang dipinjam / 0=sudah kembali
                        'akan berguna dalam pencetakan tabel peminjaman dan tabel pengembalian/riwayat rental
                        adodcRental.Recordset!lama_peminjaman = textLamaPeminjaman.Text
                        adodcRental.Recordset!tanggal_peminjaman = textTglPeminjaman.Text
                        adodcRental.Recordset!jaminan = textJaminan.Text
                        adodcRental.Recordset!harga = textTotalHarga.Text
                    adodcRental.Recordset.Update
                    
                    adodcTableMobil.Recordset!tersedia = 0 '1=tersedia untuk dipinjam | 0=tidak tersedia untuk dipinjam
                    adodcTableMobil.Recordset.Update
                    
                    adodcTableAnggota.Recordset!status_meminjam = 0 '1=boleh meminjam | 0=tidak boleh meminjam
                    adodcTableAnggota.Recordset.Update
                    
                    MsgBox "Tersimpan, Angota sudah resmi bisa menerima kunci mobil!", vbInformasi, sistem.msgTitle
                    Unload Me
                Else
                    MsgBox "Persediaan Mobil yang diminta sudah habis !", vbInformasi, sistem.msgTitle
                    Exit Function
                End If
            Else
                MsgBox "Plat Mobil tidak ditemukan !", vbInformasi, sistem.msgTitle
                Exit Function
            End If
        Else
            MsgBox "Calon Rental terdaftar sedang meminjam Mobil !", vbInformasi, sistem.msgTitle
            Exit Function
        End If
    Else
        MsgBox "Calon Rental belum terdaftar sebagai anggota !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
Exit Function
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Function

Private Sub buttonSimpan_Click()
    If comboPlatMobil.Text = "-- Pilih --" Or comboPlatMobil = "" Then
        MsgBox "Plat Mobil Belum Dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If comboNoKTP.Text = "-- Pilih --" Or comboPlatMobil = "" Then
        MsgBox "No KTP Belum Dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textLamaPeminjaman.Text = "" Then
        MsgBox "Lama Peminjaman Belum Dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textJaminan.Text = "" Then
        MsgBox "Jaminan Belum Dimasukan !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    setIdPeminjaman 'set ID untuk primary key sebelum masuk ke function savingToDatabase
    savingToDatabase 'perekaman
End Sub

Private Sub comboPlatMobil_Click()
    adodcTableMobil.Refresh
    adodcTableMobil.Recordset.Find "plat_mobil='" & comboPlatMobil.Text & "'"
    If Not adodcTableMobil.Recordset.EOF Then
        textNamaMobil.Text = adodcTableMobil.Recordset!nama_mobil
        textTipeMobil.Text = adodcTableMobil.Recordset!tipe_mobil
        textHargaHari.Text = adodcTableMobil.Recordset!harga_hari
        textLamaPeminjaman.Text = textLamaPeminjaman.Text
        textTotalHarga.Text = Val(textLamaPeminjaman) * Val(adodcTableMobil.Recordset!harga_hari)
        'Cek keberadaan file
        If Not adodcTableMobil.Recordset!alamat_gambar = "" Then
            'Cek apakah file yang dimaksud ada dalam alamat.
            If Dir(App.Path + adodcTableMobil.Recordset!alamat_gambar) <> "" Then
                'jika ada maka akan merekam sesuai rekaman di database
                imageMobil.Picture = LoadPicture(App.Path + adodcTableMobil.Recordset!alamat_gambar)
            Else
                'jika tidak ada maka akan memunculkan msgBox, mengkosonginya value gambar pada databaase2003>pegawai dan menggantinya dengan default foto
                MsgBox "Kami tidak menemukan Foto anda dalam directory kami! Informasi alamat Foto akan otomatis kami kosongkan dalam database dan kami ganti dengan Foto default kami."
                adodcTableMobil.Recordset!gambar = "" 'kosongkan
                adodcTableMobil.Recordset.Update 'update recordset
                imageMobil.Picture = LoadPicture(App.Path + "/images/photoMobil/defaultmobil.jpg") 'menggunakan gambar default
            End If
        Else
            'Jika petugas tidak mempunyai foto akan diisi dengan defaultmobil.jpg
            imageMobil.Picture = LoadPicture(App.Path + "/images/photoMobil/defaultmobil.jpg") 'menggunakan gambar default
        End If
    Else
       MsgBox "Plat Mobil tidak Ditemukan !", vbCritical, sistem.msgTitle
    End If
End Sub

Private Sub Form_Load()
    'set coneksi ke database
    setConnection

    'set posisi jendela
    Me.Left = 7000
    Me.Top = 1200
    
    'inisialisasi
    IdPeminjaman = 1
    dapatkahIdPeminjaman = False

    textTglPeminjaman = Date
    
    imageMobil.Picture = LoadPicture(App.Path + "/images/photoMobil/defaultmobil.jpg")
    
    'inisialisasi combo box
    adodcTableMobil.Refresh
    adodcTableMobil.Recordset.MoveFirst
    Do While Not adodcTableMobil.Recordset.EOF
        comboPlatMobil.AddItem adodcTableMobil.Recordset!plat_mobil
        adodcTableMobil.Recordset.MoveNext
    Loop
    
    adodcTableAnggota.Refresh
    adodcTableAnggota.Recordset.MoveFirst
    Do While Not adodcTableAnggota.Recordset.EOF
        comboNoKTP.AddItem adodcTableAnggota.Recordset!KTP
        adodcTableAnggota.Recordset.MoveNext
    Loop
    
    'set ToolTip
    buttonSimpan.ToolTipText = "Simpan"
    comboNoKTP.ToolTipText = "Masukan No KTP"
    comboPlatMobil.ToolTipText = "Masukan Plat Mobil"
    imageMobil.ToolTipText = "Gambar Mobil"
    textHargaHari.ToolTipText = "Harga / Hari"
    textJaminan.ToolTipText = "Masukan Jaminan"
    textLamaPeminjaman.ToolTipText = "Masukan Lama Peminjaman"
    textNamaMobil.ToolTipText = "Nama Mobil"
    textTglPeminjaman.ToolTipText = "Tanggal Peminjama"
    textTipeMobil.ToolTipText = "Tipe Mobil"
    textTotalHarga.ToolTipText = "Harga"
End Sub

Private Sub textLamaPeminjaman_Change()
    textTotalHarga.Text = Val(textLamaPeminjaman) * Val(adodcTableMobil.Recordset!harga_hari)
End Sub
