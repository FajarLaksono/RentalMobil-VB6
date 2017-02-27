VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form editView_peminjaman 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peninjauan / Editor"
   ClientHeight    =   7500
   ClientLeft      =   7245
   ClientTop       =   2160
   ClientWidth     =   5805
   Icon            =   "editView_peminjaman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   5805
   Begin VB.Frame framePeminjamanMobil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Peminjaman"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton buttonEditSimpan 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton buttonView 
         Caption         =   "View"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton buttonClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   6840
         Width           =   1095
      End
      Begin VB.ComboBox comboNoKTP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Text            =   "-- Pilih --"
         Top             =   4920
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
      Begin VB.ComboBox comboPlatMobil 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Text            =   "-- Pilih --"
         Top             =   480
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
      Begin VB.TextBox textTglPeminjaman 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   5640
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
      Begin VB.Timer timer 
         Left            =   1560
         Top             =   360
      End
      Begin MSAdodcLib.Adodc adodcTableAnggota 
         Height          =   330
         Left            =   120
         Top             =   6600
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
         Top             =   6900
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
      Begin MSAdodcLib.Adodc adodcRental 
         Height          =   330
         Left            =   120
         Top             =   7200
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
      Begin VB.Label labelNamaMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label labelTipeMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipe Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label labelPlatMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Plat Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label labelHargaHari 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Harga / Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label labelNoKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No KTP Peminjam"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label labelLamaPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lama Peminjaman"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label labelTglPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tgl Peminjaman"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label labelTotalHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Harga"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label labelJaminan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jaminan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label labelHari 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   5280
         Width           =   495
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
   End
End
Attribute VB_Name = "editView_peminjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim judulForm As String
Dim defaultPlat_mobil As String
Dim defaultKTP As String

Dim isKTPChanged As Boolean
Dim isPlat_mobilChanged As Boolean

Function setConnection()
    adodcTableMobil.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcTableMobil.RecordSource = "SELECT * FROM mobil" 'SQL
    
    adodcTableAnggota.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcTableAnggota.RecordSource = "SELECT * FROM anggota" 'SQL
        
    adodcRental.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcRental.RecordSource = "SELECT * FROM rental" 'SQL
End Function

Private Sub buttonClose_Click()
    Unload Me
    tableRental.Show
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

Private Sub textLamaPeminjaman_Change()
    textTotalHarga.Text = Val(textLamaPeminjaman) * Val(adodcTableMobil.Recordset!harga_hari)
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

Private Sub buttonView_Click()
    persiapanPeninjauan
    setJudul
End Sub

Private Sub Form_Load()
    setConnection

    If sistem.isEditing = True Then
        persiapanEditing
    Else
        persiapanPeninjauan
    End If
    
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
    comboNoKTP.ToolTipText = "Masukan No KTP"
    comboPlatMobil.ToolTipText = "Masukan Plat Mobil"
    imageMobil.ToolTipText = "Gambar Mobil"
    textHargaHari.ToolTipText = "Harga / Hari"
    textJaminan.ToolTipText = "Masukan Jaminan"
    textLamaPeminjaman.ToolTipText = "Masukan Lama Peminjaman"
    textNamaMobil.ToolTipText = "Nama Mobil"
    textTglPeminjaman.ToolTipText = "Tanggal Peminjaman : d/mm/yyyy"
    textTipeMobil.ToolTipText = "Tipe Mobil"
    textTotalHarga.ToolTipText = "Harga"
    
    'Set lokasi
    Me.Top = 1815
    Me.Left = 7110
    
    setConst
    muatInfo
    setJudul
    
    isKTPChanged = False
    isPlat_mobilChanged = False
End Sub

Function setConst()
    textNamaMobil.Enabled = False
    textTipeMobil.Enabled = False
    textHargaHari.Enabled = False
    textTotalHarga.Enabled = False

    textNamaMobil.BackColor = &H80000004
    textTipeMobil.BackColor = &H80000004
    textHargaHari.BackColor = &H80000004
    textTotalHarga.BackColor = &H80000004
End Function

Function persiapanEditing()
    sistem.isEditing = True
    buttonView.Visible = True
    buttonEditSimpan.Caption = "Simpan"
    
    comboPlatMobil.Enabled = True
    comboNoKTP.Enabled = True
    textLamaPeminjaman.Enabled = True
    textTglPeminjaman.Enabled = True
    textJaminan.Enabled = True
    
    comboPlatMobil.BackColor = &H80000005
    comboNoKTP.BackColor = &H80000005
    textLamaPeminjaman.BackColor = &H80000005
    textTglPeminjaman.BackColor = &H80000005
    textJaminan.BackColor = &H80000005
End Function

Function persiapanPeninjauan()
    sistem.isEditing = False
    buttonView.Visible = False
    buttonEditSimpan.Caption = "Edit"
    
    comboPlatMobil.Enabled = False
    comboNoKTP.Enabled = False
    textLamaPeminjaman.Enabled = False
    textTglPeminjaman.Enabled = False
    textJaminan.Enabled = False
    
    comboPlatMobil.BackColor = &H80000004
    comboNoKTP.BackColor = &H80000004
    textLamaPeminjaman.BackColor = &H80000004
    textTglPeminjaman.BackColor = &H80000004
    textJaminan.BackColor = &H80000004
End Function

Function setJudul()
    If sistem.isEditing = True Then
        Me.Caption = "ID Rental : " + judulForm + " - Editor"
    Else
        Me.Caption = "ID Rental : " + judulForm + " - Peninjauan"
    End If
End Function

Function muatInfo()
    'cek adakah id rental mobil yang dimaksud
    adodcRental.Refresh
    adodcRental.Recordset.Find "id_peminjaman='" & sistem.currRecord & "'"
    If Not adodcRental.Recordset.EOF Then
        judulForm = adodcRental.Recordset!id_peminjaman
        comboPlatMobil.Text = adodcRental.Recordset!plat_mobil
        comboNoKTP.Text = adodcRental.Recordset!KTP
                    
        adodcTableAnggota.Refresh
        adodcTableAnggota.Recordset.Find "KTP='" & comboNoKTP.Text & "'"
        If Not adodcTableAnggota.Recordset.EOF Then

            adodcTableMobil.Refresh
            adodcTableMobil.Recordset.Find "plat_mobil='" & comboPlatMobil.Text & "'"
            If Not adodcTableMobil.Recordset.EOF Then
                textNamaMobil.Text = adodcTableMobil.Recordset!nama_mobil
                textTipeMobil.Text = adodcTableMobil.Recordset!tipe_mobil
                textHargaHari.Text = adodcTableMobil.Recordset!harga_hari
                
                If Not adodcTableMobil.Recordset!alamat_gambar = "" Then
                    'Cek apakah file yang dimaksud ada dalam alamat.
                    If Dir(App.Path + adodcTableMobil.Recordset!alamat_gambar) <> "" Then
                        'jika ada maka akan merekam sesuai rekaman di database
                        imageMobil.Picture = LoadPicture(App.Path + adodcTableMobil.Recordset!alamat_gambar)
                    Else
                        MsgBox "Terjadi kesalahan dalam pencarian file gambar mobil !", vbCritical, "Rental Mobil"
                        imageMobil.Picture = LoadPicture(App.Path + "\images\photoMobil\defaultmobil.JPG")
                    End If
                Else
                    MsgBox "Terjadi kesalahan dalam pencarian file gambar mobil !", vbCritical, "Rental Mobil"
                    imageMobil.Picture = LoadPicture(App.Path + "\images\photoMobil\defaultmobil.JPG")
                End If
            Else
                MsgBox "Plat Mobil tidak ditemukan !", vbCritical, "Rental Mobil"
                Exit Function
            End If
        Else
            MsgBox "Calon Rental belum terdaftar sebagai anggota !", vbCritical, "Rental Mobil"
            Exit Function
        End If
        textLamaPeminjaman.Text = adodcRental.Recordset!lama_peminjaman
        textTglPeminjaman.Text = adodcRental.Recordset!tanggal_peminjaman
        textTotalHarga.Text = adodcRental.Recordset!harga
        textJaminan.Text = adodcRental.Recordset!jaminan
        
        defaultPlat_mobil = adodcRental.Recordset!plat_mobil
        defaultKTP = adodcRental.Recordset!KTP
    Else
        MsgBox "Ada masalah dalam pencarian data.", vbCritical, "Rental Mobil"
        Exit Function
    End If
End Function

Function simpanData()
    On Error GoTo errHandler

    If comboPlatMobil.Text = "-- Pilih --" Or comboPlatMobil = "" Then
        MsgBox "Plat Mobil Belum Dimasukan !"
        Exit Function
    End If
    
    If comboNoKTP.Text = "-- Pilih --" Or comboPlatMobil = "" Then
        MsgBox "No KTP Belum Dimasukan !"
        Exit Function
    End If
    
    If textLamaPeminjaman.Text = "" Then
        MsgBox "Lama Peminjaman Belum Dimasukan !"
        Exit Function
    End If
    
    If textJaminan.Text = "" Then
        MsgBox "Jaminan Belum Dimasukan !"
        Exit Function
    End If

    'cek adakah anggota rental mobil yang dimaksud
    adodcRental.Refresh
    adodcRental.Recordset.Find "id_peminjaman='" & sistem.currRecord & "'"
    If Not adodcRental.Recordset.EOF Then
        adodcTableAnggota.Refresh
        adodcTableAnggota.Recordset.Find "KTP='" & comboNoKTP.Text & "'"
        If Not adodcTableAnggota.Recordset.EOF Then
            adodcTableMobil.Refresh
            adodcTableMobil.Recordset.Find "plat_mobil='" & comboPlatMobil.Text & "'"
            If Not adodcTableMobil.Recordset.EOF Then
                'Check, dan ubah semua keperluan pada tabel database
                If comboNoKTP.Text = defaultKTP Then
                    isKTPChanged = False
                Else
                    If adodcTableAnggota.Recordset!status_meminjam > 0 Then
                        isKTPChanged = True
                    Else
                        psn = MsgBox("Ada kejanggalam pada data, No KTP yang dimaksud terekam sedang meminjam Mobil. apakah anda ingin melanjutkan tindakan ini ?", vbYesNo, sistem.msgTitle)
                        If psn = vbYes Then
                            isKTPChanged = True
                        Else
                            Exit Function
                        End If
                    End If
                End If
                
                If comboPlatMobil.Text = defaultPlat_mobil Then
                    isPlat_mobilChanged = False
                Else
                    If adodcTableMobil.Recordset!tersedia > 0 Then
                        isPlat_mobilChanged = True
                    Else
                        psn = MsgBox("Ada kejanggalam pada data, No Plat Mobil yang dimaksud terekam sedang dipinjam. apakah anda ingin melanjutkan tindakan ini ?", vbYesNo, sistem.msgTitle)
                        If psn = vbYes Then
                            isPlat_mobilChanged = True
                        Else
                            Exit Function
                        End If
                    End If
                End If
                
                If isKTPChanged = True Then
                    adodcTableAnggota.Refresh
                    adodcTableAnggota.Recordset.Find "KTP='" & defaultKTP & "'"
                    If Not adodcTableAnggota.Recordset.EOF Then
                        adodcTableAnggota.Recordset!status_meminjam = 1
                        adodcTableAnggota.Recordset.Update
                    Else
                        MsgBox "Kesalahan dalam mencari informasi KTP", vbCritical, "Rental Mobil"
                        Exit Function
                    End If
                    
                    adodcTableAnggota.Refresh
                    adodcTableAnggota.Recordset.Find "KTP='" & comboNoKTP.Text & "'"
                    If Not adodcTableAnggota.Recordset.EOF Then
                        adodcTableAnggota.Recordset!status_meminjam = 0
                        adodcTableAnggota.Recordset.Update
                    Else
                        MsgBox "Kesalahan dalam mencari informasi KTP", vbCritical, "Rental Mobil"
                        Exit Function
                    End If
                    adodcRental.Recordset!KTP = comboNoKTP.Text
                Else
                    adodcRental.Recordset!KTP = comboNoKTP.Text
                End If
                    
                    
                    
                If isPlat_mobilChanged = True Then
                    adodcTableMobil.Refresh
                    adodcTableMobil.Recordset.Find "plat_mobil='" & defaultPlat_mobil & "'"
                    If Not adodcTableMobil.Recordset.EOF Then
                        adodcTableMobil.Recordset!tersedia = 1
                        adodcTableMobil.Recordset.Update
                    Else
                        MsgBox "Kesalahan dalam mencari informasi Plat Mobil", vbCritical, "Rental Mobil"
                        Exit Function
                    End If
                    
                    adodcTableMobil.Refresh
                    adodcTableMobil.Recordset.Find "plat_mobil='" & comboPlatMobil.Text & "'"
                    If Not adodcTableMobil.Recordset.EOF Then
                        adodcTableMobil.Recordset!tersedia = 0
                        adodcTableMobil.Recordset.Update
                    Else
                        MsgBox "Kesalahan dalam mencari informasi Plat_mobil", vbCritical, "Rental Mobil"
                        Exit Function
                    End If
                    adodcRental.Recordset!plat_mobil = comboPlatMobil.Text
                Else
                    adodcRental.Recordset!plat_mobil = comboPlatMobil.Text
                End If
                    
                    adodcRental.Recordset!NIK = sistem.userNIK
                    adodcRental.Recordset!status_peminjaman = 1 'Status peminjaman, 1 = Sedang di Pinjam / 0=Sudah kembali
                    adodcRental.Recordset!lama_peminjaman = textLamaPeminjaman.Text
                    adodcRental.Recordset!tanggal_peminjaman = textTglPeminjaman.Text
                    adodcRental.Recordset!jaminan = textJaminan.Text
                    adodcRental.Recordset!harga = textTotalHarga.Text
                adodcRental.Recordset.Update
                
                MsgBox "Tersimpan.", , "Rental Mobil"
                Unload Me
                tableRental.Show
            Else
                MsgBox "Plat Mobil tidak ditemukan !", vbCritical, "Rental Mobil"
                Exit Function
            End If
        Else
            MsgBox "Calon Rental belum terdaftar sebagai anggota !", vbCritical, "Rental Mobil"
            Exit Function
        End If
    Else
        MsgBox "Ada masalah dalam pencarian data.", vbCritical, "Rental Mobil"
        Exit Function
    End If
    
Exit Function
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Function
