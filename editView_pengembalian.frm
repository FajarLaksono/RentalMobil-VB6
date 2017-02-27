VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form editView_pengembalian 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peninjauan / Editor"
   ClientHeight    =   7950
   ClientLeft      =   7140
   ClientTop       =   1830
   ClientWidth     =   6645
   Icon            =   "editView_pengembalian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   6645
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Pengembalian"
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox comboPlatMobil 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Text            =   "-- Pilih --"
         Top             =   840
         Width           =   4000
      End
      Begin VB.CommandButton buttonEditSimpan 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         Height          =   375
         Left            =   4080
         TabIndex        =   16
         Top             =   7440
         Width           =   1095
      End
      Begin VB.CommandButton buttonView 
         Caption         =   "View"
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   7440
         Width           =   1095
      End
      Begin VB.CommandButton buttonClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   7440
         Width           =   1095
      End
      Begin VB.OptionButton optionHilang 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hilang"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   3240
         Width           =   975
      End
      Begin VB.OptionButton optionRusak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rusak"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   3240
         Width           =   975
      End
      Begin VB.OptionButton optionBaik 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Baik"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   3240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ComboBox comboNoKTP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   0
         Text            =   "-- Pilih --"
         Top             =   360
         Width           =   4000
      End
      Begin VB.TextBox textLamaPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   1300
         Width           =   3405
      End
      Begin VB.TextBox textTanggalMeminjam 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   1800
         Width           =   4000
      End
      Begin VB.TextBox textTanggalPengembalian 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   2760
         Width           =   4000
      End
      Begin VB.TextBox textJaminan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   3720
         Width           =   4000
      End
      Begin VB.TextBox textHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   4200
         Width           =   4000
      End
      Begin VB.TextBox textDenda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   5160
         Width           =   4000
      End
      Begin VB.TextBox textTotalHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   5640
         Width           =   4000
      End
      Begin VB.TextBox textPotongan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         ToolTipText     =   "Uang Kembali 50% jika Mobil dikembalikan dalam setengah perjanjian"
         Top             =   4680
         Width           =   3975
      End
      Begin VB.TextBox textUangKembali 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   6360
         Width           =   3975
      End
      Begin VB.TextBox textTotalBayar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   6840
         Width           =   3975
      End
      Begin VB.TextBox textLamaPengembalian 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   2280
         Width           =   3405
      End
      Begin MSAdodcLib.Adodc adodcAnggota 
         Height          =   330
         Left            =   120
         Top             =   7560
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
      Begin MSAdodcLib.Adodc adodcMobil 
         Height          =   330
         Left            =   120
         Top             =   7320
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
         Top             =   7080
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
      Begin VB.Label labelHari2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   35
         Top             =   2300
         Width           =   495
      End
      Begin VB.Label labelKodisiMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kondisi Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label labelLamaPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lama Peminjaman"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label labelPlatMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Plat Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label labelNoKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nomer KTP Peminjam"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label labelTanggalMeminjam 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Meminjam"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label labelTanggalPengembalian 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Pengembalian"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label labelJaminan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jaminan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label labelHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Harga Awal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label labelDenda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Denda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label labelTotalHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Harga"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label labelPotongan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Potongan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label labelUangKembali 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Uang Kembali"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6360
         Width           =   2055
      End
      Begin VB.Label labelTotalbayar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Bayar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000006&
         X1              =   120
         X2              =   6360
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label labelHari 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   21
         Top             =   1330
         Width           =   495
      End
      Begin VB.Label labelLamaPengembalian 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lama Pengembalian"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1575
      End
   End
End
Attribute VB_Name = "editView_pengembalian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dapatPotongan As Boolean
Dim tempDendaValue As Double
Dim judulForm As String
Dim tanggalMeminjam As String
Dim tanggalKembali As String
Dim setengahPerjanjian As Double
Dim lamaPengembalian As Double

Dim isReady As Boolean

Function setConnection()
    adodcMobil.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcMobil.RecordSource = "SELECT plat_mobil,harga_hari FROM mobil" 'SQL
    
    adodcAnggota.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcAnggota.RecordSource = "SELECT KTP FROM anggota" 'SQL
        
    adodcRental.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcRental.RecordSource = "SELECT * FROM rental WHERE status_peminjaman=0" 'SQL
End Function

Private Sub buttonClose_Click()
    Unload Me
    tableRental.Show
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
    isReady = False

    setConnection
    If sistem.isEditing = True Then
        persiapanEditing
    Else
        persiapanPeninjauan
    End If
    
    'inisialisasi combo box
    adodcMobil.Refresh
    adodcMobil.Recordset.MoveFirst
    Do While Not adodcMobil.Recordset.EOF
        comboPlatMobil.AddItem adodcMobil.Recordset!plat_mobil
        adodcMobil.Recordset.MoveNext
    Loop
    
    adodcAnggota.Refresh
    adodcAnggota.Recordset.MoveFirst
    Do While Not adodcAnggota.Recordset.EOF
        comboNoKTP.AddItem adodcAnggota.Recordset!KTP
        adodcAnggota.Recordset.MoveNext
    Loop
    
    'Set lokasi
    Me.Top = 1815
    Me.Left = 7110
    
    setConst
    
    muatInfo
    setJudul
End Sub

Function simpanData()
    On Error GoTo errHandler
    
    If comboNoKTP = "-- Pilih --" Or comboNoKTP = "" Then
        MsgBox "Anda belum memasukan No KTP peminjam !", vbInformasi, sistem.msgTitle
        Exit Function
    End If
    
    If textDenda.Enabled = True And textDenda.Text = 0 Then
        X = MsgBox("Anda yakin mengakhiri peminjaman ini tanpa memberikanya denda ?", vbQuestion + vbYesNo, sistem.msgTitle)
        If X = vbNo Then
            textDenda.SetFocus
            Exit Function
        End If
    End If

    adodcRental.Refresh
    adodcRental.Recordset.Find "id_peminjaman='" & judulForm & "'"
    If Not adodcRental.Recordset.EOF Then
        adodcRental.Recordset!KTP = comboNoKTP.Text
        adodcRental.Recordset!plat_mobil = comboPlatMobil.Text
        
        If optionBaik.Value = True And optionRusak.Value = False And optionHilang.Value = False Then
            adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "Baik"
        End If
        
        If optionBaik.Value = False And optionRusak.Value = True And optionHilang.Value = False Then
            adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "Rusak"
        End If
        
        If optionBaik.Value = False And optionRusak.Value = False And optionHilang.Value = True Then
            adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "Hilang"
        End If

        adodcRental.Recordset!denda = textDenda.Text
        adodcRental.Recordset!harga = textHarga.Text
        adodcRental.Recordset!jaminan = textJaminan.Text
        adodcRental.Recordset!lama_peminjaman = textLamaPeminjaman.Text
        adodcRental.Recordset!lama_pengembalian = textLamaPengembalian.Text
        adodcRental.Recordset!potongan = textPotongan.Text
        adodcRental.Recordset!tanggal_peminjaman = textTanggalMeminjam.Text
        adodcRental.Recordset!tanggal_kembali = textTanggalPengembalian.Text
        adodcRental.Recordset!total_harga = textTotalHarga.Text
        adodcRental.Recordset!uang_kembali = textUangKembali.Text
        adodcRental.Recordset!total_bayar = textTotalBayar.Text
        adodcRental.Recordset.Update
        MsgBox "Tersimpan.", vbInformasi, sistem.msgTitle
        Unload Me
        tableRental.Show
    Else
        MsgBox "Terjadi kesalahan ketika penyimpanan ke database", vbCritical, sistem.msgTitle
    End If
    Exit Function
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Function

Function setJudul()
    If sistem.isEditing = True Then
        Me.Caption = "ID Rental : " + judulForm + " - Editor"
    Else
        Me.Caption = "ID Rental : " + judulForm + " - Peninjauan"
    End If
End Function

Function muatInfo()
    adodcRental.Refresh
    adodcRental.Recordset.Find "id_peminjaman='" & sistem.currRecord & "'"
    If Not adodcRental.Recordset.EOF Then
        comboNoKTP.Text = adodcRental.Recordset!KTP
        comboPlatMobil.Text = adodcRental.Recordset!plat_mobil
        
        If adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "Baik" Or adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "baik" Then
            optionBaik.Value = True
            optionRusak.Value = False
            optionHilang.Value = False
        End If
        
        If adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "Rusak" Or adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "rusak" Then
            optionBaik.Value = False
            optionRusak.Value = True
            optionHilang.Value = False
        End If
        
        If adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "Hilang" Or adodcRental.Recordset!kondisi_mobil_setelah_dipinjam = "hilang" Then
            optionBaik.Value = False
            optionRusak.Value = False
            optionHilang.Value = True
        End If

        textDenda.Text = adodcRental.Recordset!denda
        textHarga.Text = adodcRental.Recordset!harga
        textJaminan.Text = adodcRental.Recordset!jaminan
        textLamaPeminjaman.Text = adodcRental.Recordset!lama_peminjaman
        textLamaPengembalian.Text = adodcRental.Recordset!lama_pengembalian
        textPotongan.Text = adodcRental.Recordset!potongan
        textTanggalMeminjam.Text = adodcRental.Recordset!tanggal_peminjaman
        textTanggalPengembalian.Text = adodcRental.Recordset!tanggal_kembali
        textTotalHarga.Text = adodcRental.Recordset!total_harga
        textUangKembali.Text = adodcRental.Recordset!uang_kembali
        textTotalBayar.Text = adodcRental.Recordset!total_bayar
        
        judulForm = adodcRental.Recordset!id_peminjaman
        isReady = True
    Else
        MsgBox "Terjadi kesalahan ketika sedang memuat informasi", vbCritical, sistem.msgTitle
        isReady = False
    End If
    
Exit Function
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Function

Function setConst()
    textLamaPengembalian.Enabled = False
    textTotalHarga.Enabled = False
    textUangKembali.Enabled = False
    textTotalBayar.Enabled = False
    textHarga.Enabled = False
    textPotongan.Enabled = False
    
    textLamaPengembalian.BackColor = &H80000004
    textTotalHarga.BackColor = &H80000004
    textUangKembali.BackColor = &H80000004
    textTotalBayar.BackColor = &H80000004
    textHarga.BackColor = &H80000004
    textPotongan.BackColor = &H80000004
    
    'Set Tooltip
    comboNoKTP.ToolTipText = "Pilih Nomer KTP"
    comboPlatMobil.ToolTipText = "Pilih Mobil"
    optionBaik.ToolTipText = "Mobil dengan keadaan baik"
    optionRusak.ToolTipText = "Mobil dengan keadaan rusak"
    optionHilang.ToolTipText = "Mobil dengan keadaan hilang"
    textDenda.ToolTipText = "Denda untuk anggota"
    textHarga.ToolTipText = "Harga awal perjanjian dan dibayar saat tahap peminjaman/perjanjian"
    textJaminan.ToolTipText = "Jaminan"
    textLamaPeminjaman.ToolTipText = "Lama Peminjaman pada awal perjanjian"
    textLamaPengembalian.ToolTipText = "Lama Pengembalian Mobil"
    textPotongan.ToolTipText = "Potongan"
    textTanggalMeminjam.ToolTipText = "Tanggal anggota meminjam Mobil, (dd/mm/yyyy)"
    textTanggalPengembalian.ToolTipText = "Tanggal Pengembalian, (dd/mm/yyyy)"
    textTotalHarga.ToolTipText = "Total Harga"
    textUangKembali.ToolTipText = "Uang kembali jika mendapatkan potongan"
    textTotalBayar.ToolTipText = "Total Bayar pada tahap pengembalian"
    
    labelNoKTP.ToolTipText = "Pilih Nomer KTP"
    labelPlatMobil.ToolTipText = "Plat Mobil"
    optionBaik.ToolTipText = "Mobil dengan keadaan baik"
    optionRusak.ToolTipText = "Mobil dengan keadaan rusak"
    optionHilang.ToolTipText = "Mobil dengan keadaan hilang"
    labelDenda.ToolTipText = "Denda untuk anggota"
    labelHarga.ToolTipText = "Harga awal perjanjian dan dibayar saat tahap peminjaman/perjanjian"
    labelJaminan.ToolTipText = "Jaminan"
    labelLamaPeminjaman.ToolTipText = "Lama Peminjaman pada awal perjanjian"
    labelLamaPengembalian.ToolTipText = "Lama Pengembalian Mobil"
    labelPotongan.ToolTipText = "Potongan"
    labelTanggalMeminjam.ToolTipText = "Tanggal anggota meminjam Mobil, (dd/mm/yyyy)"
    labelTanggalPengembalian.ToolTipText = "Tanggal Pengembalian, (dd/mm/yyyy)"
    labelTotalHarga.ToolTipText = "Total Harga"
    labelUangKembali.ToolTipText = "Uang kembali jika mendapatkan potongan"
    labelTotalbayar.ToolTipText = "Total Bayar pada tahap pengembalian"
End Function

Function persiapanEditing()
    sistem.isEditing = True
    buttonView.Visible = True
    buttonEditSimpan.Caption = "Simpan"
    
    comboNoKTP.Enabled = True
    comboPlatMobil.Enabled = True
    
    optionBaik.Enabled = True
    optionHilang.Enabled = True
    optionRusak.Enabled = True
    
    textDenda.Enabled = True
    'textHarga.Enabled = True
    textJaminan.Enabled = True
    textLamaPeminjaman.Enabled = True
    'textLamaPengembalian.Enabled = True
    'textPotongan.Enabled = True
    textTanggalMeminjam.Enabled = True
    textTanggalPengembalian.Enabled = True
    
    comboNoKTP.BackColor = &H80000005
    comboPlatMobil.BackColor = &H80000005
    
    optionBaik.ForeColor = &H80000008
    optionHilang.ForeColor = &H80000008
    optionRusak.ForeColor = &H80000008
    
    textDenda.BackColor = &H80000005
    textHarga.BackColor = &H80000005
    textJaminan.BackColor = &H80000005
    textLamaPeminjaman.BackColor = &H80000005
    textLamaPengembalian.BackColor = &H80000005
    textPotongan.BackColor = &H80000005
    textTanggalMeminjam.BackColor = &H80000005
    textTanggalPengembalian.BackColor = &H80000005
End Function
    
Function persiapanPeninjauan()
    sistem.isEditing = False
    buttonView.Visible = False
    buttonEditSimpan.Caption = "Edit"
    
    comboNoKTP.Enabled = False
    comboPlatMobil.Enabled = False
    
    optionBaik.Enabled = False
    optionHilang.Enabled = False
    optionRusak.Enabled = False
    
    textDenda.Enabled = False
    'textHarga.Enabled = False
    textJaminan.Enabled = False
    textLamaPeminjaman.Enabled = False
    'textLamaPengembalian.Enabled = False
    'textPotongan.Enabled = False
    textTanggalMeminjam.Enabled = False
    textTanggalPengembalian.Enabled = False
    
    comboNoKTP.BackColor = &H80000004
    comboPlatMobil.BackColor = &H80000004
    
    optionBaik.ForeColor = &H80000006
    optionHilang.ForeColor = &H80000006
    optionRusak.ForeColor = &H80000006
    
    textDenda.BackColor = &H80000004
    textHarga.BackColor = &H80000004
    textJaminan.BackColor = &H80000004
    textLamaPeminjaman.BackColor = &H80000004
    textLamaPengembalian.BackColor = &H80000004
    textPotongan.BackColor = &H80000004
    textTanggalMeminjam.BackColor = &H80000004
    textTanggalPengembalian.BackColor = &H80000004
End Function

Private Sub optionBaik_Click()
    rumusPengembalian
End Sub

Private Sub optionHilang_Click()
    rumusPengembalian
End Sub

Private Sub optionRusak_Click()
    rumusPengembalian
End Sub

Private Sub textLamaPeminjaman_Change()
    adodcMobil.Refresh
    adodcMobil.Recordset.Find "plat_mobil='" & comboPlatMobil.Text & "'"
    If Not adodcMobil.Recordset.EOF Then
        textHarga.Text = Val(adodcMobil.Recordset!harga_hari) * Val(textLamaPeminjaman.Text)
    Else
        MsgBox "Terjadi kesalahan dalam mencari informasi mobil", vbInformation, sistem.msgTitle
        Exit Sub
    End If
End Sub

Private Sub textDenda_Change()
    If isReady = False Then
        Exit Sub
    End If
    
    textTotalHarga.Text = Val(textHarga.Text) + Val(textDenda.Text)
    textTotalBayar.Text = textDenda.Text
End Sub

Private Sub comboPlatMobil_Click()
    adodcMobil.Refresh
    adodcMobil.Recordset.Find "plat_mobil='" & comboPlatMobil.Text & "'"
    If Not adodcMobil.Recordset.EOF Then
        textHarga.Text = Val(adodcMobil.Recordset!harga_hari) * Val(textLamaPeminjaman.Text)
    Else
        MsgBox "Terjadi kesalahan dalam mencari informasi mobil", vbInformation, sistem.msgTitle
        Exit Sub
    End If
End Sub

Private Sub textLamaPengembalian_Change()
    rumusPengembalian
End Sub

Private Sub textTanggalPengembalian_change()
    If isReady = False Then
        Exit Sub
    End If

    On Error GoTo errHitungLamaPengembalian

    tanggalMeminjam = textTanggalMeminjam.Text
    tanggalKembali = textTanggalPengembalian.Text
    lamaPengembalian = DateDiff("d", tanggalMeminjam, tanggalKembali) 'datediff digunakan untuk menghitung jarak antara dua tanggal
    textLamaPengembalian.Text = lamaPengembalian
    
    textTanggalMeminjam.BackColor = &H80000005
    textTanggalPengembalian.BackColor = &H80000005
    Exit Sub
errHitungLamaPengembalian:
    MsgBox "Harap masukan tanggal dengan benar (DD/MM/YYYY)", vbCritical, sistem.msgTitle
    textTanggalMeminjam.BackColor = &HC0E0FF
    textTanggalPengembalian.BackColor = &HC0E0FF
End Sub

Private Sub textTanggalMeminjam_Change()
    If isReady = False Then
        Exit Sub
    End If
    
    On Error GoTo errHitungLamaPengembalian
    
    tanggalMeminjam = textTanggalMeminjam.Text
    tanggalKembali = textTanggalPengembalian.Text
    lamaPengembalian = DateDiff("d", tanggalMeminjam, tanggalKembali) 'datediff digunakan untuk menghitung jarak antara dua tanggal
    textLamaPengembalian.Text = lamaPengembalian
    
    textTanggalMeminjam.BackColor = &H80000005
    textTanggalPengembalian.BackColor = &H80000005
    Exit Sub
errHitungLamaPengembalian:
    MsgBox "Harap masukan tanggal dengan benar (DD/MM/YYYY)", vbCritical, sistem.msgTitle
    textTanggalMeminjam.BackColor = &HC0E0FF
    textTanggalPengembalian.BackColor = &HC0E0FF
End Sub

Function rumusPengembalian()
    If isReady = False Then
        Exit Function
    End If
    
    setengahPerjanjian = Val(textLamaPeminjaman.Text) / 2 'jika lama pengembalian kurang dari setengah perjanjian maka akan dapat potongan 50%
    If lamaPengembalian < setengahPerjanjian Then 'potongan
        If optionBaik.Value = True And optionRusak.Value = False And optionHilang.Value = False Then
            'Tidak Mendapatkan denda dan mendapatkan potongan
            dapatPotongan = True
            textPotongan.Text = "50 %"
            textDenda.Enabled = False
            tempDendaValue = textDenda.Text
            textDenda.BackColor = &H80000004
            textDenda.Text = 0
            
            textUangKembali.Text = Val(textHarga.Text) / 2
            textTotalHarga.Text = textUangKembali.Text
            textTotalBayar.Text = 0
            
        Else
            'tidak mendapatkan potongan dan mendapatkan denda
            dapatPotongan = False
            textPotongan.Text = "0 %"
            textDenda.Enabled = True
            textDenda.BackColor = &H80000005
            textDenda.Text = tempDendaValue
            
            textUangKembali.Text = 0
            textTotalHarga.Text = Val(textHarga.Text) + Val(textDenda.Text)
            textTotalBayar.Text = textDenda.Text
        End If
    Else
        If Not lamaPengembalian > textLamaPeminjaman.Text Then
            'tidak mendapatkan potongan dan tidak mendapatkan denda
            textPotongan.Text = "0 %" 'aman
            dapatPotongan = False
            
            textDenda.Enabled = False
            tempDendaValue = textDenda.Text
            textDenda.BackColor = &H80000004
            textDenda.Text = 0
            
            textUangKembali.Text = 0
            textTotalHarga.Text = textHarga.Text
            textTotalBayar.Text = 0
        Else
            'tidak mendapatkan potongan dan mendapatkan denda
            dapatPotongan = False
            textPotongan.Text = "0 %"
            textDenda.Enabled = True
            textDenda.BackColor = &H80000005
            textDenda.Text = tempDendaValue
            
            textUangKembali.Text = 0
            textTotalHarga.Text = Val(textHarga.Text) + Val(textDenda.Text)
            textTotalBayar.Text = textDenda.Text
        End If
    End If
End Function
