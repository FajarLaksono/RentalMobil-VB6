VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form pengembalian 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengembalian"
   ClientHeight    =   8130
   ClientLeft      =   7680
   ClientTop       =   1215
   ClientWidth     =   6900
   Icon            =   "pengembalian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   6900
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Pengembalian"
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6735
      Begin MSAdodcLib.Adodc adodcAnggota 
         Height          =   330
         Left            =   3240
         Top             =   7440
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   1680
         Top             =   7440
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
      Begin MSAdodcLib.Adodc adodcRental 
         Height          =   330
         Left            =   120
         Top             =   7440
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
      Begin VB.TextBox textTanggalPengembalian 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   2760
         Width           =   4000
      End
      Begin VB.TextBox textTanggalMeminjam 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   1800
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
      Begin VB.TextBox textPlatMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   4000
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
      Begin VB.CommandButton buttonKembalikan 
         Appearance      =   0  'Flat
         Caption         =   "Kembalikan"
         Height          =   375
         Left            =   5040
         TabIndex        =   16
         Top             =   7440
         Width           =   1335
      End
      Begin VB.Label labelHari2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   34
         Top             =   2300
         Width           =   495
      End
      Begin VB.Label labelLamaPengembalian 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lama Pengembalian"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label labelHari 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hari"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   32
         Top             =   1320
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000006&
         X1              =   120
         X2              =   6360
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label labelTotalbayar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Bayar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Label labelUangKembali 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Uang Kembali"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   6360
         Width           =   2055
      End
      Begin VB.Label labelPotongan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Potongan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label labelTotalHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Harga"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label labelDenda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Denda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label labelHarga 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Harga Awal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label labelJaminan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jaminan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label labelTanggalPengembalian 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Pengembalian"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label labelTanggalMeminjam 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Meminjam"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label labelNoKTP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nomer KTP Peminjam"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label labelPlatMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Plat Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label labelLamaPeminjaman 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lama Peminjaman"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label labelKodisiMobil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kondisi Mobil"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lama Peminjaman"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1935
      Width           =   2175
   End
End
Attribute VB_Name = "pengembalian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables
Dim dapatPotongan As Boolean
Dim tempDendaValue As Double

Dim tanggalMeminjam As String
Dim tanggalKembali As String
Dim setengahPerjanjian As Double
Dim lamaPengembalian As Double

Public Sub setConnection()
    adodcRental.ConnectionString = sistem.connectToDatabeseRentalMobil
    adodcRental.RecordSource = "select * from rental where status_peminjaman = 1" 'cari status_peminjaman = 1(sedang dipinjam) di tabel rental
    
    adodcMobil.ConnectionString = sistem.connectToDatabeseRentalMobil
    adodcMobil.RecordSource = "select * from mobil"
    
    adodcAnggota.ConnectionString = sistem.connectToDatabeseRentalMobil
    adodcAnggota.RecordSource = "select * from anggota"
End Sub

Function rumusPengembalian()
    If comboNoKTP = "-- Pilih --" Or comboNoKTP = "" Then
        Exit Function
    End If
    
    lamaPengembalian = DateDiff("d", tanggalMeminjam, tanggalKembali) 'datediff digunakan untuk menghitung jarak antara dua tanggal
    textLamaPengembalian.Text = lamaPengembalian
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

Private Sub comboNoKTP_Click()
    'mencari dan mengambil informasi peminjaman
    adodcRental.Refresh
    adodcRental.Recordset.Find "KTP='" & comboNoKTP.Text & "'"
    If Not adodcRental.Recordset.EOF Then
        textPlatMobil.Text = adodcRental.Recordset!plat_mobil
        textLamaPeminjaman = adodcRental.Recordset!lama_peminjaman
        textTanggalMeminjam.Text = adodcRental.Recordset!tanggal_peminjaman
            tanggalMeminjam = textTanggalMeminjam.Text
        textJaminan.Text = adodcRental.Recordset!jaminan
        textHarga.Text = adodcRental.Recordset!harga
    Else
       MsgBox "No KTP yang anda masukan tidak terdaftar sedang meminjam mobil.", vbInformasi, sistem.msgTitle
       Exit Sub
    End If
    
    rumusPengembalian
End Sub

Private Sub Form_Load()
    'set lokasi
    Me.Top = 840
    Me.Left = 7500
    
    'set koneksi
    setConnection
    
    'inisialisasi combo box No KTP
    adodcRental.Refresh
    adodcRental.Recordset.MoveFirst
    Do While Not adodcRental.Recordset.EOF
        comboNoKTP.AddItem adodcRental.Recordset!KTP
        adodcRental.Recordset.MoveNext
    Loop
    
    'inisialisasi
    textTanggalPengembalian.Text = Date
    tanggalKembali = textTanggalPengembalian.Text
    textPotongan.Text = "0 %"
    
    textDenda.Enabled = False
    textDenda.BackColor = &H80000004
    textDenda.Text = 0
    tempDendaValue = 0
    
    textTotalHarga.Text = 0
    textUangKembali.Text = 0
    textTotalBayar.Text = 0
    
    'set tooltip
    comboNoKTP.ToolTipText = "Pilih Nomer KTP"
    textPlatMobil.ToolTipText = "Plat Mobil"
    optionBaik.ToolTipText = "Mobil dengan keadaan baik"
    optionRusak.ToolTipText = "Mobil dengan keadaan rusak"
    optionHilang.ToolTipText = "Mobil dengan keadaan hilang"
    textDenda.ToolTipText = "Denda untuk anggota"
    textHarga.ToolTipText = "Harga awal perjanjian dan dibayar saat tahap peminjaman/perjanjian"
    textJaminan.ToolTipText = "Jaminan"
    textLamaPeminjaman.ToolTipText = "Lama Peminjaman pada awal perjanjian"
    textLamaPengembalian.ToolTipText = "Lama Pengembalian Mobil"
    textPotongan.ToolTipText = "Potongan"
    textTanggalMeminjam.ToolTipText = "Tanggal anggota meminjam Mobil"
    textTanggalPengembalian.ToolTipText = "Tanggal Pengembalian"
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
    labelTanggalMeminjam.ToolTipText = "Tanggal anggota meminjam Mobil"
    labelTanggalPengembalian.ToolTipText = "Tanggal Pengembalian"
    labelTotalHarga.ToolTipText = "Total Harga"
    labelUangKembali.ToolTipText = "Uang kembali jika mendapatkan potongan"
    labelTotalbayar.ToolTipText = "Total Bayar pada tahap pengembalian"
End Sub

Private Sub optionBaik_Click()
    rumusPengembalian
End Sub

Private Sub optionHilang_Click()
    rumusPengembalian
End Sub

Private Sub optionRusak_Click()
    rumusPengembalian
End Sub

Private Sub textDenda_Change()
    If comboNoKTP = "-- Pilih --" Or comboNoKTP = "" Then
        Exit Sub
    End If
    
    textTotalHarga.Text = Val(textHarga.Text) + Val(textDenda.Text)
    textTotalBayar.Text = textDenda.Text
End Sub

Private Sub buttonKembalikan_Click()
    On Error GoTo errHandler
    If comboNoKTP = "-- Pilih --" Or comboNoKTP = "" Then
        MsgBox "Anda belum memasukan No KTP peminjam !", vbInformasi, sistem.msgTitle
        Exit Sub
    End If
    
    If textDenda.Enabled = True And textDenda.Text = 0 Then
        X = MsgBox("Anda yakin mengakhiri peminjaman ini tanpa memberikanya denda ?", vbQuestion + vbYesNo, sistem.msgTitle)
        If X = vbNo Then
            textDenda.SetFocus
            Exit Sub
        End If
    End If
    
    'cek semua kemungkinan untuk mengembalikan
    adodcRental.Refresh
    adodcRental.Recordset.Find "KTP='" & comboNoKTP.Text & "'"
    If Not adodcRental.Recordset.EOF Then
        adodcAnggota.Refresh
        adodcAnggota.Recordset.Find "KTP='" & comboNoKTP.Text & "'"
        If Not adodcAnggota.Recordset.EOF Then
            If adodcAnggota.Recordset!status_meminjam < 1 Then
                adodcMobil.Refresh
                adodcMobil.Recordset.Find "plat_mobil='" & textPlatMobil.Text & "'"
                If Not adodcMobil.Recordset.EOF Then
                    If adodcMobil.Recordset!tersedia < 1 Then
                        adodcRental.Recordset!lama_pengembalian = textLamaPengembalian.Text
                        adodcRental.Recordset!tanggal_kembali = textTanggalPengembalian.Text
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
                        adodcRental.Recordset!total_harga = textTotalHarga.Text
                        adodcRental.Recordset!potongan = textPotongan.Text
                        adodcRental.Recordset!uang_kembali = textUangKembali.Text
                        adodcRental.Recordset!bayar_saat_pengembalian = textTotalBayar.Text
                        adodcRental.Recordset!total_bayar = Val(textTotalHarga.Text) + Val(textDenda.Text)
                        
                        adodcMobil.Recordset!tersedia = 1
                        adodcAnggota.Recordset!status_meminjam = 1
                        adodcRental.Recordset!status_meminjam = 0
                        MsgBox "Anggota sudah bisa mendapatkan Jaminan dan mobil dapat masuk ke Garasi. Terima Kasih sudah menjadi pelanggan Rental Mobil Purwokerto", vbInformasi, sistem.msgTitle
                        Unload Me
                    Else
                        MsgBox "Mobil yang terekam tidak sedang dalam peminjaman !", vbInformasi, sistem.msgTitle
                        Exit Sub
                    End If
                Else
                    MsgBox "Kesalahan dalam mencari informasi mobil di database.", vbInformasi, sistem.msgTitle
                    Exit Sub
                End If
            Else
                MsgBox "Anggota terekam tidak sedang meminjam mobil !", vbInformasi, sistem.msgTitle
                Exit Sub
            End If
        Else
            MsgBox "Kesalahan dalam mencari informasi anggota di database.", vbInformasi, sistem.msgTitle
            Exit Sub
        End If
    Else
       MsgBox "No KTP yang anda masukan tidak terdaftar sedang meminjam mobil.", vbInformasi, sistem.msgTitle
       Exit Sub
    End If
    
Exit Sub
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Sub
