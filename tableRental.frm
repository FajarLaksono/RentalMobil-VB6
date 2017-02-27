VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tableRental 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table Rental"
   ClientHeight    =   5760
   ClientLeft      =   4365
   ClientTop       =   2940
   ClientWidth     =   12960
   Icon            =   "tableRental.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   12960
   Begin VB.ComboBox comboPeminjaman 
      Height          =   315
      Left            =   6240
      TabIndex        =   4
      Text            =   "Semua data"
      Top             =   170
      Width           =   1575
   End
   Begin VB.CommandButton tombolClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   11640
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton buttonClear 
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton buttonSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox textSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton buttonView 
      Caption         =   "View"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton buttonEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton buttonDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dataGrid 
      Bindings        =   "tableRental.frx":038A
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "id_peminjaman"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "KTP"
         Caption         =   "KTP Peminjam"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "plat_mobil"
         Caption         =   "Plat Mobil"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "NIK"
         Caption         =   "Petugas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "status_peminjaman"
         Caption         =   "Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "lama_peminjaman"
         Caption         =   "lama Peminjaman"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "tanggal_peminjaman"
         Caption         =   "Tanggal Peminjaman"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "lama_pengembalian"
         Caption         =   "Lama Pengembalian"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "tanggal_kembali"
         Caption         =   "Tanggal Kembali"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "kondisi_mobil_setelah_dipinjam"
         Caption         =   "Kondisi Mobil"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "jaminan"
         Caption         =   "Jaminan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "harga"
         Caption         =   "Harga"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "denda"
         Caption         =   "Denda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "total_harga"
         Caption         =   "Total Harga"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "potongan"
         Caption         =   "Potongan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "uang_kembali"
         Caption         =   "Uang Kembali"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "bayar_saat_pengembalian"
         Caption         =   "Bayar saat pengembalian"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "total_bayar"
         Caption         =   "Total Bayar"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   329,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1425,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1590,236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1590,236
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1574,929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1484,787
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1649,764
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodcRental 
      Height          =   330
      Left            =   4830
      Top             =   165
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "KTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   195
      Width           =   735
   End
End
Attribute VB_Name = "tableRental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function setConnection()
    adodcRental.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcRental.RecordSource = "select * from rental" 'SQL
    adodcRental.Refresh
End Function

Private Sub buttonClear_Click()
    adodcRental.Refresh
    textSearch.SetFocus
    buttonClear.Enabled = False
End Sub

Private Sub buttonSearch_Click()
    If textSearch.Text = "" Then
        adodcRental.Refresh
        textSearch.SetFocus
        buttonClear.Enabled = False
        Exit Sub
    End If

    If comboPeminjaman.Text = "Semua data" Then
        adodcRental.Recordset.Filter = "KTP = '" & textSearch.Text & "'"
        buttonClear.Enabled = True
    End If
    
    If comboPeminjaman.Text = "Peminjaman" Then
        adodcRental.Recordset.Filter = ""
        
        adodcRental.Recordset.Filter = "KTP = '" & textSearch.Text & "' and  status_peminjaman= ' " & 1 & "'"
        buttonClear.Enabled = True
    End If
    
    If comboPeminjaman.Text = "Pengembalian" Then
        adodcRental.Recordset.Filter = "KTP = '" & textSearch.Text & "' and status_peminjaman= ' " & 0 & "'"
        buttonClear.Enabled = True
    End If

End Sub

Private Sub buttonEdit_Click()
    If dataGrid.Columns(4).Value < 1 Then
        'pengembalian
        sistem.isEditing = True
        sistem.currRecord = dataGrid.Columns(0).Value
        Unload Me
        editView_pengembalian.Show
    Else
        'peminjaman
        sistem.isEditing = True
        sistem.currRecord = dataGrid.Columns(0).Value
        Unload Me
        editView_peminjaman.Show
    End If
End Sub

Private Sub buttonView_Click()
    If dataGrid.Columns(4).Value < 1 Then
        'pengembalian
        sistem.isEditing = False
        sistem.currRecord = dataGrid.Columns(0).Value
        Unload Me
        editView_pengembalian.Show
    Else
        'peminjaman
        sistem.isEditing = False
        sistem.currRecord = dataGrid.Columns(0).Value
        Unload Me
        editView_peminjaman.Show
    End If
End Sub

Private Sub comboPeminjaman_Click()
    If comboPeminjaman.Text = "Semua data" Then
        adodcRental.Refresh
    End If
    
    If comboPeminjaman.Text = "Peminjaman" Then
        adodcRental.Recordset.Filter = "status_peminjaman= '" & 1 & "'"
    End If
    
    If comboPeminjaman.Text = "Pengembalian" Then
        adodcRental.Recordset.Filter = "status_peminjaman= '" & 0 & "'"
    End If
End Sub

Private Sub Form_Load()
    setConnection

    comboPeminjaman.AddItem "Semua data"
    comboPeminjaman.AddItem "Peminjaman"
    comboPeminjaman.AddItem "Pengembalian"
    
    Me.Top = 2565
    Me.Left = 4320
    
    textSearch.ToolTipText = "Mencari Berdasarkan KTP"
    buttonClear.ToolTipText = "Bersihkan Hasil Pencarian"
    buttonClear.Enabled = False
    
    colInisialisasi
End Sub

Function colInisialisasi()
    dataGrid.Columns(0).Caption = "ID"
    dataGrid.Columns(1).Caption = "KTP Peminjam"
    dataGrid.Columns(2).Caption = "Plat Mobil"
    dataGrid.Columns(3).Caption = "Petugas"
    dataGrid.Columns(4).Caption = "Status"
    dataGrid.Columns(5).Caption = "Lama Peminjaman"
    dataGrid.Columns(6).Caption = "Tanggal Peminjaman"
    dataGrid.Columns(7).Caption = "Lama Pengembalian"
    dataGrid.Columns(8).Caption = "Tanggal Kembali"
    dataGrid.Columns(9).Caption = "Kondisi Mobil"
    dataGrid.Columns(10).Caption = "Jaminan"
    dataGrid.Columns(11).Caption = "Harga"
    dataGrid.Columns(12).Caption = "Denda"
    dataGrid.Columns(13).Caption = "Total Harga"
    dataGrid.Columns(14).Caption = "Potongan"
    dataGrid.Columns(15).Caption = "Uang Kembali"
    dataGrid.Columns(16).Caption = "Bayar saat Pengembalian"
    dataGrid.Columns(17).Caption = "Total Bayar"
        
End Function

Private Sub tombolClose_Click()
    Unload Me
End Sub

Private Sub buttonDelete_Click()
    psn = MsgBox("Anda yakin menghapus data ini?", vbInformation + vbYesNo, sistem.msgTitle)
    If psn = vbYes Then
        adodcRental.Recordset.Delete
        MsgBox "Data berhasil dihapus.", vbInformation, sistem.msgTitle
    Else
        MsgBox "Data gagal dihapus.", vbExclamation, sistem.msgTitle
    End If
End Sub
