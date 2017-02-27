VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tableAnggota 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table Anggota"
   ClientHeight    =   6915
   ClientLeft      =   5040
   ClientTop       =   2070
   ClientWidth     =   12735
   Icon            =   "tableAnggota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12735
   Begin MSAdodcLib.Adodc adodcTabelAnggota 
      Height          =   330
      Left            =   5400
      Top             =   160
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
   Begin VB.CommandButton buttonClear 
      Caption         =   "X"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton buttonSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3600
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
      Top             =   100
      Width           =   2775
   End
   Begin VB.CommandButton buttonView 
      Caption         =   "View"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton buttonEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton buttonDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9720
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dataGrid 
      Bindings        =   "tableAnggota.frx":038A
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   2,5
      RowHeight       =   20
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "KTP"
         Caption         =   "KTP"
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
         DataField       =   "nama_anggota"
         Caption         =   "Nama Anggota"
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
         DataField       =   "tempat_tanggal_lahir"
         Caption         =   "Tempat, Tanggal Lahir"
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
         DataField       =   "jenis_kelamin"
         Caption         =   "Jenis Kelamin"
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
         DataField       =   "no_telp"
         Caption         =   "No. Telpone"
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
         DataField       =   "alamat"
         Caption         =   "Alamat"
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
         DataField       =   "rt_rw"
         Caption         =   "RT/RW"
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
         DataField       =   "kelDesa"
         Caption         =   "Kel / Desa"
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
         DataField       =   "kec"
         Caption         =   "Kecamatan"
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
         DataField       =   "kab"
         Caption         =   "Kabupaten"
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
         DataField       =   "kode_pos"
         Caption         =   "Kode Pos"
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
         DataField       =   "pekerjaan"
         Caption         =   "Pekerjaan"
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
         DataField       =   "no_sim"
         Caption         =   "No SIM"
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
         DataField       =   "foto_anggota"
         Caption         =   "Foto Anggota"
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
         DataField       =   "status_meminjam"
         Caption         =   "Status Meminjam"
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
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1709,858
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton buttonClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   11160
      TabIndex        =   7
      Top             =   120
      Width           =   1455
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
      TabIndex        =   8
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "tableAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function setConnection()
    adodcTabelAnggota.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcTabelAnggota.RecordSource = "select * from anggota" 'SQL
    adodcTabelAnggota.Refresh
End Function

Private Sub buttonClear_Click()
    adodcTabelAnggota.Refresh
    textSearch.SetFocus
    buttonClear.Enabled = False
End Sub

Private Sub buttonClose_Click()
    Unload Me
End Sub

Private Sub buttonDelete_Click()
    psn = MsgBox("Anda yakin menghapus data ini?", vbInformation + vbYesNo, sistem.msgTitle)
    If psn = vbYes Then
        adodcTabelAnggota.Recordset.Delete
        MsgBox "Data berhasil dihapus.", vbInformation, sistem.msgTitle
    Else
        MsgBox "Data gagal dihapus.", vbExclamation, sistem.msgTitle
    End If
End Sub

Private Sub buttonSearch_Click()
    If textSearch.Text = "" Then
        adodcRental.Refresh
        textSearch.SetFocus
        buttonClear.Enabled = False
        Exit Sub
    End If

    adodcTabelAnggota.Recordset.Filter = "KTP = '" & textSearch.Text & "'"
    buttonClear.Enabled = True
End Sub

Private Sub buttonEdit_Click()
    sistem.isEditing = True
    sistem.currRecord = dataGrid.Columns(0).Value
    Unload Me
    editView_Anggota.Show
End Sub

Private Sub buttonView_Click()
    sistem.isEditing = False
    sistem.currRecord = dataGrid.Columns(0).Value
    Unload Me
    editView_Anggota.Show
End Sub

Private Sub Form_Load()
    setConnection
    
    Me.Left = 3360
    Me.Top = 1695
    
    textSearch.ToolTipText = "Mencari Berdasarkan KTP"
    buttonClear.ToolTipText = "Bersihkan Hasil Pencarian"
    buttonClear.Enabled = False
    
    colInisialisasi
End Sub

Function colInisialisasi()
    dataGrid.Columns(0).Caption = "KTP"
    dataGrid.Columns(1).Caption = "Nama Anggota"
    dataGrid.Columns(2).Caption = "Tempat, Tanggal Lahir"
    dataGrid.Columns(3).Caption = "Jenis Kelamin"
    dataGrid.Columns(4).Caption = "No. Telpone"
    dataGrid.Columns(5).Caption = "Alamat"
    dataGrid.Columns(6).Caption = "RT/RW"
    dataGrid.Columns(7).Caption = "Kel/Desa"
    dataGrid.Columns(8).Caption = "Kecamatan"
    dataGrid.Columns(9).Caption = "Kabupaten"
    dataGrid.Columns(10).Caption = "Kode Pos"
    dataGrid.Columns(11).Caption = "Pekerjaan"
    dataGrid.Columns(12).Caption = "No SIM"
    dataGrid.Columns(13).Caption = "Foto Anggota"
    dataGrid.Columns(14).Caption = "Status Meminjam"
End Function

Private Sub dataGrid_Click()
    On Error Resume Next
End Sub
