VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tableMobil 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabel Mobil"
   ClientHeight    =   5820
   ClientLeft      =   5550
   ClientTop       =   2760
   ClientWidth     =   11295
   Icon            =   "tableMobil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11295
   Begin VB.CommandButton buttonDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton buttonEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton buttonView 
      Caption         =   "View"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   1095
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
      IMEMode         =   3  'DISABLE
      Left            =   720
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton buttonSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton buttonClear 
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton buttonClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9960
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dataGrid 
      Bindings        =   "tableMobil.frx":038A
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   3
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "plat_mobil"
         Caption         =   "plat_mobil"
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
         DataField       =   "nama_mobil"
         Caption         =   "nama_mobil"
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
         DataField       =   "tipe_mobil"
         Caption         =   "tipe_mobil"
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
         DataField       =   "harga_hari"
         Caption         =   "harga_hari"
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
         DataField       =   "tersedia"
         Caption         =   "tersedia"
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
         DataField       =   "alamat_gambar"
         Caption         =   "alamat_gambar"
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
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2250,142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1769,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2264,882
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodcDaftarMobil 
      Height          =   330
      Left            =   4825
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
      Caption         =   "PLAT"
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
      Top             =   190
      Width           =   735
   End
End
Attribute VB_Name = "tableMobil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function setConnection()
    adodcDaftarMobil.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcDaftarMobil.RecordSource = "select * from mobil" 'SQL
    adodcDaftarMobil.Refresh
End Function

Private Sub buttonDelete_Click()
    psn = MsgBox("Anda yakin menghapus data ini?", vbInformation + vbYesNo, sistem.msgTitle)
    If psn = vbYes Then
        adodcDaftarMobil.Recordset.Delete
        MsgBox "Data berhasil dihapus.", vbInformation, sistem.msgTitle
    Else
        MsgBox "Data gagal dihapus.", vbExclamation, sistem.msgTitle
    End If
End Sub

Private Sub buttonClear_Click()
    adodcDaftarMobil.Refresh
    textSearch.SetFocus
    buttonClear.Enabled = False
End Sub

Private Sub buttonClose_Click()
    Unload Me
End Sub

Private Sub buttonSearch_Click()
    If textSearch.Text = "" Then
        adodcRental.Refresh
        textSearch.SetFocus
        buttonClear.Enabled = False
        Exit Sub
    End If
    
    adodcDaftarMobil.Recordset.Filter = "plat_mobil = '" & textSearch.Text & "'"
    buttonClear.Enabled = True
End Sub

Private Sub buttonEdit_Click()
    sistem.isEditing = True
    sistem.currRecord = dataGrid.Columns(0).Value
    Unload Me
    editView_mobilBaru.Show
End Sub

Private Sub buttonView_Click()
    sistem.isEditing = False
    sistem.currRecord = dataGrid.Columns(0).Value
    Unload Me
    editView_mobilBaru.Show
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
    dataGrid.Columns(0).Caption = "Plat Mobil"
    dataGrid.Columns(1).Caption = "Nama Mobil"
    dataGrid.Columns(2).Caption = "Tipe Mobil"
    dataGrid.Columns(3).Caption = "Harga Perhari"
    dataGrid.Columns(4).Caption = "Tersedia"
    dataGrid.Columns(5).Caption = "Alamat Gambar"
End Function
