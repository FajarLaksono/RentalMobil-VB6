VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form setting 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pengaturan Akun"
   ClientHeight    =   4245
   ClientLeft      =   6690
   ClientTop       =   4170
   ClientWidth     =   7170
   Icon            =   "setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timerSetting 
      Left            =   480
      Top             =   600
   End
   Begin MSAdodcLib.Adodc adodcSetting 
      Height          =   330
      Left            =   5160
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton buttonCancle 
      Caption         =   "Cancle"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton buttonSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox textPassword2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox textPassword1 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox textJabatan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox textNama 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox textNIK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image imageKaryawan 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label labelUlang 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ulangi"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label labelGantiPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ganti Password"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label labelJabatan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Jabatan"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label labelNamaKaryawan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nama Karyawan"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label labelNIK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nomer Induk Karyawan"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables
Public alamatGambar As String
Private alamatGambarTemp1 As String

Private tanggal As String
Private waktu As String
Private namaFile As String

Function setConnection()
    adodcSetting.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcSetting.RecordSource = "select * from pegawai" 'SQL
End Function

Private Sub buttonCancle_Click()
     Unload Me
End Sub

Private Sub buttonSave_Click()
    On Error GoTo errHandler
    If Not textPassword1.Text = "" Or Not alamatGambar = alamatGambarTemp1 And Not commonDialog.FileName = "" And Not commonDialog.FileTitle = "" Then
        adodcSetting.Refresh
        adodcSetting.Recordset.Find "NIK='" & sistem.userNIK & "'"
        With adodcSetting.Recordset
            If Not adodcSetting.Recordset.EOF Then
                If Not textPassword1.Text = "" Then
                    If textPassword1.Text = textPassword2.Text Then
                        !pass = textPassword1.Text
                    Else
                        MsgBox "Pengulangan Password tidak sama, mohon periksa lagi !", vbCritical, sistem.msgTitle
                    End If
                End If
                
                If Not alamatGambar = alamatGambarTemp1 And Not commonDialog.FileName = "" And Not commonDialog.FileTitle = "" Then
                    tanggal = Format(Date, "d-mmmm-yyyy")
                    waktu = Format(Time, "h-m-s")
                    
                    namaFile = sistem.userNIK + tanggal + waktu
                    
                    FileCopy commonDialog.FileName, App.Path + "/images/photoPegawai/" + namaFile + commonDialog.FileTitle
                    alamatGambar = "/images/photoPegawai/" + namaFile + commonDialog.FileTitle
                    !gambar = alamatGambar
                    sistem.getPhoto = alamatGambar
                End If
            Else
                MsgBox "Sesuatu terjadi pada Akun anda ! tidak ditemukan NIK anda.", vbInformasi, sistem.msgTitle
            End If
            .Update
        End With
        'Reflesh main menu, hasil save yang akan ditampilkan dalam mainMenu
        Unload mainMenu
        mainMenu.Show
        Unload Me
    Else
        MsgBox "Tidak ada Perubahan pada Form Setting Anda !", vbInformasi, sistem.msgTitle
        Unload Me
    End If
Exit Sub
errHandler:
    MsgBox "Harap masukan data dengan benar !", vbCritical, sistem.msgTitle
End Sub

Private Sub Form_Load()
    'set koneksi ke database
    setConnection
    
    'ambil data dari modul sistem
    textNIK = sistem.userNIK
    textNama = sistem.getName
    textJabatan = sistem.getJabatan
    alamatGambar = App.Path + sistem.getPhoto
    imageKaryawan.Picture = LoadPicture(App.Path + sistem.getPhoto)
    alamatGambarTemp1 = App.Path + sistem.getPhoto
    
    'inisialisasi ToolTip
    buttonCancle.ToolTipText = "Cancle dan kembali ke Main menu"
    buttonSave.ToolTipText = "Simpan"
    imageKaryawan.ToolTipText = "Klik untuk ganti Foto"
    textJabatan.ToolTipText = "Jabatan : " + sistem.getJabatan
    textNama.ToolTipText = "Nama : " + sistem.getName
    textNIK.ToolTipText = "NIK : " + sistem.userNIK
    textPassword1.ToolTipText = "Ganti Password"
    textPassword2.ToolTipText = "Ulangin untuk komfirmasi Password anda"
End Sub

Private Sub imageKaryawan_Click()
    'dibawah adalah cara untuk mendapatkan file, kami buat hanya bisa memilih file JPEG untuk gambar
    commonDialog.FileName = "" 'inisialisasi
    commonDialog.Filter = "JPEG Files|*.jpg|All Files|*.*" 'set format pengambilan file
    commonDialog.ShowOpen 'buka dialog pemilihan dari windows
    alamatGambar = commonDialog.FileName 'ambil alamat gambar ke variabel alamatGambar
    'dibawah adalah manipulasi string, check apakah file yang diambil adalah "" / 0
    If Len(Trim(alamatGambar)) < 1 Then
        Exit Sub
    End If
    'set imageKaryawan sesuai gambar
    imageKaryawan.Picture = LoadPicture(alamatGambar)
End Sub
