VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login - Rental Mobil Purwokerto V0.5"
   ClientHeight    =   5580
   ClientLeft      =   5805
   ClientTop       =   3360
   ClientWidth     =   9915
   ClipControls    =   0   'False
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":038A
   ScaleHeight     =   5580
   ScaleWidth      =   9915
   Begin VB.Timer loadingBarTime 
      Left            =   1920
      Top             =   4920
   End
   Begin MSAdodcLib.Adodc adodcLogin 
      Height          =   330
      Left            =   240
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "AdodcLogin"
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
   Begin VB.TextBox textPassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "*"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton loginButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "Login"
      Height          =   350
      Left            =   2580
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   3050
      UseMaskColor    =   -1  'True
      Width           =   800
   End
   Begin VB.TextBox textNIK 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   0
      Text            =   "NIK"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Shape loadingBar 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Label labelNotif 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NIK yang dimasukan Salah !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   9495
   End
   Begin VB.Shape shapeLoginForm 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H80000004&
      BorderStyle     =   6  'Inside Solid
      Height          =   1575
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   2895
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim un_touched As Boolean
Dim un_touched_pass As Boolean
'un_toched (alternatif tool tip) akan digunakan untuk mengetahui apakah textBox NIK dan Password kosong atau isi
'Jika kosong akan diisi dengan "NIK" dan " * " pada textNIK dan textPassword
'jika textBox sudah terisi maka akan tetap menampilkan apa yang user isikan

Function setConnection()
    adodcLogin.ConnectionString = sistem.connectToDatabeseRentalMobil 'mengambil string pada modul sistem
    adodcLogin.RecordSource = "select * from pegawai" 'SQL
End Function

Function clear_unInputForm()
    'function ini akan memeriksa apakah textBox Kosong
    'Dan akan diganti menjadi "NIK" atau " * "
    If textNIK.Text = "" Then
        textNIK.Text = "NIK"
        un_touched = True
    End If
    
    If textPassword.Text = "" Then
        textPassword.Text = "*"
        un_touched_pass = True
    End If
End Function

Private Sub Form_Load()
    setConnection

    'inisialisasi variabel un_touched
    un_touched = True
    un_touched_pass = True
    
    'inisialisasi Label Notif
    labelNotif.Caption = ""
    
    'Mengatur ToolTip
    textNIK.ToolTipText = "Masukan NIK"
    textPassword.ToolTipText = "Masukan Password"
    loginButton.ToolTipText = "Masuk"
End Sub

Private Sub textNIK_GotFocus()
    'akan segera mengkosongkan jika user fokus pada textbox dan untouched bernilai true
    If un_touched = True Then
        textNIK.Text = ""
        un_touched = False
    End If
End Sub

Private Sub textNIK_LostFocus()
    clear_unInputForm 'Memanggil function di atas
End Sub

Private Sub textPassword_GotFocus()
    'akan segera mengkosongkan jika user fokus pada textbox dan untouched bernilai true
    If un_touched_pass = True Then
        textPassword.Text = ""
        un_touched_pass = False
    End If
End Sub

Private Sub textPassword_LostFocus()
    clear_unInputForm 'Memanggil function di atas
End Sub

Private Sub loginButton_Click()
    loadingBarTime_Timer 'Loading Barberjalan
    
    'Cek jika form sudah terisi dengan benar
    If textNIK = "" Or textNIK = "NIK" Then
        labelNotif.Caption = "Masukan NIK anda !"
        textNIK.SetFocus
        loadingBarInisialisasi
        Exit Sub
    End If
    
    If textPassword.Text = "" Or textPassword = "*" Then
        labelNotif.Caption = "Masukan Password Anda !"
        textPassword.SetFocus
        loadingBarInisialisasi
        Exit Sub
    End If
    On Error GoTo errHandler
    adodcLogin.Refresh 'merefresh database yang terupdate
    adodcLogin.Recordset.Find "NIK='" & textNIK.Text & "'" 'mencari berdasarkan nik yang dimaksud
    If Not adodcLogin.Recordset.EOF Then
        'cek password
        If textPassword.Text = adodcLogin.Recordset!pass Then
            labelNotif.Caption = "Login Berhasil !"
            tellToSistem
            mainMenu.Show
            Unload Me
        Else
            labelNotif.Caption = "Password Salah !"
            textNIK.Text = ""
            textPassword.Text = ""
            textNIK.SetFocus
            Exit Sub
        End If
    Else
       labelNotif.Caption = "NIK tidak Ditemukan !"
       Exit Sub
    End If
    
Exit Sub
errHandler:
    labelNotif.Caption = "Harap Masukan NIK dan Password anda dengan benar !"
    loadingBarInisialisasi
End Sub

Function tellToSistem()
    'ini akan mencatat semua informasi petugas dan memasukanya ke dalam sistem
    'ini dapat mengurangi baris kode kedepanya, dan tidak perlu untuk membuka, mencari dan mengambil database
    sistem.userNIK = adodcLogin.Recordset!NIK
    sistem.getName = adodcLogin.Recordset!nama
    sistem.getJabatan = adodcLogin.Recordset!jabatan
    
    If Not adodcLogin.Recordset!gambar = "" Then
        'Cek apakah file yang dimaksud ada dalam alamat.
        If Dir(App.Path + adodcLogin.Recordset!gambar) <> "" Then
            'jika ada maka akan merekam sesuai rekaman di database
            sistem.getPhoto = (adodcLogin.Recordset!gambar)
        Else
            'jika tidak ada, maka akan memunculkan msgBox, mengkosonginya value gambar pada databaase2003>pegawai dan menggantinya dengan default foto
            MsgBox "Kami tidak menemukan Foto anda dalam directory kami! Informasi alamat Foto akan otomatis kami kosongkan dalam database dan kami ganti dengan Foto default kami.", vbInformasi, sistem.msgTitle
            adodcLogin.Recordset!gambar = "" 'kosongkan
            adodcLogin.Recordset.Update 'update recordset
            sistem.getPhoto = ("\images\photoPegawai\defaultprofile.JPG") 'menggunakan gambar default
        End If
    Else
        'Jika petugas tidak mempunyai foto akan diisi dengan defaultprofile.jpg
        sistem.getPhoto = ("\images\photoPegawai\defaultprofile.JPG")
    End If
End Function

Private Sub loadingBarTime_Timer()
    'loading bar untuk form login
    Dim i As Integer
    For i = 1 To 10000
        loadingBar.Width = i
    Next i
End Sub

Function loadingBarInisialisasi()
    loadingBar.Width = 1 'taruh loading bar pada awal mulai
End Function
