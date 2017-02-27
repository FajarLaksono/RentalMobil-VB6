VERSION 5.00
Begin VB.Form about 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4080
   ClientLeft      =   8310
   ClientTop       =   4170
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2816.089
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox listBoxCreditsList 
      Height          =   1815
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      Picture         =   "about.frx":038A
      TabIndex        =   0
      Top             =   3585
      Width           =   1260
   End
   Begin VB.Image picIcon 
      Height          =   855
      Left            =   120
      Picture         =   "about.frx":09D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   2401.958
      Y2              =   2401.958
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Creator :"
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rental Mobil Purwokerto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   2112.067
      Y2              =   2112.067
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.5"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Top             =   720
      Width           =   3885
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me 'keluar
End Sub

Private Sub Form_Load()
    'Daftar Pembuat
    listBoxCreditsList.AddItem "(12155313) Ali Nur Rachman"
    listBoxCreditsList.AddItem "(12155462) Deden Triana"
    listBoxCreditsList.AddItem "(12151247) Fajar Aziz Laksono"
    listBoxCreditsList.AddItem "(12155291) Feti Adi Saputro"
    listBoxCreditsList.AddItem "(12150235) Mochammad Badruttamam"
    listBoxCreditsList.AddItem "(12150292) Rena Averonal"
    listBoxCreditsList.AddItem "(12155149) Syam Raka Febrian"
    listBoxCreditsList.AddItem "(12152257) Viky Feri Andre"
    
    'Mengatur Tool Tip
    picIcon.ToolTipText = "Icon Rental Mobil Purwokerto"
    lblTitle.ToolTipText = "Rental Mobil Purwokerto"
    lblVersion.ToolTipText = "Versi Software"
    listBoxCreditsList.ToolTipText = "Daftar pembuat software Rental Mobil Purwokerto"
End Sub
