VERSION 5.00
Begin VB.Form loadingBarForMainMenu 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "loading"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   -1125
   ClientWidth     =   20250
   Icon            =   "loadingBarMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   90
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   Begin VB.Timer LoadingBarMainMenuTime 
      Left            =   3120
      Top             =   0
   End
   Begin VB.Shape loadingBarMainMenu 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "loadingBarForMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    'Pojok Kiri Atas
    
    loadingBarInisialisasi
End Sub

Public Sub LoadingBarMainMenuTime_Timer()
    Dim i As Integer 'indek pengulangan
    For i = 1 To 25000
        loadingBarMainMenu.Width = i
    Next i
End Sub

Function loadingBarInisialisasi()
    loadingBarMainMenu.Width = 1 'taruh loading bar ke posisi awal
End Function
