VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmBackupRestore"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7125
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRestore 
      Caption         =   "btnRestore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
   End
   Begin VB.DirListBox dirRestore 
      Height          =   2565
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.FileListBox fileRestore 
      Height          =   2625
      Left            =   2640
      Pattern         =   "*.backup"
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
   Begin MSComctlLib.ProgressBar barRestore 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label lblRestore 
      Alignment       =   2  'Center
      Caption         =   "lblRestore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RestorePath  As String
Public BackFile     As String

Private Sub Form_Load()

    RestorePath = strGetDesktopPath()
    Me.dirRestore.Path = RestorePath
    Me.fileRestore.Path = Me.dirRestore.Path
    Me.barRestore.Min = 0
    Me.barRestore.Max = 250
    'бг
    Me.Caption = "Зареди База Данни"
    Me.btnRestore.Caption = "Зареди"
    Me.lblRestore.Caption = "Изберете архивен файл"
End Sub

Private Sub dirRestore_Change()

    Me.fileRestore.Path = Me.dirRestore.Path
End Sub

Private Sub btnRestore_Click()

    Dim i As Integer

    Me.barRestore.Value = 0
    
    MousePointer = vbHourglass
    
    BackFile = Me.fileRestore.FileName
    RestorePath = Me.dirRestore.Path & "\" & BackFile
    If Dir(RestorePath) <> "" Then
        Shell "c:\progra~2\postgr~1\9.2\bin\pg_restore -i -h localhost -p 5432 -d postgres -U postgres -w -F t -v -c " & RestorePath & ""
        Do While isRunning("pg_restore.exe") = True
            Me.barRestore.Value = Me.barRestore.Value + 1
            Sleep 100
        Loop
        For i = Me.barRestore.Value To Me.barRestore.Max
            Me.barRestore.Value = i
            Sleep 10
        Next i
        If isRunning("pg_restore.exe") = False Then
            MousePointer = vbDefault
            MsgBox "Данните за заредени", vbOKOnly Or vbInformation, "Заредено"
            Unload Me
        Else
            Me.barRestore.Value = 0
            MousePointer = vbDefault
            'бг
            MsgBox "Грешка при зареждане на база данни", vbOKOnly Or vbCritical, "Грешка"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    AdminPanel.Show
End Sub
