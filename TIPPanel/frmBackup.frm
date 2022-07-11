VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmBackup"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar barBackup 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.CommandButton btnBackup 
      Caption         =   "btnBackup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.DirListBox dirRestore 
      Height          =   2115
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblBackup 
      Alignment       =   2  'Center
      Caption         =   "lblBackup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SavePath As String
Public BackFile As String

Private Sub Form_Load()

    Me.dirRestore.Path = BackPath
    Me.barBackup.Min = 0
    Me.barBackup.Max = 250
    Me.Caption = MsgArhBx
    Me.btnBackup.Caption = MsgArhBx
    Me.lblBackup.Caption = lblChooseFolder
End Sub

Private Sub btnBackup_Click()

    Dim i As Integer

    BackFile = "\TIPbkp" & Format(Now, "YYYYMMDD-HHMMSS") & ".backup"
    SavePath = Me.dirRestore.Path & BackFile
    Shell "c:\progra~1\postgr~1\9.2\bin\pg_dump.exe -i -h localhost -p 5432 -U postgres -w -E sql_ascii -F t -b -v -f " & SavePath & " postgres"
    
    MousePointer = vbHourglass
    
    Me.barBackup.Value = 0
    
    Do While isRunning("pg_dump.exe") = True
        Sleep 100
        Me.barBackup.Value = Me.barBackup.Value + 1
    Loop
    If isRunning("pg_dump.exe") = False Then
        If Dir(SavePath) <> "" Then
            For i = Me.barBackup.Value To Me.barBackup.Max
                Me.barBackup.Value = i
                Sleep 10
            Next i
            MousePointer = vbDefault
            MsgBox MsgArhivReady, vbOKOnly Or vbInformation, MsgArhBx
            End
        Else
            Me.barBackup.Value = 0
            MousePointer = vbDefault
            MsgBox MsgArhivError, vbOKOnly Or vbCritical, MsgErrBx
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
End Sub
