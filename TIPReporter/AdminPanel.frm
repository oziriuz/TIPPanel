VERSION 5.00
Begin VB.Form AdminPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AdminPanel"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   Icon            =   "AdminPanel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRestore 
      Cancel          =   -1  'True
      Caption         =   "Зареди База Данни"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton btnNewMachine 
      Caption         =   "btnNewMachine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton btnCompanyInfo 
      Caption         =   "btnCompanyInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton btnForm3 
      Caption         =   "btnForm3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton btnForm2 
      Caption         =   "btnForm2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton btnAllow 
      Caption         =   "btnAllow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton btnForm1 
      Caption         =   "btnForm1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton btnLogout 
      Caption         =   "btnLogout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   3600
      Width           =   3255
   End
End
Attribute VB_Name = "AdminPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnNewMachine_Click()
    frmDBFirst.Show
    
End Sub

Private Sub btnRestore_Click()
    frmRestore.Show
End Sub

Private Sub Form_Load()
    
    Me.Caption = frmAdPanel
    Me.btnLogout.Caption = UniExit
    Me.btnCompanyInfo.Caption = uniComInfo
    Me.btnForm1.Caption = uniForm1
    Me.btnForm2.Caption = uniForm2
    Me.btnForm3.Caption = uniForm3
    Me.btnAllow.Caption = uniSettings
    Me.btnNewMachine.Caption = "Добави връзка с машина"
    
    AddMach = True
End Sub

Private Sub btnCompanyInfo_Click()
    frmComInfo.Show
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
End Sub

Private Sub btnForm1_Click()
    setForm1.Show
    Unload setForm2
    Unload frmAllow
    Unload setForm3
    Unload frmComInfo
End Sub

Private Sub btnForm2_Click()
    setForm2.Show
    Unload frmComInfo
    Unload frmAllow
    Unload setForm1
    Unload setForm3
End Sub

Private Sub btnForm3_Click()
    setForm3.Show
    Unload frmComInfo
    Unload frmAllow
    Unload setForm1
    Unload setForm2
End Sub

Private Sub btnAllow_Click()
    frmAllow.Show
    Unload frmComInfo
    Unload setForm1
    Unload setForm2
    Unload setForm3
End Sub

Private Sub btnLogout_Click()
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmAllow
    Unload frmComInfo
    Unload Me
    Call frmStartRep.Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub
