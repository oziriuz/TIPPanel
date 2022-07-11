VERSION 5.00
Begin VB.Form frmLabPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmLabPass"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   Icon            =   "frmLabPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "btnSave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtLabPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblLabPass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblLabPass"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmLabPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Caption = frmLabPassCap
    Me.lblLabPass.Caption = lblLabPassCap
    Me.btnSave.Caption = uniSave
    Me.txtLabPass.MaxLength = 15
End Sub

Private Sub txtLabPass_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub btnSave_Click()

    SaveSetting PlacePass, PlacePassAdd, "Lab", Me.txtLabPass
    Unload Me
End Sub

