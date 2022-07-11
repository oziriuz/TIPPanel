VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmPass"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Caption = lblPassCap
    btnOK.Caption = UniOK
    Me.txtPass.MaxLength = 20
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Public Sub btnOK_Click()

    Dim PassCheck As String

    PassCheck = GetSetting(PlacePass, PlacePassAdd, "Lab", ErrRes)

    If PassCheck = txtPass.Text Then
        If FlagButRec = 1 Then Call SvNwRecBut
        If FlagButRec = 2 Then Call DelRecBut
        Unload Me
    Else
        MsgBox MsgWrongPass, vbOKOnly Or vbCritical, MsgErrBx
        Unload Me
    End If
End Sub

