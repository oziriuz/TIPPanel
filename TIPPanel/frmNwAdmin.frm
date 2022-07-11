VERSION 5.00
Begin VB.Form frmNwAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmNwAdmin"
   ClientHeight    =   2445
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4095
   Icon            =   "frmNwAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1444.587
   ScaleMode       =   0  'User
   ScaleWidth      =   3844.982
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassConf 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtAdmin 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblPassConf 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPassConf"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPass"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblAdmin 
      Alignment       =   1  'Right Justify
      Caption         =   "lblAdmin"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmNwAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    frmNwAdmin.Caption = frmNewAd
    lblAdmin.Caption = uniNm
    lblPass.Caption = lblPassCap
    lblPassConf.Caption = lblPassConfCap
    btnOK.Caption = UniOK
    btnCancel.Caption = UniCancel
End Sub

Private Sub txtAdmin_GotFocus()

    txtAdmin.SelStart = 0
    txtAdmin.SelLength = Len(txtAdmin.Text)
End Sub

Private Sub txtAdmin_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 8 'латиница A-Z, и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtPass_GotFocus()

    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
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

Private Sub txtPassConf_GotFocus()

    txtPassConf.SelStart = 0
    txtPassConf.SelLength = Len(txtPassConf.Text)
End Sub

Private Sub txtPassConf_Change()

    If Len(txtPassConf.Text) <> 0 And Len(txtPass.Text) <> 0 And Len(txtAdmin.Text) <> 0 Then
        btnOK.Enabled = True
    Else
        btnOK.Enabled = False
    End If
End Sub

Private Sub txtPassConf_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub btnOK_Click()

    Const MaxChars = 15
    
    Dim Admin           As String
    Dim AdminPass       As String
    Dim AdminPassConf   As String
    Dim cn              As New ADODB.Connection
    Dim rs              As New Recordset
    Dim comIns          As String

    Admin = txtAdmin
    AdminPass = txtPass
    AdminPassConf = txtPassConf

    If Len(Admin) > MaxChars Or Len(AdminPass) > MaxChars Then
        MsgBox MsgMaxL15, vbOKOnly Or vbCritical, MsgErrBx
        GoTo EndFlag
    Else
    End If
    If AdminPass <> AdminPassConf Then
        MsgBox MsgPassNotConf, vbOKOnly Or vbCritical, MsgErrBx
    Else
'------------------------------Start PostgreSQL--------------------------------------
        cn.Open ConStr
        
        MousePointer = vbHourglass
        
        comIns = "INSERT INTO admin_data VALUES(" & 1 & ",'" & Admin & "','" & AdminPass & "')"
        Set rs = cn.Execute(comIns)
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
        MsgBox MsgAdminSuccess & vbNewLine & Admin, vbOKOnly Or vbInformation
    End If
EndFlag:
    frmStart.Show
    Unload Me
End Sub

Private Sub btnCancel_Click()
    
    frmStart.Show
    Unload Me
End Sub

