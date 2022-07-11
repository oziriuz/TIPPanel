VERSION 5.00
Begin VB.Form frmEdAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmEdAdmin"
   ClientHeight    =   3540
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5055
   Icon            =   "frmEdAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2091.548
   ScaleMode       =   0  'User
   ScaleWidth      =   4746.371
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassNew 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1815
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtAdminNew 
      Height          =   375
      Left            =   1815
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtPassConf 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1815
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtAdminOld 
      Height          =   375
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtPassOld 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
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
      Left            =   600
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblPassNew 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPassNew"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblAdminNew 
      Alignment       =   1  'Right Justify
      Caption         =   "lblAdminNew"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblPassOld 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPassOld"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblAdminOld 
      Alignment       =   1  'Right Justify
      Caption         =   "lblAdminOld"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmEdAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    frmEdAdmin.Caption = btnEditAd
    lblAdminOld.Caption = uniNm
    lblPassOld.Caption = lblPassCap
    lblAdminNew.Caption = uniNewNm
    lblPassNew.Caption = lblPassCap
    lblPassConf.Caption = lblPassConfCap
    btnOK.Caption = UniOK
    btnCancel.Caption = UniCancel
    Me.txtAdminOld.MaxLength = 15
    Me.txtPassOld.MaxLength = 15
    Me.txtAdminNew.MaxLength = 15
    Me.txtPassNew.MaxLength = 15
    Me.txtPassConf.MaxLength = 15
End Sub

Private Sub txtAdminOld_GotFocus()

    txtAdminOld.SelStart = 0
    txtAdminOld.SelLength = Len(txtAdminOld.Text)
End Sub

Private Sub txtAdminOld_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 8 'латиница A-Z, и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtPassOld_GotFocus()

    txtPassOld.SelStart = 0
    txtPassOld.SelLength = Len(txtPassOld.Text)
End Sub

Private Sub txtPassOld_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtAdminNew_GotFocus()

    txtAdminNew.SelStart = 0
    txtAdminNew.SelLength = Len(txtAdminNew.Text)
End Sub

Private Sub txtAdminNew_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 8 'латиница A-Z, и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtPassNew_GotFocus()

    txtPassNew.SelStart = 0
    txtPassNew.SelLength = Len(txtPassNew.Text)
End Sub

Private Sub txtPassNew_KeyPress(KeyAscii As Integer)

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

Private Sub txtPassConf_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtPassConf_Change()

    If Len(txtPassConf.Text) <> 0 And Len(txtPassNew.Text) <> 0 And Len(txtAdminNew.Text) <> 0 And Len(txtPassOld.Text) <> 0 And Len(txtAdminOld.Text) <> 0 Then
        btnOK.Enabled = True
    Else
        btnOK.Enabled = False
    End If
End Sub

Private Sub btnOK_Click()
    
    Dim cn      As New ADODB.Connection
    Dim rs      As Recordset
    Dim comm    As String
    Dim comEdit As String
    
'------------------------------Start PostgreSQL--------------------------------------
    cn.Open ConStr

    comm = "SELECT * FROM admin_data LIMIT 1"
    comEdit = "UPDATE admin_data SET a_name = '" & Me.txtAdminNew.Text & "', a_pass ='" & Me.txtPassNew.Text & "' WHERE a_num = 1"
    
    Set rs = cn.Execute(comm)
    If Me.txtAdminOld.Text <> rs!a_name Or Me.txtPassOld.Text <> rs!a_pass Then
        MousePointer = vbDefault
        MsgBox MsgWrong, vbOKOnly Or vbCritical, MsgErrBx
        rs.Close
        Set rs = Nothing
        cn.Close
        Set cn = Nothing
        GoTo FlagLog
    Else
        If Me.txtPassNew.Text <> Me.txtPassConf.Text Then
            MousePointer = vbDefault
            MsgBox MsgPassNotConf, vbOKOnly Or vbCritical, MsgErrBx
            rs.Close
            Set rs = Nothing
            cn.Close
            Set cn = Nothing
            GoTo FlagLog
        ElseIf Me.txtAdminOld.Text = rs!a_name And Me.txtPassOld.Text = rs!a_pass And Me.txtPassNew.Text = Me.txtPassConf.Text Then
            Set rs = cn.Execute(comEdit)
            MousePointer = vbDefault
            MsgBox MsgAdminSuccess & vbNewLine & Me.txtAdminNew.Text, vbOKOnly Or vbInformation
        End If
    End If
                
    rs.Close
    Set rs = Nothing
    cn.Close 'затваряме връзката
    Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
FlagLog:
    Unload Me
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

