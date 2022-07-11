VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmLogin"
   ClientHeight    =   1725
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4080
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1019.187
   ScaleMode       =   0  'User
   ScaleWidth      =   3830.899
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOper 
      Height          =   345
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPass"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblOper 
      Alignment       =   1  'Right Justify
      Caption         =   "lblOper"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Login          As String
Public RootUser       As Boolean
Public AdminSuccess   As Boolean
Public LoginSucceeded As Boolean

Private Sub Form_Load()

    frmLogin.Caption = UniEnter
    lblOper.Caption = lblOperCap
    lblPass.Caption = lblPassCap
    btnOK.Caption = UniOK
    btnCancel.Caption = UniCancel
    OperName = ""
    AdminSuccess = False
    LoginSucceeded = False
    RootUser = False
End Sub

Private Sub txtOper_GotFocus()

    txtOper.SelStart = 0
    txtOper.SelLength = Len(txtOper.Text)
End Sub

Private Sub txtOper_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 43 '-
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtPass_GotFocus()

    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
End Sub

Private Sub txtPass_Change()

    If Len(txtPass.Text) <> 0 And Len(txtOper.Text) <> 0 Then
        btnOK.Enabled = True
    Else
        btnOK.Enabled = False
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 43 '-
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub btnOK_Click()

    Dim Root            As String
    Dim Admin           As String
    Dim AdminPass       As String
    Dim OperPass        As String
    Dim LogNum          As Boolean
    Dim cn              As New ADODB.Connection
    Dim rs              As New Recordset
    Dim comm            As String
    Dim counter         As Long
    Dim comIns          As String
    
    Root = "rootUseR"
    Login = txtOper
    LogNum = IsNumeric(Login)
'------------------------------Start PostgreSQL--------------------------------------
    cn.Open ConStr
    
    MousePointer = vbHourglass
    
    comm = "SELECT * FROM admin_data ORDER BY a_num ASC LIMIT 1"
    Set rs = cn.Execute(comm)
    
    Admin = rs!a_name
    AdminPass = rs!a_pass
    
    Select Case Login
        Case Root
            If txtPass = "rootpass" Then
                LoginSucceeded = True
                AdminSuccess = True
                RootUser = True
                OperName = "extreme"
                rs.Close
                Set rs = Nothing
                cn.Close 'затваряме връзката
                Set cn = Nothing
                GoTo Ready
            Else
                LoginSucceeded = False
                AdminSuccess = False
                RootUser = False
                MousePointer = vbDefault
                MsgBox MsgWrong, vbOKOnly Or vbCritical, MsgErrBx
                rs.Close
                Set rs = Nothing
                cn.Close 'затваряме връзката
                Set cn = Nothing
                GoTo Again
            End If
        Case Admin
            If txtPass = AdminPass Then
                LoginSucceeded = True
                AdminSuccess = True
                RootUser = False
                OperName = uniAdmin
                rs.Close
                Set rs = Nothing
                cn.Close 'затваряме връзката
                MousePointer = vbDefault
                Set cn = Nothing
                GoTo Ready
            Else
                LoginSucceeded = False
                AdminSuccess = False
                RootUser = False
                MousePointer = vbDefault
                MsgBox MsgWrong, vbOKOnly Or vbCritical, MsgErrBx
                rs.Close
                Set rs = Nothing
                cn.Close 'затваряме връзката
                Set cn = Nothing
                GoTo Again
            End If
        Case Else
            If LogNum = True And Login <> "." And Login <> vbBack Then
                If 0 < Login < 10 Then
                    comm = "SELECT * FROM oper_data WHERE o_num = " & Login & ""
                    Set rs = cn.Execute(comm)
                    If Not rs.BOF And Not rs.EOF Then
                        OperPass = rs!o_pass
                        OperName = rs!o_name
                    Else
                        MousePointer = vbDefault
                        MsgBox MsgWrong, vbOKOnly Or vbCritical, MsgErrBx
                        rs.Close
                        Set rs = Nothing
                        cn.Close 'затваряме връзката
                        Set cn = Nothing
                        GoTo Again
                    End If
                    rs.Close
                    Set rs = Nothing
                    cn.Close 'затваряме връзката
                    Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
                    If OperPass = Me.txtPass.Text Then
                        LoginSucceeded = True
                        AdminSuccess = False
                        RootUser = False
                        GoTo Ready
                    Else
                    End If
                Else
                End If
                GoTo ErrMsg
            Else
ErrMsg:
                RootUser = False
                LoginSucceeded = False
                AdminSuccess = False
                MousePointer = vbDefault
                MsgBox MsgWrong, vbOKOnly Or vbCritical, MsgErrBx
                Close
            End If
    End Select
Again:
    Unload Me
    frmStart.Show
Ready:
    If LoginSucceeded = True Then
'------------------------------Start PostgreSQL--------------------------------------
        cn.Open ConStr

        If MachineNumber = 1 Then comm = "SELECT * FROM entry_log ORDER BY log_num DESC LIMIT 1"
        If MachineNumber = 2 Then comm = "SELECT * FROM entry_log2 ORDER BY log_num DESC LIMIT 1"
        Set rs = cn.Execute(comm)
        If Not rs.BOF And Not rs.EOF Then
            counter = Val(rs!log_num) + 1
        Else
            counter = 1
        End If
        If MachineNumber = 1 Then comIns = "INSERT INTO entry_log (log_num, log_name, log_enter_date, log_enter) VALUES  (" & counter & ",'" & OperName & "','" & Format(Now, "DD-MM-YYYY") & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "')"
        If MachineNumber = 2 Then comIns = "INSERT INTO entry_log2 (log_num, log_name, log_enter_date, log_enter) VALUES  (" & counter & ",'" & OperName & "','" & Format(Now, "DD-MM-YYYY") & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "')"
        Set rs = cn.Execute(comIns)
        rs.Close
        Set rs = Nothing
        cn.Close 'затваряме връзката
        Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
        Unload Me
        DispPanel.Show
    Else
    End If
End Sub

Private Sub btnCancel_Click()
    
    LoginSucceeded = False
    frmStart.Show
    Unload Me
End Sub

