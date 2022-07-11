VERSION 5.00
Begin VB.Form frmEdOper 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmEdOper"
   ClientHeight    =   3660
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4575
   Icon            =   "frmEdOper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2162.448
   ScaleMode       =   0  'User
   ScaleWidth      =   4295.676
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnDelete 
      Caption         =   "btnDelete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtFamily 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.ComboBox cmbOper 
      Height          =   315
      ItemData        =   "frmEdOper.frx":08CA
      Left            =   1320
      List            =   "frmEdOper.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtPassConf 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblPassConf 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPassConf"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPass"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblFam 
      Alignment       =   1  'Right Justify
      Caption         =   "lblFam"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "lblName"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblOper 
      Alignment       =   1  'Right Justify
      Caption         =   "lblOper"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmEdOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbOper_Click()

    Dim cn   As New ADODB.Connection
    Dim rs   As Recordset
    Dim comm As String
    
    If Me.cmbOper.ListIndex > -1 Then
'------------------------------Start PostgreSQL--------------------------------------
        cn.Open ConStr

        comm = "SELECT * FROM oper_data WHERE o_num = " & Val(Me.cmbOper.Text) & ""
        Set rs = cn.Execute(comm)
        If Not rs.BOF And Not rs.EOF Then
            Me.txtName = Left$(rs!o_name, InStr(rs!o_name, " ") - 1)
            Me.txtFamily = Mid$(rs!o_name, InStr(rs!o_name, " ") + 1, Len(rs!o_name))
        End If
        rs.Close
        Set rs = Nothing
        cn.Close 'затваряме връзката
        Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
    End If
End Sub

Private Sub Form_Load()

    Dim cn   As New ADODB.Connection
    Dim rs   As Recordset
    Dim comm As String

    frmEdOper.Caption = btnEditOp
    lblOper.Caption = uniCode
    lblName.Caption = uniNm
    lblFam.Caption = uniFam
    lblPass.Caption = lblPassCap
    lblPassConf.Caption = lblPassConfCap
    btnDelete.Caption = uniDel
    btnOK.Caption = UniOK
    btnCancel.Caption = UniCancel
    Me.txtName.MaxLength = 15
    Me.txtFamily.MaxLength = 15
    Me.txtPass.MaxLength = 15
    Me.txtPassConf.MaxLength = 15
    
'------------------------------Start PostgreSQL--------------------------------------
    cn.Open ConStr
    
    MousePointer = vbHourglass
    
    comm = "SELECT o_num FROM oper_data ORDER BY o_num ASC"
    Set rs = cn.Execute(comm)
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            Me.cmbOper.AddItem rs!o_num
            rs.MoveNext
        Loop
    Else
        MousePointer = vbDefault
        MsgBox MsgNoOps, vbOKOnly Or vbCritical, MsgErrBx
        rs.Close
        Set rs = Nothing
        cn.Close 'затваряме връзката
        Set cn = Nothing
        GoTo EndSub
    End If
    rs.Close
    Set rs = Nothing
    cn.Close 'затваряме връзката
    MousePointer = vbDefault
    Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
EndSub:
End Sub

Private Sub txtName_GotFocus()

    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtName_Change()

    If Len(cmbOper.Text) <> 0 Then
        btnDelete.Enabled = True
    Else
        btnDelete.Enabled = False
    End If
End Sub

Private Sub txtFamily_GotFocus()

    txtFamily.SelStart = 0
    txtFamily.SelLength = Len(txtFamily.Text)
End Sub

Private Sub txtFamily_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
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

    If Len(txtPassConf.Text) <> 0 And Len(txtPass.Text) <> 0 Then
        btnOK.Enabled = True
    Else
        btnOK.Enabled = False
    End If
End Sub

Private Sub btnDelete_Click()

    Dim cn   As New ADODB.Connection
    Dim rs   As Recordset
    Dim comm As String
    
    If Me.cmbOper.ListIndex > -1 Then
'------------------------------Start PostgreSQL--------------------------------------
        cn.Open ConStr
        
        MousePointer = vbHourglass
        
        comm = "DELETE FROM oper_data WHERE o_num = " & Val(Me.cmbOper.Text) & ""
        Set rs = cn.Execute(comm)
        rs.Close
        Set rs = Nothing
        cn.Close 'затваряме връзката
        MousePointer = vbDefault
        Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
    Else
        MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        GoTo EndSub
    End If
EndSub:
    Unload Me
End Sub

Private Sub btnOK_Click()

    Dim cn      As New ADODB.Connection
    Dim rs      As Recordset
    Dim comEdit As String
    
    If Len(Me.txtName.Text) > 0 And Len(Me.txtFamily.Text) > 0 And Len(Me.txtPass.Text) > 0 Then
        If Me.txtPass.Text = Me.txtPassConf.Text Then
'------------------------------Start PostgreSQL--------------------------------------
            cn.Open ConStr
            
            MousePointer = vbHourglass
            
            comEdit = "UPDATE oper_data SET o_name = '" & Me.txtName.Text & " " & Me.txtFamily.Text & "', o_pass ='" & Me.txtPass.Text & "' WHERE o_num = " & Val(Me.cmbOper.Text) & ""
            Set rs = cn.Execute(comEdit)
            rs.Close
            Set rs = Nothing
            cn.Close 'затваряме връзката
            MousePointer = vbDefault
            Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
            MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
            GoTo FlagEnd
        Else
            MsgBox MsgPassNotConf, vbOKOnly Or vbCritical, MsgErrBx
            txtPass.SetFocus
            txtPass.SelStart = 0
            txtPass.SelLength = Len(txtPass.Text)
            GoTo FlagEnd
        End If
    End If
FlagEnd:
    Unload Me
End Sub

Private Sub btnCancel_Click()
    
    Unload Me
End Sub

