VERSION 5.00
Begin VB.Form frmNwOper 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmNwOper"
   ClientHeight    =   3450
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3870
   Icon            =   "frmNwOper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2038.373
   ScaleMode       =   0  'User
   ScaleWidth      =   3633.72
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbOper 
      Height          =   315
      ItemData        =   "frmNwOper.frx":08CA
      Left            =   1320
      List            =   "frmNwOper.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFamily 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtPassConf 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblPassConf 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPassConf"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      Caption         =   "lblPass"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblFam 
      Alignment       =   1  'Right Justify
      Caption         =   "lblFam"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "lblName"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblOper 
      Alignment       =   1  'Right Justify
      Caption         =   "lblOper"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmNwOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'екран за въвеждане на нов оператор

    Dim i       As Integer
    Dim cn      As New ADODB.Connection
    Dim rs      As Recordset
    Dim comm    As String
    
    frmNwOper.Caption = btnCreateOp
    lblOper.Caption = uniCode
    lblName.Caption = uniNm
    lblFam.Caption = uniFam
    lblPass.Caption = lblPassCap
    lblPassConf.Caption = lblPassConfCap
    btnOK.Caption = UniOK
    btnCancel.Caption = UniCancel
    Me.txtName.MaxLength = 15
    Me.txtFamily.MaxLength = 15
    Me.txtPass.MaxLength = 15
    Me.txtPassConf.MaxLength = 15

    i = 1 'брояч
        
    cmbOper.Clear 'почистваме комбото
'------------------------------Start PostgreSQL--------------------------------------
    cn.Open ConStr

    comm = "SELECT * FROM oper_data ORDER BY o_num ASC"
    Set rs = cn.Execute(comm)
    If Not rs.BOF And Not rs.EOF Then
        For i = 1 To 9
            rs.MoveFirst
            Do While Not rs.EOF
                If i <> rs!o_num Then
                    rs.MoveNext
                    If rs.EOF Then
                        Me.cmbOper.AddItem i
                    End If
                Else
                    rs.MoveNext
                    Exit Do
                End If
            Loop
        Next i
    Else
        For i = 1 To 9
            Me.cmbOper.AddItem i
        Next i
    End If
    rs.Close
    Set rs = Nothing
    cn.Close 'затваряме връзката
    Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
End Sub

Private Sub cmbOper_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
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

Private Sub txtPass_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
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

'разрешава бутона ОК след като се попълнят всички полета
Private Sub txtPassConf_Change()

    If Len(txtPassConf.Text) <> 0 And Len(txtPass.Text) <> 0 And Len(txtName.Text) <> 0 And Len(txtFamily.Text) <> 0 And Len(cmbOper.Text) <> 0 Then
        btnOK.Enabled = True
    Else
        btnOK.Enabled = False
    End If
End Sub

Private Sub btnOK_Click()

    Dim OperName    As String
    Dim cn          As New ADODB.Connection
    Dim rs          As Recordset
    Dim comm        As String
    Dim comIns      As String
    
    OperName = Me.txtName.Text & " " & Me.txtFamily.Text

    '------------------------------Start PostgreSQL--------------------------------------
    cn.Open ConStr
    
    MousePointer = vbHourglass
    
    comm = "SELECT * FROM oper_data ORDER BY o_num ASC"
    comIns = "INSERT INTO oper_data VALUES('" & Me.cmbOper.Text & "','" & OperName & "','" & Me.txtPass.Text & "')"
    Set rs = cn.Execute(comm)
    If Not rs.BOF And Not rs.EOF Then rs.MoveFirst
    Do While Not rs.EOF
        If OperName = rs!o_name Then 'проверка за съвпадение на имена
            MousePointer = vbDefault
            MsgBox MsgSameNmFamOp, vbOKOnly Or vbCritical, MsgErrBx
            txtName.SetFocus
            txtName.SelStart = 0
            txtName.SelLength = Len(txtName.Text)
            rs.Close
            Set rs = Nothing
            cn.Close 'затваряме връзката
            Set cn = Nothing
            GoTo FlagEnd 'отиваме на лога
        Else
        End If
        rs.MoveNext
    Loop
    If Me.txtPass.Text = Me.txtPassConf.Text Then 'проверка за потвърждение на паролата
        Set rs = cn.Execute(comIns)
        MousePointer = vbDefault
        MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave 'съобщение за успешен запис
    Else
        MousePointer = vbDefault
        MsgBox MsgPassNotConf, vbOKOnly Or vbCritical, MsgErrBx
        txtPass.SetFocus
        txtPass.SelStart = 0
        txtPass.SelLength = Len(txtPass.Text)
        rs.Close
        Set rs = Nothing
        cn.Close 'затваряме връзката
        Set cn = Nothing
        GoTo FlagEnd 'отиваме на лога
    End If
    rs.Close
    Set rs = Nothing
    cn.Close 'затваряме връзката
    MousePointer = vbDefault
    Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
    Unload Me
FlagEnd:
End Sub

Private Sub btnCancel_Click()
    
    Unload Me 'изход от прозореца
End Sub

