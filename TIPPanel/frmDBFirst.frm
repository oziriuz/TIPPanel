VERSION 5.00
Begin VB.Form frmDBFirst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDBFirst"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4800
   Icon            =   "frmDBFirst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblEnterPass 
      Alignment       =   2  'Center
      Caption         =   "lblEnterPass"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblEnterIP 
      Alignment       =   2  'Center
      Caption         =   "lblEnterIP"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmDBFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Caption = frmDBdata
    Me.lblEnterIP.Caption = lblEntIP
    Me.lblEnterPass.Caption = lblEntPass
    Me.btnOK.Caption = UniOK
    Me.btnCancel.Caption = UniExit
    Me.txtIP.Text = "127.0.0.1"
    Me.txtPass.Text = ""
    Me.txtIP.MaxLength = 15
    Me.txtPass.MaxLength = 30
    frmStart.Hide
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub btnOK_Click()

    Dim cn              As ADODB.Connection
    Dim rs              As Recordset
    Dim intEmpFile      As Integer
    
    intEmpFile = FreeFile
    ConStr = "PROVIDER=PostgreSQL;" & "DATA SOURCE=" & Me.txtIP.Text & ";" & "LOCATION=" & DbaseName & ";" & "USER ID=" & DbaseUser & ";" & "PASSWORD=" & Me.txtPass.Text & ";"
    
    On Error Resume Next
    
    Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT tablename FROM pg_tables WHERE tablename ='pg_statistic';")
    rs.MoveFirst
    If rs!tablename <> "pg_statistic" Then
        MousePointer = vbDefault
        MsgBox MsgNoDBConn, vbOKOnly Or vbCritical, MsgErrBx
        frmDBFirst.Show
    Else
        Open DBSetFile For Output As #intEmpFile
        Write #intEmpFile, Me.txtIP.Text, Me.txtPass.Text
        Close #intEmpFile
        MousePointer = vbDefault
        MsgBox MsgDBConnEst, vbOKOnly Or vbInformation, UniEnter
        Call frmStart.Form_Load
        frmStart.Show
        Me.Hide
    End If
End Sub

Private Sub btnCancel_Click()

    Unload Me
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
End Sub

