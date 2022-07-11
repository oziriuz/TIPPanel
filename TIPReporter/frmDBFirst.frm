VERSION 5.00
Begin VB.Form frmDBFirst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDBFirst"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMachine 
      Height          =   375
      Left            =   1200
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
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblMachine 
      Alignment       =   2  'Center
      Caption         =   "lblMachine"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4575
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

Private Sub Form_Load()
    Me.Caption = frmDBdata
    Me.lblEnterIP.Caption = lblEntIP
    Me.lblMachine.Caption = "Машина Име:"
    Me.btnOK.Caption = UniOK
    Me.btnCancel.Caption = UniExit
    
    Me.txtIP.Text = "127.0.0.1"
    Me.txtIP.MaxLength = 15
    frmStartRep.Hide
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
        If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If
End Sub

Private Sub btnOK_Click()
    Dim intEmpFileNbr1 As Integer
    
    intEmpFileNbr1 = FreeFile
    
    ConStr = "PROVIDER=PostgreSQL;" _
            & "DATA SOURCE=" & Me.txtIP.Text & ";" _
            & "LOCATION=" & DbaseName & ";" _
            & "USER ID=" & DbaseUser & ";" _
            & "PASSWORD=" & PassConnStr & ";"
    
    On Error Resume Next
    
    MachName = Me.txtMachine.Text
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 30
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT * FROM pg_tables WHERE tablename ='pg_statistic';")
    rs.MoveFirst
    If rs!tablename <> "pg_statistic" Then
        MousePointer = vbDefault
        MsgBox MsgNoDBConn, vbOKOnly Or vbCritical, MsgErrBx
        frmDBFirst.Show
    Else
        Open DBSetFile For Append As #intEmpFileNbr1
        Write #intEmpFileNbr1, Me.txtIP.Text, Me.txtMachine
        Close #intEmpFileNbr1
        MousePointer = vbDefault
        MsgBox MsgDBConnEst, vbOKOnly Or vbInformation, UniEnter
        Call frmStartRep.Form_Load
        frmStartRep.Show
        Me.Hide
    End If
    MousePointer = vbDefault
End Sub

Private Sub btnCancel_Click()
    Unload Me
    If AddMach = False Then
        End
    End If
        
End Sub

Private Sub txtMachine_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+
        
            'пропуска кода
        Case 46 '.
        
            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub
