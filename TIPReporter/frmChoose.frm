VERSION 5.00
Begin VB.Form frmChoose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmChoose"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox cmbChoose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      Caption         =   "lblChoose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
    End
End Sub

Private Sub btnOK_Click()
    
    intEmpFileNbr1 = FreeFile
    
    Open DBSetFile For Input As #intEmpFileNbr1
    Do Until EOF(intEmpFileNbr1)
        Input #intEmpFileNbr1, IPConnStr, MachName
        If MachName = Me.cmbChoose.Text Then Exit Do
    Loop
    Close #intEmpFileNbr1
    
    If Me.cmbChoose.ListIndex <> -1 Then
        Choice = True
    Else
        MsgBox "Изберете машина за връзка!", vbCritical Or vbOKOnly, MsgErrBx
        GoTo EndSub
    End If
    
    Call frmStartRep.Form_Load
    frmStartRep.Show
    Me.Hide
EndSub:
End Sub

Private Sub Form_Load()
    Me.Caption = "Избор на машина"
    Me.lblChoose.Caption = "Изберете машина за връзка:"
    Me.btnOK.Caption = UniOK
    Me.btnCancel.Caption = UniExit
    
    Choice = False
    
    intEmpFileNbr1 = FreeFile
    
    Open DBSetFile For Input As #intEmpFileNbr1
    Do Until EOF(intEmpFileNbr1)
        Input #intEmpFileNbr1, IPConnStr, MachName
        Me.cmbChoose.AddItem MachName
    Loop
    Close #intEmpFileNbr1

End Sub
