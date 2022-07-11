VERSION 5.00
Begin VB.Form frmMixCap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ ���������"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMixCap 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblMixCap 
      BackStyle       =   0  'Transparent
      Caption         =   "��������� �� �������:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   3105
   End
End
Attribute VB_Name = "frmMixCap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PointLook33 As Boolean

Private Sub btnCancel_Click()

    Unload Me
End Sub

Private Sub btnOK_Click()

    Dim SyncItemValuesConfig(1 To 1)    As Variant
    Dim SyncItemSrvHandlesConfig(1)     As Long
    Dim SyncItemSrvErrConfig()          As Long

    For i = 0 To DispPanel.ItemCountConfig - 1
        SyncItemSrvHandlesConfig(i + 1) = handyConfig(i + 1)
        SyncItemValuesConfig(i + 1) = CSng(Me.txtMixCap.Text)
    Next i
    DispPanel.ConGroupConfig.SyncWrite DispPanel.ItemCountConfig, SyncItemSrvHandlesConfig, _
    SyncItemValuesConfig, SyncItemSrvErrConfig
    MixCap = CSng(rDs(frmOPC.Config(0)))
    Unload Me
End Sub

Private Sub Form_Load()
    Me.btnOK.Caption = UniOK
    Me.btnCancel.Caption = UniCancel
    Me.txtMixCap.MaxLength = 4
    
    Me.txtMixCap.Text = CSng(rDs(frmOPC.Config(0).Text))
End Sub

Private Sub txtMixCap_GotFocus()
    txtMixCap.SelStart = 0
    txtMixCap.SelLength = Len(txtMixCap.Text)

    If InStr(txtMixCap.Text, DecSep) <> 0 Then
        PointLook33 = True
    Else
        PointLook33 = False
    End If

End Sub

Private Sub txtMixCap_Change()

    If InStr(txtMixCap.Text, DecSep) <> 0 Then
        PointLook33 = True
    Else
        PointLook33 = False
    End If

End Sub

Private Sub txtMixCap_KeyPress(KeyAscii As Integer)

    If InStr(txtMixCap.Text, DecSep) <> 0 Then
        PointLook33 = True
    Else
        PointLook33 = False
    End If

    If txtMixCap.SelLength = Len(txtMixCap.Text) Then
        PointLook33 = False
    Else
    End If

    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If

    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook33 = True Then
        KeyAscii = 0
    Else

        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook33 = True
        Else
        End If
    End If

End Sub

