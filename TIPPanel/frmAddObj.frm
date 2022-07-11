VERSION 5.00
Begin VB.Form frmAddObj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmAddObj"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAddObj 
      Caption         =   "btnAddObj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtObjClnt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox txtKmClnt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblClntObj 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblClntObj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblKmClnt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblKmClnt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblKm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "km"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   270
   End
End
Attribute VB_Name = "frmAddObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Caption = uniObj
    Me.lblClntObj = uniObj
    Me.lblKmClnt = uniKm
    Me.txtObjClnt.MaxLength = 50
    Me.txtObjClnt.Text = ""
    Me.txtKmClnt.MaxLength = 5
    Me.txtKmClnt.Text = 0
    Me.btnAddObj.Caption = uniSave
End Sub

Private Sub txtObjClnt_GotFocus()

    txtObjClnt.SelStart = 0
    txtObjClnt.SelLength = Len(txtObjClnt.Text)
End Sub

Private Sub txtObjClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtKmClnt_GotFocus()

    txtKmClnt.SelStart = 0
    txtKmClnt.SelLength = Len(txtKmClnt.Text)
End Sub

Private Sub txtKmClnt_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub btnAddObj_Click()

    Dim itmX      As MSComctlLib.ListItem
    Dim cn        As ADODB.Connection
    Dim rs        As Recordset
    Dim rs1       As Recordset
    Dim comIns    As String
    Dim WorkCount As Long
    
    If Len(Me.txtObjClnt) > 0 Then
'------------------------------Start PostgreSQL----------------------------------
        Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
        
        MousePointer = vbHourglass
        
        WorkCount = 1
        
        Set rs = cn.Execute("SELECT w_num FROM worksites ORDER BY w_num DESC;")
        If Not rs.BOF And Not rs.EOF Then
            WorkCount = Val(rs!w_num) + 1
        End If
        Set rs1 = cn.Execute("SELECT c_num FROM clients WHERE c_num = " & Val(DispPanel.txtClnt) & ";")
        If rs1.EOF Or rs1.BOF Then
            MsgBox MsgNoClnt, vbOKOnly Or vbCritical, MsgErrBx
            rs.Close
            Set rs = Nothing
            cn.Close
            Set cn = Nothing
            Unload Me
            GoTo EndSub
        End If
        Set rs = cn.Execute("SELECT w_name FROM worksites WHERE w_cnum = '" & Val(DispPanel.txtClnt.Text) & "' AND w_name = '" & Me.txtObjClnt.Text & "'") 'клиент
        If Not rs.BOF And Not rs.EOF Then
            MousePointer = vbDefault
            MsgBox MsgNewName, vbOKOnly Or vbCritical, MsgErrBx
            rs.Close
            Set rs = Nothing
            cn.Close
            Set cn = Nothing
            GoTo EndSub
        End If
            
        'ако няма съвпадения правим запис
        comIns = "INSERT INTO worksites VALUES(" & WorkCount & ",'" & Val(DispPanel.txtClnt) & "','" & Me.txtObjClnt & "','" & Val(Me.txtKmClnt) & "', 'true')"
        Set rs = cn.Execute(comIns)
                
        'затваряме базата данни и прекратяваме функцията
        rs.Close
        Set rs = Nothing
        cn.Close
        Set cn = Nothing
'------------------------------Start PostgreSQL----------------------------------

        MousePointer = vbDefault
        
        MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
        
        Set itmX = DispPanel.lstObj.ListItems.Add(1, , Me.txtObjClnt)
            itmX.SubItems(1) = Val(Me.txtKmClnt)
        Unload Me
    Else
        MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
    End If
EndSub:
End Sub

