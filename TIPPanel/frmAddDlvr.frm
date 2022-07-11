VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddDlvr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmAddDlvr"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6570
   Icon            =   "frmAddDlvr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSearch 
      Caption         =   "btnSearch"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton btnSvDlvr 
      Caption         =   "btnSvDlvr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtDlvrQuant 
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
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtDlvrDocNo 
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
      TabIndex        =   4
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtDlvrDocType 
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
      TabIndex        =   3
      Top             =   2280
      Width           =   3495
   End
   Begin VB.ComboBox cmbDlvrMat 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3840
      Width           =   4095
   End
   Begin VB.TextBox txtDlvr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cmbDlvrSupName 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox txtDlvrSupBG 
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
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtDlvrSup 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker queDlvrDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   65535
      CalendarForeColor=   -2147483639
      CustomFormat    =   "dd.MM.yyy"
      Format          =   115015683
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin VB.Label lblKg 
      BackStyle       =   0  'Transparent
      Caption         =   "kg"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label lblDlvrDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrDate"
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
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblDlvrQuant 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrQuant"
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
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lblDlvrDocNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrDocNo"
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
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblDlvrDocType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrDocType"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lblDlvrMat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrMat"
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
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblDlvr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvr"
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
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblDlvrSupBG 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrSupBG"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblDlvrSupName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrSupName"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblDlvrSup 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDlvrSup"
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
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmAddDlvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dlvr        As Long
Dim DlvrMat     As String
Dim DlvrSupName As String
Dim DlvrSupBG   As String
Dim DlvrDocType As String
Dim DlvrDocNo   As String
Dim DlvrDate    As String
Dim StampDate   As String
Dim DlvrQuant   As Single
Dim PointLook5  As Boolean

Private Sub Form_Load()

    Dim cn      As ADODB.Connection
    Dim rs      As Recordset

    Me.Caption = uniDlvr & " " & uniMats
    lblDlvr.Caption = uniCode
    lblDlvrMat.Caption = uniMat
    lblDlvrSup.Caption = uniSup
    lblDlvrSupName.Caption = uniFirm
    lblDlvrSupBG.Caption = uniBG
    lblDlvrDocType.Caption = uniTypeDoc
    lblDlvrDocNo.Caption = uniNoDoc
    lblDlvrDate.Caption = uniDateDlvr
    lblDlvrQuant.Caption = uniQuant
    btnSvDlvr.Caption = uniSave
    Me.btnSearch.Caption = uniSrchBG
    Me.txtDlvrDocNo.MaxLength = 20
    Me.txtDlvrDocType.MaxLength = 20
    Me.txtDlvrSupBG.MaxLength = 15
    Me.txtDlvrQuant.MaxLength = 15
    
    Me.queDlvrDate = Now
    
    DecSep = GetDecimalSep()
    
    If InStr(txtDlvrQuant.Text, DecSep) = 0 Then PointLook5 = False
'------------------------------Start PostgreSQL----------------------------------
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    'зареждане на последен номер на доставка
    Set rs = cn.Execute("SELECT del_num FROM deliveries ORDER BY del_num DESC LIMIT 1;")
    If Not rs.EOF And Not rs.BOF Then
        txtDlvr.Text = Format(rs!del_num + 1, "0000000")
    Else
        txtDlvr.Text = "0000001"
    End If
    
    'зареждане на комбо с имена на материали без водата
    Set rs = cn.Execute("SELECT m_name, m_type FROM materials_bc" & MachineNumber & " ORDER BY m_name ASC;")
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    Do While Not rs.EOF
        If Val(rs!m_type) <> 2 Then frmAddDlvr.cmbDlvrMat.AddItem rs!m_name
        rs.MoveNext
    Loop
    
    'зареждане на комбо с имена на доставчици
    Set rs = cn.Execute("SELECT * FROM suppliers WHERE s_show = '1' ORDER BY s_name ASC;")
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    Do While Not rs.EOF
        frmAddDlvr.cmbDlvrSupName.AddItem rs!s_name
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    MousePointer = vbDefault
    
    If DispPanel.lstMat.ListItems.count > 0 And DispPanel.txtMat <> "" And DispPanel.cmbMatType.Text <> uniWat Then
        frmAddDlvr.cmbDlvrMat.Text = DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(1).Text
    Else
    End If
End Sub

Private Sub cmbDlvrSupName_Click()

    Call LoadDlvrSup
End Sub

Private Sub txtDlvrSupBG_GotFocus()

    txtDlvrSupBG.SelStart = 0
    txtDlvrSupBG.SelLength = Len(txtDlvrSupBG.Text)
End Sub

Private Sub txtDlvrSupBG_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtDlvrDocType_GotFocus()

    txtDlvrDocType.SelStart = 0
    txtDlvrDocType.SelLength = Len(txtDlvrDocType.Text)
End Sub

Private Sub txtDlvrDocType_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица ј-я
        Case 224 To 255 'кирилица а-€
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtDlvrDocNo_GotFocus()

    txtDlvrDocNo.SelStart = 0
    txtDlvrDocNo.SelLength = Len(txtDlvrDocNo.Text)
End Sub

Private Sub txtDlvrDocNo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица ј-я
        Case 224 To 255 'кирилица а-€
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtDlvrQuant_GotFocus()

    txtDlvrQuant.SelStart = 0
    txtDlvrQuant.SelLength = Len(txtDlvrQuant.Text)

    If InStr(txtDlvrQuant.Text, DecSep) <> 0 Then
        PointLook5 = True
    Else
        PointLook5 = False
    End If
End Sub

Private Sub txtDlvrQuant_Change()

    If InStr(txtDlvrQuant.Text, DecSep) <> 0 Then
        PointLook5 = True
    Else
        PointLook5 = False
    End If
End Sub

Private Sub txtDlvrQuant_KeyPress(KeyAscii As Integer)

    If InStr(txtDlvrQuant.Text, DecSep) <> 0 Then
        PointLook5 = True
    Else
        PointLook5 = False
    End If
    If txtDlvrQuant.SelLength = Len(txtDlvrQuant.Text) Then
        PointLook5 = False
    Else
    End If
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If
    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook5 = True Then
        KeyAscii = 0
    Else
        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook5 = True
        Else
        End If
    End If
End Sub

Private Sub btnSearch_Click()

    Call LoadDlvrSupBG
End Sub

Private Sub btnSvDlvr_Click()

    Dim cn              As ADODB.Connection
    Dim rs              As Recordset
    Dim comIns          As String
    Dim MachineOther    As Integer

    If Len(txtDlvr.Text) > 0 And Len(cmbDlvrMat.Text) > 0 And Len(txtDlvrSup.Text) > 0 And Len(txtDlvrSupBG.Text) > 0 And Len(cmbDlvrSupName.Text) > 0 And Len(txtDlvrDocType.Text) > 0 And Len(txtDlvrDocNo.Text) > 0 And Len(txtDlvrQuant.Text) > 0 Then
        Dlvr = txtDlvr.Text
        DlvrMat = cmbDlvrMat.Text
        DlvrSupName = cmbDlvrSupName.Text
        DlvrSupBG = txtDlvrSupBG.Text
        DlvrDocType = txtDlvrDocType.Text
        DlvrDocNo = txtDlvrDocNo.Text
        DlvrDate = Format(queDlvrDate.Value, "DD.MM.YYYY - HH:MM:SS")
        StampDate = Format(queDlvrDate.Value, "DD-MM-YYYY")
        DlvrQuant = ARound(CSng(rDs(txtDlvrQuant.Text)) / 1000, 5)

        If DayToday < StampDate Then
            MsgBox MsgCantSvFutDelivery, vbOKOnly Or vbCritical, MsgErrBx
            GoTo EndSub
        Else
        End If
        If MachineNumber = 1 Then MachineOther = 2
        If MachineNumber = 2 Then MachineOther = 1
        
'------------------------------Start PostgreSQL----------------------------------
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        'запис на доставката в таблица доставки
        comIns = "INSERT INTO deliveries VALUES(" & Dlvr & ",'" & DlvrMat & "','" & DlvrSupName & "','" & DlvrSupBG & "','" & DlvrDocType & "','" & DlvrDocNo & "','" & DlvrDate & "','" & StampDate & "','" & DlvrQuant & "','" & OperName & "')"
        Set rs = cn.Execute(comIns)
    
        'добав€не на доставеното количество към материала и в двете таблици
        Set rs = cn.Execute("SELECT m_del, m_name FROM materials_bc" & MachineNumber & " WHERE m_name = '" & DlvrMat & "';") 'маркираме по име
        Set rs = cn.Execute("UPDATE materials_bc" & MachineNumber & " SET m_del = '" & CSng(rDs(rs!m_del)) + CSng(rDs(DlvrQuant)) & "'WHERE m_name = '" & DlvrMat & "';") 'корекци€ по име
    
        Set rs = cn.Execute("SELECT m_del, m_name FROM materials_bc" & MachineOther & " WHERE m_name = '" & DlvrMat & "';") 'маркираме по име
        Set rs = cn.Execute("UPDATE materials_bc" & MachineOther & " SET m_del = '" & CSng(rDs(rs!m_del)) + CSng(rDs(DlvrQuant)) & "'WHERE m_name = '" & DlvrMat & "';") 'корекци€ по име
    
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
'--------------------------End PostgreSQL------------------------------------------
        MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
        Unload Me
        Call OpenMaterials
    Else
        MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
    End If
EndSub:
End Sub

