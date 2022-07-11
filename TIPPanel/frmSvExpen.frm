VERSION 5.00
Begin VB.Form frmSvExpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmSvExpen"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6570
   Icon            =   "frmSvExpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSvExpen 
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
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtExpenQuant 
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
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox cmbExpenMat 
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
      TabIndex        =   0
      Top             =   480
      Width           =   4095
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
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblExpenQuant 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblExpenQuant"
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
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblExpenMat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblExpenMat"
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
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmSvExpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Expen      As Long
Dim ExpenMat   As String
Dim ExpenDate  As String
Dim ExpenQuant As Single
Dim PointLook6 As Boolean

Private Sub Form_Load()
    
    Dim cn6 As ADODB.Connection
    Dim rs6 As Recordset
    
    Me.Caption = uniEnterExp
    lblExpenMat.Caption = uniMat
    lblExpenQuant.Caption = uniQuant
    btnSvExpen.Caption = uniSave
    
    Me.txtExpenQuant.MaxLength = 10
    
    DecSep = GetDecimalSep()
    
    If InStr(txtExpenQuant.Text, DecSep) = 0 Then PointLook6 = False
    
'------------------------------Start PostgreSQL----------------------------------
    Set cn6 = New ADODB.Connection
        cn6.ConnectionTimeout = 10
        cn6.Open ConStr
    
    MousePointer = vbHourglass
    
    'зареждане на последен номер на разход
    Set rs6 = cn6.Execute("SELECT row_num FROM other_expen ORDER BY row_num DESC LIMIT 1;")
    If Not rs6.EOF And Not rs6.BOF Then
        Expen = Val(rs6!row_num) + 1
    Else
        Expen = 1
    End If
    
    'зареждане на комбо с имена на други материали
    Set rs6 = cn6.Execute("SELECT m_name FROM materials_bc" & MachineNumber & " WHERE m_type = '4' ORDER BY m_num ASC;")
    If Not rs6.EOF And Not rs6.BOF Then rs6.MoveFirst
    Do While Not rs6.EOF
        Me.cmbExpenMat.AddItem rs6!m_name
        rs6.MoveNext
    Loop
    
    rs6.Close
    Set rs6 = Nothing
    cn6.Close
    MousePointer = vbDefault
    Set cn6 = Nothing
'--------------------------End PostgreSQL------------------------------------------

    On Error Resume Next

    If DispPanel.lstMat.ListItems.count > 0 And DispPanel.txtMat <> "" And DispPanel.cmbMatType.Text = uniOther Then
        Me.cmbExpenMat.Text = DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(1).Text
    Else
    End If
End Sub

Private Sub btnSvExpen_Click()
    
    Dim cn6    As ADODB.Connection
    Dim rs6    As Recordset
    Dim comIns As String
    Dim lastRowSv As Long
        
    If Len(Me.cmbExpenMat.Text) > 0 And Len(Me.txtExpenQuant.Text) > 0 And CSng(rDs(Me.txtExpenQuant.Text)) > 0 Then
        ExpenMat = Me.cmbExpenMat.Text
        ExpenQuant = ARound(CSng(rDs(Me.txtExpenQuant.Text)) / 1000, 5)
        ExpenDate = Format(Now, "DD.MM.YYYY - HH:MM:SS")
'------------------------------Start PostgreSQL----------------------------------
        Set cn6 = New ADODB.Connection
            cn6.ConnectionTimeout = 10
            cn6.Open ConStr
            
        MousePointer = vbHourglass
        
        'маркираме материала по име
        Set rs6 = cn6.Execute("SELECT m_del, m_sold FROM materials_bc" & MachineNumber & " WHERE m_name = '" & ExpenMat & "';")
        
        'проверка за достатъчна наличност
        If (CSng(rDs(rs6!m_del)) - CSng(rDs(rs6!m_sold))) < CSng(rDs(ExpenQuant)) Then
            MousePointer = vbDefault
            MsgBox MsgNotEnQuant, vbOKOnly Or vbCritical, MsgErrBx
            rs6.Close
            Set rs6 = Nothing
            cn6.Close
            Set cn6 = Nothing
            GoTo EndSub
        Else
            'запис на разхода в таблица други разходи
            comIns = "INSERT INTO other_expen VALUES(" & Expen & ",'" & ExpenMat & "','" & ExpenQuant & "','" & OperName & "','" & ExpenDate & "')"
            Set rs6 = cn6.Execute(comIns)
            'маркираме материала по име за да вземем старата стойност
            Set rs6 = cn6.Execute("SELECT m_name, m_sold FROM materials_bc" & MachineNumber & " WHERE m_name = '" & ExpenMat & "';")
            'корекция на разхода в таблица материали
            Set rs6 = cn6.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & CSng(rDs(rs6!m_sold)) + CSng(rDs(ExpenQuant)) & "'WHERE m_name = '" & ExpenMat & "';") 'корекция по име
            Set rs6 = cn6.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & ExpenMat & "' AND stamp_date = '" & DayToday & "';")
            If Not rs6.EOF And Not rs6.BOF Then
                Set rs6 = cn6.Execute("UPDATE daily_expenses SET mat_sold = '" & CSng(rDs(rs6!mat_sold)) + CSng(rDs(ExpenQuant)) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & ExpenMat & "' AND stamp_date = '" & DayToday & "';")
            Else
                'откриваме последния запис
                Set rs6 = cn6.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                If Not rs6.EOF And Not rs6.BOF Then
                    lastRowSv = Val(rs6!row_num) + 1
                Else
                    lastRowSv = 1
                End If
                Set rs6 = cn6.Execute("INSERT INTO daily_expenses VALUES(" & lastRowSv & ",'" & ExpenMat & "','" & CSng(rDs(ExpenQuant)) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
            End If
        End If
        rs6.Close
        Set rs6 = Nothing
        cn6.Close
        MousePointer = vbDefault
        Set cn6 = Nothing
'--------------------------End PostgreSQL------------------------------------------
        MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
        Unload Me
        Call OpenMaterials
    Else
        MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
    End If
EndSub:
End Sub

Private Sub txtExpenQuant_GotFocus()

    txtExpenQuant.SelStart = 0
    txtExpenQuant.SelLength = Len(txtExpenQuant.Text)

    If InStr(txtExpenQuant.Text, DecSep) <> 0 Then
        PointLook6 = True
    Else
        PointLook6 = False
    End If
End Sub

Private Sub txtExpenQuant_Change()

    If InStr(txtExpenQuant.Text, DecSep) <> 0 Then
        PointLook6 = True
    Else
        PointLook6 = False
    End If
End Sub

Private Sub txtExpenQuant_KeyPress(KeyAscii As Integer)

    If InStr(txtExpenQuant.Text, DecSep) <> 0 Then
        PointLook6 = True
    Else
        PointLook6 = False
    End If
    If txtExpenQuant.SelLength = Len(txtExpenQuant.Text) Then
        PointLook6 = False
    Else
    End If
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If
    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook6 = True Then
        KeyAscii = 0
    Else
        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook6 = True
        Else
        End If
    End If
End Sub

