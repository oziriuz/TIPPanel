VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRevision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmRevision"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11205
   Icon            =   "frmRevision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbOper 
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
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   8040
      Width           =   3015
   End
   Begin VB.TextBox txtSupervisor 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   8640
      Width           =   3015
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "btnSave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "btnCancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   8640
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdRev 
      Height          =   7095
      Left            =   360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12515
      _Version        =   393216
      Cols            =   6
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstRevPrnt 
      Height          =   2775
      Left            =   11400
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblRInfo 
      Alignment       =   2  'Center
      Caption         =   "lblRInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   7920
      Width           =   5055
   End
   Begin VB.Label lblOper 
      Alignment       =   1  'Right Justify
      Caption         =   "lblOper"
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
      Left            =   360
      TabIndex        =   7
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label lblSupervisor 
      Alignment       =   1  'Right Justify
      Caption         =   "lblSupervisor"
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
      Left            =   360
      TabIndex        =   6
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Label lblRev 
      Caption         =   "lblRev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   10335
   End
End
Attribute VB_Name = "frmRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colXH      As MSComctlLib.ColumnHeader
Dim itmXH      As MSComctlLib.ListItem
Dim PointLook7 As Boolean

Private Sub Form_Load()

    Dim cngr            As ADODB.Connection
    Dim rsgr            As Recordset
    Dim rsgrOther       As Recordset
    Dim MachineOther    As Integer
    Dim igr             As Integer
    Dim MatC            As Integer
    
    Me.Caption = uniRevision
    Me.lblRev.Caption = lblRevision
    Me.lblOper.Caption = uniDisp
    Me.lblSupervisor.Caption = uniRevisor
    Me.txtSupervisor.MaxLength = 30
    Me.btnSave.Caption = uniSave
    Me.btnCancel.Caption = UniCancel
    Me.lblRInfo = lblRevInfo
    Me.lstRevPrnt.ColumnHeaders.Clear
    Me.lstRevPrnt.ListItems.Clear
    
    If MachineNumber = 1 Then MachineOther = 2
    If MachineNumber = 2 Then MachineOther = 1
        
    Set colXH = Me.lstRevPrnt.ColumnHeaders.Add()
        colXH.Text = uniCode
        colXH.Width = 700
    Set colXH = Me.lstRevPrnt.ColumnHeaders.Add()
        colXH.Text = uniMat
        colXH.Width = 700
    Set colXH = Me.lstRevPrnt.ColumnHeaders.Add()
        colXH.Text = uniDelivered
        colXH.Width = 700
    Set colXH = Me.lstRevPrnt.ColumnHeaders.Add()
        colXH.Text = uniSold
        colXH.Width = 700
    Set colXH = Me.lstRevPrnt.ColumnHeaders.Add()
        colXH.Text = uniHave
        colXH.Width = 700
    Set colXH = Me.lstRevPrnt.ColumnHeaders.Add()
        colXH.Text = uniNewa & " " & uniHave
        colXH.Width = 700
    
    igr = 0
    MatC = 0
'------------------------------Start PostgreSQL----------------------------------
    Set cngr = New ADODB.Connection
        cngr.ConnectionTimeout = 10
        cngr.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rsgr = cngr.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_type <> '2' ORDER BY m_num ASC;")
    Set rsgrOther = cngr.Execute("SELECT * FROM materials_bc" & MachineOther & " WHERE m_type <> '2' ORDER BY m_num ASC;")
    
    If Not rsgr.EOF And Not rsgr.BOF Then
        If Not rsgrOther.EOF And Not rsgrOther.BOF Then rsgrOther.MoveFirst
        rsgr.MoveFirst
        Me.btnSave.Enabled = True
    Else
        MousePointer = vbDefault
        MsgBox MsgNoMat, vbOKOnly Or vbCritical, MsgErrBx
        rsgr.Close
        Set rsgr = Nothing
        cngr.Close
        Set cngr = Nothing
        Me.btnSave.Enabled = False
        GoTo EndSub
    End If
    Do While Not rsgr.EOF
        MatC = MatC + 1
        rsgr.MoveNext
    Loop
    rsgr.MoveFirst
    rsgrOther.MoveFirst
    With Me.grdRev
        .Rows = MatC + 1
        .Cols = 6
        .TextMatrix(0, 0) = uniCode
        .TextMatrix(0, 1) = uniMat
        .TextMatrix(0, 2) = uniDelivered
        .TextMatrix(0, 3) = uniSold
        .TextMatrix(0, 4) = uniHave
        .TextMatrix(0, 5) = uniNewa & " " & uniHave
        If Not rsgr.EOF And Not rsgr.BOF Then rsgr.MoveFirst
        If Not rsgrOther.EOF And Not rsgrOther.BOF Then rsgrOther.MoveFirst
        Do While Not rsgr.EOF Or Not rsgrOther.EOF
            igr = igr + 1
            .Rows = igr + 1
            .TextMatrix(igr, 0) = rsgr!m_num
            .TextMatrix(igr, 1) = rsgr!m_name
            .TextMatrix(igr, 2) = rDs(rsgr!m_del)
            .TextMatrix(igr, 3) = rDs(rsgr!m_sold)
            .TextMatrix(igr, 4) = CSng(rDs(rsgr!m_del)) - CSng(rDs(rsgr!m_sold)) - CSng(rDs(rsgrOther!m_sold))
            .TextMatrix(igr, 5) = ""
            rsgr.MoveNext
            rsgrOther.MoveNext
        Loop
    End With
    
    'зареждане на комбото с оператори
    Set rsgr = cngr.Execute("SELECT o_name FROM oper_data ORDER BY o_num ASC;")
    If Not rsgr.EOF And Not rsgr.BOF Then rsgr.MoveFirst
    Do While Not rsgr.EOF
        Me.cmbOper.AddItem rsgr!o_name
        rsgr.MoveNext
    Loop
    rsgr.Close
    Set rsgr = Nothing
    cngr.Close
    MousePointer = vbDefault
    Set cngr = Nothing
'--------------------------End PostgreSQL------------------------------------------
    FlexGrid_AutoSizeColumns grdRev, Me
EndSub:
End Sub

Private Sub grdRev_KeyPress(KeyAscii As Integer)
 
    Dim sTemp As String
 
    With grdRev
        sTemp = .TextMatrix(.Row, 5)
        If InStr(sTemp, DecSep) <> 0 Then
            PointLook7 = True
        Else
            PointLook7 = False
        End If
        Select Case KeyAscii
            Case 13  'Enter
                .TextMatrix(.Row, 5) = sTemp
                If .Row = .Rows - 1 Then
                    sTemp = ""
                    Exit Sub
                Else
                    .Row = .Row + 1
                    .Col = 5 'Make this zero if First Col also is edittable
                    sTemp = .TextMatrix(.Row, 5)
                End If
            Case 8 ' backspace
                If Len(sTemp) > 0 Then
                    sTemp = Left$(sTemp, Len(sTemp) - 1)
                End If
            Case 27 ' escape
                sTemp = ""
            Case 0 To 31
                KeyAscii = 0
            Case Else
                If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
                    KeyAscii = 0
                Else
                End If
                If Len(sTemp) > 10 Then
                    KeyAscii = 0
                Else
                End If
                If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And (PointLook7 = True Or Len(sTemp) = 0) Then
                    KeyAscii = 0
                Else
                    If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
                        KeyAscii = Asc(DecSep)
                        PointLook7 = True
                    Else
                    End If
                End If
                sTemp = sTemp & Chr$(KeyAscii)
        End Select
        .TextMatrix(.Row, 5) = sTemp
    End With
End Sub

Private Sub txtSupervisor_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 45 '-
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub btnSave_Click()

    Dim cngr         As ADODB.Connection
    Dim rsgr         As Recordset
    Dim rsgrOther    As Recordset
    Dim comInsgr     As String
    Dim comEditgr    As String
    Dim LastRS       As Long
    Dim LastRevision As Long
    Dim tempName     As String
    Dim tempQold     As String
    Dim tempQnew     As String
    Dim RevDate      As String
    Dim irow         As Integer
    Dim response     As Integer
    Dim MachineOther As Integer
        
    If MachineNumber = 1 Then MachineOther = 2
    If MachineNumber = 2 Then MachineOther = 1
        
    RevDate = Format(Now, "DD.MM.YYYY - HH:MM:SS")
        
    If Len(Me.cmbOper.Text) > 0 And Len(Me.txtSupervisor.Text) > 0 Then
'------------------------------Start PostgreSQL----------------------------------
        Set cngr = New ADODB.Connection
        cngr.ConnectionTimeout = 10
        cngr.Open ConStr
        MousePointer = vbHourglass
        
        Set rsgr = cngr.Execute("SELECT * FROM revision ORDER BY row_num DESC LIMIT 1;")
        
        If Not rsgr.BOF And Not rsgr.EOF Then
            LastRS = Val(rsgr!row_num) + 1
            LastRevision = Val(rsgr!rev_num) + 1
        Else
            LastRS = 1
            LastRevision = 1
        End If
    
        With Me.grdRev
            For irow = 1 To Me.grdRev.Rows - 1
                Set itmXH = Me.lstRevPrnt.ListItems.Add(1, , Format(.TextMatrix(irow, 0)))
                    itmXH.SubItems(1) = .TextMatrix(irow, 1)
                    itmXH.SubItems(2) = .TextMatrix(irow, 2)
                    itmXH.SubItems(3) = .TextMatrix(irow, 3)
                    itmXH.SubItems(4) = .TextMatrix(irow, 4)
                If .TextMatrix(irow, 5) = "" Then
                    itmXH.SubItems(5) = 0
                Else
                    itmXH.SubItems(5) = .TextMatrix(irow, 5)
                End If
                tempName = .TextMatrix(irow, 1)
                tempQold = .TextMatrix(irow, 4)
                tempQnew = .TextMatrix(irow, 5)
                If tempQnew = "" Then tempQnew = "0"
                If tempName <> "" Then
                    comInsgr = "INSERT INTO revision VALUES(" & LastRS & "," & LastRevision & ",'" & tempName & "','" & tempQold & "','" & tempQnew & "','" & Me.cmbOper.Text & "','" & Me.txtSupervisor.Text & "','" & RevDate & "','" & DayToday & "')"
                    Set rsgr = cngr.Execute(comInsgr)
                    comEditgr = "UPDATE materials_bc" & MachineNumber & " SET m_del = '" & tempQnew & "', m_sold = '0' WHERE m_name ='" & tempName & "'"
                    Set rsgr = cngr.Execute(comEditgr)
                    comEditgr = "UPDATE materials_bc" & MachineOther & " SET m_del = '" & tempQnew & "', m_sold = '0' WHERE m_name ='" & tempName & "'"
                    Set rsgrOther = cngr.Execute(comEditgr)
                    LastRS = LastRS + 1
                Else
                End If
            Next irow
        End With
        response = MsgBox(MsgClWatRev, vbYesNo Or vbQuestion, uniWat)
        If response = vbYes Then
            Set rsgr = cngr.Execute("SELECT m_name, m_sold FROM materials_bc" & MachineNumber & " WHERE m_type = '2'")
            comEditgr = "UPDATE materials_bc" & MachineNumber & " SET m_sold = '0' WHERE m_type = '2'"
            cngr.Execute comEditgr
            Set rsgrOther = cngr.Execute("SELECT m_name, m_sold FROM materials_bc" & MachineOther & " WHERE m_type = '2'")
            comEditgr = "UPDATE materials_bc" & MachineOther & " SET m_sold = '0' WHERE m_type = '2'"
            cngr.Execute comEditgr
        Else
        End If
        rsgr.Close
        rsgrOther.Close
        Set rsgr = Nothing
        Set rsgrOther = Nothing
        cngr.Close
        MousePointer = vbDefault
        Set cngr = Nothing
'--------------------------End PostgreSQL------------------------------------------
        Call PrintLVPic(Me.lstRevPrnt, 1, True, True, True, uniRevision & " " & LastRevision)
        Unload Me
        Call OpenMaterials
    Else
        MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
    End If
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

