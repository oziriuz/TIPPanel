VERSION 5.00
Begin VB.Form frmNameSilos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmNameSilos"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbWat 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox cmbWat 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2295
   End
   Begin VB.ComboBox cmbIM 
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ComboBox cmbChem 
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   7080
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cmbChem 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   7080
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ComboBox cmbChem 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   7080
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2775
   End
   Begin VB.ComboBox cmbChem 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   7080
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
      Width           =   2775
   End
   Begin VB.ComboBox cmbChem 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7080
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox cmbScr 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   3840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cmbScr 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cmbScr 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox cmbIM 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ComboBox cmbIM 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox cmbIM 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox cmbIM 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox cmbChem 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   7080
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   240
      Width           =   2775
   End
   Begin VB.ComboBox cmbScr 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   2295
   End
   Begin VB.ComboBox cmbIM 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
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
      Left            =   5400
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton btnSaveSilos 
      Caption         =   "btnSaveSilos"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblWat2 
      Alignment       =   1  'Right Justify
      Caption         =   "lblWat2"
      Height          =   255
      Left            =   3120
      TabIndex        =   37
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblIM6 
      Alignment       =   1  'Right Justify
      Caption         =   "lblIM6"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblChem6 
      Alignment       =   1  'Right Justify
      Caption         =   "lblChem6"
      Height          =   255
      Left            =   6360
      TabIndex        =   32
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblChem5 
      Alignment       =   1  'Right Justify
      Caption         =   "lblChem5"
      Height          =   255
      Left            =   6360
      TabIndex        =   31
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblChem4 
      Alignment       =   1  'Right Justify
      Caption         =   "lblChem4"
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblChem3 
      Alignment       =   1  'Right Justify
      Caption         =   "lblChem3"
      Height          =   255
      Left            =   6360
      TabIndex        =   29
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblChem2 
      Alignment       =   1  'Right Justify
      Caption         =   "lblChem2"
      Height          =   255
      Left            =   6360
      TabIndex        =   28
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblChem1 
      Alignment       =   1  'Right Justify
      Caption         =   "lblChem1"
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblWat1 
      Alignment       =   1  'Right Justify
      Caption         =   "lblWat1"
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblScr4 
      Alignment       =   1  'Right Justify
      Caption         =   "lblScr4"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblScr3 
      Alignment       =   1  'Right Justify
      Caption         =   "lblScr3"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblScr2 
      Alignment       =   1  'Right Justify
      Caption         =   "lblScr2"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblScr1 
      Alignment       =   1  'Right Justify
      Caption         =   "lblScr1"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblIM5 
      Alignment       =   1  'Right Justify
      Caption         =   "lblIM5"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblIM4 
      Alignment       =   1  'Right Justify
      Caption         =   "lblIM4"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblIM3 
      Alignment       =   1  'Right Justify
      Caption         =   "lblIM3"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblIM2 
      Alignment       =   1  'Right Justify
      Caption         =   "lblIM2"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblIM1 
      Alignment       =   1  'Right Justify
      Caption         =   "lblIM1"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmNameSilos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim intEmpFile  As Integer
    Dim i           As Integer
    Dim cn          As ADODB.Connection
    Dim rs          As Recordset

    intEmpFile = FreeFile
    
    Call LoadMat
    
    frmNameSilos.Caption = frmNmSilos
    lblIM1.Caption = uniIMShort & "1"
    lblIM2.Caption = uniIMShort & "2"
    lblIM3.Caption = uniIMShort & "3"
    lblIM4.Caption = uniIMShort & "4"
    lblIM5.Caption = uniIMShort & "5"
    lblIM6.Caption = uniIMShort & "6"
    lblScr1.Caption = uniCemShort & "1"
    lblScr2.Caption = uniCemShort & "2"
    lblScr3.Caption = uniCemShort & "3"
    lblScr4.Caption = uniCemShort & "4"
    lblWat1.Caption = uniWat & "1"
    lblWat1.Caption = uniWat & "2"
    lblChem1.Caption = uniChemShort & "1"
    lblChem2.Caption = uniChemShort & "2"
    lblChem3.Caption = uniChemShort & "3"
    lblChem4.Caption = uniChemShort & "4"
    lblChem5.Caption = uniChemShort & "5"
    lblChem6.Caption = uniChemShort & "6"
    btnSaveSilos.Caption = uniSave
    btnCancel.Caption = UniCancel
    
    Select Case ns1
        Case 1
            cmbIM(0).Visible = True
            cmbIM(1).Visible = False
            cmbIM(2).Visible = False
            cmbIM(3).Visible = False
            cmbIM(4).Visible = False
            cmbIM(5).Visible = False
            lblIM2.Caption = ""
            lblIM3.Caption = ""
            lblIM4.Caption = ""
            lblIM5.Caption = ""
            lblIM6.Caption = ""
        Case 2
            cmbIM(0).Visible = True
            cmbIM(1).Visible = True
            cmbIM(2).Visible = False
            cmbIM(3).Visible = False
            cmbIM(4).Visible = False
            cmbIM(5).Visible = False
            lblIM3.Caption = ""
            lblIM4.Caption = ""
            lblIM5.Caption = ""
            lblIM6.Caption = ""
        Case 3
            cmbIM(0).Visible = True
            cmbIM(1).Visible = True
            cmbIM(2).Visible = True
            cmbIM(3).Visible = False
            cmbIM(4).Visible = False
            cmbIM(5).Visible = False
            lblIM4.Caption = ""
            lblIM5.Caption = ""
            lblIM6.Caption = ""
        Case 4
            cmbIM(0).Visible = True
            cmbIM(1).Visible = True
            cmbIM(2).Visible = True
            cmbIM(3).Visible = True
            cmbIM(4).Visible = False
            cmbIM(5).Visible = False
            lblIM5.Caption = ""
            lblIM6.Caption = ""
        Case 5
            cmbIM(0).Visible = True
            cmbIM(1).Visible = True
            cmbIM(2).Visible = True
            cmbIM(3).Visible = True
            cmbIM(4).Visible = True
            cmbIM(5).Visible = False
            lblIM6.Caption = ""
        Case 6
            cmbIM(0).Visible = True
            cmbIM(1).Visible = True
            cmbIM(2).Visible = True
            cmbIM(3).Visible = True
            cmbIM(4).Visible = True
            cmbIM(5).Visible = True
        Case Else
            cmbIM(0).Visible = False
            cmbIM(1).Visible = False
            cmbIM(2).Visible = False
            cmbIM(3).Visible = False
            cmbIM(4).Visible = False
            cmbIM(5).Visible = False
            lblIM1.Caption = ""
            lblIM2.Caption = ""
            lblIM3.Caption = ""
            lblIM4.Caption = ""
            lblIM5.Caption = ""
            lblIM6.Caption = ""
    End Select
    
    Select Case ns3
        Case 1
            cmbScr(0).Visible = True
            cmbScr(1).Visible = False
            cmbScr(2).Visible = False
            cmbScr(3).Visible = False
            lblScr2.Caption = ""
            lblScr3.Caption = ""
            lblScr4.Caption = ""
        Case 2
            cmbScr(0).Visible = True
            cmbScr(1).Visible = True
            cmbScr(2).Visible = False
            cmbScr(3).Visible = False
            lblScr3.Caption = ""
            lblScr4.Caption = ""
        Case 3
            cmbScr(0).Visible = True
            cmbScr(1).Visible = True
            cmbScr(2).Visible = True
            cmbScr(3).Visible = False
            lblScr4.Caption = ""
        Case 4
            cmbScr(0).Visible = True
            cmbScr(1).Visible = True
            cmbScr(2).Visible = True
            cmbScr(3).Visible = True
        Case Else
            cmbScr(0).Visible = False
            cmbScr(1).Visible = False
            cmbScr(2).Visible = False
            cmbScr(3).Visible = False
            lblScr1.Caption = ""
            lblScr2.Caption = ""
            lblScr3.Caption = ""
            lblScr4.Caption = ""
    End Select
    
    Select Case ns2
        Case 1
            cmbWat(0).Visible = True
            cmbWat(1).Visible = False
            lblWat2.Caption = ""
        Case 2
            cmbWat(0).Visible = True
            cmbWat(1).Visible = True
        Case Else
            cmbWat(0).Visible = False
            cmbWat(1).Visible = False
            lblWat1.Caption = ""
            lblWat2.Caption = ""
    End Select
    
    Select Case ns4
        Case 1
            cmbChem(0).Visible = True
            cmbChem(1).Visible = False
            cmbChem(2).Visible = False
            cmbChem(3).Visible = False
            cmbChem(4).Visible = False
            cmbChem(5).Visible = False
            lblChem2.Caption = ""
            lblChem3.Caption = ""
            lblChem4.Caption = ""
            lblChem5.Caption = ""
            lblChem6.Caption = ""
        Case 2
            cmbChem(0).Visible = True
            cmbChem(1).Visible = True
            cmbChem(2).Visible = False
            cmbChem(3).Visible = False
            cmbChem(4).Visible = False
            cmbChem(5).Visible = False
            lblChem3.Caption = ""
            lblChem4.Caption = ""
            lblChem5.Caption = ""
            lblChem6.Caption = ""
        Case 3
            cmbChem(0).Visible = True
            cmbChem(1).Visible = True
            cmbChem(2).Visible = True
            cmbChem(3).Visible = False
            cmbChem(4).Visible = False
            cmbChem(5).Visible = False
            lblChem4.Caption = ""
            lblChem5.Caption = ""
            lblChem6.Caption = ""
        Case 4
            cmbChem(0).Visible = True
            cmbChem(1).Visible = True
            cmbChem(2).Visible = True
            cmbChem(3).Visible = True
            cmbChem(4).Visible = False
            cmbChem(5).Visible = False
            lblChem5.Caption = ""
            lblChem6.Caption = ""
        Case 5
            cmbChem(0).Visible = True
            cmbChem(1).Visible = True
            cmbChem(2).Visible = True
            cmbChem(3).Visible = True
            cmbChem(4).Visible = True
            cmbChem(5).Visible = False
            lblChem6.Caption = ""
        Case 6
            cmbChem(0).Visible = True
            cmbChem(1).Visible = True
            cmbChem(2).Visible = True
            cmbChem(3).Visible = True
            cmbChem(4).Visible = True
            cmbChem(5).Visible = True
        Case Else
            cmbChem(0).Visible = False
            cmbChem(1).Visible = False
            cmbChem(2).Visible = False
            cmbChem(3).Visible = False
            cmbChem(4).Visible = False
            cmbChem(5).Visible = False
            lblChem1.Caption = ""
            lblChem2.Caption = ""
            lblChem3.Caption = ""
            lblChem4.Caption = ""
            lblChem5.Caption = ""
            lblChem6.Caption = ""
    End Select

    For i = 0 To 4
        cmbIM(i).Clear
        cmbIM(i).AddItem uniEmpty
    Next i
    For i = 0 To 3
        cmbScr(i).Clear
        cmbScr(i).AddItem uniEmpty
    Next i
    For i = 0 To 1
        cmbWat(i).Clear
        cmbWat(i).AddItem uniEmpty
    Next i
    For i = 0 To 5
        cmbChem(i).Clear
        cmbChem(i).AddItem uniEmpty
    Next i
'------------------------------Start PostgreSQL----------------------------------
    Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT * FROM materials_bc" & MachineNumber & " ORDER BY m_num ASC;")
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    Do While Not rs.EOF
        If Val(rs!m_type) >= 0 Then
            If ns1 = 0 Then GoTo WOim
            If Val(rs!m_type) = 0 Then
                For i = 0 To 5
                    cmbIM(i).AddItem rs!m_name
                Next i
                If Val(Mid$(rs!m_load, 1, 2)) = 1 Then cmbIM(0).Text = rs!m_name Else
                If Val(Mid$(rs!m_load, 3, 2)) = 1 Then cmbIM(1).Text = rs!m_name Else
                If Val(Mid$(rs!m_load, 5, 2)) = 1 Then cmbIM(2).Text = rs!m_name Else
                If Val(Mid$(rs!m_load, 7, 2)) = 1 Then cmbIM(3).Text = rs!m_name Else
                If Val(Mid$(rs!m_load, 9, 2)) = 1 Then cmbIM(4).Text = rs!m_name Else
                If Val(Mid$(rs!m_load, 11, 2)) = 1 Then cmbIM(5).Text = rs!m_name
            End If
WOim:
            If ns3 = 0 Then GoTo WOscr
            If Val(rs!m_type) = 1 Then
                For i = 0 To 3
                    cmbScr(i).AddItem rs!m_name
                Next i
                If Val(Mid$(rs!m_load, 1, 2)) = 1 Then cmbScr(0).Text = rs!m_name
                If Val(Mid$(rs!m_load, 3, 2)) = 1 Then cmbScr(1).Text = rs!m_name
                If Val(Mid$(rs!m_load, 5, 2)) = 1 Then cmbScr(2).Text = rs!m_name
                If Val(Mid$(rs!m_load, 7, 2)) = 1 Then cmbScr(3).Text = rs!m_name
            End If
WOscr:
            If ns2 = 0 Then GoTo WOwat
            If Val(rs!m_type) = 2 Then
                For i = 0 To 1
                    cmbWat(i).AddItem rs!m_name
                Next i
                If Val(Mid$(rs!m_load, 1, 2)) = 1 Then cmbWat(0).Text = rs!m_name
                If Val(Mid$(rs!m_load, 3, 2)) = 1 Then cmbWat(1).Text = rs!m_name
            End If
WOwat:

            If ns4 = 0 Then GoTo WOchem
            If Val(rs!m_type) = 3 Then
                For i = 0 To 5
                    cmbChem(i).AddItem rs!m_name
                Next i
                If Val(Mid$(rs!m_load, 1, 2)) = 1 Then cmbChem(0).Text = rs!m_name
                If Val(Mid$(rs!m_load, 3, 2)) = 1 Then cmbChem(1).Text = rs!m_name
                If Val(Mid$(rs!m_load, 5, 2)) = 1 Then cmbChem(2).Text = rs!m_name
                If Val(Mid$(rs!m_load, 7, 2)) = 1 Then cmbChem(3).Text = rs!m_name
                If Val(Mid$(rs!m_load, 9, 2)) = 1 Then cmbChem(4).Text = rs!m_name
                If Val(Mid$(rs!m_load, 11, 2)) = 1 Then cmbChem(5).Text = rs!m_name
            End If
WOchem:
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
'--------------------------End PostgreSQL------------------------------------------
    For i = 0 To ns1 - 1
        If cmbIM(i).ListIndex = -1 Then cmbIM(i).Text = uniEmpty
    Next i
    For i = 0 To ns3 - 1
        If cmbScr(i).ListIndex = -1 Then cmbScr(i).Text = uniEmpty
    Next i
    For i = 0 To ns2 - 1
        If cmbWat(i).ListIndex = -1 Then cmbWat(i).Text = uniEmpty
    Next i
    For i = 0 To ns4 - 1
        If cmbChem(i).ListIndex = -1 Then cmbChem(i).Text = uniEmpty
    Next i

    If Dir(SilosFile) = "" Then
        Open SilosFile For Output As #intEmpFile
        Close #intEmpFile
    Else
        Open SilosFile For Input As #intEmpFile
        Do Until EOF(intEmpFile)
            Input #intEmpFile, IM(1), IM(2), IM(3), IM(4), IM(5), IM(6), Scr(1), Scr(2), Scr(3), Scr(4), Wat(1), Wat(2), Chem(1), Chem(2), Chem(3), Chem(4), Chem(5), Chem(6)
        Loop
        Close #intEmpFile
    End If
    For i = 0 To ns1 - 1
        cmbIM(i).Refresh
    Next i
    For i = 0 To ns3 - 1
        cmbScr(i).Refresh
    Next i
    For i = 0 To ns2 - 1
        cmbWat(i).Refresh
    Next i
    For i = 0 To ns4 - 1
        cmbChem(i).Refresh
    Next i
End Sub

Public Sub btnSaveSilos_Click()

    Dim intEmpFile          As Integer
    Dim i                   As Integer
    Dim cn                  As New ADODB.Connection
    Dim rs                  As New ADODB.Recordset
    Dim comIns(2 To 19)     As String
    Dim comEdit(2 To 19)    As String
    Dim rwind               As Integer
    
    intEmpFile = FreeFile
    
    For i = 0 To ns1 - 1
        IM(i + 1) = cmbIM(i).List(cmbIM(i).ListIndex)
    Next i
    For i = 0 To ns3 - 1
        Scr(i + 1) = cmbScr(i).List(cmbScr(i).ListIndex)
    Next i
    For i = 0 To ns2 - 1
        Wat(i + 1) = cmbWat(i).List(cmbWat(i).ListIndex)
    Next i
    For i = 0 To ns4 - 1
        Chem(i + 1) = cmbChem(i).List(cmbChem(i).ListIndex)
    Next i
    
    Kill SilosFile
    
    Open SilosFile For Output As #intEmpFile
    Write #intEmpFile, IM(1), IM(2), IM(3), IM(4), IM(5), IM(6), Scr(1), Scr(2), Scr(3), Scr(4), Wat(1), Wat(2), Chem(1), Chem(2), Chem(3), Chem(4), Chem(5), Chem(6)
    Close #intEmpFile
'-----------------------Start postgreSQL-----------------------------------
    Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
    
    MousePointer = vbHourglass
    
    comIns(2) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(2,'" & IM(1) & "')"
    comIns(3) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(3,'" & IM(2) & "')"
    comIns(4) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(4,'" & IM(3) & "')"
    comIns(5) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(5,'" & IM(4) & "')"
    comIns(6) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(6,'" & IM(5) & "')"
    comIns(7) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(7,'" & IM(6) & "')"
    comIns(8) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(8,'" & Scr(1) & "')"
    comIns(9) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(9,'" & Scr(2) & "')"
    comIns(10) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(10,'" & Scr(3) & "')"
    comIns(11) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(11,'" & Scr(4) & "')"
    comIns(12) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(12,'" & Wat(1) & "')"
    comIns(13) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(13,'" & Wat(2) & "')"
    comIns(14) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(14,'" & Chem(1) & "')"
    comIns(15) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(15,'" & Chem(2) & "')"
    comIns(16) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(16,'" & Chem(3) & "')"
    comIns(17) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(17,'" & Chem(4) & "')"
    comIns(18) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(18,'" & Chem(5) & "')"
    comIns(19) = "INSERT INTO settings_bc" & MachineNumber & " VALUES(19,'" & Chem(6) & "')"
    comEdit(2) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & IM(1) & "' WHERE ind = 2"
    comEdit(3) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & IM(2) & "' WHERE ind = 3"
    comEdit(4) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & IM(3) & "' WHERE ind = 4"
    comEdit(5) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & IM(4) & "' WHERE ind = 5"
    comEdit(6) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & IM(5) & "' WHERE ind = 6"
    comEdit(7) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & IM(6) & "' WHERE ind = 7"
    comEdit(8) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Scr(1) & "' WHERE ind = 8"
    comEdit(9) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Scr(2) & "' WHERE ind = 9"
    comEdit(10) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Scr(3) & "' WHERE ind = 10"
    comEdit(11) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Scr(4) & "' WHERE ind = 11"
    comEdit(12) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Wat(1) & "' WHERE ind = 12"
    comEdit(13) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Wat(2) & "' WHERE ind = 13"
    comEdit(14) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Chem(1) & "' WHERE ind = 14"
    comEdit(15) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Chem(2) & "' WHERE ind = 15"
    comEdit(16) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Chem(3) & "' WHERE ind = 16"
    comEdit(17) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Chem(4) & "' WHERE ind = 17"
    comEdit(18) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Chem(5) & "' WHERE ind = 18"
    comEdit(19) = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & Chem(6) & "' WHERE ind = 19"
                    
    Set rs = cn.Execute("SELECT * FROM settings_bc" & MachineNumber & " WHERE ind >= 2 AND ind <=19;")
    rwind = 1
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            rwind = rwind + 1
            cn.Execute comEdit(rwind)
            rs.MoveNext
        Loop
        If rs.EOF Then
            For i = rwind + 1 To 19
                cn.Execute comIns(i)
            Next i
        End If
    Else
        For rwind = 2 To 19
            cn.Execute comIns(rwind)
        Next rwind
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
'--------------------------End PostgreSQL-----------------------------------
    Unload Me
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

