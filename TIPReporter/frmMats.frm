VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMats"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10095
   Icon            =   "frmMats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnBack 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   7440
      Width           =   735
   End
   Begin VB.ComboBox cmbMat 
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
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "btnLoad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "btnExport"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "btnPrint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   7440
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstMats 
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11245
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
   Begin VB.Label lblMat 
      Alignment       =   1  'Right Justify
      Caption         =   "lblMat"
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
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmMats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstMats)
End Sub

Private Sub Form_Load()
'зареждане на меню материали
    
    Dim itmX As ListItem
    Dim types(0 To 4) As String

    Me.Caption = uniMats
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    Me.btnLoad.Caption = btLoad
    Me.lblMat.Caption = uniMat
    
    Me.lstMats.ColumnHeaders.Clear
    Me.lstMats.ListItems.Clear
    
    Me.cmbMat.Clear
    
    Set colx = Me.lstMats.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstMats.ColumnHeaders.Add()
        colx.Text = uniNm
        colx.Width = 3000
    
    Set colx = Me.lstMats.ColumnHeaders.Add()
        colx.Text = uniType
        colx.Width = 2000
    
    Set colx = Me.lstMats.ColumnHeaders.Add()
        colx.Text = uniLoad
        colx.Width = 2000

    Set colx = Me.lstMats.ColumnHeaders.Add()
        colx.Text = uniDelivered
        colx.Width = 1500

    Set colx = Me.lstMats.ColumnHeaders.Add()
        colx.Text = uniSold
        colx.Width = 1500
    
    Set colx = Me.lstMats.ColumnHeaders.Add()
        colx.Text = uniHave
        colx.Width = 1500
    
    types(0) = uniIM
    types(1) = uniConMat
    types(2) = uniWat
    types(3) = uniChem
    types(4) = uniOther
    
'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT * FROM materials_bc" & MachineNumber & " ORDER BY m_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbMat.AddItem rsR!m_name
        
        Set itmX = Me.lstMats.ListItems.Add(1, , Format(rsR!m_num, "000"))
            itmX.SubItems(1) = rsR!m_name
            If Val(rsR!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsR!m_type))
            itmX.SubItems(3) = rsR!m_load
            If Val(rsR!m_type) <> 2 Then
                itmX.SubItems(4) = rDs(rsR!m_del)
            Else
                itmX.SubItems(4) = "---------"
            End If
            itmX.SubItems(5) = rDs(rsR!m_sold)
            If Val(rsR!m_type) <> 2 Then
                itmX.SubItems(6) = CSng(rDs(rsR!m_del)) - CSng(rDs(rsR!m_sold))
            Else
                itmX.SubItems(6) = "---------"
            End If
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    AutoColW Me.lstMats
End Sub

Private Sub cmbMat_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbMat, KeyAscii, True)
End Sub

Private Sub btnLoad_Click()

If Me.cmbMat.Text <> "" Then

    Dim itmX As ListItem
    Dim types(0 To 4) As String
    
    types(0) = uniIM
    types(1) = uniConMat
    types(2) = uniWat
    types(3) = uniChem
    types(4) = uniOther
    
    Me.lstMats.ListItems.Clear
    
'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & Me.cmbMat.Text & "' ORDER BY m_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Set itmX = Me.lstMats.ListItems.Add(1, , Format(rsR!m_num, "000"))
            itmX.SubItems(1) = rsR!m_name
            If Val(rsR!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsR!m_type))
            itmX.SubItems(3) = rsR!m_load
            If Val(rsR!m_type) <> 2 Then
                itmX.SubItems(4) = rDs(rsR!m_del)
            Else
                itmX.SubItems(4) = "---------"
            End If
            itmX.SubItems(5) = rDs(rsR!m_sold)
            If Val(rsR!m_type) <> 2 Then
                itmX.SubItems(6) = CSng(rDs(rsR!m_del)) - CSng(rDs(rsR!m_sold))
            Else
                itmX.SubItems(6) = "---------"
            End If
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    AutoColW Me.lstMats
Else
    Call Form_Load
End If
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstMats, 1, True, True, True, uniMats)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

