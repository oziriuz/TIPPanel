VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClnts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmClnts"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
   Icon            =   "frmClnts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11670
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
      Left            =   10560
      TabIndex        =   6
      Top             =   7800
      Width           =   735
   End
   Begin VB.ComboBox cmbClnt 
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
      Left            =   4920
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
      Left            =   8400
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
      Left            =   6600
      TabIndex        =   2
      Top             =   7800
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
      Left            =   2760
      TabIndex        =   1
      Top             =   7800
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstClnts 
      Height          =   6735
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11880
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
   Begin VB.Label lblClnt 
      Alignment       =   1  'Right Justify
      Caption         =   "lblClnt"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmClnts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'зареждане на меню клиенти
    
    Dim itmX As ListItem
    
    Me.Caption = uniClnts
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    Me.btnLoad.Caption = btLoad
    Me.lblClnt.Caption = uniClnt
    
    Me.lstClnts.ColumnHeaders.Clear
    Me.lstClnts.ListItems.Clear
    
    Me.cmbClnt.Clear
    
    Set colx = Me.lstClnts.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstClnts.ColumnHeaders.Add()
        colx.Text = uniFirm
        colx.Width = 2000
    
    Set colx = Me.lstClnts.ColumnHeaders.Add()
        colx.Text = uniBG
        colx.Width = 1400
    
    Set colx = Me.lstClnts.ColumnHeaders.Add()
        colx.Text = uniMOL
        colx.Width = 2000

    Set colx = Me.lstClnts.ColumnHeaders.Add()
        colx.Text = uniAdd
        colx.Width = 2000

    Set colx = Me.lstClnts.ColumnHeaders.Add()
        colx.Text = uniTel
        colx.Width = 1300
    
'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT * FROM clients ORDER BY c_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbClnt.AddItem rsR!c_name
        
        Set itmX = Me.lstClnts.ListItems.Add(1, , Format(rsR!c_num, "0000"))
            itmX.SubItems(1) = rsR!c_name
            itmX.SubItems(2) = rsR!c_bg
            itmX.SubItems(3) = rsR!c_mol
            itmX.SubItems(4) = rsR!c_add
            itmX.SubItems(5) = rsR!c_tel
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    AutoColW Me.lstClnts
End Sub

Private Sub cmbClnt_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbClnt, KeyAscii, True)
End Sub

Private Sub btnLoad_Click()

If Me.cmbClnt.Text <> "" Then

    Me.lstClnts.ListItems.Clear
        
'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT * FROM clients WHERE c_name = '" & Me.cmbClnt.Text & "' ORDER BY c_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Set itmX = Me.lstClnts.ListItems.Add(1, , Format(rsR!c_num, "0000"))
            itmX.SubItems(1) = rsR!c_name
            itmX.SubItems(2) = rsR!c_bg
            itmX.SubItems(3) = rsR!c_mol
            itmX.SubItems(4) = rsR!c_add
            itmX.SubItems(5) = rsR!c_tel
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    AutoColW Me.lstClnts
Else
    Call Form_Load
End If
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstClnts, 2, True, True, True, uniClnts)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstClnts)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

