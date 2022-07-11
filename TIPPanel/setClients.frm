VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form setClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "setClients"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11535
   Icon            =   "setClients.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6720
      TabIndex        =   2
      Top             =   6960
      Width           =   1455
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
      Left            =   3480
      TabIndex        =   1
      Top             =   6960
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstSetClnt 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10610
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
   Begin VB.Label lblCetClnt 
      Caption         =   "lblSetClnt"
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
      TabIndex        =   3
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "setClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim colx    As MSComctlLib.ColumnHeader
    Dim itmX    As MSComctlLib.ListItem
    Dim cn2     As ADODB.Connection
    Dim rs2     As Recordset
    
    Me.Caption = uniSettings & " " & uniClnts
    Me.lblCetClnt.Caption = lblSetClients
    Me.btnSave.Caption = uniSave
    Me.btnCancel.Caption = UniCancel
    Me.lstSetClnt.ColumnHeaders.Clear
    Me.lstSetClnt.ListItems.Clear
    Me.lstSetClnt.Checkboxes = True
    
    Set colx = Me.lstSetClnt.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 900
    Set colx = Me.lstSetClnt.ColumnHeaders.Add()
        colx.Text = uniFirm
        colx.Width = 2000
    Set colx = Me.lstSetClnt.ColumnHeaders.Add()
        colx.Text = uniBG
        colx.Width = 1400
    Set colx = Me.lstSetClnt.ColumnHeaders.Add()
        colx.Text = uniMOL
        colx.Width = 2000
    Set colx = Me.lstSetClnt.ColumnHeaders.Add()
        colx.Text = uniAdd
        colx.Width = 2000
    Set colx = Me.lstSetClnt.ColumnHeaders.Add()
        colx.Text = uniTel
        colx.Width = 1300
'------------------------------Start PostgreSQL----------------------------------
    Set cn2 = New ADODB.Connection
        cn2.ConnectionTimeout = 10
        cn2.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rs2 = cn2.Execute("SELECT * FROM clients ORDER BY c_num ASC;")
    If Not rs2.EOF And Not rs2.BOF Then rs2.MoveFirst
    Do While Not rs2.EOF
        Set itmX = Me.lstSetClnt.ListItems.Add(1, , Format(rs2!c_num, "0000"))
            itmX.SubItems(1) = rs2!c_name
            itmX.SubItems(2) = rs2!c_bg
            itmX.SubItems(3) = rs2!c_mol
            itmX.SubItems(4) = rs2!c_add
            itmX.SubItems(5) = rs2!c_tel
        If rs2!c_show = True Then itmX.Checked = True
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
    cn2.Close
    Set cn2 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault
    
    If Me.lstSetClnt.ListItems.count > 0 Then
        AutoColW Me.lstSetClnt
    Else
    End If
End Sub

Private Sub btnSave_Click()

    Dim cn2         As ADODB.Connection
    Dim rs2         As Recordset
    Dim comEdit     As String
    Dim itmX        As MSComctlLib.ListItem
    
'------------------------------Start PostgreSQL----------------------------------
    Set cn2 = New ADODB.Connection
        cn2.ConnectionTimeout = 10
        cn2.Open ConStr
        
    MousePointer = vbHourglass
    
    For Each itmX In Me.lstSetClnt.ListItems
        comEdit = "UPDATE clients SET c_show = '" & itmX.Checked & "' WHERE c_num =" & Val(itmX.Text) & ""
        Set rs2 = cn2.Execute(comEdit)
    Next
    If Not rs2 Is Nothing Then rs2.Close
    Set rs2 = Nothing
    cn2.Close
    Set cn2 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault

    Call OpenClients
    Unload Me
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

