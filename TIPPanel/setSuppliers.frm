VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form setSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "setSuppliers"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11535
   Icon            =   "setSuppliers.frx":0000
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   6960
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstSetSup 
      Height          =   6015
      Left            =   240
      TabIndex        =   3
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
   Begin VB.Label lblSetSup 
      Caption         =   "lblSetSup"
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
      TabIndex        =   2
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "setSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim colx    As MSComctlLib.ColumnHeader
    Dim itmX    As MSComctlLib.ListItem
    Dim cn5     As ADODB.Connection
    Dim rs5     As Recordset
    
    Me.Caption = uniSettings & " " & uniSups
    Me.lblSetSup.Caption = lblSetSuppliers
    Me.btnSave.Caption = uniSave
    Me.btnCancel.Caption = UniCancel
    Me.lstSetSup.ColumnHeaders.Clear
    Me.lstSetSup.ListItems.Clear
    Me.lstSetSup.Checkboxes = True
    
    Set colx = Me.lstSetSup.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 700
    Set colx = Me.lstSetSup.ColumnHeaders.Add()
    colx.Text = uniFirm
    colx.Width = 3000
    Set colx = Me.lstSetSup.ColumnHeaders.Add()
    colx.Text = uniBG
    colx.Width = 1400
    Set colx = Me.lstSetSup.ColumnHeaders.Add()
    colx.Text = uniMOL
    colx.Width = 2500
    Set colx = Me.lstSetSup.ColumnHeaders.Add()
    colx.Text = uniAdd
    colx.Width = 2500
    Set colx = Me.lstSetSup.ColumnHeaders.Add()
    colx.Text = uniTel
    colx.Width = 1300
    Set colx = Me.lstSetSup.ColumnHeaders.Add()
    colx.Text = uniNote
    colx.Width = 2200
'------------------------------Start PostgreSQL----------------------------------
    Set cn5 = New ADODB.Connection
        cn5.ConnectionTimeout = 10
        cn5.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rs5 = cn5.Execute("SELECT * FROM suppliers ORDER BY s_num ASC;")
    If Not rs5.EOF And Not rs5.BOF Then rs5.MoveFirst
    Do While Not rs5.EOF
        Set itmX = Me.lstSetSup.ListItems.Add(1, , Format(rs5!s_num, "000"))
            itmX.SubItems(1) = rs5!s_name
            itmX.SubItems(2) = rs5!s_bg
            itmX.SubItems(3) = rs5!s_mol
            itmX.SubItems(4) = rs5!s_add
            itmX.SubItems(5) = rs5!s_tel
            itmX.SubItems(6) = rs5!s_note
        If rs5!s_show = True Then itmX.Checked = True
        rs5.MoveNext
    Loop
    rs5.Close
    Set rs5 = Nothing
    cn5.Close
    Set cn5 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault
    
    If Me.lstSetSup.ListItems.count > 0 Then
        AutoColW Me.lstSetSup
    Else
    End If
End Sub

Private Sub btnSave_Click()

    Dim cn5     As ADODB.Connection
    Dim rs5     As Recordset
    Dim comEdit As String
    Dim itmX    As MSComctlLib.ListItem
'------------------------------Start PostgreSQL----------------------------------
    Set cn5 = New ADODB.Connection
        cn5.ConnectionTimeout = 10
        cn5.Open ConStr
        
    MousePointer = vbHourglass
    
    For Each itmX In Me.lstSetSup.ListItems
        comEdit = "UPDATE suppliers SET s_show = '" & itmX.Checked & "' WHERE s_num =" & Val(itmX.Text) & ""
        Set rs5 = cn5.Execute(comEdit)
    Next
    If Not rs5 Is Nothing Then rs5.Close
    Set rs5 = Nothing
    cn5.Close
    Set cn5 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault

    Call OpenSuppliers
    Unload Me
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

