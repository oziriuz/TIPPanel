VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form setDrivers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "setDrivers"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11535
   Icon            =   "setDrivers.frx":0000
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
   Begin MSComctlLib.ListView lstSetDrv 
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
   Begin VB.Label lblSetDrv 
      Caption         =   "lblSetDrv"
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
Attribute VB_Name = "setDrivers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim colx    As MSComctlLib.ColumnHeader
    Dim itmX    As MSComctlLib.ListItem
    Dim cn3     As ADODB.Connection
    Dim rs3     As Recordset
    
    Me.Caption = uniSettings & " " & uniDrvs
    Me.lblSetDrv.Caption = lblSetDrivers
    Me.btnSave.Caption = uniSave
    Me.btnCancel.Caption = UniCancel
    Me.lstSetDrv.ColumnHeaders.Clear
    Me.lstSetDrv.ListItems.Clear
    Me.lstSetDrv.Checkboxes = True
    
    Set colx = Me.lstSetDrv.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 700
    Set colx = Me.lstSetDrv.ColumnHeaders.Add()
        colx.Text = uniNm
        colx.Width = 4000
    Set colx = Me.lstSetDrv.ColumnHeaders.Add()
        colx.Text = uniDrvReg
        colx.Width = 1200
    Set colx = Me.lstSetDrv.ColumnHeaders.Add()
        colx.Text = uniCapacity
        colx.Width = 1000
    Set colx = Me.lstSetDrv.ColumnHeaders.Add()
        colx.Text = uniMod
        colx.Width = 1500
    Set colx = Me.lstSetDrv.ColumnHeaders.Add()
        colx.Text = uniTel
        colx.Width = 1300
    Set colx = Me.lstSetDrv.ColumnHeaders.Add()
        colx.Text = uniNote
        colx.Width = 3000
'------------------------------Start PostgreSQL----------------------------------
    Set cn3 = New ADODB.Connection
        cn3.ConnectionTimeout = 10
        cn3.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rs3 = cn3.Execute("SELECT * FROM drivers ORDER BY d_num ASC;")
    If Not rs3.EOF And Not rs3.BOF Then rs3.MoveFirst
    Do While Not rs3.EOF
        Set itmX = Me.lstSetDrv.ListItems.Add(1, , Format(rs3!d_num, "0000"))
            itmX.SubItems(1) = rs3!d_name
            itmX.SubItems(2) = rs3!d_reg
            itmX.SubItems(3) = rDs(rs3!d_cap)
            itmX.SubItems(4) = rs3!d_mod
            itmX.SubItems(5) = rs3!d_tel
            itmX.SubItems(6) = rs3!d_note
        If rs3!d_show = True Then itmX.Checked = True
        rs3.MoveNext
    Loop
    rs3.Close
    Set rs3 = Nothing
    cn3.Close
    Set cn3 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault
    
    If Me.lstSetDrv.ListItems.count > 0 Then
        AutoColW Me.lstSetDrv
    Else
    End If
End Sub

Private Sub btnSave_Click()

    Dim cn3     As ADODB.Connection
    Dim rs3     As Recordset
    Dim comEdit As String
    Dim itmX    As MSComctlLib.ListItem
    
'------------------------------Start PostgreSQL----------------------------------
    Set cn3 = New ADODB.Connection
        cn3.ConnectionTimeout = 10
        cn3.Open ConStr
        
    MousePointer = vbHourglass
    
    For Each itmX In Me.lstSetDrv.ListItems
        comEdit = "UPDATE drivers SET d_show = '" & itmX.Checked & "' WHERE d_num =" & Val(itmX.Text) & ""
        Set rs3 = cn3.Execute(comEdit)
    Next
    If Not rs3 Is Nothing Then rs3.Close
    Set rs3 = Nothing
    cn3.Close
    Set cn3 = Nothing
    '--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault

    Call OpenDrivers
    Unload Me
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

