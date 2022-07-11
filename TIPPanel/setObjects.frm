VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form setObjects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "setObjects"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "setObjects.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   7110
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
      Left            =   4200
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
      Left            =   1440
      TabIndex        =   1
      Top             =   6960
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstSetObj 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Label lblCetObj 
      Caption         =   "lblSetObj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   6375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "setObjects"
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
    
    Me.Caption = uniSettings & " " & uniObj
    Me.lblCetObj.Caption = lblSetObjects
    Me.btnSave.Caption = uniSave
    Me.btnCancel.Caption = UniCancel
    Me.lstSetObj.ColumnHeaders.Clear
    Me.lstSetObj.ListItems.Clear
    Me.lstSetObj.Checkboxes = True

    Set colx = Me.lstSetObj.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 900
    
    Set colx = Me.lstSetObj.ColumnHeaders.Add()
        colx.Text = uniObj
        colx.Width = 2000
    
    Set colx = Me.lstSetObj.ColumnHeaders.Add()
        colx.Text = uniKmShort
        colx.Width = 800
'------------------------------Start PostgreSQL----------------------------------
    Set cn2 = New ADODB.Connection
        cn2.ConnectionTimeout = 10
        cn2.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rs2 = cn2.Execute("SELECT * FROM worksites WHERE w_cnum = '" & Val(DispPanel.txtClnt.Text) & "' ORDER BY w_name DESC;")
    If Not rs2.EOF And Not rs2.BOF Then rs2.MoveFirst
    Do While Not rs2.EOF
        Set itmX = Me.lstSetObj.ListItems.Add(1, , rs2!w_num)
        If rs2!w_show = True Then itmX.Checked = True
        itmX.SubItems(1) = rs2!w_name
        itmX.SubItems(2) = rs2!w_km
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
    cn2.Close
    Set cn2 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault
    
    If Me.lstSetObj.ListItems.count > 0 Then
        AutoColW Me.lstSetObj
    Else
    End If
End Sub

Private Sub btnSave_Click()

    Dim cn2     As ADODB.Connection
    Dim rs2     As Recordset
    Dim comEdit As String
    Dim itmX    As MSComctlLib.ListItem
    
'------------------------------Start PostgreSQL----------------------------------
    
    Set cn2 = New ADODB.Connection
        cn2.ConnectionTimeout = 10
        cn2.Open ConStr
        
    MousePointer = vbHourglass
    
    For Each itmX In Me.lstSetObj.ListItems
        comEdit = "UPDATE worksites SET w_show = '" & itmX.Checked & "' WHERE w_num =" & Val(itmX.Text) & ""
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

