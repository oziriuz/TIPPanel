VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form setRecepies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "setRecepies"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11535
   Icon            =   "setRecepies.frx":0000
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
   Begin MSComctlLib.ListView lstSetRec 
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
   Begin VB.Label lblSetRec 
      Caption         =   "lblSetRec"
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
Attribute VB_Name = "setRecepies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Const colw = 1300
    
    Dim colx    As MSComctlLib.ColumnHeader
    Dim itmX    As MSComctlLib.ListItem
    Dim count   As Integer
    Dim RecSet  As Recipe
    Dim cn4     As ADODB.Connection
    Dim rs4     As Recordset
    Dim n       As Integer
    
    Me.Caption = uniSettings & " " & uniRecs
    Me.lblSetRec.Caption = lblSetRecepies
    Me.btnSave.Caption = uniSave
    Me.btnCancel.Caption = UniCancel
    Me.lstSetRec.Checkboxes = True
    Me.lstSetRec.ColumnHeaders.Clear
    Me.lstSetRec.ListItems.Clear

    'настройка на заглавките на таблицата
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 700
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniNm
        colx.Width = 1600
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniRecType
        colx.Width = 1200
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniClass
        colx.Width = 1500
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniClassK
        colx.Width = 1700
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniClassV
        colx.Width = 1700
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniClassH
        colx.Width = 1700
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniClassP
        colx.Width = 1700
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniEDM
        colx.Width = 1200
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniTimePourShort
        colx.Width = 500
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniTimeMixShort
        colx.Width = 500
    'запис на имената на течките в таблицата
    For count = 1 To ns1
        Set colx = Me.lstSetRec.ColumnHeaders.Add()
            colx.Text = IM(count)
            colx.Width = colw
    Next count
    For count = 1 To ns3
        Set colx = Me.lstSetRec.ColumnHeaders.Add()
            colx.Text = Scr(count)
            colx.Width = colw
    Next count
    For count = 1 To ns2
        Set colx = Me.lstSetRec.ColumnHeaders.Add()
            colx.Text = Wat(count)
            colx.Width = colw
    Next count
    For count = 1 To ns4
        Set colx = Me.lstSetRec.ColumnHeaders.Add()
            colx.Text = Chem(count)
            colx.Width = colw
    Next count
    Set colx = Me.lstSetRec.ColumnHeaders.Add()
        colx.Text = uniTotalKg
        colx.Width = 1200

    Set RecSet = New Recipe

'------------------------------Start PostgreSQL----------------------------------
    Set cn4 = New ADODB.Connection
        cn4.ConnectionTimeout = 10
        cn4.Open ConStr
        
    MousePointer = vbHourglass
    
    Set rs4 = cn4.Execute("SELECT * FROM recepies ORDER BY r_num ASC;")
    If Not rs4.EOF And Not rs4.BOF Then rs4.MoveFirst
    Do While Not rs4.EOF
        RecSet.Code = Val(rs4!r_num)
        RecSet.Title = rs4!r_name
        RecSet.Kind = rs4!r_type
        RecSet.Class = rs4!r_class
        RecSet.ClassK = rs4!r_classk
        RecSet.ClassV = rs4!r_classv
        RecSet.ClassH = rs4!r_classh
        RecSet.ClassP = rs4!r_classp
        RecSet.EDM = Val(rs4!r_edm)
        RecSet.Tpour = Val(rs4!r_tpour)
        RecSet.Tmix = Val(rs4!r_tmix)
        RecSet.initIM(1) = Val(rs4!init_im1)
        RecSet.kgIM(1) = Val(rs4!kg_im1)
        RecSet.initIM(2) = Val(rs4!init_im2)
        RecSet.kgIM(2) = Val(rs4!kg_im2)
        RecSet.initIM(3) = Val(rs4!init_im3)
        RecSet.kgIM(3) = Val(rs4!kg_im3)
        RecSet.initIM(4) = Val(rs4!init_im4)
        RecSet.kgIM(4) = Val(rs4!kg_im4)
        RecSet.initIM(5) = Val(rs4!init_im5)
        RecSet.kgIM(5) = Val(rs4!kg_im5)
        RecSet.initScr(1) = Val(rs4!init_scr1)
        RecSet.kgScr(1) = Val(rs4!kg_scr1)
        RecSet.initScr(2) = Val(rs4!init_scr2)
        RecSet.kgScr(2) = Val(rs4!kg_scr2)
        RecSet.initScr(3) = Val(rs4!init_scr3)
        RecSet.kgScr(3) = Val(rs4!kg_scr3)
        RecSet.initScr(4) = Val(rs4!init_scr4)
        RecSet.kgScr(4) = Val(rs4!kg_scr4)
        RecSet.initWat(1) = Val(rs4!init_wat1)
        RecSet.kgWat(1) = Val(rs4!kg_wat1)
        RecSet.initWat(2) = Val(rs4!init_wat2)
        RecSet.kgWat(2) = Val(rs4!kg_wat2)
        RecSet.initChem(1) = Val(rs4!init_chem1)
        RecSet.kgChem(1) = CSng(rDs(rs4!kg_chem1))
        RecSet.initChem(2) = Val(rs4!init_chem2)
        RecSet.kgChem(2) = CSng(rDs(rs4!kg_chem2))
        RecSet.initChem(3) = Val(rs4!init_chem3)
        RecSet.kgChem(3) = CSng(rDs(rs4!kg_chem3))
        RecSet.initChem(4) = Val(rs4!init_chem4)
        RecSet.kgChem(4) = CSng(rDs(rs4!kg_chem4))
        RecSet.initChem(5) = Val(rs4!init_chem5)
        RecSet.kgChem(5) = CSng(rDs(rs4!kg_chem5))
        RecSet.initChem(6) = Val(rs4!init_chem6)
        RecSet.kgChem(6) = CSng(rDs(rs4!kg_chem6))
        RecSet.Visible = rs4!r_show
        
        Set itmX = Me.lstSetRec.ListItems.Add(1, , Format(RecSet.Code, "0000"))
            itmX.SubItems(1) = RecSet.Title
            itmX.SubItems(2) = RecSet.Kind
            itmX.SubItems(3) = RecSet.Class
            itmX.SubItems(4) = RecSet.ClassK
            itmX.SubItems(5) = RecSet.ClassV
            itmX.SubItems(6) = RecSet.ClassH
            itmX.SubItems(7) = RecSet.ClassP
            itmX.SubItems(8) = RecSet.EDM
            itmX.SubItems(9) = RecSet.Tpour
            itmX.SubItems(10) = RecSet.Tmix
        For n = 0 To ns1 - 1
            itmX.SubItems(11 + n) = RecSet.AllkgIM(n + 1)
        Next n
        For n = 0 To ns3 - 1
            itmX.SubItems(11 + ns1 + n) = RecSet.AllkgScr(n + 1)
        Next n
        For n = 0 To ns2 - 1
            itmX.SubItems(11 + ns1 + ns3 + n) = RecSet.kgWat(n + 1)
        Next n
        For n = 0 To ns4 - 1
            itmX.SubItems(11 + ns1 + ns3 + ns2 + n) = RecSet.AllkgChem(n + 1)
        Next n
            itmX.SubItems(11 + ns1 + ns3 + ns2 + ns4) = CSng(rDs(rs4!kg_total))
        If RecSet.Visible = True Then itmX.Checked = True
        rs4.MoveNext
    Loop
    rs4.Close
    Set rs4 = Nothing
    cn4.Close
    Set cn4 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault
    
    If Me.lstSetRec.ListItems.count > 0 Then
        AutoColW Me.lstSetRec
    Else
    End If
End Sub

Private Sub btnSave_Click()

    Dim cn4     As ADODB.Connection
    Dim rs4     As Recordset
    Dim comEdit As String
    Dim itmX    As MSComctlLib.ListItem
    
'------------------------------Start PostgreSQL----------------------------------
    Set cn4 = New ADODB.Connection
        cn4.ConnectionTimeout = 10
        cn4.Open ConStr
        
    MousePointer = vbHourglass
    
    For Each itmX In Me.lstSetRec.ListItems
        comEdit = "UPDATE recepies SET r_show = '" & itmX.Checked & "' WHERE r_num =" & Val(itmX.Text) & ""
        Set rs4 = cn4.Execute(comEdit)
    Next
    If Not rs4 Is Nothing Then rs4.Close
    Set rs4 = Nothing
    cn4.Close
    Set cn4 = Nothing
'--------------------------End PostgreSQL------------------------------------------
    MousePointer = vbDefault

    Call OpenRecepies
    Unload Me
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

