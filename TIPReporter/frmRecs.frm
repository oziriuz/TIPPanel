VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmRecs"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15975
   Icon            =   "frmRecs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   15975
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
      Left            =   14880
      TabIndex        =   6
      Top             =   7680
      Width           =   735
   End
   Begin VB.ComboBox cmbRec 
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
      Left            =   9240
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
      Left            =   12720
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
      Left            =   9120
      TabIndex        =   2
      Top             =   7680
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
      Left            =   4560
      TabIndex        =   1
      Top             =   7680
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstRecs 
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   11668
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
   Begin VB.Label lblRec 
      Alignment       =   1  'Right Justify
      Caption         =   "lblRec"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmRecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Const colw = 1300
    
    Dim colx As ColumnHeader
    Dim itmX As ListItem
    
    Dim IM(0 To 6) As String
    Dim Scr(0 To 4) As String
    Dim Wat(0 To 2) As String
    Dim Chem(0 To 6) As String

Private Sub Form_Load()
'зареждане на меню рецепти
    Dim Rec As Recipe
    Dim count As Integer
    
    Me.Caption = uniRecs
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    Me.btnLoad.Caption = btLoad
    Me.lblRec.Caption = uniRec
    
    If MachineNumber = 1 Then
        ns1 = n1s1
        ns2 = n1s2
        ns3 = n1s3
        ns4 = n1s4
    ElseIf MachineNumber = 2 Then
        ns1 = n2s1
        ns2 = n2s2
        ns3 = n2s3
        ns4 = n2s4
    End If
    
'почистване на таблицата
    Me.lstRecs.ColumnHeaders.Clear
    Me.lstRecs.ListItems.Clear

    Me.cmbRec.Clear
    
'настройка на заглавките на таблицата
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 700
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniNm
        colx.Width = 1600
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniRecType
        colx.Width = 1200
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniClass
        colx.Width = 1500
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniClassK
        colx.Width = 1700
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniClassV
        colx.Width = 1700
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniClassH
        colx.Width = 1700
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniClassP
        colx.Width = 1700
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniEDM
        colx.Width = 1200
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniTimePourShort
        colx.Width = 500
        
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniTimeMixShort
        colx.Width = 500
    
    IM(0) = ""
    Scr(0) = ""
    Wat(0) = ""
    Chem(0) = ""

'прочитане на имената на материалите

'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New ADODB.Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT * FROM settings_bc1 WHERE ind >= 2 AND ind <=7 ORDER BY ind;")
    
    If Not rsR.BOF And Not rsR.EOF Then rsR.MoveFirst
    
    ind = 0
    
    Do While Not rsR.EOF
        ind = ind + 1
        IM(ind) = rsR!im_num
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT * FROM settings_bc1 WHERE ind >= 8 AND ind <=11 ORDER BY ind;")
    
    If Not rsR.BOF And Not rsR.EOF Then rsR.MoveFirst
    
    ind = 0
    
    Do While Not rsR.EOF
        ind = ind + 1
        Scr(ind) = rsR!im_num
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT * FROM settings_bc1 WHERE ind >= 12 AND ind <= 13 ORDER BY ind;")
    
    If Not rsR.BOF And Not rsR.EOF Then rsR.MoveFirst
    
    ind = 0
    
    Do While Not rsR.EOF
        ind = ind + 1
        Wat(ind) = rsR!im_num
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT * FROM settings_bc1 WHERE ind >= 14 AND ind <=19 ORDER BY ind;")
    
    If Not rsR.BOF And Not rsR.EOF Then rsR.MoveFirst
    
    ind = 0
    
    Do While Not rsR.EOF
        ind = ind + 1
        Chem(ind) = rsR!im_num
        rsR.MoveNext
    Loop
    
'запис на имената на течките в таблицата
    For count = 1 To ns1
        Set colx = Me.lstRecs.ColumnHeaders.Add()
            colx.Text = IM(count)
            colx.Width = colw
    Next count
    
    For count = 1 To ns3
        Set colx = Me.lstRecs.ColumnHeaders.Add()
            colx.Text = Scr(count)
            colx.Width = colw
    Next count
    
    For count = 1 To ns2
        Set colx = Me.lstRecs.ColumnHeaders.Add()
            colx.Text = Wat(count)
            colx.Width = colw
    Next count
    
    For count = 1 To ns4
        Set colx = Me.lstRecs.ColumnHeaders.Add()
            colx.Text = Chem(count)
            colx.Width = colw
    Next count
    
    Set colx = Me.lstRecs.ColumnHeaders.Add()
        colx.Text = uniTotalKg
        colx.Width = 1200

'зареждане на рецепти
    Set Rec = New Recipe
    
    Set rsR = cnR.Execute("SELECT * FROM recepies ORDER BY r_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbRec.AddItem rsR!r_name
    
        Rec.Code = Val(rsR!r_num)
        Rec.Title = rsR!r_name
        Rec.Kind = rsR!r_type
        Rec.Class = rsR!r_class
        Rec.ClassK = rsR!r_classk
        Rec.ClassV = rsR!r_classv
        Rec.ClassH = rsR!r_classh
        Rec.ClassP = rsR!r_classp
        Rec.EDM = Val(rsR!r_edm)
        Rec.Tpour = Val(rsR!r_tpour)
        Rec.Tmix = Val(rsR!r_tmix)
        Rec.initIM(1) = Val(rsR!init_im1)
        Rec.kgIM(1) = Val(rsR!kg_im1)
        Rec.initIM(2) = Val(rsR!init_im2)
        Rec.kgIM(2) = Val(rsR!kg_im2)
        Rec.initIM(3) = Val(rsR!init_im3)
        Rec.kgIM(3) = Val(rsR!kg_im3)
        Rec.initIM(4) = Val(rsR!init_im4)
        Rec.kgIM(4) = Val(rsR!kg_im4)
        Rec.initIM(5) = Val(rsR!init_im5)
        Rec.kgIM(5) = Val(rsR!kg_im5)
        Rec.initIM(6) = Val(rsR!init_im6)
        Rec.kgIM(6) = Val(rsR!kg_im6)
        Rec.initScr(1) = Val(rsR!init_scr1)
        Rec.kgScr(1) = Val(rsR!kg_scr1)
        Rec.initScr(2) = Val(rsR!init_scr2)
        Rec.kgScr(2) = Val(rsR!kg_scr2)
        Rec.initScr(3) = Val(rsR!init_scr3)
        Rec.kgScr(3) = Val(rsR!kg_scr3)
        Rec.initScr(4) = Val(rsR!init_scr4)
        Rec.kgScr(4) = Val(rsR!kg_scr4)
        Rec.initWat(1) = Val(rsR!init_wat1)
        Rec.kgWat(1) = Val(rsR!kg_wat1)
        Rec.initWat(2) = Val(rsR!init_wat2)
        Rec.kgWat(2) = Val(rsR!kg_wat2)
        Rec.initChem(1) = Val(rsR!init_chem1)
        Rec.kgChem(1) = CSng(rDs(rsR!kg_chem1))
        Rec.initChem(2) = Val(rsR!init_chem2)
        Rec.kgChem(2) = CSng(rDs(rsR!kg_chem2))
        Rec.initChem(3) = Val(rsR!init_chem3)
        Rec.kgChem(3) = CSng(rDs(rsR!kg_chem3))
        Rec.initChem(4) = Val(rsR!init_chem4)
        Rec.kgChem(4) = CSng(rDs(rsR!kg_chem4))
        Rec.initChem(5) = Val(rsR!init_chem5)
        Rec.kgChem(5) = CSng(rDs(rsR!kg_chem5))
        Rec.initChem(6) = Val(rsR!init_chem6)
        Rec.kgChem(6) = CSng(rDs(rsR!kg_chem6))
        Rec.kgTotal = CSng(rDs(rsR!kg_total))
        
        Set itmX = Me.lstRecs.ListItems.Add(1, , Format(Rec.Code, "0000"))
            itmX.SubItems(1) = Rec.Title
            itmX.SubItems(2) = Rec.Kind
            itmX.SubItems(3) = Rec.Class
            itmX.SubItems(4) = Rec.ClassK
            itmX.SubItems(5) = Rec.ClassV
            itmX.SubItems(6) = Rec.ClassH
            itmX.SubItems(7) = Rec.ClassP
            itmX.SubItems(8) = Rec.EDM
            itmX.SubItems(9) = Rec.Tpour
            itmX.SubItems(10) = Rec.Tmix
        
        For n = 0 To ns1 - 1
            itmX.SubItems(11 + n) = Rec.AllkgIM(n + 1)
        Next n
        
        For n = 0 To ns3 - 1
            itmX.SubItems(11 + ns1 + n) = Rec.AllkgScr(n + 1)
        Next n
        
        For n = 0 To ns2 - 1
            itmX.SubItems(11 + ns1 + ns3 + n) = Rec.AllkgWat(n + 1)
        Next n
        
        For n = 0 To ns4 - 1
            itmX.SubItems(11 + ns1 + ns3 + ns2 + n) = Rec.AllkgChem(n + 1)
        Next n
        
        itmX.SubItems(11 + ns1 + ns3 + ns2 + ns4) = Rec.kgTotal
        rsR.MoveNext
    Loop
    
    Set Rec = Nothing
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    AutoColW Me.lstRecs
End Sub

Private Sub cmbRec_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbRec, KeyAscii, True)
End Sub

Private Sub btnLoad_Click()

If Me.cmbRec.Text <> "" Then
    Dim Rec As Recipe
    
    Me.lstRecs.ListItems.Clear
    
    IM(0) = ""
    Scr(0) = ""
    Wat(0) = ""
    Chem(0) = ""

'прочитане на имената на материалите

'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New ADODB.Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set Rec = New Recipe

'зареждане на рецепти
    Set rsR = cnR.Execute("SELECT * FROM recepies WHERE r_name = '" & Me.cmbRec.Text & "' ORDER BY r_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Rec.Code = Val(rsR!r_num)
        Rec.Title = rsR!r_name
        Rec.Kind = rsR!r_type
        Rec.Class = rsR!r_class
        Rec.ClassK = rsR!r_classk
        Rec.ClassV = rsR!r_classv
        Rec.ClassH = rsR!r_classh
        Rec.ClassP = rsR!r_classp
        Rec.EDM = Val(rsR!r_edm)
        Rec.Tpour = Val(rsR!r_tpour)
        Rec.Tmix = Val(rsR!r_tmix)
        Rec.initIM(1) = Val(rsR!init_im1)
        Rec.kgIM(1) = Val(rsR!kg_im1)
        Rec.initIM(2) = Val(rsR!init_im2)
        Rec.kgIM(2) = Val(rsR!kg_im2)
        Rec.initIM(3) = Val(rsR!init_im3)
        Rec.kgIM(3) = Val(rsR!kg_im3)
        Rec.initIM(4) = Val(rsR!init_im4)
        Rec.kgIM(4) = Val(rsR!kg_im4)
        Rec.initIM(5) = Val(rsR!init_im5)
        Rec.kgIM(5) = Val(rsR!kg_im5)
        Rec.initIM(6) = Val(rsR!init_im6)
        Rec.kgIM(6) = Val(rsR!kg_im6)
        Rec.initScr(1) = Val(rsR!init_scr1)
        Rec.kgScr(1) = Val(rsR!kg_scr1)
        Rec.initScr(2) = Val(rsR!init_scr2)
        Rec.kgScr(2) = Val(rsR!kg_scr2)
        Rec.initScr(3) = Val(rsR!init_scr3)
        Rec.kgScr(3) = Val(rsR!kg_scr3)
        Rec.initScr(4) = Val(rsR!init_scr4)
        Rec.kgScr(4) = Val(rsR!kg_scr4)
        Rec.initWat(1) = Val(rsR!init_wat1)
        Rec.kgWat(1) = Val(rsR!kg_wat1)
        Rec.initWat(2) = Val(rsR!init_wat2)
        Rec.kgWat(2) = Val(rsR!kg_wat2)
        Rec.initChem(1) = Val(rsR!init_chem1)
        Rec.kgChem(1) = CSng(rDs(rsR!kg_chem1))
        Rec.initChem(2) = Val(rsR!init_chem2)
        Rec.kgChem(2) = CSng(rDs(rsR!kg_chem2))
        Rec.initChem(3) = Val(rsR!init_chem3)
        Rec.kgChem(3) = CSng(rDs(rsR!kg_chem3))
        Rec.initChem(4) = Val(rsR!init_chem4)
        Rec.kgChem(4) = CSng(rDs(rsR!kg_chem4))
        Rec.initChem(5) = Val(rsR!init_chem5)
        Rec.kgChem(5) = CSng(rDs(rsR!kg_chem5))
        Rec.initChem(6) = Val(rsR!init_chem6)
        Rec.kgChem(6) = CSng(rDs(rsR!kg_chem6))
        Rec.kgTotal = CSng(rDs(rsR!kg_total))
        
        Set itmX = Me.lstRecs.ListItems.Add(1, , Format(Rec.Code, "0000"))
            itmX.SubItems(1) = Rec.Title
            itmX.SubItems(2) = Rec.Kind
            itmX.SubItems(3) = Rec.Class
            itmX.SubItems(4) = Rec.ClassK
            itmX.SubItems(5) = Rec.ClassV
            itmX.SubItems(6) = Rec.ClassH
            itmX.SubItems(7) = Rec.ClassP
            itmX.SubItems(8) = Rec.EDM
            itmX.SubItems(9) = Rec.Tpour
            itmX.SubItems(10) = Rec.Tmix
        
        
        For n = 0 To ns1 - 1
            itmX.SubItems(11 + n) = Rec.AllkgIM(n + 1)
        Next n
        
        For n = 0 To ns3 - 1
            itmX.SubItems(11 + ns1 + n) = Rec.AllkgScr(n + 1)
        Next n
        
        For n = 0 To ns2 - 1
            itmX.SubItems(11 + ns1 + ns3 + n) = Rec.AllkgWat(n + 1)
        Next n
        
        For n = 0 To ns4 - 1
            itmX.SubItems(11 + ns1 + ns3 + ns2 + n) = Rec.AllkgChem(n + 1)
        Next n
        
        itmX.SubItems(11 + ns1 + ns3 + ns4 + ns2) = Rec.kgTotal
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    AutoColW Me.lstRecs
Else
    Call Form_Load
End If
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstRecs, 2, True, True, True, uniRecs)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstRecs)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

