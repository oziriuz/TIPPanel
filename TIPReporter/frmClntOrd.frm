VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClntOrd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmClntOrd"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13215
   Icon            =   "frmClntOrd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   13215
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
      Left            =   12120
      TabIndex        =   16
      Top             =   8160
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
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Width           =   3255
   End
   Begin VB.ComboBox cmbObj 
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
      Left            =   3960
      TabIndex        =   10
      Top             =   600
      Width           =   3255
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
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox cmbClass 
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
      Left            =   10680
      TabIndex        =   8
      Top             =   600
      Width           =   2175
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
      Left            =   7560
      TabIndex        =   5
      Top             =   8160
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
      Left            =   3360
      TabIndex        =   4
      Top             =   8160
      Width           =   2295
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
      Left            =   9600
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstClntOrd 
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   9763
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
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   65535
      CalendarForeColor=   -2147483639
      CustomFormat    =   "dd.MM.yyy"
      Format          =   112459779
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   65535
      CalendarForeColor=   -2147483639
      CustomFormat    =   "dd.MM.yyy"
      Format          =   112459779
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
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
      Left            =   2280
      TabIndex        =   17
      Top             =   7680
      Width           =   8895
   End
   Begin VB.Label lblClnt 
      Alignment       =   2  'Center
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
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblObj 
      Alignment       =   2  'Center
      Caption         =   "lblObj"
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
      Left            =   4080
      TabIndex        =   14
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
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
      Left            =   7680
      TabIndex        =   13
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      Caption         =   "lblClass"
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
      Left            =   10800
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblEnd 
      Alignment       =   1  'Right Justify
      Caption         =   "lblEnd"
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
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblStart 
      Alignment       =   1  'Right Justify
      Caption         =   "lblStart"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmClntOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Me.Caption = repClntOrd
    Me.lblClnt.Caption = uniClnt
    Me.lblObj.Caption = uniObj
    Me.lblRec.Caption = uniRec
    Me.lblClass.Caption = uniClass
    Me.lblStart.Caption = lblStDate
    Me.lblEnd.Caption = lblEndDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    
    Me.lblNote.Caption = "Заявките се зареждат за двете машини общо!"
    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False
    
    Me.dtStart = Now
    Me.dtEnd = Now
    
    If frmStartRep.chMach1.Value = 1 Then
        MachineNumber = 1
    Else
        If frmStartRep.chMach2.Value = 1 Then
            MachineNumber = 2
        Else
            frmStartRep.chMach1.Value = 1
            MachineNumber = 1
        End If
    End If
    
    Me.lstClntOrd.ColumnHeaders.Clear
    Me.lstClntOrd.ListItems.Clear
    
    Me.cmbClnt.Clear
    Me.cmbObj.Clear
    Me.cmbRec.Clear
    Me.cmbClass.Clear
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniDate
        colx.Width = 1200
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniClnt
        colx.Width = 1200
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniObj
        colx.Width = 1000
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniRec & " " & uniNm
        colx.Width = 1000
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniClass
        colx.Width = 1000
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniOrdered & " " & uniQ
        colx.Width = 1000
    
    Set colx = Me.lstClntOrd.ColumnHeaders.Add()
        colx.Text = uniMade & " " & uniQ
        colx.Width = 1000
    
    AutoColW Me.lstClntOrd
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (order_clnt_name) order_clnt_name FROM orders ORDER BY order_clnt_name ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbClnt.AddItem rsR!order_clnt_name
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (order_clnt_obj) order_clnt_obj FROM orders ORDER BY order_clnt_obj ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbObj.AddItem rsR!order_clnt_obj
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (order_rec_name) order_rec_name FROM orders ORDER BY order_rec_name ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbRec.AddItem rsR!order_rec_name
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (order_rec_class) order_rec_class FROM orders ORDER BY order_rec_class ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbClass.AddItem rsR!order_rec_class
        rsR.MoveNext
    Loop
    
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------
   
    Dim i, j As Integer
        
    For i = 0 To Me.cmbClnt.listCount - 2 Step 1
        For j = Me.cmbClnt.listCount - 1 To i + 1 Step -1
            If Me.cmbClnt.List(i) = Me.cmbClnt.List(j) Then
                Me.cmbClnt.RemoveItem (j)
            End If
        Next
    Next
        
    For i = 0 To Me.cmbObj.listCount - 2 Step 1
        For j = Me.cmbObj.listCount - 1 To i + 1 Step -1
            If Me.cmbObj.List(i) = Me.cmbObj.List(j) Then
                Me.cmbObj.RemoveItem (j)
            End If
        Next
    Next
    
    For i = 0 To Me.cmbRec.listCount - 2 Step 1
        For j = Me.cmbRec.listCount - 1 To i + 1 Step -1
            If Me.cmbRec.List(i) = Me.cmbRec.List(j) Then
                Me.cmbRec.RemoveItem (j)
            End If
        Next
    Next

    For i = 0 To Me.cmbClass.listCount - 2 Step 1
        For j = Me.cmbClass.listCount - 1 To i + 1 Step -1
            If Me.cmbClass.List(i) = Me.cmbClass.List(j) Then
                Me.cmbClass.RemoveItem (j)
            End If
        Next
    Next
   
End Sub

Private Sub cmbClnt_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbClnt, KeyAscii, True)
EndSub:
End Sub

Private Sub cmbObj_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbObj, KeyAscii, True)
EndSub:
End Sub

Private Sub cmbRec_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbRec, KeyAscii, True)
EndSub:
End Sub

Private Sub cmbClass_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbClass, KeyAscii, True)
EndSub:
End Sub

Private Sub btnLoad_Click()

Me.lstClntOrd.ListItems.Clear

'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
    Dim commR As String
    Dim DayStart As String
    Dim DayEnd As String
    Dim rcCounter As Integer
    
    DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
    DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    rcCounter = 0
    
    'маркираме набор от записи
    'основен стринг без филтри
    commR = "SELECT order_num, order_date, order_clnt_name, order_clnt_obj, order_rec_name, order_rec_class, order_q, order_qmade FROM orders WHERE stamp_date >= '" & DayStart _
    & "' AND stamp_date <= '" & DayEnd & "'"
    
    'стринг за добавка при търсенето за обект
    commR1 = " AND order_clnt_name = '" & Me.cmbClnt.Text & "'"
    
    'стринг за добавка при търсенето за обект
    commR2 = " AND order_clnt_obj = '" & Me.cmbObj.Text & "'"
    
    'стринг за добавка при търсенето за име рецепта
    commR3 = " AND order_rec_name = '" & Me.cmbRec.Text & "'"
    
    'стринг за добавка при търсенето за клас якост
    commR4 = " AND order_rec_class = '" & Me.cmbClass.Text & "'"
    
    'стринг за сортиране по номер
    commRord = " ORDER BY order_num ASC;"
    
    'резултен стринг за търсенето при добавен филтър за обект
    If Me.cmbClnt.Text <> "" Then commR = commR & commR1
    
    'резултен стринг за търсенето при добавен филтър за обект
    If Me.cmbObj.Text <> "" Then commR = commR & commR2
    
    'резултен стринг за търсенето при добавен филтър за име рецепта
    If Me.cmbRec.Text <> "" Then commR = commR & commR3
    
    'резултен стринг за търсенето при добавен филтър за клас на якост
    If Me.cmbClass.Text <> "" Then commR = commR & commR4
    
    'добавка на стринга за сортиране
    commR = commR + commRord
        
    Set rsR = cnR.Execute(commR)
    
    'отиваме на първия запис
    If Not rsR.EOF And Not rsR.BOF Then
        rsR.MoveFirst
        Me.btnPrint.Enabled = True
        Me.btnExport.Enabled = True
    Else
        Me.btnPrint.Enabled = False
        Me.btnExport.Enabled = False
        MousePointer = vbDefault
        MsgBox MsgNoRecords, vbOKOnly Or vbInformation, MsgErrNoRec
            
        rsR.Close
        Set rsR = Nothing
        cnR.Close
        Set cnR = Nothing
        Exit Sub
    End If

    Do While Not rsR.EOF
        rcCounter = rcCounter + 1
        Set itmX = Me.lstClntOrd.ListItems.Add(1, , Format(rsR!order_num, "0000000"))
            itmX.SubItems(1) = rsR!order_date
            itmX.SubItems(2) = rsR!order_clnt_name
            itmX.SubItems(3) = rsR!order_clnt_obj
            itmX.SubItems(4) = rsR!order_rec_name
            itmX.SubItems(5) = rsR!order_rec_class
            itmX.SubItems(6) = CSng(rDs(rsR!order_q))
            itmX.SubItems(7) = CSng(rDs(rsR!order_qmade))
            
        rsR.MoveNext
    Loop
    
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------

     'след като прекъснем връзката сумираме тоталите
    Dim i As Integer
    Dim tOrdVol As Single
    Dim tOrdMadeVol As Single
    
    tOrdVol = 0
    tOrdMadeVol = 0
    
    For i = 1 To Me.lstClntOrd.ListItems.count
        tOrdVol = tOrdVol + CSng(rDs(Me.lstClntOrd.ListItems.Item(i).SubItems(6)))
        tOrdMadeVol = tOrdMadeVol + CSng(rDs(Me.lstClntOrd.ListItems.Item(i).SubItems(7)))
    Next i
    
    'първо въвеждаме един празен ред
    Set itmX = Me.lstClntOrd.ListItems.Add(1, , "X")
    
    'след него въвеждаме тотали
    Set itmX = Me.lstClntOrd.ListItems.Add(1, , "XXXX")
        itmX.SubItems(1) = uniTotal & ": " & rcCounter
        itmX.SubItems(6) = tOrdVol
        itmX.SubItems(7) = tOrdMadeVol
   
    AutoColW Me.lstClntOrd
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstClntOrd, 2, True, True, True, repClntOrd)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstClntOrd)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

