VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliveries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDeliveries"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10695
   Icon            =   "frmDeliveries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10695
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
      Left            =   9600
      TabIndex        =   12
      Top             =   7800
      Width           =   735
   End
   Begin VB.ComboBox cmbSup 
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
      TabIndex        =   9
      Top             =   600
      Width           =   3255
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
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   3015
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
      TabIndex        =   5
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
      Left            =   2520
      TabIndex        =   4
      Top             =   7800
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
      Left            =   7080
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstDeliveries 
      Height          =   5655
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9975
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
      Left            =   7080
      TabIndex        =   1
      Top             =   600
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
      Format          =   7471107
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   600
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
      Format          =   7471107
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin VB.Label lblSup 
      Alignment       =   2  'Center
      Caption         =   "lblSup"
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
      TabIndex        =   11
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblMat 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   10
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
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
      Left            =   8880
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmDeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Me.Caption = uniDlvrs
    Me.lblSup.Caption = uniSup
    Me.lblMat.Caption = uniMat
    Me.lblStart.Caption = lblStDate
    Me.lblEnd.Caption = lblEndDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
            
    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False

    Me.dtStart = Now
    Me.dtEnd = Now
    
    Me.lstDeliveries.ColumnHeaders.Clear
    Me.lstDeliveries.ListItems.Clear
    
    Me.cmbSup.Clear
    Me.cmbMat.Clear
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniDate
        colx.Width = 1200
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniSup
        colx.Width = 1200
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniTypeDoc
        colx.Width = 1000
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniNr
        colx.Width = 1000
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniMat
        colx.Width = 1000
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniQ & " [t]"
        colx.Width = 1000
    
    Set colx = Me.lstDeliveries.ColumnHeaders.Add()
        colx.Text = uniDisp
        colx.Width = 1000
    
    AutoColW Me.lstDeliveries
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (del_sup_name) del_sup_name FROM deliveries ORDER BY del_sup_name ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbSup.AddItem rsR!del_sup_name
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (del_mat) del_mat FROM deliveries ORDER BY del_mat ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbMat.AddItem rsR!del_mat
        rsR.MoveNext
    Loop
    
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------
   
End Sub

Private Sub cmbMat_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbMat, KeyAscii, True)
End Sub

Private Sub cmbSup_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbSup, KeyAscii, True)
End Sub

Private Sub btnLoad_Click()

Me.lstDeliveries.ListItems.Clear

'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
    Dim commR As String
    Dim DayStart As String
    Dim DayEnd As String
    
    DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
    DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме набор от записи
    'основен стринг без филтри
    commR = "SELECT * FROM deliveries WHERE stamp_date >= '" & DayStart _
    & "' AND stamp_date <= '" & DayEnd & "'"
    
    'стринг за добавка при търсенето за доставчик
    commR1 = " AND del_sup_name = '" & Me.cmbSup.Text & "'"
    
    'стринг за добавка при търсенето за материал
    commR2 = " AND del_mat = '" & Me.cmbMat.Text & "'"
    
    'стринг за сортиране по номер
    commRord = " ORDER BY del_num ASC;"
    
    'резултен стринг за търсенето при добавен филтър за обект
    If Me.cmbSup.Text <> "" Then commR = commR & commR1
    
    'резултен стринг за търсенето при добавен филтър за обект
    If Me.cmbMat.Text <> "" Then commR = commR & commR2
    
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
        Set itmX = Me.lstDeliveries.ListItems.Add(1, , rsR!del_num)
            itmX.SubItems(1) = rsR!del_date
            itmX.SubItems(2) = rsR!del_sup_name
            itmX.SubItems(3) = rsR!del_doc_type
            itmX.SubItems(4) = rsR!del_doc_num
            itmX.SubItems(5) = rsR!del_mat
            itmX.SubItems(6) = CSng(rDs(rsR!del_q))
            itmX.SubItems(7) = rsR!del_op
            
        rsR.MoveNext
    Loop
    
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------

    AutoColW Me.lstDeliveries
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstDeliveries, 2, True, True, True, uniDlvrs)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstDeliveries)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

