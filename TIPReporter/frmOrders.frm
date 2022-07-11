VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmOrders"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13830
   Icon            =   "frmOrders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13830
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
      Left            =   12720
      TabIndex        =   3
      Top             =   7680
      Width           =   735
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
      Left            =   7920
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
      Left            =   3600
      TabIndex        =   1
      Top             =   7680
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstOrders 
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12515
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
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
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'зареждане на меню заявки

    Me.Caption = uniOrds
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport

    Dim colx As ColumnHeader
    Dim itmX As ListItem
    
    Me.lstOrders.ColumnHeaders.Clear
    Me.lstOrders.ListItems.Clear
    
    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 1000
    
    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniDateOrd
        colx.Width = 1750
    
    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniDateReady
        colx.Width = 1750
    
    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniOrdered
        colx.Width = 1100
        colx.Tag = "number"
    
    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniMade
        colx.Width = 1100

    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniRec & " " & uniCode
        colx.Width = 1200

    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniNm & " " & uniRec
        colx.Width = 1200
    
    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniClass
        colx.Width = 1300
    
    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniClnt & " " & uniCode
        colx.Width = 1100

    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniNm & " " & uniClnt
        colx.Width = 1300

    Set colx = Me.lstOrders.ColumnHeaders.Add()
        colx.Text = uniObj
        colx.Width = 1300
        
'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
'зареждане на всички заявки в ListView
    Set rsR = cnR.Execute("SELECT * FROM orders ORDER BY order_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Set itmX = Me.lstOrders.ListItems.Add(1, , Format(rsR!order_num, "0000000"))
            itmX.SubItems(1) = rsR!order_date
            itmX.SubItems(2) = rsR!order_date_que
            itmX.SubItems(3) = rDs(rsR!order_q)
            itmX.SubItems(4) = rDs(rsR!order_qmade)
            itmX.SubItems(5) = Format(rsR!order_rec, "0000")
            itmX.SubItems(6) = rsR!order_rec_name
            itmX.SubItems(7) = rsR!order_rec_class
            itmX.SubItems(8) = Format(rsR!order_clnt, "0000")
            itmX.SubItems(9) = rsR!order_clnt_name
            itmX.SubItems(10) = rsR!order_clnt_obj
        
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------

    AutoColW Me.lstOrders
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstOrders, 2, True, True, True, uniOrds)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstOrders)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

