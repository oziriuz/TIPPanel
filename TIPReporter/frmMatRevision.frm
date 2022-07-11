VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMatRevision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMatRevision"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9270
   Icon            =   "frmMatRevision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   9270
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
      Left            =   8160
      TabIndex        =   8
      Top             =   8520
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
      Left            =   5160
      TabIndex        =   5
      Top             =   8520
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
      Left            =   1800
      TabIndex        =   4
      Top             =   8520
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
      Left            =   6000
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin MSComctlLib.ListView lstMatRevision 
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   12515
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
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   375
      Left            =   360
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
      Left            =   2505
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
      Left            =   2475
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
      Left            =   345
      TabIndex        =   6
      Top             =   255
      Width           =   1455
   End
End
Attribute VB_Name = "frmMatRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Me.Caption = uniRevisions
    Me.lblStart.Caption = lblStDate
    Me.lblEnd.Caption = lblEndDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport

    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False
    
    Me.dtStart = Now
    Me.dtEnd = Now
    
    Me.lstMatRevision.ColumnHeaders.Clear
    Me.lstMatRevision.ListItems.Clear
    
    Set colx = Me.lstMatRevision.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstMatRevision.ColumnHeaders.Add()
        colx.Text = uniDate
        colx.Width = 1000
    
    Set colx = Me.lstMatRevision.ColumnHeaders.Add()
        colx.Text = uniMat
        colx.Width = 1000
    
    Set colx = Me.lstMatRevision.ColumnHeaders.Add()
        colx.Text = uniOld & " " & uniQ & " [t]"
        colx.Width = 1000

    Set colx = Me.lstMatRevision.ColumnHeaders.Add()
        colx.Text = uniNew & " " & uniQ & " [t]"
        colx.Width = 1000

    Set colx = Me.lstMatRevision.ColumnHeaders.Add()
        colx.Text = uniDisp
        colx.Width = 1000

    Set colx = Me.lstMatRevision.ColumnHeaders.Add()
        colx.Text = uniRevisor
        colx.Width = 1000

    AutoColW Me.lstMatRevision
    
End Sub

Private Sub btnLoad_Click()

    Me.lstMatRevision.ListItems.Clear
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
    Dim commR As String
    Dim tempRev As Integer
    Dim ind As Integer
    Dim frstFlag As Boolean
    Dim DayStart As String
    Dim DayEnd As String
    
    DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
    DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме набор от записи ако има избрано име
    commR = "SELECT * FROM revision WHERE stamp_date >= '" & DayStart _
    & "' AND stamp_date <= '" & DayEnd _
    & "' ORDER BY row_num ASC;"
        
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
    
    frstFlag = True
        
    Do While Not rsR.EOF
        If frstFlag = False And tempRev <> rsR!rev_num Then
            ind = ind + 1
            Set itmX = Me.lstMatRevision.ListItems.Add(ind, , "")
        End If
        tempRev = rsR!rev_num
        frstFlag = False
        ind = ind + 1
        Set itmX = Me.lstMatRevision.ListItems.Add(ind, , Format(rsR!rev_num, "0000"))
            itmX.SubItems(1) = rsR!rev_date
            itmX.SubItems(2) = rsR!rev_matname
            itmX.SubItems(3) = CSng(rDs(rsR!rev_matqold))
            itmX.SubItems(4) = CSng(rDs(rsR!rev_matqnew))
            itmX.SubItems(5) = rsR!rev_op
            itmX.SubItems(6) = rsR!rev_supervisor
            
        rsR.MoveNext
    Loop
    
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close 'затваряме връзката
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------
    
    AutoColW Me.lstMatRevision
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstMatRevision, 1, True, True, True, uniRevisions)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstMatRevision)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

