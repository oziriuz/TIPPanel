VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperWork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmOperWork"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   Icon            =   "frmOperWork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   7695
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
      Left            =   6600
      TabIndex        =   10
      Top             =   7800
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
      Left            =   4080
      TabIndex        =   6
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
      Left            =   1320
      TabIndex        =   5
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
      Left            =   4440
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.ComboBox cmbOper 
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
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstOperWork 
      Height          =   5655
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9975
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
      Left            =   4080
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
      Format          =   103481347
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
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
      Format          =   103481347
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
      Left            =   5880
      TabIndex        =   9
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
      Left            =   4080
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblOper 
      Alignment       =   2  'Center
      Caption         =   "lblOper"
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
      TabIndex        =   7
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmOperWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за 1 или всички (ако няма избран) оператори - филтър по име от лог таблицата
'всеки ред от справката изобразява един запис от зададен период от формата
    
    Me.Caption = uniLog
    Me.lblOper.Caption = uniDisp
    Me.lblStart.Caption = lblStDate
    Me.lblEnd.Caption = lblEndDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    
    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False
    
    Me.dtStart = Now
    Me.dtEnd = Now
    
    Me.lstOperWork.ColumnHeaders.Clear
    Me.lstOperWork.ListItems.Clear
    
    Me.cmbOper.Clear
    
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
    
    Set colx = Me.lstOperWork.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstOperWork.ColumnHeaders.Add()
        colx.Text = uniNm
        colx.Width = 1000
    
    Set colx = Me.lstOperWork.ColumnHeaders.Add()
        colx.Text = UniEnter
        colx.Width = 1000
    
    Set colx = Me.lstOperWork.ColumnHeaders.Add()
        colx.Text = UniExit
        colx.Width = 1000

    AutoColW Me.lstOperWork
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    If MachineNumber = 1 Then
        Set rsR = cnR.Execute("SELECT DISTINCT ON (log_name) log_name FROM entry_log ORDER BY log_name ASC;")
    ElseIf MachineNumber = 2 Then
        Set rsR = cnR.Execute("SELECT DISTINCT ON (log_name) log_name FROM entry_log2 ORDER BY log_name ASC;")
    End If
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbOper.AddItem rsR!log_name
        rsR.MoveNext
    Loop
    
    If frmStartRep.chMach1.Value = 1 And frmStartRep.chMach2.Value = 1 Then
        Set rsR = cnR.Execute("SELECT DISTINCT ON (log_name) log_name FROM entry_log2 ORDER BY log_name ASC;")
    
        If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
        Do While Not rsR.EOF
            Me.cmbOper.AddItem rsR!log_name
            rsR.MoveNext
        Loop
    End If
    
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------

    Dim i, j As Integer
        
    For i = 0 To Me.cmbOper.listCount - 2 Step 1
        For j = Me.cmbOper.listCount - 1 To i + 1 Step -1
            If Me.cmbOper.List(i) = Me.cmbOper.List(j) Then
                Me.cmbOper.RemoveItem (j)
            End If
        Next
    Next

End Sub

Private Sub cmbOper_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbOper, KeyAscii, True)
EndSub:
End Sub

Private Sub btnLoad_Click()

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

    Me.lstOperWork.ListItems.Clear
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
    Dim commR As String
    Dim rcCounter As Integer
    Dim DayStart As String
    Dim DayEnd As String
    
    DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
    DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr

AgainOther:

    MousePointer = vbHourglass

    If MachineNumber = 1 Then
        If Me.cmbOper.Text <> "" Then
            'маркираме набор от записи ако има избрано име
            commR = "SELECT log_name, log_enter, log_exit FROM entry_log WHERE log_name = '" _
            & Me.cmbOper.Text & "' AND log_enter_date >= '" & DayStart _
            & "' AND log_enter_date <= '" & DayEnd _
            & "' ORDER BY log_num ASC;"
        Else
            'маркираме набор от записи ако няма избрано име
            commR = "SELECT log_name, log_enter, log_exit FROM entry_log WHERE  log_enter_date >= '" _
            & DayStart & "' AND log_enter_date <= '" & DayEnd _
            & "' ORDER BY log_num ASC;"
        End If
    ElseIf MachineNumber = 2 Then
        If Me.cmbOper.Text <> "" Then
            'маркираме набор от записи ако има избрано име
            commR = "SELECT log_name, log_enter, log_exit FROM entry_log2 WHERE log_name = '" _
            & Me.cmbOper.Text & "' AND log_enter_date >= '" & DayStart _
            & "' AND log_enter_date <= '" & DayEnd _
            & "' ORDER BY log_num ASC;"
        Else
            'маркираме набор от записи ако няма избрано име
            commR = "SELECT log_name, log_enter, log_exit FROM entry_log2 WHERE  log_enter_date >= '" _
            & DayStart & "' AND log_enter_date <= '" & DayEnd _
            & "' ORDER BY log_num ASC;"
        End If
    End If
    
    Set rsR = cnR.Execute(commR)
    
    'отиваме на първия запис
    If Not rsR.EOF And Not rsR.BOF Then
        rsR.MoveFirst
        Me.btnPrint.Enabled = True
        Me.btnExport.Enabled = True
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperWork.ListItems.Add(rcCounter, , "")
            itmX.SubItems(1) = "Машина " & MachineNumber
        have = have + 1
        If ag = False Then firstM = True
    Else
        If ag = True Then
            If firstM = False Then
                Me.btnPrint.Enabled = False
                Me.btnExport.Enabled = False
                MousePointer = vbDefault
                MsgBox MsgNoRecords, vbOKOnly Or vbInformation, MsgErrNoRec

                rsR.Close
                Set rsR = Nothing
                cnR.Close
                Set cnR = Nothing
                GoTo EndSub
            Else
                MousePointer = vbDefault
                rsR.Close
                Set rsR = Nothing
                cnR.Close
                Set cnR = Nothing
                GoTo EndSub
            End If
        Else
            If frmStartRep.chMach2.Value = 1 Then
                rsR.Close
                Set rsR = Nothing
                ag = True
                MachineNumber = 2
                ind = 0
                GoTo AgainOther
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
        End If
    End If
    
    If ag = False Then rcCounter = 1 'нулираме брояча на редовете в ListView
    
    Do While Not rsR.EOF
        rcCounter = rcCounter + 1 'първи ред ++
        ind = ind + 1
        Set itmX = Me.lstOperWork.ListItems.Add(rcCounter, , Format(ind, "0000"))
            If rsR!log_name <> "Null" Then itmX.SubItems(1) = rsR!log_name
            If rsR!log_enter <> "Null" Then itmX.SubItems(2) = rsR!log_enter
            If rsR!log_exit <> "Null" Then itmX.SubItems(3) = rsR!log_exit
            
        rsR.MoveNext 'местим на следващ запис
    Loop
    
    If ag = False And MachineNumber = 1 And frmStartRep.chMach2.Value = 1 Then
        MachineNumber = 2
        ag = True
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperWork.ListItems.Add(rcCounter, , "")
        ind = 0
        GoTo AgainOther
    End If

    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close 'затваряме връзката
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------
EndSub:
    AutoColW Me.lstOperWork
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstOperWork, 1, True, True, True, uniLog & "  (" & Me.cmbOper & ")")
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstOperWork)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

