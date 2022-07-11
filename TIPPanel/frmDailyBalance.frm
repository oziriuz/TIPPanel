VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDailyBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDailyBalance"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   15510
   StartUpPosition =   1  'CenterOwner
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
      Left            =   8640
      TabIndex        =   1
      Top             =   4680
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
      TabIndex        =   0
      Top             =   4680
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstBalance 
      Height          =   3615
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   6376
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
End
Attribute VB_Name = "frmDailyBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstBalance)
End Sub

Private Sub Form_Load()
'зареждане на дневен отчет
    
    Dim MachineOther            As String
    Dim itmX                    As MSComctlLib.ListItem
    Dim colx                    As MSComctlLib.ColumnHeader
    Dim materialNames(1 To 100) As String
    Dim materialToday(1 To 100) As Single
    Dim materialNow(1 To 100)   As Single
    Dim materialPlus(1 To 100)  As Single
    Dim materialStart(1 To 100) As Single
    Dim cnBal                   As ADODB.Connection
    Dim rsBal                   As Recordset
    Dim rsBalOther              As Recordset
    Dim comm                    As String
    Dim i                       As Integer
    Dim icount                  As Integer
    
    Me.Caption = "Дневен отчет - складова наличност"
    Me.btnPrint.Caption = "Печат"
    Me.btnExport.Caption = "Експорт"
    
    If MachineNumber = 1 Then MachineOther = 2
    If MachineNumber = 2 Then MachineOther = 1

    Me.lstBalance.ColumnHeaders.Clear
    Me.lstBalance.ListItems.Clear
    
    Set colx = Me.lstBalance.ColumnHeaders.Add()
        colx.Text = "Материали"
        colx.Width = 800
        
'------------------------------Start PostgreSQL----------------------------------
    Set cnBal = New ADODB.Connection 'връзка с база данни
        cnBal.ConnectionTimeout = 10
        cnBal.Open ConStr 'отваряме връзката
        
    MousePointer = vbHourglass
    
    comm = "SELECT * FROM daily_expenses WHERE stamp_date = '" & Format(Now, "DD-MM-YYYY") & "' ORDER BY row_num ASC;"
    
    Set rsBal = cnBal.Execute(comm) 'маркираме всички замеси от намерената експедиция
    If Not rsBal.BOF And Not rsBal.EOF Then
        rsBal.MoveFirst 'отиваме в началото на маркираните замеси
    Else
        GoTo EndSub
    End If

    'изчитане на дневния разход от базата данни
    i = 1
    Do While Not rsBal.EOF
        Set colx = Me.lstBalance.ColumnHeaders.Add()
            colx.Text = rsBal!mat_name & " [t]"
            colx.Width = 3000
            materialNames(i) = rsBal!mat_name
            materialToday(i) = CSng(rDs(rsBal!mat_sold))
        rsBal.MoveNext
        i = i + 1
    Loop
    
    icount = i - 1
    
    'определяне на крайното салдо към момента
    Set itmX = Me.lstBalance.ListItems.Add(1, , "Крайно салдо") 'запис в ListView
    For i = 1 To icount
        Set rsBal = cnBal.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & materialNames(i) & "';")
        Set rsBalOther = cnBal.Execute("SELECT * FROM materials_bc" & MachineOther & " WHERE m_name = '" & materialNames(i) & "';")
        If Not rsBal.EOF And Not rsBal.BOF Then rsBal.MoveFirst
        If Not rsBalOther.EOF And Not rsBalOther.BOF Then
            rsBalOther.MoveFirst
            materialNow(i) = CSng(rDs(rsBal!m_del)) - CSng(rDs(rsBal!m_sold)) - CSng(rDs(rsBalOther!m_sold))
            If rsBal!m_type = "3" Then
                materialNow(i) = ARound(materialNow(i), 5)
            Else
                materialNow(i) = ARound(materialNow(i), 3)
            End If
            itmX.SubItems(i) = materialNow(i)
        Else
            materialNow(i) = 0
            itmX.SubItems(i) = materialNow(i)
        End If
    Next i
    
    rsBal.Close 'затваряме записите
    rsBalOther.Close 'затваряме записите
        
    'визуализиране на дневния разход
    Set itmX = Me.lstBalance.ListItems.Add(1, , "Разход") 'запис в ListView
    For i = 1 To icount
        itmX.SubItems(i) = materialToday(i)
    Next i
    
    'определяне на приходите
    Set itmX = Me.lstBalance.ListItems.Add(1, , "Приход") 'запис в ListView
    For i = 1 To icount
        Set rsBal = cnBal.Execute("SELECT * FROM deliveries WHERE del_mat = '" & materialNames(i) & "' AND stamp_date = '" & Format(Now, "DD-MM-YYYY") & "';")
        If Not rsBal.EOF And Not rsBal.BOF Then
            rsBal.MoveFirst
            materialPlus(i) = CSng(rDs(rsBal!del_q))
            itmX.SubItems(i) = materialPlus(i)
        Else
            materialPlus(i) = 0
            itmX.SubItems(i) = materialPlus(i)
        End If
    Next i
    
    rsBal.Close 'затваряме записите
    Set rsBal = Nothing
    cnBal.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnBal = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------

    'определяне на началното салдо
    Set itmX = Me.lstBalance.ListItems.Add(1, , "Начално салдо") 'запис в ListView
    For i = 1 To icount
        materialStart(i) = materialNow(i) - materialPlus(i) + materialToday(i)
        itmX.SubItems(i) = materialStart(i)
    Next i

    AutoColW Me.lstBalance
    
    GoTo EndNormal
EndSub:
    MsgBox "Няма информация за днес!", vbOKOnly Or vbInformation, "Няма данни"
    MousePointer = vbDefault
    ErrDaily = True
    Exit Sub
EndNormal:
    ErrDaily = False
End Sub

Private Sub btnPrint_Click()

    Call PrintLVPic(Me.lstBalance, 2, True, True, False, "Дневен отчет - складова наличност")
End Sub

