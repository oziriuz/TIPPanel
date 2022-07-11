VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMatSold 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMatSold"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   Icon            =   "frmMatSold.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
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
      Top             =   7920
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
      Left            =   3960
      TabIndex        =   6
      Top             =   7920
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
      Left            =   1440
      TabIndex        =   5
      Top             =   7920
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
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstMatSold 
      Height          =   5775
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10186
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
      Format          =   47316995
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
      Format          =   47316995
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
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmMatSold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Me.Caption = repMatSold
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
    
    Me.lstMatSold.ColumnHeaders.Clear
    Me.lstMatSold.ListItems.Clear
    
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
    
    Me.cmbMat.Clear
    
    Set colx = Me.lstMatSold.ColumnHeaders.Add()
        colx.Text = uniNr
        colx.Width = 800
    
    Set colx = Me.lstMatSold.ColumnHeaders.Add()
        colx.Text = uniDate
        colx.Width = 1000
    
    Set colx = Me.lstMatSold.ColumnHeaders.Add()
        colx.Text = uniMat
        colx.Width = 1000
    
    Set colx = Me.lstMatSold.ColumnHeaders.Add()
        colx.Text = uniSold
        colx.Width = 1000

    AutoColW Me.lstMatSold
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (m_name) m_name FROM materials_bc" & MachineNumber & " ORDER BY m_name ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbMat.AddItem rsR!m_name
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

    Me.lstMatSold.ListItems.Clear
    
    If Me.cmbMat.Text <> "" Then
    
'-----------------------Start postgreSQL-----------------------------------
        Dim cnR As New ADODB.Connection
        Dim rsR As New Recordset
        Dim commR As String
        Dim rcCounter As Integer
        Dim SubTotal As Single
        Dim SumTotal As Single
        Dim chDate As String
        Dim tempDate As String
        Dim frstFlag As Boolean
        Dim DayStart As String
        Dim DayEnd As String
    
        DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
        DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
AgainOther:
    
        cnR.ConnectionTimeout = 30
        cnR.Open ConStr
        MousePointer = vbHourglass
    
        'маркираме набор от записи ако има избрано име
        commR = "SELECT stamp_date, im1_name, im1i, im2_name, im2i, im3_name, im3i, im4_name, im4i, im5_name, im5i, im6_name, im6i, cem1_name, cem1i, cem2_name, cem2i, cem3_name, cem3i, cem4_name, cem4i, wat1_name, wat1i, wat2_name, wat2i, chem1_name, chem1i, chem2_name, chem2i, chem3_name, chem3i, chem4_name, chem4i, chem5_name, chem5i, chem6_name, chem6i FROM mix_result_bc" & MachineNumber & " WHERE stamp_date >= '" & DayStart _
        & "' AND stamp_date <= '" & DayEnd _
        & "' ORDER BY mix_num ASC;"
        
        Set rsR = cnR.Execute(commR)
    
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            Me.btnPrint.Enabled = True
            Me.btnExport.Enabled = True
            rcCounter = rcCounter + 1
            Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "")
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
                    cnR.Close
                    Set cnR = Nothing
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
        
        SubTotal = 0
        SumTotal = 0
        frstFlag = True
        
        Do While Not rsR.EOF
            Select Case Me.cmbMat.Text
            
                Case rsR!im1_name
                    SubTotal = SubTotal + CSng(rDs(rsR!im1i))
            
                Case rsR!im2_name
                    SubTotal = SubTotal + CSng(rDs(rsR!im2i))
            
                Case rsR!im3_name
                    SubTotal = SubTotal + CSng(rDs(rsR!im3i))
            
                Case rsR!im4_name
                    SubTotal = SubTotal + CSng(rDs(rsR!im4i))
            
                Case rsR!im5_name
                    SubTotal = SubTotal + CSng(rDs(rsR!im5i))
            
                Case rsR!im6_name
                    SubTotal = SubTotal + CSng(rDs(rsR!im6i))
            
                Case rsR!cem1_name
                    SubTotal = SubTotal + CSng(rDs(rsR!cem1i))
            
                Case rsR!cem2_name
                    SubTotal = SubTotal + CSng(rDs(rsR!cem2i))
            
                Case rsR!cem3_name
                    SubTotal = SubTotal + CSng(rDs(rsR!cem3i))
            
                Case rsR!cem4_name
                    SubTotal = SubTotal + CSng(rDs(rsR!cem4i))
            
                Case rsR!wat1_name
                    SubTotal = SubTotal + CSng(rDs(rsR!wat1i))
            
                Case rsR!wat2_name
                    SubTotal = SubTotal + CSng(rDs(rsR!wat2i))
            
                Case rsR!chem1_name
                    SubTotal = SubTotal + CSng(rDs(rsR!chem1i))
            
                Case rsR!chem2_name
                    SubTotal = SubTotal + CSng(rDs(rsR!chem2i))
            
                Case rsR!chem3_name
                    SubTotal = SubTotal + CSng(rDs(rsR!chem3i))
            
                Case rsR!chem4_name
                    SubTotal = SubTotal + CSng(rDs(rsR!chem4i))
            
                Case rsR!chem5_name
                    SubTotal = SubTotal + CSng(rDs(rsR!chem5i))
            
                Case rsR!chem6_name
                    SubTotal = SubTotal + CSng(rDs(rsR!chem6i))
            
            End Select
            
            
            
            tempDate = rsR!stamp_date
            
            rsR.MoveNext 'местим на следващ запис
            If Not rsR.EOF Then
                If chDate <> rsR!stamp_date Then
                    chDate = rsR!stamp_date
                    If Not rsR.EOF And frstFlag = False Then
                        rcCounter = rcCounter + 1
                        ind = ind + 1
                        Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , Format(ind, "0000"))
                            itmX.SubItems(1) = tempDate
                            itmX.SubItems(2) = Me.cmbMat.Text
                            itmX.SubItems(3) = SubTotal / 1000
                            
                        SumTotal = SumTotal + SubTotal
                        SubTotal = 0
                    End If
                End If
            ElseIf rsR.EOF Then
                rcCounter = rcCounter + 1
                ind = ind + 1
                Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , Format(ind, "0000"))
                    itmX.SubItems(1) = tempDate
                    itmX.SubItems(2) = Me.cmbMat.Text
                    itmX.SubItems(3) = SubTotal / 1000
                    
                SumTotal = SumTotal + SubTotal
            End If
            
            frstFlag = False
        Loop
    
        MousePointer = vbDefault
        rsR.Close
        Set rsR = Nothing
        cnR.Close 'затваряме връзката
        Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------

        'след като прекъснем връзката сумираме тоталите
        Dim uniD As String
    
        'първо въвеждаме един празен ред
        rcCounter = rcCounter + 1
        Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "---------")
            itmX.SubItems(1) = "---------"
            itmX.SubItems(2) = "---------"
            itmX.SubItems(3) = "---------"
    
        'след него въвеждаме тотали
        If ind = 1 Then
            uniD = uniDay
        Else
            uniD = uniDays
        End If
        
        rcCounter = rcCounter + 1
        Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "")
            itmX.SubItems(1) = uniTotal & ind & " " & uniD
            itmX.SubItems(3) = SumTotal / 1000
        indT = indT + ind
        SumTotalT = SumTotalT + SumTotal
        
        If ag = False And MachineNumber = 1 And frmStartRep.chMach2.Value = 1 Then
            MachineNumber = 2
            ag = True
            rcCounter = rcCounter + 1
            Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "")
            ind = 0
            GoTo AgainOther
        End If
        
    If have = 2 Then
        'първо въвеждаме един празен ред
        rcCounter = rcCounter + 1
        Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "")
            
        rcCounter = rcCounter + 1
        Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "")
            itmX.SubItems(1) = "Всичко"
            
        rcCounter = rcCounter + 1
        Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "---------")
            itmX.SubItems(1) = "---------"
            itmX.SubItems(2) = "---------"
            itmX.SubItems(3) = "---------"
        
        'след него въвеждаме тотали
        rcCounter = rcCounter + 1
        Set itmX = Me.lstMatSold.ListItems.Add(rcCounter, , "")
            itmX.SubItems(1) = indT
            itmX.SubItems(3) = SumTotalT / 1000
    End If
    Else
        MsgBox msgChooseFilter, vbOKOnly Or vbInformation, MsgErrBx
    End If
EndSub:
    AutoColW Me.lstMatSold
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstMatSold, 1, True, True, True, repMatSold)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstMatSold)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

