VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmOperAll"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670
   Icon            =   "frmOperAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8670
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
      Left            =   7560
      TabIndex        =   8
      Top             =   5160
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
      Left            =   4680
      TabIndex        =   5
      Top             =   5160
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
      Left            =   1680
      TabIndex        =   4
      Top             =   5160
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
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstOperAll 
      Height          =   3735
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6588
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
      Left            =   600
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
      Format          =   48037891
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   2760
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
      Format          =   48037891
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
      Left            =   2760
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
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmOperAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за всички оператори - филтър по период от резултатната таблица зададен от формата
'всеки ред от справката изобразява един оператор

    Me.Caption = repOperAll
    Me.lblStart.Caption = lblStDate
    Me.lblEnd.Caption = lblEndDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport

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
    
    Me.lstOperAll.ColumnHeaders.Clear
    Me.lstOperAll.ListItems.Clear
    
    Set colx = Me.lstOperAll.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstOperAll.ColumnHeaders.Add()
        colx.Text = uniNm
        colx.Width = 1200
    
    Set colx = Me.lstOperAll.ColumnHeaders.Add()
        colx.Text = uniDays
        colx.Width = 1000
    
    Set colx = Me.lstOperAll.ColumnHeaders.Add()
        colx.Text = uniExpeds
        colx.Width = 1000
    
    Set colx = Me.lstOperAll.ColumnHeaders.Add()
        colx.Text = uniMixes
        colx.Width = 1000
    
    Set colx = Me.lstOperAll.ColumnHeaders.Add()
        colx.Text = uniQ
        colx.Width = 1000
    
    If rQinForm = 1 Then
        Set colx = Me.lstOperAll.ColumnHeaders.Add()
            colx.Text = uniQ & " " & uniMade
            colx.Width = 1000
    End If
    
    AutoColW Me.lstOperAll
    
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

    Me.lstOperAll.ListItems.Clear

'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
    Dim commR As String
    Dim tempOper As String
    Dim tempDate As String
    Dim tempExped As Integer
    Dim tempVol As Single
    Dim tempExVol As Single
    Dim rcCounter As Integer
    Dim expCounter As Integer
    Dim mixCounter As Integer
    Dim dayCounter As Integer
    Dim DayStart As String
    Dim DayEnd As String
    
    DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
    DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")

AgainOther:
    
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме набор от записи
    commR = "SELECT exp_num, time_exp_start, name_op, exp_q, total_vol FROM mix_result_bc" & MachineNumber & " WHERE stamp_date >= '" & DayStart _
    & "' AND stamp_date <= '" & DayEnd _
    & "' ORDER BY name_op, mix_num, exp_num, stamp_date ASC;"

    Set rsR = cnR.Execute(commR)
    
    'отиваме на първия запис
    If Not rsR.EOF And Not rsR.BOF Then
        rsR.MoveFirst
        Me.btnPrint.Enabled = True
        Me.btnExport.Enabled = True
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "")
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
    
    'старт на групирането
beginCount:
    rcCounter = rcCounter + 1 'първи ред ++
    If Not rsR.EOF And Not rsR.BOF Then
        tempOper = rsR!name_op 'вземаме първия срещнат оператор
    End If
    
    'нулираме променливите носещи данните за таблицата
    expCounter = 0
    mixCounter = 0
    expCounter = 0
    dayCounter = 0
    tempExVol = 0
    tempVol = 0
    tempDate = ""
    
    Do While Not rsR.EOF
        If tempOper = rsR!name_op Then 'ако записа отговаря на името - при първи обход винаги
            
            If tempDate <> Left(rsR!time_exp_start, 10) Then dayCounter = dayCounter + 1
            
            tempDate = Left(rsR!time_exp_start, 10)
            'ако експедицията е различна от номера си броим експедиция +1 и количество +
            'при първия обход на цикъла променливата е празна и броим +1 номер и количеството +
            If tempExped <> rsR!exp_num Then
                tempExVol = tempExVol + CSng(rDs(rsR!exp_q))
                expCounter = expCounter + 1
            Else
            End If
            
            'записваме номера на експедицията като нова
            'така при втори обход на цикъла вече не е празна променлива
            'записваме новите стойности след предходната проверка
            'с цел да сравняваме изменението на номера от базата данни
            tempExped = rsR!exp_num
            
            'сумираме обемите от всеки запис
            tempVol = tempVol + CSng(rDs(rsR!total_vol))
            
            'броим всеки запис като замес
            mixCounter = mixCounter + 1
        Else
            Exit Do 'ако името е различно от първото срещнато прекъсваме цикъла
        End If
        rsR.MoveNext 'местим на следващ запис
    Loop
    
'    If tempOper <> "" Then
        'при прекъсване на цикъла попълваме в ListView
        ind = ind + 1
        Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , Format(ind, "0000"))
            itmX.SubItems(1) = tempOper
            itmX.SubItems(2) = dayCounter
            itmX.SubItems(3) = expCounter
            itmX.SubItems(4) = mixCounter
            itmX.SubItems(5) = tempExVol
            If rQinForm = 1 Then
                itmX.SubItems(6) = ARound(tempVol, 3)
            End If
'    End If
    
    'ако не сме стигнали до края на маркираните записи то се връщаме на флага за начало да вземем новото име
    If Not rsR.EOF Then GoTo beginCount
    
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------

     'след като прекъснем връзката сумираме тоталите
    Dim i As Integer
    Dim tExp As Long
    Dim Tmix As Long
    Dim tExpVol As Single
    Dim tVol As Single
    Dim tDay As Integer
    
    tExp = 0
    Tmix = 0
    tExpVol = 0
    tVol = 0
    
    For i = rcCounter - ind + 1 To Me.lstOperAll.ListItems.count
        tDay = tDay + Val(Me.lstOperAll.ListItems.Item(i).SubItems(2))
        tExp = tExp + Val(Me.lstOperAll.ListItems.Item(i).SubItems(3))
        Tmix = Tmix + Val(Me.lstOperAll.ListItems.Item(i).SubItems(4))
        tExpVol = tExpVol + CSng(rDs(Me.lstOperAll.ListItems.Item(i).SubItems(5)))
        If rQinForm = 1 Then
            tVol = tVol + CSng(rDs(Me.lstOperAll.ListItems.Item(i).SubItems(6)))
        End If
    Next i
    
    'първо въвеждаме един празен ред
    rcCounter = rcCounter + 1
    Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "---------")
        itmX.SubItems(1) = "---------"
        itmX.SubItems(2) = "---------"
        itmX.SubItems(3) = "---------"
        itmX.SubItems(4) = "---------"
        itmX.SubItems(5) = "---------"
    
    'след него въвеждаме тотали
    rcCounter = rcCounter + 1
    Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "XXXX")
        itmX.SubItems(1) = uniTotal
        itmX.SubItems(2) = tDay
        itmX.SubItems(3) = tExp
        itmX.SubItems(4) = Tmix
        itmX.SubItems(5) = tExpVol
        If rQinForm = 1 Then
            itmX.SubItems(6) = ARound(tVol, 3)
        End If
        tDayT = tDayT + tDay
        tExpT = tExpT + tExp
        TmixT = TmixT + Tmix
        tExpVolT = tExpVolT + tExpVol
        tVolT = tVolT + tVol
        
    If ag = False And MachineNumber = 1 And frmStartRep.chMach2.Value = 1 Then
        MachineNumber = 2
        ag = True
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "")
        ind = 0
        GoTo AgainOther
    End If

    If have = 2 Then
        'първо въвеждаме един празен ред
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "")
            
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "")
            itmX.SubItems(1) = "Всичко"
            
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "---------")
            itmX.SubItems(1) = "---------"
            itmX.SubItems(2) = "---------"
            itmX.SubItems(3) = "---------"
            itmX.SubItems(4) = "---------"
            itmX.SubItems(5) = "---------"
        
        'след него въвеждаме тотали
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperAll.ListItems.Add(rcCounter, , "XXXX")
            itmX.SubItems(1) = uniTotal
            itmX.SubItems(2) = tDayT
            itmX.SubItems(3) = tExpT
            itmX.SubItems(4) = TmixT
            itmX.SubItems(5) = tExpVolT
            If rQinForm = 1 Then
                itmX.SubItems(6) = ARound(tVolT, 3)
            End If
    End If
    
EndSub:
    AutoColW Me.lstOperAll

End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstOperAll, 1, True, True, True, repOperAll)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstOperAll)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

