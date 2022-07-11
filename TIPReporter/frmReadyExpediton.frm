VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReadyExpedition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmReadyExpedition"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16575
   Icon            =   "frmReadyExpediton.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   16575
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
      Left            =   15480
      TabIndex        =   8
      Top             =   8640
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
      Left            =   9000
      TabIndex        =   5
      Top             =   8640
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
      Left            =   5280
      TabIndex        =   4
      Top             =   8640
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
      Left            =   13320
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin MSComctlLib.ListView lstReadyExpedition 
      Height          =   7215
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   12726
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
      Left            =   8880
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
      Format          =   47316995
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   11040
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
      Left            =   11040
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
      Left            =   8880
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmReadyExpedition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за всички клиенти
'всеки ред от справката изобразява една експедиция през дадения период от формата

    Me.Caption = repDailyExped
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
    
    Me.lstReadyExpedition.ColumnHeaders.Clear
    Me.lstReadyExpedition.ListItems.Clear
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'номер експедиция
        colx.Text = uniEx & " " & uniNr
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'дата и час експедиция
        colx.Text = uniDate & " - " & uniHour
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'кол. на експедицията
        colx.Text = uniEx & " " & uniQ
        colx.Width = 1000

    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'номер заявка/дата
        colx.Text = uniOrd & " " & uniNr & "/" & uniDate
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'кол. по заявка
        colx.Text = uniOrdered & "  " & uniQ
        colx.Width = 1000

    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'клиент
        colx.Text = uniClnt
        colx.Width = 1000

    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'обект
        colx.Text = uniObj
        colx.Width = 1000

    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'разстояние до обект
        colx.Text = uniKmShort
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'име рецепта
        colx.Text = uniRec
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'клас якост
        colx.Text = uniClass
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'клас консистенция
        colx.Text = uniClassK
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'клас въздействие
        colx.Text = uniClassV
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'клас с-не на хлориди
        colx.Text = uniClassH
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'водоплътност
        colx.Text = uniClassP
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'водач
        colx.Text = uniDrv
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'кола номер - капацитет
        colx.Text = uniDrvReg & " - " & uniCap
        colx.Width = 1000
    
    Set colx = Me.lstReadyExpedition.ColumnHeaders.Add() 'оператор/диспечер
        colx.Text = uniDisp
        colx.Width = 1000
    
    Me.lstReadyExpedition.SortKey = 1
    
    AutoColW Me.lstReadyExpedition
    
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

    Me.lstReadyExpedition.ListItems.Clear

        Dim Temp As Result
        Set Temp = New Result

'-----------------------Start postgreSQL-----------------------------------
        Dim cnR As New ADODB.Connection
        Dim rsR As New Recordset
        Dim commR As String
        Dim frstFlag As Boolean
        Dim expCounter As Integer
        Dim DayStart As String
        Dim DayEnd As String
    
        DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
        DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
        cnR.ConnectionTimeout = 30
        cnR.Open ConStr
        MousePointer = vbHourglass

AgainOther:

        'стринг за маркираме набор от записи само по клиент
        commR = "SELECT time_mix_ready, name_op, exp_num, exp_q, ord_num, ord_date, ord_q, name_clnt, obj_clnt, km_clnt, name_rec, class_rec, classk_rec, classv_rec, classh_rec, classp_rec, name_drv, reg_drv, cap_drv, total_vol FROM mix_result_bc" & MachineNumber & " WHERE stamp_date >= '" & DayStart _
        & "' AND stamp_date <= '" & DayEnd & "' ORDER BY mix_num, stamp_date ASC;"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)
    
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            Me.btnPrint.Enabled = True
            Me.btnExport.Enabled = True
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
                    ag = True
                    MachineNumber = 2
'                    ind = 0
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
        
        'старт на групирането
beginCount:
    
        frstFlag = True
    
        'нулираме променливите носещи данните за таблицата
        expCounter = 0
        mixCounter = 0
        expCounter = 0
        tempVol = 0

        Do While Not rsR.EOF
            'ако експедицията е различна от номера си броим експедиция +1
            'при първия обход на цикъла променливата е празна и броим +1 номер
            'при първия обход не показваме данните в ListView - само събираме информация
            If Temp.ExpNum <> Val(rsR!exp_num) Then
                'при смяна на номера зареждаме данни в ListView
                'и нулираме необходимите променливи
                If frstFlag = False Then
fillList:
                    Set itmX = Me.lstReadyExpedition.ListItems.Add(1, , "M" & MachineNumber & "-" & Temp.ExpNum)
                        itmX.SubItems(1) = Temp.MixReadyTime
                        If rQinForm = 0 Then
                            itmX.SubItems(2) = Temp.ExpQuant
                        Else
                            itmX.SubItems(2) = ARound(Temp.TotalQuant, 3)
                        End If
                        itmX.SubItems(3) = Temp.OrderCode & " / " & Temp.OrderDate 'с датата
                        itmX.SubItems(4) = Temp.OrderQuant
                        itmX.SubItems(5) = Temp.ClntTitle
                        itmX.SubItems(6) = Temp.ClntWorksite
                        itmX.SubItems(7) = Temp.WorksiteDist
                        itmX.SubItems(8) = Temp.RecTitle
                        itmX.SubItems(9) = Temp.RecClass
                        itmX.SubItems(10) = Temp.RecClassK
                        itmX.SubItems(11) = Temp.RecClassV
                        itmX.SubItems(12) = Temp.RecClassH
                        itmX.SubItems(13) = Temp.RecClassP
                        itmX.SubItems(14) = Temp.DrvTitle
                        itmX.SubItems(15) = Temp.DrvCarNum & " - " & Temp.DrvCapacity
                        itmX.SubItems(16) = Temp.DispName
                    If rsR.EOF Then Exit Do
                End If
            
                Temp.TotalQuant = 0
                expCounter = expCounter + 1
                frstFlag = False
            Else
            End If
            
            'записваме номера на експедицията като нова
            'така при втори обход на цикъла вече не е празна променлива
            Temp.ExpNum = Val(rsR!exp_num)
        
            'записваме новите стойности след предходната проверка
            'с цел да сравняваме изменението на номера от базата данни
            Temp.MixReadyTime = rsR!time_mix_ready
            Temp.ExpQuant = CSng(rDs(rsR!exp_q))
            Temp.OrderCode = rsR!ord_num
            Temp.OrderDate = Left(rsR!ord_date, 10)
            Temp.OrderQuant = CSng(rDs(rsR!ord_q))
            Temp.ClntTitle = rsR!name_clnt
            Temp.ClntWorksite = rsR!obj_clnt
            Temp.WorksiteDist = Val(rsR!km_clnt)
            Temp.RecTitle = rsR!name_rec
            Temp.RecClass = rsR!class_rec
            Temp.RecClassK = rsR!classk_rec
            Temp.RecClassV = rsR!classv_rec
            Temp.RecClassH = rsR!classh_rec
            Temp.RecClassP = rsR!classp_rec
            Temp.DrvTitle = rsR!name_drv
            Temp.DrvCarNum = rsR!reg_drv
            Temp.DrvCapacity = CSng(rDs(rsR!cap_drv))
            Temp.DispName = rsR!name_op
        
            'сумираме обемите от всеки запис
            Temp.TotalQuant = Temp.TotalQuant + CSng(rDs(rsR!total_vol))
            
            rsR.MoveNext 'местим на следващ запис
            If rsR.EOF Then GoTo fillList
        Loop
    
        If ag = False And MachineNumber = 1 And frmStartRep.chMach2.Value = 1 Then
            MachineNumber = 2
            ag = True
            GoTo AgainOther
        End If
        
        MousePointer = vbDefault
        rsR.Close
        Set rsR = Nothing
        cnR.Close
        Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------
    
        'след като прекъснем връзката сумираме тоталите
        Dim i As Long
        Dim tExpVol As Single
    
        tExpVol = 0
    
        For i = 1 To Me.lstReadyExpedition.ListItems.count
            tExpVol = tExpVol + CSng(rDs(Me.lstReadyExpedition.ListItems.Item(i).SubItems(2)))
        Next i
    
        'първо въвеждаме един празен ред
        Set itmX = Me.lstReadyExpedition.ListItems.Add(1, , "")
            itmX.SubItems(1) = "X"
            
        'след него въвеждаме тотали
        Set itmX = Me.lstReadyExpedition.ListItems.Add(1, , "")
            itmX.SubItems(1) = uniTotal & Me.lstReadyExpedition.ListItems.count - 2 & " " & uniEx
            itmX.SubItems(2) = ARound(tExpVol, 3)
EndSub:
    AutoColW Me.lstReadyExpedition
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstReadyExpedition, 2, True, True, True, repDailyExped)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstReadyExpedition)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Temp = Nothing
    frmStartRep.Show
End Sub

