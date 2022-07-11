VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClntAllExpedition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmClntAllExpedition"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15735
   Icon            =   "frmClntAllExpediton.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   15735
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
      Left            =   14640
      TabIndex        =   5
      Top             =   8880
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
      Left            =   8880
      TabIndex        =   2
      Top             =   8880
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
      Left            =   4680
      TabIndex        =   1
      Top             =   8880
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
      Left            =   12480
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   360
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
      Format          =   98893827
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   360
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
      Format          =   98893827
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin MSComctlLib.ListView lstClntAllExpedition 
      Height          =   7695
      Left            =   360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13573
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
      Left            =   9000
      TabIndex        =   4
      Top             =   480
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
      Left            =   5400
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmClntAllExpedition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за 1 клиент - филтър по име от резултатната таблица
'всеки ред от справката изобразява една експедиция през дадения период от формата
    
    Me.Caption = "Всички клиенти по експедиции"
    Me.lblStart.Caption = lblStDate
    Me.lblEnd.Caption = lblEndDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    
    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False
    
    Me.dtStart = Now
    Me.dtEnd = Now
    
    
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
    
    Me.lstClntAllExpedition.ColumnHeaders.Clear
    Me.lstClntAllExpedition.ListItems.Clear
        
    Set colx = Me.lstClntAllExpedition.ColumnHeaders.Add() 'клиент
        colx.Text = uniClnt
        colx.Width = 800
        
    Set colx = Me.lstClntAllExpedition.ColumnHeaders.Add() 'номер бележка
        colx.Text = uniEx & " " & uniNr
        colx.Width = 800

    Set colx = Me.lstClntAllExpedition.ColumnHeaders.Add() 'обект
        colx.Text = uniObj
        colx.Width = 1000

    Set colx = Me.lstClntAllExpedition.ColumnHeaders.Add() 'дата експедиция
        colx.Text = uniDate
        colx.Width = 1000
    
'-----------------------Start postgreSQL-----------------------------------
        Dim cnR As New ADODB.Connection
        Dim rsR As New Recordset
        Dim commR As String
        Dim DayStart As String
        Dim DayEnd As String
        
        Dim tempRec(1 To 50) As String
    
        DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
        DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
        cnR.ConnectionTimeout = 30
        cnR.Open ConStr
        MousePointer = vbHourglass
            
        'намираме всички клиенти за маркирания период
        Set rsR = cnR.Execute("SELECT DISTINCT ON (class_rec) class_rec FROM mix_result_bc" & MachineNumber & " WHERE stamp_date >= '" & DayStart & "' AND stamp_date <= '" & DayEnd & "' ORDER BY class_rec ASC")
    
        If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst

        Do While Not rsR.EOF
            Set colx = Me.lstClntAllExpedition.ColumnHeaders.Add() 'клас рецепта
                colx.Text = rsR!class_rec
                colx.Width = 1000
            reccounter = reccounter + 1
            tempRec(reccounter) = rsR!class_rec
            rsR.MoveNext
        Loop
        
    Set colx = Me.lstClntAllExpedition.ColumnHeaders.Add() 'дата експедиция
        colx.Text = "Общо"
        colx.Width = 1000
        
'    Me.lstClntAllExpedition.SortKey = 0
    
    AutoColW Me.lstClntAllExpedition
    

    Me.lstClntAllExpedition.ListItems.Clear
'    If Me.cmbClnt.Text <> "" Then
        
        Dim Temp As Result
        Set Temp = New Result
'-----------------------Start postgreSQL-----------------------------------
        Dim commR1 As String
        Dim commR2 As String
        Dim commR3 As String
        Dim commR4 As String
        Dim commR5 As String
        Dim commRord As String
        Dim frstFlag As Boolean
        Dim expCounter As Integer
    
        DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
        DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
        MousePointer = vbHourglass
    
AgainOther:
        If endofclients = True Then GoTo fillList
        
        'намираме всички клиенти за маркирания период
        Set rsR = cnR.Execute("SELECT DISTINCT ON (name_clnt) name_clnt FROM mix_result_bc" & MachineNumber & " WHERE stamp_date >= '" & DayStart & "' AND stamp_date <= '" & DayEnd & "' ORDER BY name_clnt DESC")
        
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
                    GoTo SumTotals
                End If
            Else
                If frmStartRep.chMach2.Value = 1 Then
                    ag = True
                    MachineNumber = 2
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
        
        Do While Not rsR.EOF
            For i = 1 To indclnt
                rsR.MoveNext
            Next i
            If rsR.EOF Then
                nowClnt = ""
                Exit Do
            Else
            End If
            If nowClnt = rsR!name_clnt Then
                rsR.MoveNext
            Else
                nowClnt = rsR!name_clnt
                indclnt = indclnt + 1
                rsR.MoveNext
                If rsR.EOF Then endofclients = True
                Exit Do
            End If
        Loop

        'стринг за маркираме набор от записи само по клиент
        commR1 = "SELECT time_mix_ready, name_op, exp_num, exp_q, ord_num, ord_date, ord_q, obj_clnt, km_clnt, name_rec, class_rec, classk_rec, classv_rec, classh_rec, classp_rec, name_drv, reg_drv, cap_drv, total_vol FROM mix_result_bc" & MachineNumber & " WHERE name_clnt = '" _
        & nowClnt & "' AND stamp_date >= '" & DayStart & "' AND stamp_date <= '" & DayEnd & "' ORDER BY mix_num DESC"

        'записваме стринг за търсене само по основен филтър - име клиент
        commR = commR1
        
        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)
        
        If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
        'отиваме на първия запис
        'старт на групирането
beginCount:
    
        frstFlag = True
    
        'нулираме променливите носещи данните за таблицата
        expCounter = 0
        mixCounter = 0
        Temp.TotalQuant = 0

        Do While Not rsR.EOF
            'ако експедицията е различна от номера си броим експедиция +1
            'при първия обход на цикъла променливата е празна и броим +1 номер
            'при първия обход не показваме данните в ListView - само събираме информация
            If Temp.ExpNum <> Val(rsR!exp_num) Then
                'при смяна на номера зареждаме данни в ListView
                'и нулираме необходимите променливи
                If frstFlag = False Then
fillList:
                    Set itmX = Me.lstClntAllExpedition.ListItems.Add(1, , nowClnt)
                        itmX.SubItems(1) = "M" & MachineNumber & "-" & Temp.ExpNum
                        itmX.SubItems(2) = Temp.ClntWorksite
                        itmX.SubItems(3) = Temp.MixReadyTime ' датата
                        For i = 1 To reccounter
                            If Temp.RecClass = tempRec(i) Then
                                itmX.SubItems(i + 3) = ARound(Temp.ExpQuant, 3)
                            End If
                        Next i
                        itmX.SubItems(i + 3) = ARound(Temp.ExpQuant, 3)
'                        itmX.SubItems(5) = Temp.ClntWorksite
'                        itmX.SubItems(6) = Temp.WorksiteDist
'                        itmX.SubItems(7) = Temp.RecTitle
'                        itmX.SubItems(8) = Temp.RecClass
'                        itmX.SubItems(9) = Temp.RecClassK
'                        itmX.SubItems(10) = Temp.RecClassV
'                        itmX.SubItems(11) = Temp.RecClassH
'                        itmX.SubItems(12) = Temp.RecClassP
'                        itmX.SubItems(13) = Temp.DrvTitle
'                        itmX.SubItems(14) = Temp.DrvCarNum & " - " & Temp.DrvCapacity
'                        itmX.SubItems(15) = Temp.DispName
                        myCount = myCount + 1
                    If rsR.EOF Then
                        MousePointer = vbDefault
                        Exit Do
                    End If
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
            Temp.OrderCode = Val(rsR!ord_num)
            Temp.OrderDate = Left(rsR!ord_date, 10)
            Temp.OrderQuant = CSng(rDs(rsR!ord_q))
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
    
        If endofclients = False Then GoTo AgainOther
        
        If ag = False And MachineNumber = 1 And frmStartRep.chMach2.Value = 1 Then
            MachineNumber = 2
            ag = True
            endofclients = False
            indclnt = 0
            GoTo AgainOther
        End If
        
        MousePointer = vbDefault
        rsR.Close
        Set rsR = Nothing
        cnR.Close
        Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------
SumTotals:
        'след като прекъснем връзката сумираме тоталите
'        Dim I As Long
        Dim Tmix As Long
        Dim tExpVol As Single
        Dim tExpRec(1 To 50) As Single
    
        tExp = 0
        Tmix = 0
        tExpVol = 0
    
        For i = 1 To Me.lstClntAllExpedition.ListItems.count
            For nn = 1 To reccounter
                tExpRec(nn) = tExpRec(nn) + CSng(rDs(Me.lstClntAllExpedition.ListItems.Item(i).SubItems(nn + 3)))
            Next nn
            
            tExpVol = tExpVol + CSng(rDs(Me.lstClntAllExpedition.ListItems.Item(i).SubItems(reccounter + 4)))
        Next i
                
        'след него въвеждаме тотали
        Set itmX = Me.lstClntAllExpedition.ListItems.Add(myCount + 1, , uniTotal)
            itmX.SubItems(1) = Me.lstClntAllExpedition.ListItems.count - 1 & " " & uniEx
            For nn = 1 To reccounter
                itmX.SubItems(nn + 3) = ARound(tExpRec(nn), 3)
            Next nn
            itmX.SubItems(reccounter + 4) = ARound(tExpVol, 3)
'    Else
'        MsgBox msgChooseFilter, vbOKOnly Or vbInformation, MsgErrBx
'    End If
EndSub:
    AutoColW Me.lstClntAllExpedition
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstClntAllExpedition, 1, True, True, True, "Всички клиенти по експедиции")
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstClntAllExpedition)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Temp = Nothing
    frmStartRep.Show
End Sub

