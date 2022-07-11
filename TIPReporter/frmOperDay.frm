VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperDay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmOperDay"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   Icon            =   "frmOperDay.frx":0000
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
      Left            =   3960
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
      Left            =   1440
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
      ItemData        =   "frmOperDay.frx":08CA
      Left            =   360
      List            =   "frmOperDay.frx":08CC
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstOperDay 
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
Attribute VB_Name = "frmOperDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за 1 оператор - филтър по име от резултатната таблица
'всеки ред от справката изобразява един ден от зададен период от формата
    
    Me.Caption = repOperDay
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
    
    Me.lstOperDay.ColumnHeaders.Clear
    Me.lstOperDay.ListItems.Clear
    
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
    
    Me.cmbOper.Clear
    
    Set colx = Me.lstOperDay.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstOperDay.ColumnHeaders.Add()
        colx.Text = uniDate
        colx.Width = 1200
    
    Set colx = Me.lstOperDay.ColumnHeaders.Add()
        colx.Text = uniExpeds
        colx.Width = 1000
    
    Set colx = Me.lstOperDay.ColumnHeaders.Add()
        colx.Text = uniMixes
        colx.Width = 1000
    
    Set colx = Me.lstOperDay.ColumnHeaders.Add()
        colx.Text = uniQ
        colx.Width = 1000
    
    If rQinForm = 1 Then
        Set colx = Me.lstOperDay.ColumnHeaders.Add()
            colx.Text = uniQ & " " & uniMade
            colx.Width = 1000
    End If
    
    AutoColW Me.lstOperDay
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (name_op) name_op FROM mix_result_bc" & MachineNumber & " ORDER BY name_op ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbOper.AddItem rsR!name_op
        rsR.MoveNext
    Loop
    
    If frmStartRep.chMach1.Value = 1 And frmStartRep.chMach2.Value = 1 Then
        Set rsR = cnR.Execute("SELECT DISTINCT ON (name_op) name_op FROM mix_result_bc2 ORDER BY name_op ASC;")
    
        If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
        Do While Not rsR.EOF
            Me.cmbOper.AddItem rsR!name_op
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
    
    Me.lstOperDay.ListItems.Clear
    
    If Me.cmbOper.Text <> "" Then
    
'-----------------------Start postgreSQL-----------------------------------
        Dim cnR As New ADODB.Connection
        Dim rsR As New Recordset
        Dim commR As String
        Dim tempDate As String
        Dim tempExped As Integer
        Dim tempVol As Single
        Dim tempExVol As Single
        Dim rcCounter As Integer
        Dim expCounter As Integer
        Dim mixCounter As Integer
        Dim DayStart As String
        Dim DayEnd As String
    
        DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
        DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
            
AgainOther:
    
        cnR.ConnectionTimeout = 30
        cnR.Open ConStr
        MousePointer = vbHourglass
    
        'маркираме набор от записи
        commR = "SELECT time_exp_start, exp_num, exp_q, total_vol FROM mix_result_bc" & MachineNumber & " WHERE name_op = '" _
        & Me.cmbOper.Text & "' AND stamp_date >= '" & DayStart _
        & "' AND stamp_date <= '" & DayEnd _
        & "' ORDER BY mix_num, stamp_date ASC;"

        Set rsR = cnR.Execute(commR)
    
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            Me.btnPrint.Enabled = True
            Me.btnExport.Enabled = True
            rcCounter = rcCounter + 1
            Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "")
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
            tempDate = Left(rsR!time_exp_start, 10) 'вземаме първата срещната дата
        End If
    
        'нулираме променливите носещи данните за таблицата
        expCounter = 0
        mixCounter = 0
        tempExVol = 0
        tempVol = 0
    
        Do While Not rsR.EOF
            If tempDate = Left(rsR!time_exp_start, 10) Then 'ако записа отговаря на датата
            
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
                Exit Do 'ако датата е различна от първата срещната прекъсваме цикъла
            End If
            rsR.MoveNext 'местим на следващ запис
        Loop
    
        'при прекъсване на цикъла попълваме в ListView
        If tempDate <> "" Then
            ind = ind + 1
            Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , Format(ind, "0000"))
                itmX.SubItems(1) = tempDate
                itmX.SubItems(2) = expCounter
                itmX.SubItems(3) = mixCounter
                itmX.SubItems(4) = tempExVol
                If rQinForm = 1 Then
                    itmX.SubItems(5) = ARound(tempVol, 3)
                End If
        End If
        
        'ако не сме стигнали до края на маркираните записи то се връщаме на флага за начало да вземем новата дата
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
        Dim uniD As String
        
        tExp = 0
        Tmix = 0
        tExpVol = 0
        tVol = 0
    
        For i = rcCounter - ind + 1 To Me.lstOperDay.ListItems.count
            tExp = tExp + Val(Me.lstOperDay.ListItems.Item(i).SubItems(2))
            Tmix = Tmix + Val(Me.lstOperDay.ListItems.Item(i).SubItems(3))
            tExpVol = tExpVol + CSng(rDs(Me.lstOperDay.ListItems.Item(i).SubItems(4)))
            If rQinForm = 1 Then
                tVol = tVol + CSng(rDs(Me.lstOperDay.ListItems.Item(i).SubItems(5)))
            End If
        Next i
    
        'първо въвеждаме един празен ред
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "---------")
            itmX.SubItems(1) = "---------"
            itmX.SubItems(2) = "---------"
            itmX.SubItems(3) = "---------"
            itmX.SubItems(4) = "---------"
        
        'след него въвеждаме тотали
        If Me.lstOperDay.ListItems.count = 2 Then
            uniD = uniDay
        Else
            uniD = uniDays
        End If
        rcCounter = rcCounter + 1
        Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "XXXX")
            itmX.SubItems(1) = uniTotal & ind & " " & uniD
            itmX.SubItems(2) = tExp
            itmX.SubItems(3) = Tmix
            itmX.SubItems(4) = tExpVol
            If rQinForm = 1 Then
                itmX.SubItems(5) = ARound(tVol, 3)
            End If
            indT = indT + ind
            tExpT = tExpT + tExp
            TmixT = TmixT + Tmix
            tExpVolT = tExpVolT + tExpVol
            tVolT = tVolT + tVol
            
        If ag = False And MachineNumber = 1 And frmStartRep.chMach2.Value = 1 Then
            MachineNumber = 2
            ag = True
            rcCounter = rcCounter + 1
            Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "")
            ind = 0
            GoTo AgainOther
        End If
        
        If have = 2 Then
            'първо въвеждаме един празен ред
            rcCounter = rcCounter + 1
            Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "")
            
            rcCounter = rcCounter + 1
            Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "")
                itmX.SubItems(1) = "Всичко"
            
            rcCounter = rcCounter + 1
            Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "---------")
                itmX.SubItems(1) = "---------"
                itmX.SubItems(2) = "---------"
                itmX.SubItems(3) = "---------"
                itmX.SubItems(4) = "---------"
        
            'след него въвеждаме тотали
            If Me.lstOperDay.ListItems.count = 2 Then
                uniD = uniDay
            Else
                uniD = uniDays
            End If
            rcCounter = rcCounter + 1
            Set itmX = Me.lstOperDay.ListItems.Add(rcCounter, , "XXXX")
                itmX.SubItems(1) = uniTotal & indT & " " & uniD
                itmX.SubItems(2) = tExpT
                itmX.SubItems(3) = TmixT
                itmX.SubItems(4) = tExpVolT
                If rQinForm = 1 Then
                    itmX.SubItems(5) = ARound(tVolT, 3)
                End If
        End If
        
    Else
        MsgBox msgChooseFilter, vbOKOnly Or vbInformation, MsgErrBx
    End If
EndSub:
    AutoColW Me.lstOperDay
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstOperDay, 1, True, True, True, repOperDay & "  (" & Me.cmbOper & ")")
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstOperDay)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

