VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrvDay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDrvDay"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   Icon            =   "frmDrvDay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11415
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
      Left            =   10320
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
      Left            =   6600
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
      Left            =   2520
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
      Left            =   8160
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cmbDrv 
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
   Begin MSComctlLib.ListView lstDrvDay 
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11245
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
      Left            =   4200
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
      Left            =   6120
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
      Left            =   6120
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
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblDrv 
      Alignment       =   2  'Center
      Caption         =   "lblDrv"
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
Attribute VB_Name = "frmDrvDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за 1 оператор - филтър по име от резултатната таблица
'всеки ред от справката изобразява един ден от зададен период от формата
    
    Me.Caption = repDrvDay
    Me.lblDrv.Caption = uniDrv
    Me.lblStart.Caption = lblStDate
    Me.lblEnd.Caption = lblEndDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    
    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False
    
    Me.dtStart = Now
    Me.dtEnd = Now
    
    Me.lstDrvDay.ColumnHeaders.Clear
    Me.lstDrvDay.ListItems.Clear
    
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
    
    Me.cmbDrv.Clear
    
    Set colx = Me.lstDrvDay.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 800
    
    Set colx = Me.lstDrvDay.ColumnHeaders.Add()
        colx.Text = uniDate
        colx.Width = 1200
    
    Set colx = Me.lstDrvDay.ColumnHeaders.Add()
        colx.Text = uniDrvReg & " - " & uniCapacity
        colx.Width = 1000
    
    Set colx = Me.lstDrvDay.ColumnHeaders.Add()
        colx.Text = uniKmShort
        colx.Width = 1000
    
    Set colx = Me.lstDrvDay.ColumnHeaders.Add()
        colx.Text = uniExpeds
        colx.Width = 1000
    
    Set colx = Me.lstDrvDay.ColumnHeaders.Add()
        colx.Text = uniQ
        colx.Width = 1000
    
    If rQinForm = 1 Then
        Set colx = Me.lstDrvDay.ColumnHeaders.Add()
            colx.Text = uniQ & " " & uniMade
            colx.Width = 1000
    End If
    
    Me.lstDrvDay.SortKey = 1
    
    AutoColW Me.lstDrvDay
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (name_drv) name_drv FROM mix_result_bc" & MachineNumber & " ORDER BY name_drv ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbDrv.AddItem rsR!name_drv
        rsR.MoveNext
    Loop
    
    If frmStartRep.chMach1.Value = 1 And frmStartRep.chMach2.Value = 1 Then
        Set rsR = cnR.Execute("SELECT DISTINCT ON (name_drv) name_drv FROM mix_result_bc2 ORDER BY name_drv ASC;")
    
        If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
        Do While Not rsR.EOF
            Me.cmbDrv.AddItem rsR!name_drv
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
        
    For i = 0 To Me.cmbDrv.listCount - 2 Step 1
        For j = Me.cmbDrv.listCount - 1 To i + 1 Step -1
            If Me.cmbDrv.List(i) = Me.cmbDrv.List(j) Then
                Me.cmbDrv.RemoveItem (j)
            End If
        Next
    Next

End Sub

Private Sub cmbDrv_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbDrv, KeyAscii, True)
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

    Me.lstDrvDay.ListItems.Clear
    
    If Me.cmbDrv.Text <> "" Then

'-----------------------Start postgreSQL-----------------------------------
        Dim cnR As New ADODB.Connection
        Dim rsR As New Recordset
        Dim commR As String
        Dim tempDate As String
        Dim tempDrvCap As String
        Dim tempDrvReg As String
        Dim tempExped As Integer
        Dim tempVol As Single
        Dim tempExVol As Single
        Dim rcCounter As Integer
        Dim expCounter As Integer
        Dim tempKm As Long
        Dim DayStart As String
        Dim DayEnd As String
    
        DayStart = Format(Me.dtStart.Value, "DD-MM-YYYY")
        DayEnd = Format(Me.dtEnd.Value, "DD-MM-YYYY")
    
        cnR.ConnectionTimeout = 30
        cnR.Open ConStr
        MousePointer = vbHourglass

AgainOther:

        'маркираме набор от записи
        commR = "SELECT exp_num, time_mix_ready, reg_drv, cap_drv, km_clnt, exp_q, total_vol FROM mix_result_bc" & MachineNumber & " WHERE name_drv = '" _
        & Me.cmbDrv.Text & "' AND stamp_date >= '" & DayStart _
        & "' AND stamp_date <= '" & DayEnd _
        & "' ORDER BY mix_num, stamp_date ASC;"

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
        
        rcCounter = 0 'нулираме брояча на редовете в ListView
    
        'старт на групирането
beginCount:
        rcCounter = rcCounter + 1 'първи ред ++
        If Not rsR.EOF And Not rsR.BOF Then
            tempDate = Left(rsR!time_mix_ready, 10) 'вземаме първата срещната дата
        End If
    
        'нулираме променливите носещи данните за таблицата
        expCounter = 0
        tempKm = 0
        expCounter = 0
        tempExVol = 0
        tempVol = 0
        tempDrvReg = ""
        tempDrvCap = ""
    
        Do While Not rsR.EOF
            If tempDate = Left(rsR!time_mix_ready, 10) Then 'ако записа отговаря на датата
            
                'ако експедицията е различна от номера си броим експедиция +1 и количество +
                'при първия обход на цикъла променливата е празна и броим +1 номер и количеството +
                If tempExped <> rsR!exp_num Then
                    tempExVol = tempExVol + CSng(rDs(rsR!exp_q))
                    tempKm = tempKm + CSng(rDs(rsR!km_clnt)) 'сумираме км от всяка експедиция
                    expCounter = expCounter + 1
                Else
                End If
            
                'записваме номера на експедицията като нова
                'така при втори обход на цикъла вече не е празна променлива
                'записваме новите стойности след предходната проверка
                'с цел да сравняваме изменението на номера от базата данни
                tempExped = rsR!exp_num
                tempDrvReg = rsR!reg_drv
                tempDrvCap = CSng(rDs(rsR!cap_drv))
            
                'сумираме обемите от всеки запис
                tempVol = tempVol + CSng(rDs(rsR!total_vol))
            
            Else
                Exit Do 'ако датата е различна от първата срещната прекъсваме цикъла
            End If
            rsR.MoveNext 'местим на следващ запис
        Loop
        
        If tempDate <> "" Then
            'при прекъсване на цикъла попълваме в ListView
            Set itmX = Me.lstDrvDay.ListItems.Add(1, , "M" & MachineNumber)
                itmX.SubItems(1) = tempDate
                itmX.SubItems(2) = tempDrvReg & " - " & tempDrvCap
                itmX.SubItems(3) = tempKm
                itmX.SubItems(4) = expCounter
                itmX.SubItems(5) = tempExVol
                If rQinForm = 1 Then
                    itmX.SubItems(6) = ARound(tempVol, 3)
                End If
        End If
        
        'ако не сме стигнали до края на маркираните записи то се връщаме на флага за начало да вземем новата дата
        If Not rsR.EOF Then GoTo beginCount
        
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
        Dim i As Integer
        Dim tExp As Long
        Dim tKm As Long
        Dim tExpVol As Single
        Dim tVol As Single
        Dim uniD As String
    
        tExp = 0
        Tmix = 0
        tExpVol = 0
        tVol = 0
    
        For i = 1 To Me.lstDrvDay.ListItems.count
            tKm = tKm + Val(Me.lstDrvDay.ListItems.Item(i).SubItems(3))
            tExp = tExp + Val(Me.lstDrvDay.ListItems.Item(i).SubItems(4))
            tExpVol = tExpVol + CSng(rDs(Me.lstDrvDay.ListItems.Item(i).SubItems(5)))
            If rQinForm = 1 Then
                tVol = tVol + CSng(rDs(Me.lstDrvDay.ListItems.Item(i).SubItems(6)))
            End If
        Next i
    
        'след него въвеждаме тотали
        If Me.lstDrvDay.ListItems.count = 1 Then
            uniD = uniDay
        Else
            uniD = uniDays
        End If
        Set itmX = Me.lstDrvDay.ListItems.Add(1, , "XXXX")
            itmX.SubItems(1) = uniTotal & Me.lstDrvDay.ListItems.count - 1 & " " & uniD
            itmX.SubItems(3) = tKm
            itmX.SubItems(4) = tExp
            itmX.SubItems(5) = tExpVol
            If rQinForm = 1 Then
                itmX.SubItems(6) = ARound(tVol, 3)
            End If
    Else
        MsgBox msgChooseFilter, vbOKOnly Or vbInformation, MsgErrBx
    End If
EndSub:
    AutoColW Me.lstDrvDay
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstDrvDay, 1, True, True, True, repDrvDay & "  (" & Me.cmbDrv & ")")
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstDrvDay)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

