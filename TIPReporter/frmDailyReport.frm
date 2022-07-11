VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDailyReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDailyReport"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17175
   Icon            =   "frmDailyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   17175
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
      Left            =   16080
      TabIndex        =   4
      Top             =   9000
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
      Left            =   9600
      TabIndex        =   2
      Top             =   9000
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
      TabIndex        =   1
      Top             =   9000
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
      Left            =   13920
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin MSComctlLib.ListView lstDailyReport 
      Height          =   7815
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   13785
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
   Begin MSComCtl2.DTPicker dtDay 
      Height          =   375
      Left            =   12120
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
      Format          =   52887555
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin VB.Label lblDay 
      Alignment       =   1  'Right Justify
      Caption         =   "lblDay"
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
      Left            =   10440
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmDailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за всички клиенти
'всеки ред от справката изобразява една експедиция през дадения период от формата
    
    Me.Caption = repDailyRep
    Me.lblDay.Caption = uniDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport

    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False
    
    Me.dtDay = Now
        
End Sub

Private Sub btnLoad_Click()
    
    Dim colx As ColumnHeader
    
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
         
'-----------------------Start postgreSQL-----------------------------------
        Dim cnR As New ADODB.Connection
        Dim rsR As New Recordset
        Dim commR As String
        
        Dim tempIM(1 To 20) As String
        Dim tempScr(1 To 20) As String
        Dim tempChem(1 To 20) As String
        
        Dim tempQIM(1 To 20) As Single
        Dim tempQScr(1 To 20) As Single
        Dim tempQChem(1 To 20) As Single
        
        Dim totalQIM(1 To 20) As Single
        Dim totalQScr(1 To 20) As Single
        Dim totalQChem(1 To 20) As Single
        
        Dim nowClnt As String
        Dim nowObj As String
        Dim nowRec As String
        Dim nowRegDrv As String
        
        Dim Day As String
    
        Day = Format(Me.dtDay.Value, "DD-MM-YYYY")
        
        cnR.ConnectionTimeout = 30
        cnR.Open ConStr
        MousePointer = vbHourglass
    
        'стринг за маркиране на материали
        commR = "SELECT * FROM materials_bc" & MachineNumber & " WHERE m_type = '1';"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)
    
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            i = 1
        Else
        End If
        Do While Not rsR.EOF
            tempScr(i) = rsR!m_name
            i = i + 1
            rsR.MoveNext
        Loop
        
        'стринг за маркиране на материали
        commR = "SELECT * FROM materials_bc" & MachineNumber & " WHERE m_type = '0';"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)
    
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            i = 1
        Else
        End If
        Do While Not rsR.EOF
            tempIM(i) = rsR!m_name
            i = i + 1
            rsR.MoveNext
        Loop
    
        'стринг за маркиране на материали
        commR = "SELECT * FROM materials_bc" & MachineNumber & " WHERE m_type = '3';"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)
    
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            i = 1
        Else
        End If
        Do While Not rsR.EOF
            tempChem(i) = rsR!m_name
            i = i + 1
            rsR.MoveNext
        Loop
    
    Me.lstDailyReport.ColumnHeaders.Clear
    Me.lstDailyReport.ListItems.Clear
    
    Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , "Клиент")  'клиент
        colx.Width = 1
    
    Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , "Обект") 'обект
        colx.Width = 1
    
    Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , "Рецепта") 'рецепта име
        colx.Width = 1

    Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , "Количество") 'произведена кол.
        colx.Width = 1
    
    Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , "Транспорт") 'кола
        colx.Width = 1
    
    i = 1
    Do While tempScr(i) <> ""
        Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , tempScr(i) & " [t]") 'цим
            colx.Width = 1
        i = i + 1
    Loop
    
    i = 1
    Do While tempIM(i) <> ""
        Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , tempIM(i) & " [t]") 'им
            colx.Width = 1
        i = i + 1
    Loop
    
    i = 1
    Do While tempChem(i) <> ""
        Set colx = Me.lstDailyReport.ColumnHeaders.Add(, , tempChem(i) & " [t]") 'хд
            colx.Width = 1
        i = i + 1
    Loop
        
StartSorting:

        If endofclients = True Then
            GoTo ObjSorting
        Else
            EndOfObjects = False
            indobj = 0
            indrec = 0
            indreg = 0
        End If

        'стринг за маркиране на 1 клиент
        commR = "SELECT DISTINCT ON (name_clnt) name_clnt FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' ORDER BY name_clnt ASC;"

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
                    GoTo StartSorting
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
                If rsR.EOF Then endofclients = True
            Else
                nowClnt = rsR!name_clnt
                indclnt = indclnt + 1
                rsR.MoveNext
                If rsR.EOF Then endofclients = True
                Exit Do
            End If
        Loop
    
ObjSorting:

        If EndOfObjects = True And endofclients = True Then
            GoTo RecSorting
        Else
            EndOfRecs = False
            indrec = 0
            indreg = 0
        End If

        'стринг за маркиране на 1 обект от настоящия клиент nowClnt
        commR = "SELECT DISTINCT ON (obj_clnt) obj_clnt FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' AND name_clnt = '" & nowClnt & "' ORDER BY obj_clnt ASC;"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)
    
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            Me.btnPrint.Enabled = True
            Me.btnExport.Enabled = True
        Else
        End If

        Do While Not rsR.EOF
            For i = 1 To indobj
                rsR.MoveNext
            Next i
            If rsR.EOF Then
                nowObj = ""
                Exit Do
            Else
            End If
            If nowObj = rsR!obj_clnt Then
                rsR.MoveNext
                If rsR.EOF Then EndOfObjects = True
            Else
                nowObj = rsR!obj_clnt
                indobj = indobj + 1
                rsR.MoveNext
                If rsR.EOF Then EndOfObjects = True
                Exit Do
            End If
        Loop
                
RecSorting:

        If EndOfRecs = True And EndOfObjects = True And endofclients = True Then
            GoTo RegDrvSorting
        Else
            EndOfRegDrv = False
            indreg = 0
        End If

        'стринг за маркиране на 1 рецепта от избрания обект nowObj от настоящия клиент nowClnt
        commR = "SELECT DISTINCT ON (name_rec) name_rec FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' AND name_clnt = '" & nowClnt & "' AND obj_clnt = '" & nowObj & "' ORDER BY name_rec ASC;"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)

        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
        Else
            EndOfRecs = True
        End If

        Do While Not rsR.EOF
            For i = 1 To indrec
                rsR.MoveNext
            Next i
            If rsR.EOF Then
                nowRec = ""
                EndOfRecs = True
                Exit Do
            Else
            End If
            If nowRec = rsR!name_rec Then
                rsR.MoveNext
                If rsR.EOF Then EndOfRecs = True
            Else
                nowRec = rsR!name_rec
                indrec = indrec + 1
                rsR.MoveNext
                If rsR.EOF Then EndOfRecs = True
                Exit Do
            End If
        Loop
        
RegDrvSorting:

        If EndOfRegDrv = True And EndOfRecs = True And EndOfObjects = True And endofclients = True Then
            GoTo CalculateQ
        Else
        End If

        'стринг за маркиране на 1 кола от рецепта nowRec от избрания обект nowObj от настоящия клиент nowClnt
        commR = "SELECT DISTINCT ON (reg_drv) reg_drv FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' AND name_clnt = '" & nowClnt & "' AND obj_clnt = '" & nowObj & "' AND name_rec = '" & nowRec & "' ORDER BY reg_drv ASC;"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)

        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
        Else
        End If

        Do While Not rsR.EOF
            For i = 1 To indreg
                rsR.MoveNext
            Next i
            If rsR.EOF Then
                nowRegDrv = ""
                EndOfRegDrv = True
                Exit Do
            Else
            End If
            If nowRegDrv = rsR!reg_drv Then
                rsR.MoveNext
                If rsR.EOF Then EndOfRegDrv = True
            Else
                nowRegDrv = rsR!reg_drv
                indreg = indreg + 1
                rsR.MoveNext
                If rsR.EOF Then EndOfRegDrv = True
                Exit Do
            End If
        Loop

CalculateQ:

        'стринг за маркиране на количеството на експедициите
        commR = "SELECT DISTINCT ON (exp_num) exp_q FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' AND name_clnt = '" & nowClnt & "' AND obj_clnt = '" & nowObj & "' AND name_rec = '" & nowRec & "' AND reg_drv = '" & nowRegDrv & "' ORDER BY exp_num ASC;"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)

        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
        Else
        End If
        
        tempQ = 0
        
        Do While Not rsR.EOF
            tempQ = tempQ + CSng(rDs(rsR!exp_q))
            totalQ = totalQ + CSng(rDs(rsR!exp_q))
            rsR.MoveNext
        Loop

CalculateQScr:

        'стринг за маркиране на количеството на циментите
        commR = "SELECT cem1_name, cem1i, cem2_name, cem2i, cem3_name, cem3i, cem4_name, cem4i FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' AND name_clnt = '" & nowClnt & "' AND obj_clnt = '" & nowObj & "' AND name_rec = '" & nowRec & "' AND reg_drv = '" & nowRegDrv & "';"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)

        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
        Else
        End If
        
        i = 1
        Do While tempScr(i) <> ""
            tempQScr(i) = 0
            i = i + 1
        Loop
        
        Do While Not rsR.EOF
            i = 1
            Do While tempScr(i) <> ""
                If tempScr(i) = rsR!cem1_name Then
                    tempQScr(i) = tempQScr(i) + CSng(rDs(rsR!cem1i))
                    totalQScr(i) = totalQScr(i) + CSng(rDs(rsR!cem1i))
                End If
                If tempScr(i) = rsR!cem2_name Then
                    tempQScr(i) = tempQScr(i) + CSng(rDs(rsR!cem2i))
                    totalQScr(i) = totalQScr(i) + CSng(rDs(rsR!cem2i))
                End If
                If tempScr(i) = rsR!cem3_name Then
                    tempQScr(i) = tempQScr(i) + CSng(rDs(rsR!cem3i))
                    totalQScr(i) = totalQScr(i) + CSng(rDs(rsR!cem3i))
                End If
                If tempScr(i) = rsR!cem4_name Then
                    tempQScr(i) = tempQScr(i) + CSng(rDs(rsR!cem4i))
                    totalQScr(i) = totalQScr(i) + CSng(rDs(rsR!cem4i))
                End If
                i = i + 1
            Loop
            rsR.MoveNext
        Loop

CalculateQIM:

        'стринг за маркиране на количеството на им
        commR = "SELECT im1_name, im1i, im2_name, im2i, im3_name, im3i, im4_name, im4i, im5_name, im5i, im6_name, im6i FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' AND name_clnt = '" & nowClnt & "' AND obj_clnt = '" & nowObj & "' AND name_rec = '" & nowRec & "' AND reg_drv = '" & nowRegDrv & "';"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)

        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
        Else
        End If
        
        i = 1
        Do While tempIM(i) <> ""
            tempQIM(i) = 0
            i = i + 1
        Loop
        
        Do While Not rsR.EOF
            i = 1
            Do While tempIM(i) <> ""
                If tempIM(i) = rsR!im1_name Then
                    tempQIM(i) = tempQIM(i) + CSng(rDs(rsR!im1i))
                    totalQIM(i) = totalQIM(i) + CSng(rDs(rsR!im1i))
                End If
                If tempIM(i) = rsR!im2_name Then
                    tempQIM(i) = tempQIM(i) + CSng(rDs(rsR!im2i))
                    totalQIM(i) = totalQIM(i) + CSng(rDs(rsR!im2i))
                End If
                If tempIM(i) = rsR!im3_name Then
                    tempQIM(i) = tempQIM(i) + CSng(rDs(rsR!im3i))
                    totalQIM(i) = totalQIM(i) + CSng(rDs(rsR!im3i))
                End If
                If tempIM(i) = rsR!im4_name Then
                    tempQIM(i) = tempQIM(i) + CSng(rDs(rsR!im4i))
                    totalQIM(i) = totalQIM(i) + CSng(rDs(rsR!im4i))
                End If
                If tempIM(i) = rsR!im5_name Then
                    tempQIM(i) = tempQIM(i) + CSng(rDs(rsR!im5i))
                    totalQIM(i) = totalQIM(i) + CSng(rDs(rsR!im5i))
                End If
                If tempIM(i) = rsR!im6_name Then
                    tempQIM(i) = tempQIM(i) + CSng(rDs(rsR!im6i))
                    totalQIM(i) = totalQIM(i) + CSng(rDs(rsR!im6i))
                End If
            i = i + 1
            Loop
            rsR.MoveNext
        Loop

CalculateQChem:

        'стринг за маркиране на количеството на им
        commR = "SELECT chem1_name, chem1i, chem2_name, chem2i, chem3_name, chem3i, chem4_name, chem4i, chem5_name, chem5i, chem6_name, chem6i FROM mix_result_bc" & MachineNumber & " WHERE stamp_date = '" & Day & "' AND name_clnt = '" & nowClnt & "' AND obj_clnt = '" & nowObj & "' AND name_rec = '" & nowRec & "' AND reg_drv = '" & nowRegDrv & "';"

        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)

        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
        Else
        End If
        
        i = 1
        Do While tempChem(i) <> ""
            tempQChem(i) = 0
            i = i + 1
        Loop
        
        Do While Not rsR.EOF
            i = 1
            Do While tempChem(i) <> ""
                If tempChem(i) = rsR!chem1_name Then
                    tempQChem(i) = tempQChem(i) + CSng(rDs(rsR!chem1i))
                    totalQChem(i) = totalQChem(i) + CSng(rDs(rsR!chem1i))
                End If
                If tempChem(i) = rsR!chem2_name Then
                    tempQChem(i) = tempQChem(i) + CSng(rDs(rsR!chem2i))
                    totalQChem(i) = totalQChem(i) + CSng(rDs(rsR!chem2i))
                End If
                If tempChem(i) = rsR!chem3_name Then
                    tempQChem(i) = tempQChem(i) + CSng(rDs(rsR!chem3i))
                    totalQChem(i) = totalQChem(i) + CSng(rDs(rsR!chem3i))
                End If
                If tempChem(i) = rsR!chem4_name Then
                    tempQChem(i) = tempQChem(i) + CSng(rDs(rsR!chem4i))
                    totalQChem(i) = totalQChem(i) + CSng(rDs(rsR!chem4i))
                End If
                If tempChem(i) = rsR!chem5_name Then
                    tempQChem(i) = tempQChem(i) + CSng(rDs(rsR!chem5i))
                    totalQChem(i) = totalQChem(i) + CSng(rDs(rsR!chem5i))
                End If
                If tempChem(i) = rsR!chem6_name Then
                    tempQChem(i) = tempQChem(i) + CSng(rDs(rsR!chem6i))
                    totalQChem(i) = totalQChem(i) + CSng(rDs(rsR!chem6i))
                End If
            i = i + 1
            Loop
            rsR.MoveNext
        Loop

    If nowClnt <> "" Then
        ind = ind + 1
        Set itmX = Me.lstDailyReport.ListItems.Add(ind, , nowClnt)
            itmX.SubItems(1) = nowObj
            itmX.SubItems(2) = nowRec
            itmX.SubItems(3) = ARound(tempQ, 3)
            itmX.SubItems(4) = nowRegDrv
            i = 1
            counting = 0
            countingn = 0
            Do While tempScr(i) <> ""
                tempQScr(i) = ARound(tempQScr(i) / 1000, 3)
                itmX.SubItems(i + 4) = tempQScr(i)
                counting = counting + 1
                i = i + 1
            Loop
            i = 1
            Do While tempIM(i) <> ""
                tempQIM(i) = ARound(tempQIM(i) / 1000, 3)
                itmX.SubItems(i + counting + 4) = tempQIM(i)
                countingn = countingn + 1
                i = i + 1
            Loop
            i = 1
            Do While tempChem(i) <> ""
                tempQChem(i) = ARound(tempQChem(i) / 1000, 5)
                itmX.SubItems(i + counting + countingn + 4) = tempQChem(i)
                i = i + 1
            Loop
            
            If EndOfRegDrv = False Then GoTo RegDrvSorting
            If EndOfRecs = False Then GoTo RecSorting
            If EndOfObjects = False Then GoTo ObjSorting
            
            If EndOfRegDrv = True And EndOfRecs = True And EndOfObjects = True And endofclients = True Then
                GoTo EndSorting
            Else
                GoTo StartSorting
            End If
    End If

EndSorting:
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing

        Set itmX = Me.lstDailyReport.ListItems.Add(ind + 1, , "Общо разход:")
            itmX.SubItems(3) = ARound(totalQ, 3)
            i = 1
            counting = 0
            countingn = 0
            Do While tempScr(i) <> ""
                totalQScr(i) = ARound(totalQScr(i) / 1000, 3)
                itmX.SubItems(i + 4) = totalQScr(i)
                counting = counting + 1
                i = i + 1
            Loop
            i = 1
            Do While tempIM(i) <> ""
                totalQIM(i) = ARound(totalQIM(i) / 1000, 3)
                itmX.SubItems(i + counting + 4) = totalQIM(i)
                countingn = countingn + 1
                i = i + 1
            Loop
            i = 1
            Do While tempChem(i) <> ""
                totalQChem(i) = ARound(totalQChem(i) / 1000, 5)
                itmX.SubItems(i + counting + countingn + 4) = totalQChem(i)
                i = i + 1
            Loop

    AutoColW Me.lstDailyReport
EndSub:
    MousePointer = vbDefault
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstDailyReport, 2, True, True, True, "Дневна справка Бетонов възел")
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstDailyReport)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Temp = Nothing
    frmStartRep.Show
End Sub

