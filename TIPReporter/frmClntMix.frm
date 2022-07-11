VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClntMix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmClntMix"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16710
   Icon            =   "frmClntMix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   16710
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
      Left            =   15600
      TabIndex        =   14
      Top             =   9000
      Width           =   735
   End
   Begin VB.ComboBox cmbClass 
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
      Left            =   9480
      TabIndex        =   10
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox cmbRec 
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
      Left            =   7080
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox cmbObj 
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
      Left            =   3720
      TabIndex        =   8
      Top             =   600
      Width           =   3255
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
      Left            =   9240
      TabIndex        =   5
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
      Left            =   5160
      TabIndex        =   4
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
      Left            =   13440
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cmbClnt 
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
   Begin MSComctlLib.ListView lstClntMix 
      Height          =   7575
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   13361
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
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
      Left            =   11640
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
      Format          =   104267779
      CurrentDate     =   41487.3333333333
      MaxDate         =   45291
      MinDate         =   41487
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      Caption         =   "lblClass"
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
      Left            =   9600
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "lblRec"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblObj 
      Alignment       =   2  'Center
      Caption         =   "lblObj"
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
      Left            =   3840
      TabIndex        =   11
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
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
      Left            =   11760
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblClnt 
      Alignment       =   2  'Center
      Caption         =   "lblClnt"
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
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmClntMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'справка за 1 клиент - филтър по име от резултатната таблица
'всеки ред от справката изобразява една експедиция през дадения период от формата

    Dim colx As ColumnHeader
    
    Me.Caption = repClntMix
    Me.lblClnt.Caption = uniClnt
    Me.lblObj.Caption = uniObj
    Me.lblRec.Caption = uniRec
    Me.lblClass.Caption = uniClass
    Me.lblDay.Caption = uniDate
    Me.btnLoad.Caption = btLoad
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    
    Me.btnPrint.Enabled = False
    Me.btnExport.Enabled = False
    
    Me.dtDay = Now
    
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
    
    Me.lstClntMix.ColumnHeaders.Clear
    Me.lstClntMix.ListItems.Clear
    
    Me.cmbClnt.Clear
    Me.cmbObj.Clear
    Me.cmbRec.Clear
    Me.cmbClass.Clear
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'номер експ./дата; празно
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'обем експ.; номер замес от експ.
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'празно; час замес
        colx.Width = 1

    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'номер заявка/дата; им1
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'заявено кол.; им2
        colx.Width = 1

    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'обект; им3
        colx.Width = 1

    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'разстояние; им4
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'празно; им5
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'рецепта име; им6
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'рецепта тип; ц1
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'клас якост; ц2
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'клас конс.; ц3
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'клас възд.; ц4
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'клас хл.; вода1
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'празно; вода2
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'водоплътност; хд1
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'празно; хд2
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'водач; хд3
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'кола-кап; хд4
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'празно; хд5
        colx.Width = 1
    
    Set colx = Me.lstClntMix.ColumnHeaders.Add() 'диспечер; хд6
        colx.Width = 1
        
    If rQinForm = 1 Then
        Set colx = Me.lstClntMix.ColumnHeaders.Add() 'празно; реален обем на замес
            colx.Width = 1
    End If
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (name_clnt) name_clnt FROM mix_result_bc" & MachineNumber & " ORDER BY name_clnt ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbClnt.AddItem rsR!name_clnt
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (obj_clnt)  obj_clnt FROM mix_result_bc" & MachineNumber & " ORDER BY obj_clnt ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbObj.AddItem rsR!obj_clnt
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (name_rec) name_rec FROM mix_result_bc" & MachineNumber & " ORDER BY name_rec ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbRec.AddItem rsR!name_rec
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (class_rec) class_rec FROM mix_result_bc" & MachineNumber & " ORDER BY class_rec ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbClass.AddItem rsR!class_rec
        rsR.MoveNext
    Loop
    

'    If frmStartRep.chMach1.Value = 1 And frmStartRep.chMach2.Value = 1 Then
'        Set rsR = cnR.Execute("SELECT DISTINCT ON (name_clnt) name_clnt, obj_clnt, name_rec, class_rec, name_drv FROM mix_result_bc" & MachineNumber & " ORDER BY name_clnt ASC;")
'
'        If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
'
'        Do While Not rsR.EOF
'            Me.cmbClnt.AddItem rsR!name_clnt
'            Me.cmbObj.AddItem rsR!obj_clnt
'            Me.cmbRec.AddItem rsR!name_rec
'            Me.cmbClass.AddItem rsR!class_rec
'            rsR.MoveNext
'        Loop
'    End If
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------

    Dim I, j As Integer
        
    For I = 0 To Me.cmbClnt.listCount - 2 Step 1
        For j = Me.cmbClnt.listCount - 1 To I + 1 Step -1
            If Me.cmbClnt.List(I) = Me.cmbClnt.List(j) Then
                Me.cmbClnt.RemoveItem (j)
            End If
        Next
    Next
        
    For I = 0 To Me.cmbObj.listCount - 2 Step 1
        For j = Me.cmbObj.listCount - 1 To I + 1 Step -1
            If Me.cmbObj.List(I) = Me.cmbObj.List(j) Then
                Me.cmbObj.RemoveItem (j)
            End If
        Next
    Next
    
    For I = 0 To Me.cmbRec.listCount - 2 Step 1
        For j = Me.cmbRec.listCount - 1 To I + 1 Step -1
            If Me.cmbRec.List(I) = Me.cmbRec.List(j) Then
                Me.cmbRec.RemoveItem (j)
            End If
        Next
    Next

    For I = 0 To Me.cmbClass.listCount - 2 Step 1
        For j = Me.cmbClass.listCount - 1 To I + 1 Step -1
            If Me.cmbClass.List(I) = Me.cmbClass.List(j) Then
                Me.cmbClass.RemoveItem (j)
            End If
        Next
    Next

End Sub

Private Sub cmbClnt_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbClnt, KeyAscii, True)
EndSub:
End Sub

Private Sub cmbObj_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbObj, KeyAscii, True)
EndSub:
End Sub

Private Sub cmbRec_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbRec, KeyAscii, True)
EndSub:
End Sub

Private Sub cmbClass_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbClass, KeyAscii, True)
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

    Me.lstClntMix.ListItems.Clear
    
    If Me.cmbClnt.Text <> "" Then
        
        Dim Temp As Result
        Set Temp = New Result
        
'-----------------------Start postgreSQL-----------------------------------
        Dim cnR As New ADODB.Connection
        Dim rsR As New Recordset
        Dim commR As String
        Dim commR1 As String
        Dim commR2 As String
        Dim commR3 As String
        Dim commR4 As String
        Dim commRord As String
        Dim frstFlag As Boolean
        Dim begFlag As Boolean
    
        Dim iT As Integer
    
        Dim subT(1 To 19) As Single
    
        Dim allT(1 To 19) As Single
    
        Dim expCounter As Integer
        Dim mixCounter As Integer
        Dim tmixCounter As Integer
        Dim Day As String
        Dim ind As Integer
    
        Day = Format(Me.dtDay.Value, "DD-MM-YYYY")
        ind = 0
        
AgainOther:

        cnR.ConnectionTimeout = 30
        cnR.Open ConStr
        MousePointer = vbHourglass

        'стринг за маркираме набор от записи само по клиент
        commR1 = "SELECT * FROM mix_result_bc" & MachineNumber & " WHERE name_clnt = '" _
        & Me.cmbClnt.Text & "' AND stamp_date = '" & Day & ""

        'стринг за добавка при търсенето за обект
        commR2 = "' AND obj_clnt = '" & Me.cmbObj.Text & ""
    
        'стринг за добавка при търсенето за име рецепта
        commR3 = "' AND name_rec = '" & Me.cmbRec.Text & ""
    
        'стринг за добавка при търсенето за клас якост
        commR4 = "' AND class_rec = '" & Me.cmbClass.Text & ""
    
        'стринг за добавка при сортирането
        commRord = "' ORDER BY mix_num, stamp_date ASC;"
    
        'записваме стринг за търсене само по основен филтър - име клиент
        commR = commR1
    
        'резултен стринг за търсенето при добавен филтър за обект
        If Me.cmbObj.Text <> "" Then commR = commR & commR2
    
        'резултен стринг за търсенето при добавен филтър за име рецепта
        If Me.cmbRec.Text <> "" Then commR = commR & commR3
    
        'резултен стринг за търсенето при добавен филтър за клас на якост
        If Me.cmbClass.Text <> "" Then commR = commR & commR4
    
        'добавка на стринга за сортиране
        commR = commR + commRord
    
        'изпълняваме търсенето
        Set rsR = cnR.Execute(commR)
        
        'отиваме на първия запис
        If Not rsR.EOF And Not rsR.BOF Then
            rsR.MoveFirst
            Me.btnPrint.Enabled = True
            Me.btnExport.Enabled = True
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "Машина " & MachineNumber)
            have = have + 1
            If ag = True Then
                tempTExVol = 0
                Dim k As Integer
                For k = 1 To 19
                    allT(k) = 0
                Next k
            Else
                firstM = True
            End If
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

        'старт на групирането
        begFlag = True
        frstFlag = True
    
        'нулираме променливите носещи данните за таблицата
        expCounter = 0
        mixCounter = 0
        tmixCounter = 0
        Temp.TotalQuant = 0
    
        Do While Not rsR.EOF
            'изчитане на данните
            If Temp.ExpNum <> Val(rsR!exp_num) Then frstFlag = True
        
            Temp.ExpNum = Val(rsR!exp_num)
            Temp.StampDate = rsR!stamp_date
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
        
            Temp.IMname(1) = rsR!im1_name
            Temp.IMname(2) = rsR!im2_name
            Temp.IMname(3) = rsR!im3_name
            Temp.IMname(4) = rsR!im4_name
            Temp.IMname(5) = rsR!im5_name
            Temp.IMname(6) = rsR!im6_name
            Temp.SCRname(1) = rsR!cem1_name
            Temp.SCRname(2) = rsR!cem2_name
            Temp.SCRname(3) = rsR!cem3_name
            Temp.SCRname(4) = rsR!cem4_name
            Temp.WATname(1) = rsR!wat1_name
            Temp.WATname(2) = rsR!wat2_name
            Temp.CHEMname(1) = rsR!chem1_name
            Temp.CHEMname(2) = rsR!chem2_name
            Temp.CHEMname(3) = rsR!chem3_name
            Temp.CHEMname(4) = rsR!chem4_name
            Temp.CHEMname(5) = rsR!chem5_name
            Temp.CHEMname(6) = rsR!chem6_name
            Temp.IMmeasured(1) = rsR!im1i
            Temp.IMmeasured(2) = rsR!im2i
            Temp.IMmeasured(3) = rsR!im3i
            Temp.IMmeasured(4) = rsR!im4i
            Temp.IMmeasured(5) = rsR!im5i
            Temp.IMmeasured(6) = rsR!im6i
            Temp.SCRmeasured(1) = rsR!cem1i
            Temp.SCRmeasured(2) = rsR!cem2i
            Temp.SCRmeasured(3) = rsR!cem3i
            Temp.SCRmeasured(4) = rsR!cem4i
            Temp.WATmeasured(1) = rsR!wat1i
            Temp.WATmeasured(2) = rsR!wat2i
            Temp.CHEMmeasured(1) = CSng(rDs(rsR!chem1i))
            Temp.CHEMmeasured(2) = CSng(rDs(rsR!chem2i))
            Temp.CHEMmeasured(3) = CSng(rDs(rsR!chem3i))
            Temp.CHEMmeasured(4) = CSng(rDs(rsR!chem4i))
            Temp.CHEMmeasured(5) = CSng(rDs(rsR!chem5i))
            Temp.CHEMmeasured(6) = CSng(rDs(rsR!chem6i))
            
            If rQinForm = 1 Then
                Temp.TotalQuant = CSng(rDs(rsR!total_vol))
            End If
        
            'зареждане на данни за всяка една експедиция в ListView
            If frstFlag = True Then
                If begFlag = False Then
            
                    'междинни тотали на експедиците
                    ind = ind + 1
                    Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
                        itmX.SubItems(2) = uniMixes & ": " & mixCounter
                        For iT = ind - mixCounter To ind - 1
                            subT(1) = subT(1) + Val(Me.lstClntMix.ListItems(iT).SubItems(3))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(3) <> "" Then itmX.SubItems(3) = subT(1)
                        allT(1) = allT(1) + subT(1)
                        subT(1) = 0
                
                        For iT = ind - mixCounter To ind - 1
                            subT(2) = subT(2) + Val(Me.lstClntMix.ListItems(iT).SubItems(4))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(4) <> "" Then itmX.SubItems(4) = subT(2)
                        allT(2) = allT(2) + subT(2)
                        subT(2) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(3) = subT(3) + Val(Me.lstClntMix.ListItems(iT).SubItems(5))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(5) <> "" Then itmX.SubItems(5) = subT(3)
                        allT(3) = allT(3) + subT(3)
                        subT(3) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(4) = subT(4) + Val(Me.lstClntMix.ListItems(iT).SubItems(6))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(6) <> "" Then itmX.SubItems(6) = subT(4)
                        allT(4) = allT(4) + subT(4)
                        subT(4) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(5) = subT(5) + Val(Me.lstClntMix.ListItems(iT).SubItems(7))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(7) <> "" Then itmX.SubItems(7) = subT(5)
                        allT(5) = allT(5) + subT(5)
                        subT(5) = 0
                        
                        For iT = ind - mixCounter To ind - 1
                            subT(6) = subT(6) + Val(Me.lstClntMix.ListItems(iT).SubItems(8))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(8) <> "" Then itmX.SubItems(8) = subT(6)
                        allT(6) = allT(6) + subT(6)
                        subT(6) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(7) = subT(7) + Val(Me.lstClntMix.ListItems(iT).SubItems(9))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(9) <> "" Then itmX.SubItems(9) = subT(7)
                        allT(7) = allT(7) + subT(7)
                        subT(7) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(8) = subT(8) + Val(Me.lstClntMix.ListItems(iT).SubItems(10))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(10) <> "" Then itmX.SubItems(10) = subT(8)
                        allT(8) = allT(8) + subT(8)
                        subT(8) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(9) = subT(9) + Val(Me.lstClntMix.ListItems(iT).SubItems(11))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(11) <> "" Then itmX.SubItems(11) = subT(9)
                        allT(9) = allT(9) + subT(9)
                        subT(9) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(10) = subT(10) + Val(Me.lstClntMix.ListItems(iT).SubItems(12))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(12) <> "" Then itmX.SubItems(12) = subT(10)
                        allT(10) = allT(10) + subT(10)
                        subT(10) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(11) = subT(11) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(13)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(13) <> "" Then itmX.SubItems(13) = subT(11)
                        allT(11) = allT(11) + subT(11)
                        subT(11) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(12) = subT(12) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(14)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(14) <> "" Then itmX.SubItems(14) = subT(12)
                        allT(12) = allT(12) + subT(12)
                        subT(12) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(13) = subT(13) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(15)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(15) <> "" Then itmX.SubItems(15) = subT(13)
                        allT(13) = allT(13) + subT(13)
                        subT(13) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(14) = subT(14) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(16)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(16) <> "" Then itmX.SubItems(16) = subT(14)
                        allT(14) = allT(14) + subT(14)
                        subT(14) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(15) = subT(15) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(17)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(17) <> "" Then itmX.SubItems(17) = subT(15)
                        allT(15) = allT(15) + subT(15)
                        subT(15) = 0
                    
                        For iT = ind - mixCounter To ind - 1
                            subT(16) = subT(16) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(18)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(18) <> "" Then itmX.SubItems(18) = subT(16)
                        allT(16) = allT(16) + subT(16)
                        subT(16) = 0
                        
                        For iT = ind - mixCounter To ind - 1
                            subT(16) = subT(17) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(19)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(18) <> "" Then itmX.SubItems(19) = subT(17)
                        allT(17) = allT(17) + subT(17)
                        subT(17) = 0
                        
                        For iT = ind - mixCounter To ind - 1
                            subT(17) = subT(18) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(20)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(19) <> "" Then itmX.SubItems(20) = subT(18)
                        allT(18) = allT(18) + subT(18)
                        subT(18) = 0
                        
                        If rQinForm = 1 Then
                            For iT = ind - mixCounter To ind - 1
                                subT(19) = subT(19) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(21)))
                            Next iT
                            If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(21) <> "" Then itmX.SubItems(21) = ARound(subT(19), 3)
                            allT(19) = allT(19) + subT(19)
                            subT(19) = 0
                        End If
                    
                    'празен ред
                    ind = ind + 1
                    Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
                End If
            
                expCounter = expCounter + 1
                tmixCounter = tmixCounter + mixCounter
                mixCounter = 0
                frstFlag = False
                begFlag = False
                
                'заглавки за експедиция
                ind = ind + 1
                Set itmX = Me.lstClntMix.ListItems.Add(ind, , uniEx & " " & uniNr & " / " & uniDate)
                    itmX.SubItems(1) = uniQ & " " & uniEx
                    itmX.SubItems(2) = uniOrd & " " & uniNr & " / " & uniDate
                    itmX.SubItems(3) = uniOrdered & " " & uniQ
                    itmX.SubItems(4) = uniObj
                    itmX.SubItems(5) = uniKmShort
                    itmX.SubItems(7) = uniRec
                    itmX.SubItems(8) = uniRecType
                    itmX.SubItems(9) = uniClass
                    itmX.SubItems(10) = uniClassK
                    itmX.SubItems(11) = uniClassV
                    itmX.SubItems(12) = uniClassH
                    itmX.SubItems(13) = uniClassP
                    itmX.SubItems(16) = uniDrv
                    itmX.SubItems(17) = uniDrvReg & " - " & uniCap
                    itmX.SubItems(19) = uniDisp
            
                'данни за експедиция
                ind = ind + 1
                Set itmX = Me.lstClntMix.ListItems.Add(ind, , "M" & MachineNumber & "-" & Temp.ExpNum & " / " & Temp.StampDate)
                    itmX.SubItems(1) = Temp.ExpQuant
                    itmX.SubItems(2) = Temp.OrderCode & " / " & Temp.OrderDate 'с датата
                    itmX.SubItems(3) = Temp.OrderQuant
                    itmX.SubItems(4) = Temp.ClntWorksite
                    itmX.SubItems(5) = Temp.WorksiteDist
                    itmX.SubItems(7) = Temp.RecTitle
                    itmX.SubItems(8) = Temp.RecKind
                    itmX.SubItems(9) = Temp.RecClass
                    itmX.SubItems(10) = Temp.RecClassK
                    itmX.SubItems(11) = Temp.RecClassV
                    itmX.SubItems(12) = Temp.RecClassH
                    itmX.SubItems(13) = Temp.RecClassP
                    itmX.SubItems(16) = Temp.DrvTitle
                    itmX.SubItems(17) = Temp.DrvCarNum & " - " & Temp.DrvCapacity
                    itmX.SubItems(19) = Temp.DispName
                
                tempTExVol = tempTExVol + Temp.ExpQuant
                
                'заглавки за замесите
                ind = ind + 1
                Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
                    itmX.SubItems(1) = uniMix & " " & uniNr
                    itmX.SubItems(2) = uniHourOut
                    If Temp.IMname(1) <> "0" And Temp.IMname(1) <> uniEmpty And Temp.IMname(1) <> "" Then
                        itmX.SubItems(3) = Temp.IMname(1)
                    Else
                    End If
                    If Temp.IMname(2) <> "0" And Temp.IMname(2) <> uniEmpty And Temp.IMname(2) <> "" Then
                        itmX.SubItems(4) = Temp.IMname(2)
                    Else
                    End If
                    If Temp.IMname(3) <> "0" And Temp.IMname(3) <> uniEmpty And Temp.IMname(3) <> "" Then
                        itmX.SubItems(5) = Temp.IMname(3)
                    Else
                    End If
                    If Temp.IMname(4) <> "0" And Temp.IMname(4) <> uniEmpty And Temp.IMname(4) <> "" Then
                        itmX.SubItems(6) = Temp.IMname(4)
                    Else
                    End If
                    If Temp.IMname(5) <> "0" And Temp.IMname(5) <> uniEmpty And Temp.IMname(5) <> "" Then
                        itmX.SubItems(7) = Temp.IMname(5)
                    Else
                    End If
                    If Temp.IMname(6) <> "0" And Temp.IMname(6) <> uniEmpty And Temp.IMname(6) <> "" Then
                        itmX.SubItems(8) = Temp.IMname(6)
                    Else
                    End If
                    If Temp.SCRname(1) <> "0" And Temp.SCRname(1) <> uniEmpty And Temp.SCRname(1) <> "" Then
                        itmX.SubItems(9) = Temp.SCRname(1)
                    Else
                    End If
                    If Temp.SCRname(2) <> "0" And Temp.SCRname(2) <> uniEmpty And Temp.SCRname(2) <> "" Then
                        itmX.SubItems(10) = Temp.SCRname(2)
                    Else
                    End If
                    If Temp.SCRname(3) <> "0" And Temp.SCRname(3) <> uniEmpty And Temp.SCRname(3) <> "" Then
                        itmX.SubItems(11) = Temp.SCRname(3)
                    Else
                    End If
                    If Temp.SCRname(4) <> "0" And Temp.SCRname(4) <> uniEmpty And Temp.SCRname(4) <> "" Then
                        itmX.SubItems(12) = Temp.SCRname(4)
                    Else
                    End If
                    If Temp.WATname(1) <> "0" And Temp.WATname(1) <> uniEmpty And Temp.WATname(1) <> "" Then
                        itmX.SubItems(13) = Temp.WATname(1)
                    Else
                    End If
                    If Temp.WATname(2) <> "0" And Temp.WATname(2) <> uniEmpty And Temp.WATname(2) <> "" Then
                        itmX.SubItems(14) = Temp.WATname(2)
                    Else
                    End If
                    If Temp.CHEMname(1) <> "0" And Temp.CHEMname(1) <> uniEmpty And Temp.CHEMname(1) <> "" Then
                        itmX.SubItems(15) = Temp.CHEMname(1)
                    Else
                    End If
                    If Temp.CHEMname(2) <> "0" And Temp.CHEMname(2) <> uniEmpty And Temp.CHEMname(2) <> "" Then
                        itmX.SubItems(16) = Temp.CHEMname(2)
                    Else
                    End If
                    If Temp.CHEMname(3) <> "0" And Temp.CHEMname(3) <> uniEmpty And Temp.CHEMname(3) <> "" Then
                        itmX.SubItems(17) = Temp.CHEMname(3)
                    Else
                    End If
                    If Temp.CHEMname(4) <> "0" And Temp.CHEMname(4) <> uniEmpty And Temp.CHEMname(4) <> "" Then
                        itmX.SubItems(18) = Temp.CHEMname(4)
                    Else
                    End If
                    If Temp.CHEMname(5) <> "0" And Temp.CHEMname(5) <> uniEmpty And Temp.CHEMname(5) <> "" Then
                        itmX.SubItems(19) = Temp.CHEMname(5)
                    Else
                    End If
                    If Temp.CHEMname(6) <> "0" And Temp.CHEMname(6) <> uniEmpty And Temp.CHEMname(6) <> "" Then
                        itmX.SubItems(20) = Temp.CHEMname(6)
                    Else
                    End If
                    If rQinForm = 1 Then
                        itmX.SubItems(21) = uniQ
                    End If
                
            End If
            
            mixCounter = mixCounter + 1
        
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
                If rsR!avstat = True Then
                    itmX.SubItems(1) = mixCounter & " (k)"
                Else
                    itmX.SubItems(1) = mixCounter
                End If
                itmX.SubItems(2) = Right(Temp.MixReadyTime, 8)
                If Temp.IMname(1) <> "0" And Temp.IMname(1) <> uniEmpty And Temp.IMname(1) <> "" Then
                    itmX.SubItems(3) = Temp.IMmeasured(1)
                Else
                End If
                If Temp.IMname(2) <> "0" And Temp.IMname(2) <> uniEmpty And Temp.IMname(2) <> "" Then
                    itmX.SubItems(4) = Temp.IMmeasured(2)
                Else
                End If
                If Temp.IMname(3) <> "0" And Temp.IMname(3) <> uniEmpty And Temp.IMname(3) <> "" Then
                    itmX.SubItems(5) = Temp.IMmeasured(3)
                Else
                End If
                If Temp.IMname(4) <> "0" And Temp.IMname(4) <> uniEmpty And Temp.IMname(4) <> "" Then
                    itmX.SubItems(6) = Temp.IMmeasured(4)
                Else
                End If
                If Temp.IMname(5) <> "0" And Temp.IMname(5) <> uniEmpty And Temp.IMname(5) <> "" Then
                    itmX.SubItems(7) = Temp.IMmeasured(5)
                Else
                End If
                If Temp.IMname(6) <> "0" And Temp.IMname(6) <> uniEmpty And Temp.IMname(6) <> "" Then
                    itmX.SubItems(8) = Temp.IMmeasured(6)
                Else
                End If
                If Temp.SCRname(1) <> "0" And Temp.SCRname(1) <> uniEmpty And Temp.SCRname(1) <> "" Then
                    itmX.SubItems(9) = Temp.SCRmeasured(1)
                Else
                End If
                If Temp.SCRname(2) <> "0" And Temp.SCRname(2) <> uniEmpty And Temp.SCRname(2) <> "" Then
                    itmX.SubItems(10) = Temp.SCRmeasured(2)
                Else
                End If
                If Temp.SCRname(3) <> "0" And Temp.SCRname(3) <> uniEmpty And Temp.SCRname(3) <> "" Then
                    itmX.SubItems(11) = Temp.SCRmeasured(3)
                Else
                End If
                If Temp.SCRname(4) <> "0" And Temp.SCRname(4) <> uniEmpty And Temp.SCRname(4) <> "" Then
                    itmX.SubItems(12) = Temp.SCRmeasured(4)
                Else
                End If
                If Temp.WATname(1) <> "0" And Temp.WATname(1) <> uniEmpty And Temp.WATname(1) <> "" Then
                    itmX.SubItems(13) = Temp.WATmeasured(1)
                Else
                End If
                If Temp.WATname(2) <> "0" And Temp.WATname(2) <> uniEmpty And Temp.WATname(2) <> "" Then
                    itmX.SubItems(14) = Temp.WATmeasured(2)
                Else
                End If
                If Temp.CHEMname(1) <> "0" And Temp.CHEMname(1) <> uniEmpty And Temp.CHEMname(1) <> "" Then
                    itmX.SubItems(15) = Temp.CHEMmeasured(1)
                Else
                End If
                If Temp.CHEMname(2) <> "0" And Temp.CHEMname(2) <> uniEmpty And Temp.CHEMname(2) <> "" Then
                    itmX.SubItems(16) = Temp.CHEMmeasured(2)
                Else
                End If
                If Temp.CHEMname(3) <> "0" And Temp.CHEMname(3) <> uniEmpty And Temp.CHEMname(3) <> "" Then
                    itmX.SubItems(17) = Temp.CHEMmeasured(3)
                Else
                End If
                If Temp.CHEMname(4) <> "0" And Temp.CHEMname(4) <> uniEmpty And Temp.CHEMname(4) <> "" Then
                    itmX.SubItems(18) = Temp.CHEMmeasured(4)
                Else
                End If
                If Temp.CHEMname(5) <> "0" And Temp.CHEMname(5) <> uniEmpty And Temp.CHEMname(5) <> "" Then
                    itmX.SubItems(19) = Temp.CHEMmeasured(5)
                Else
                End If
                If Temp.CHEMname(6) <> "0" And Temp.CHEMname(6) <> uniEmpty And Temp.CHEMname(6) <> "" Then
                    itmX.SubItems(20) = Temp.CHEMmeasured(6)
                Else
                End If
                If rQinForm = 1 Then
                    itmX.SubItems(21) = ARound(Temp.TotalQuant, 3)
                End If
            
            rsR.MoveNext 'местим на следващ запис
        
            If rsR.EOF Then
                'междинни тотали на последната експедиция
                ind = ind + 1
                Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
                    itmX.SubItems(2) = uniMixes & ": " & mixCounter
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(1) = subT(1) + Val(Me.lstClntMix.ListItems(iT).SubItems(3))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(3) <> "" Then itmX.SubItems(3) = subT(1)
                    allT(1) = allT(1) + subT(1)
                    subT(1) = 0
                
                    For iT = ind - mixCounter To ind - 1
                        subT(2) = subT(2) + Val(Me.lstClntMix.ListItems(iT).SubItems(4))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(4) <> "" Then itmX.SubItems(4) = subT(2)
                    allT(2) = allT(2) + subT(2)
                    subT(2) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(3) = subT(3) + Val(Me.lstClntMix.ListItems(iT).SubItems(5))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(5) <> "" Then itmX.SubItems(5) = subT(3)
                    allT(3) = allT(3) + subT(3)
                    subT(3) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(4) = subT(4) + Val(Me.lstClntMix.ListItems(iT).SubItems(6))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(6) <> "" Then itmX.SubItems(6) = subT(4)
                    allT(4) = allT(4) + subT(4)
                    subT(4) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(5) = subT(5) + Val(Me.lstClntMix.ListItems(iT).SubItems(7))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(7) <> "" Then itmX.SubItems(7) = subT(5)
                    allT(5) = allT(5) + subT(5)
                    subT(5) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(6) = subT(6) + Val(Me.lstClntMix.ListItems(iT).SubItems(8))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(8) <> "" Then itmX.SubItems(8) = subT(6)
                    allT(6) = allT(6) + subT(6)
                    subT(6) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(7) = subT(7) + Val(Me.lstClntMix.ListItems(iT).SubItems(9))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(9) <> "" Then itmX.SubItems(9) = subT(7)
                    allT(7) = allT(7) + subT(7)
                    subT(7) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(8) = subT(8) + Val(Me.lstClntMix.ListItems(iT).SubItems(10))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(10) <> "" Then itmX.SubItems(10) = subT(8)
                    allT(8) = allT(8) + subT(8)
                    subT(8) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(9) = subT(9) + Val(Me.lstClntMix.ListItems(iT).SubItems(11))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(11) <> "" Then itmX.SubItems(11) = subT(9)
                    allT(9) = allT(9) + subT(9)
                    subT(9) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(10) = subT(10) + Val(Me.lstClntMix.ListItems(iT).SubItems(12))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(12) <> "" Then itmX.SubItems(12) = subT(10)
                    allT(10) = allT(10) + subT(10)
                    subT(10) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(11) = subT(11) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(13)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(13) <> "" Then itmX.SubItems(13) = subT(11)
                    allT(11) = allT(11) + subT(11)
                    subT(11) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(12) = subT(12) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(14)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(14) <> "" Then itmX.SubItems(14) = subT(12)
                    allT(12) = allT(12) + subT(12)
                    subT(12) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(13) = subT(13) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(15)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(15) <> "" Then itmX.SubItems(15) = subT(13)
                    allT(13) = allT(13) + subT(13)
                    subT(13) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(14) = subT(14) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(16)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(16) <> "" Then itmX.SubItems(16) = subT(14)
                    allT(14) = allT(14) + subT(14)
                    subT(14) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(15) = subT(15) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(17)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(17) <> "" Then itmX.SubItems(17) = subT(15)
                    allT(15) = allT(15) + subT(15)
                    subT(15) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(16) = subT(16) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(18)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(18) <> "" Then itmX.SubItems(18) = subT(16)
                    allT(16) = allT(16) + subT(16)
                    subT(16) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(17) = subT(17) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(19)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(19) <> "" Then itmX.SubItems(19) = subT(17)
                    allT(17) = allT(17) + subT(17)
                    subT(17) = 0
                    
                    For iT = ind - mixCounter To ind - 1
                        subT(18) = subT(18) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(20)))
                    Next iT
                    If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(20) <> "" Then itmX.SubItems(20) = subT(18)
                    allT(18) = allT(18) + subT(18)
                    subT(18) = 0
                    
                    If rQinForm = 1 Then
                        For iT = ind - mixCounter To ind - 1
                            subT(19) = subT(19) + CSng(rDs(Me.lstClntMix.ListItems(iT).SubItems(21)))
                        Next iT
                        If Me.lstClntMix.ListItems(ind - mixCounter - 1).SubItems(21) <> "" Then itmX.SubItems(21) = ARound(subT(19), 3)
                        allT(19) = allT(19) + subT(19)
                        subT(19) = 0
                    End If
            
                tmixCounter = tmixCounter + mixCounter
            
                'празен ред
                ind = ind + 1
                Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
            End If
        Loop
        
        MousePointer = vbDefault
        rsR.Close
        Set rsR = Nothing
        cnR.Close
        Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------
    
        'след като прекъснем връзката сумираме тоталите
    
    
        'първо въвеждаме един празен ред
        ind = ind + 1
        Set itmX = Me.lstClntMix.ListItems.Add(ind, , "-------------------------")
    
        'заглавки за тоталите
        ind = ind + 1
        Set itmX = Me.lstClntMix.ListItems.Add(ind, , uniTotal)
            itmX.SubItems(1) = uniQ & " " & uniEx
            itmX.SubItems(2) = uniMixes
            If Temp.IMname(1) <> "0" And Temp.IMname(1) <> uniEmpty And Temp.IMname(1) <> "" Then
                itmX.SubItems(3) = Temp.IMname(1)
            Else
            End If
            If Temp.IMname(2) <> "0" And Temp.IMname(2) <> uniEmpty And Temp.IMname(2) <> "" Then
                itmX.SubItems(4) = Temp.IMname(2)
            Else
            End If
            If Temp.IMname(3) <> "0" And Temp.IMname(3) <> uniEmpty And Temp.IMname(3) <> "" Then
                itmX.SubItems(5) = Temp.IMname(3)
            Else
            End If
            If Temp.IMname(4) <> "0" And Temp.IMname(4) <> uniEmpty And Temp.IMname(4) <> "" Then
                itmX.SubItems(6) = Temp.IMname(4)
            Else
            End If
            If Temp.IMname(5) <> "0" And Temp.IMname(5) <> uniEmpty And Temp.IMname(5) <> "" Then
                itmX.SubItems(7) = Temp.IMname(5)
            Else
            End If
            If Temp.IMname(6) <> "0" And Temp.IMname(6) <> uniEmpty And Temp.IMname(6) <> "" Then
                itmX.SubItems(8) = Temp.IMname(6)
            Else
            End If
            If Temp.SCRname(1) <> "0" And Temp.SCRname(1) <> uniEmpty And Temp.SCRname(1) <> "" Then
                itmX.SubItems(9) = Temp.SCRname(1)
            Else
            End If
            If Temp.SCRname(2) <> "0" And Temp.SCRname(2) <> uniEmpty And Temp.SCRname(2) <> "" Then
                itmX.SubItems(10) = Temp.SCRname(2)
            Else
            End If
            If Temp.SCRname(3) <> "0" And Temp.SCRname(3) <> uniEmpty And Temp.SCRname(3) <> "" Then
                itmX.SubItems(11) = Temp.SCRname(3)
            Else
            End If
            If Temp.SCRname(4) <> "0" And Temp.SCRname(4) <> uniEmpty And Temp.SCRname(4) <> "" Then
                itmX.SubItems(12) = Temp.SCRname(4)
            Else
            End If
            If Temp.WATname(1) <> "0" And Temp.WATname(1) <> uniEmpty And Temp.WATname(1) <> "" Then
                itmX.SubItems(13) = Temp.WATname(1)
            Else
            End If
            If Temp.WATname(2) <> "0" And Temp.WATname(2) <> uniEmpty And Temp.WATname(2) <> "" Then
                itmX.SubItems(14) = Temp.WATname(2)
            Else
            End If
            If Temp.CHEMname(1) <> "0" And Temp.CHEMname(1) <> uniEmpty And Temp.CHEMname(1) <> "" Then
                itmX.SubItems(15) = Temp.CHEMname(1)
            Else
            End If
            If Temp.CHEMname(2) <> "0" And Temp.CHEMname(2) <> uniEmpty And Temp.CHEMname(2) <> "" Then
                itmX.SubItems(16) = Temp.CHEMname(2)
            Else
            End If
            If Temp.CHEMname(3) <> "0" And Temp.CHEMname(3) <> uniEmpty And Temp.CHEMname(3) <> "" Then
                itmX.SubItems(17) = Temp.CHEMname(3)
            Else
            End If
            If Temp.CHEMname(4) <> "0" And Temp.CHEMname(4) <> uniEmpty And Temp.CHEMname(4) <> "" Then
                itmX.SubItems(18) = Temp.CHEMname(4)
            Else
            End If
            If Temp.CHEMname(5) <> "0" And Temp.CHEMname(5) <> uniEmpty And Temp.CHEMname(5) <> "" Then
                itmX.SubItems(19) = Temp.CHEMname(5)
            Else
            End If
            If Temp.CHEMname(6) <> "0" And Temp.CHEMname(6) <> uniEmpty And Temp.CHEMname(6) <> "" Then
                itmX.SubItems(20) = Temp.CHEMname(6)
            Else
            End If
            
            If rQinForm = 1 Then
                itmX.SubItems(21) = uniQ
            End If
    
        'след него въвеждаме тотали
        ind = ind + 1
        If MachineNumber = 1 Then
            indT1 = ind
        ElseIf MachineNumber = 2 Then
            indT2 = ind
        End If
        Set itmX = Me.lstClntMix.ListItems.Add(ind, , expCounter & " " & uniEx)
            itmX.SubItems(1) = tempTExVol
            itmX.SubItems(2) = tmixCounter
            If Me.lstClntMix.ListItems(ind - 1).SubItems(3) <> "" Then itmX.SubItems(3) = allT(1)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(4) <> "" Then itmX.SubItems(4) = allT(2)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(5) <> "" Then itmX.SubItems(5) = allT(3)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(6) <> "" Then itmX.SubItems(6) = allT(4)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(7) <> "" Then itmX.SubItems(7) = allT(5)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(8) <> "" Then itmX.SubItems(8) = allT(6)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(9) <> "" Then itmX.SubItems(9) = allT(7)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(10) <> "" Then itmX.SubItems(10) = allT(8)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(11) <> "" Then itmX.SubItems(11) = allT(9)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(12) <> "" Then itmX.SubItems(12) = allT(10)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(13) <> "" Then itmX.SubItems(13) = allT(11)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(14) <> "" Then itmX.SubItems(14) = allT(12)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(15) <> "" Then itmX.SubItems(15) = allT(13)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(16) <> "" Then itmX.SubItems(16) = allT(14)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(17) <> "" Then itmX.SubItems(17) = allT(15)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(18) <> "" Then itmX.SubItems(18) = allT(16)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(19) <> "" Then itmX.SubItems(19) = allT(17)
            If Me.lstClntMix.ListItems(ind - 1).SubItems(20) <> "" Then itmX.SubItems(20) = allT(18)
            If rQinForm = 1 Then
                If Me.lstClntMix.ListItems(ind - 1).SubItems(21) <> "" Then itmX.SubItems(21) = ARound(allT(19), 3)
            End If
        
        tempTExVolT = tempTExVolT + tempTExVol
        tmixCounterT = tmixCounterT + tmixCounter
        
        If ag = False And MachineNumber = 1 And frmStartRep.chMach2.Value = 1 Then
            MachineNumber = 2
            ag = True
            
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
            GoTo AgainOther
        End If
        
        If have = 2 Then
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "")
            ind = ind + 1
            Set itmX = Me.lstClntMix.ListItems.Add(ind, , "ОБЩО ЗА ДВЕТЕ")
                itmX.SubItems(1) = uniQ & " " & uniEx & ": " & tempTExVolT
                itmX.SubItems(2) = uniMixes & ": " & tmixCounterT
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(3)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(3))) > 0 Then _
                itmX.SubItems(3) = Val(Me.lstClntMix.ListItems(indT1).SubItems(3)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(3))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(4)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(4))) > 0 Then _
                itmX.SubItems(4) = Val(Me.lstClntMix.ListItems(indT1).SubItems(4)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(4))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(5)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(5))) > 0 Then _
                itmX.SubItems(5) = Val(Me.lstClntMix.ListItems(indT1).SubItems(5)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(5))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(6)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(6))) > 0 Then _
                itmX.SubItems(6) = Val(Me.lstClntMix.ListItems(indT1).SubItems(6)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(6))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(7)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(7))) > 0 Then _
                itmX.SubItems(7) = Val(Me.lstClntMix.ListItems(indT1).SubItems(7)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(7))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(8)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(8))) > 0 Then _
                itmX.SubItems(8) = Val(Me.lstClntMix.ListItems(indT1).SubItems(8)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(8))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(9)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(9))) > 0 Then _
                itmX.SubItems(9) = Val(Me.lstClntMix.ListItems(indT1).SubItems(9)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(9))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(10)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(10))) > 0 Then _
                itmX.SubItems(10) = Val(Me.lstClntMix.ListItems(indT1).SubItems(10)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(10))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(11)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(11))) > 0 Then _
                itmX.SubItems(11) = Val(Me.lstClntMix.ListItems(indT1).SubItems(11)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(11))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(12)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(12))) > 0 Then _
                itmX.SubItems(12) = Val(Me.lstClntMix.ListItems(indT1).SubItems(12)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(12))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(13)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(13))) > 0 Then _
                itmX.SubItems(13) = Val(Me.lstClntMix.ListItems(indT1).SubItems(13)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(13))
                If (Val(Me.lstClntMix.ListItems(indT1).SubItems(14)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(14))) > 0 Then _
                itmX.SubItems(14) = Val(Me.lstClntMix.ListItems(indT1).SubItems(14)) + Val(Me.lstClntMix.ListItems(indT2).SubItems(14))
                
                If (CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(15))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(15)))) > 0 Then _
                itmX.SubItems(15) = CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(15))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(15)))
                
                If (CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(16))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(16)))) > 0 Then _
                itmX.SubItems(16) = CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(16))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(16)))
                
                If (CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(17))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(17)))) > 0 Then _
                itmX.SubItems(17) = CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(17))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(17)))
                
                If (CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(18))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(18)))) > 0 Then _
                itmX.SubItems(18) = CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(18))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(18)))
                
                If (CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(19))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(19)))) > 0 Then _
                itmX.SubItems(19) = CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(19))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(19)))
                
                If (CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(20))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(20)))) > 0 Then _
                itmX.SubItems(20) = CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(20))) + CSng(Me.lstClntMix.ListItems(indT2).SubItems(20))
                
                If rQinForm = 1 Then
                    If CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(21))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(21))) > 0 Then
                        itmX.SubItems(21) = CSng(rDs(Me.lstClntMix.ListItems(indT1).SubItems(21))) + CSng(rDs(Me.lstClntMix.ListItems(indT2).SubItems(21)))
                    End If
                End If
        End If
    Else
        MsgBox msgChooseFilter, vbOKOnly Or vbInformation, MsgErrBx
    End If
EndSub:
    AutoColW Me.lstClntMix
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstClntMix, 2, False, True, True, repClntMix & "  (" & Me.cmbClnt & ")")
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstClntMix)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Temp = Nothing
    frmStartRep.Show
End Sub

