VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClntExpedition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmClntExpedition"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15735
   Icon            =   "frmClntExpediton.frx":0000
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
      TabIndex        =   18
      Top             =   8880
      Width           =   735
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
      Left            =   12120
      TabIndex        =   13
      Top             =   720
      Width           =   3255
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
      TabIndex        =   12
      Top             =   720
      Width           =   2415
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
      Left            =   6840
      TabIndex        =   11
      Top             =   720
      Width           =   2415
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
      Left            =   3600
      TabIndex        =   10
      Top             =   720
      Width           =   3015
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1320
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
      Top             =   720
      Width           =   3015
   End
   Begin MSComctlLib.ListView lstClntExpedition 
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2040
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11668
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
      Left            =   7080
      TabIndex        =   2
      Top             =   1440
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
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   10680
      TabIndex        =   3
      Top             =   1440
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
      Left            =   12240
      TabIndex        =   17
      Top             =   360
      Width           =   2775
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
      TabIndex        =   16
      Top             =   360
      Width           =   1935
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
      Left            =   6960
      TabIndex        =   15
      Top             =   360
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
      Left            =   3720
      TabIndex        =   14
      Top             =   360
      Width           =   2535
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
      TabIndex        =   9
      Top             =   1560
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
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
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
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmClntExpedition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'������� �� 1 ������ - ������ �� ��� �� ������������ �������
'����� ��� �� ��������� ���������� ���� ���������� ���� ������� ������ �� �������
    
    Me.Caption = repClntExped
    Me.lblClnt.Caption = uniClnt
    Me.lblObj.Caption = uniObj
    Me.lblRec.Caption = uniRec
    Me.lblClass.Caption = uniClass
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
    
    Me.lstClntExpedition.ColumnHeaders.Clear
    Me.lstClntExpedition.ListItems.Clear
    
    Me.cmbClnt.Clear
    Me.cmbObj.Clear
    Me.cmbRec.Clear
    Me.cmbClass.Clear
    Me.cmbDrv.Clear
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '����� ���
        colx.Text = uniEx & " " & uniNr
        colx.Width = 800
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '����� ����������/���� � ��� ����������
        colx.Text = uniDate
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���. �� ������������
        colx.Text = uniEx & " " & uniQ
        colx.Width = 1000

    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '����� ������/����
        colx.Text = uniOrd & " " & uniNr & "/" & uniDate
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���. �� ������
        colx.Text = uniOrdered & "  " & uniQ
        colx.Width = 1000

    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '�����
        colx.Text = uniObj
        colx.Width = 1000

    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���������� �� �����
        colx.Text = uniKmShort
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '��� �������
        colx.Text = uniRec
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���� �����
        colx.Text = uniClass
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���� ������������
        colx.Text = uniClassK
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���� �����������
        colx.Text = uniClassV
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���� �-�� �� �������
        colx.Text = uniClassH
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '������������
        colx.Text = uniClassP
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '�����
        colx.Text = uniDrv
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '���� ����� - ���������
        colx.Text = uniDrvReg & " - " & uniCap
        colx.Width = 1000
    
    Set colx = Me.lstClntExpedition.ColumnHeaders.Add() '��������/��������
        colx.Text = uniDisp
        colx.Width = 1000
        
    Me.lstClntExpedition.SortKey = 1
    
    AutoColW Me.lstClntExpedition
    
'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
        
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    '��������� �������������� �� ��������
    Set rsR = cnR.Execute("SELECT DISTINCT ON (name_clnt) name_clnt FROM mix_result_bc" & MachineNumber & " ORDER BY name_clnt ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbClnt.AddItem rsR!name_clnt
        rsR.MoveNext
    Loop
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (obj_clnt) obj_clnt FROM mix_result_bc" & MachineNumber & " ORDER BY obj_clnt ASC;")
    
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
    
    Set rsR = cnR.Execute("SELECT DISTINCT ON (name_drv) name_drv FROM mix_result_bc" & MachineNumber & " ORDER BY name_drv ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbDrv.AddItem rsR!name_drv
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
'            Me.cmbDrv.AddItem rsR!name_drv
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

    For I = 0 To Me.cmbDrv.listCount - 2 Step 1
        For j = Me.cmbDrv.listCount - 1 To I + 1 Step -1
            If Me.cmbDrv.List(I) = Me.cmbDrv.List(j) Then
                Me.cmbDrv.RemoveItem (j)
            End If
        Next
    Next

End Sub

Private Sub cmbClnt_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbClnt, KeyAscii, True)
End Sub

Private Sub cmbObj_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbObj, KeyAscii, True)
End Sub

Private Sub cmbRec_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbRec, KeyAscii, True)
End Sub

Private Sub cmbClass_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbClass, KeyAscii, True)
End Sub

Private Sub cmbDrv_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbDrv, KeyAscii, True)
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

    Me.lstClntExpedition.ListItems.Clear
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
        Dim commR5 As String
        Dim commRord As String
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

        '������ �� ��������� ����� �� ������ ���� �� ������
        commR1 = "SELECT time_mix_ready, name_op, exp_num, exp_q, ord_num, ord_date, ord_q, obj_clnt, km_clnt, name_rec, class_rec, classk_rec, classv_rec, classh_rec, classp_rec, name_drv, reg_drv, cap_drv, total_vol FROM mix_result_bc" & MachineNumber & " WHERE name_clnt = '" _
        & Me.cmbClnt.Text & "' AND stamp_date >= '" & DayStart & "' AND stamp_date <= '" & DayEnd & ""

        '������ �� ������� ��� ��������� �� �����
        commR2 = "' AND obj_clnt = '" & Me.cmbObj.Text & ""
    
        '������ �� ������� ��� ��������� �� ��� �������
        commR3 = "' AND name_rec = '" & Me.cmbRec.Text & ""
    
        '������ �� ������� ��� ��������� �� ���� �����
        commR4 = "' AND class_rec = '" & Me.cmbClass.Text & ""
    
        '������ �� ������� ��� ��������� �� �����
        commR5 = "' AND name_drv = '" & Me.cmbDrv.Text & ""
    
        '������ �� ������� ��� �����������
        commRord = "' ORDER BY mix_num, stamp_date ASC;"
    
        '��������� ������ �� ������� ���� �� ������� ������ - ��� ������
        commR = commR1
    
        '�������� ������ �� ��������� ��� ������� ������ �� �����
        If Me.cmbObj.Text <> "" Then commR = commR & commR2
    
        '�������� ������ �� ��������� ��� ������� ������ �� ��� �������
        If Me.cmbRec.Text <> "" Then commR = commR & commR3
    
        '�������� ������ �� ��������� ��� ������� ������ �� ���� �� �����
        If Me.cmbClass.Text <> "" Then commR = commR & commR4
    
        '�������� ������ �� ��������� ��� ������� ������ �� �����
        If Me.cmbDrv.Text <> "" Then commR = commR & commR5
    
        '������� �� ������� �� ���������
        commR = commR + commRord
    
        '����������� ���������
        Set rsR = cnR.Execute(commR)
    
        '������� �� ������ �����
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
        '����� �� �����������
beginCount:
    
        frstFlag = True
    
        '�������� ������������ ������ ������� �� ���������
        expCounter = 0
        mixCounter = 0
        Temp.TotalQuant = 0

        Do While Not rsR.EOF
            '��� ������������ � �������� �� ������ �� ����� ���������� +1
            '��� ������ ����� �� ������ ������������ � ������ � ����� +1 �����
            '��� ������ ����� �� ��������� ������� � ListView - ���� �������� ����������
            If Temp.ExpNum <> Val(rsR!exp_num) Then
                '��� ����� �� ������ ��������� ����� � ListView
                '� �������� ������������ ����������
                If frstFlag = False Then
fillList:
                    Set itmX = Me.lstClntExpedition.ListItems.Add(1, , "M" & MachineNumber & "-" & Temp.ExpNum)
                        itmX.SubItems(1) = Temp.MixReadyTime
                        If rQinForm = 0 Then
                            itmX.SubItems(2) = Temp.ExpQuant
                        Else
                            itmX.SubItems(2) = ARound(Temp.TotalQuant, 3)
                        End If
                        itmX.SubItems(3) = Temp.OrderCode & " / " & Temp.OrderDate '� ������
                        itmX.SubItems(4) = Temp.OrderQuant
                        itmX.SubItems(5) = Temp.ClntWorksite
                        itmX.SubItems(6) = Temp.WorksiteDist
                        itmX.SubItems(7) = Temp.RecTitle
                        itmX.SubItems(8) = Temp.RecClass
                        itmX.SubItems(9) = Temp.RecClassK
                        itmX.SubItems(10) = Temp.RecClassV
                        itmX.SubItems(11) = Temp.RecClassH
                        itmX.SubItems(12) = Temp.RecClassP
                        itmX.SubItems(13) = Temp.DrvTitle
                        itmX.SubItems(14) = Temp.DrvCarNum & " - " & Temp.DrvCapacity
                        itmX.SubItems(15) = Temp.DispName
                    If rsR.EOF Then Exit Do
                End If
            
                Temp.TotalQuant = 0
                expCounter = expCounter + 1
                frstFlag = False
            Else
            End If
            
            '��������� ������ �� ������������ ���� ����
            '���� ��� ����� ����� �� ������ ���� �� � ������ ����������
            Temp.ExpNum = Val(rsR!exp_num)
        
            '��������� ������ ��������� ���� ����������� ��������
            '� ��� �� ���������� ����������� �� ������ �� ������ �����
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
        
            '�������� ������� �� ����� �����
            Temp.TotalQuant = Temp.TotalQuant + CSng(rDs(rsR!total_vol))
            
            rsR.MoveNext '������ �� ������� �����
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
    
        '���� ���� ��������� �������� �������� ��������
        Dim I As Long
        Dim tKm As Long
        Dim Tmix As Long
        Dim tExpVol As Single
    
        tExp = 0
        Tmix = 0
        tExpVol = 0
    
        For I = 1 To Me.lstClntExpedition.ListItems.count
            tExpVol = tExpVol + CSng(rDs(Me.lstClntExpedition.ListItems.Item(I).SubItems(2)))
            tKm = tKm + Val(Me.lstClntExpedition.ListItems.Item(I).SubItems(6))
        Next I
    
        '����� ��������� ���� ������ ���
        Set itmX = Me.lstClntExpedition.ListItems.Add(1, , "")
            itmX.SubItems(1) = "X"
            
        '���� ���� ��������� ������
        Set itmX = Me.lstClntExpedition.ListItems.Add(1, , "")
            itmX.SubItems(1) = uniTotal & Me.lstClntExpedition.ListItems.count - 2 & " " & uniEx
            itmX.SubItems(2) = ARound(tExpVol, 3)
            itmX.SubItems(6) = tKm
    Else
        MsgBox msgChooseFilter, vbOKOnly Or vbInformation, MsgErrBx
    End If
EndSub:
    AutoColW Me.lstClntExpedition
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstClntExpedition, 2, True, True, True, repClntExped & "  (" & Me.cmbClnt & ")")
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstClntExpedition)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Temp = Nothing
    frmStartRep.Show
End Sub

