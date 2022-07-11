VERSION 5.00
Begin VB.Form frmStartRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmStartRep"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13695
   Icon            =   "frmStartRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNotes2 
      Caption         =   "btnNotes2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   33
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton btnMats2 
      Caption         =   "btnMats2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   32
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CheckBox chMach2 
      Caption         =   "Машина 2"
      Height          =   255
      Left            =   9600
      TabIndex        =   31
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chMach1 
      Caption         =   "Машина 1"
      Height          =   255
      Left            =   7800
      TabIndex        =   30
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "btnExit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   0
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton btnNotes 
      Caption         =   "btnNotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   21
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton btnMats 
      Caption         =   "btnMats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   20
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton btnSups 
      Caption         =   "btnSups"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   19
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton btnDrvs 
      Caption         =   "btnDrvs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   18
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton btnRecs 
      Caption         =   "btnRecs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   16
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton btnClnts 
      Caption         =   "btnClnts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   17
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton btnOrders 
      Caption         =   "btnOrders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   15
      Top             =   840
      Width           =   2175
   End
   Begin VB.PictureBox picReport 
      Height          =   3135
      Left            =   7560
      Picture         =   "frmStartRep.frx":08CA
      ScaleHeight     =   3075
      ScaleWidth      =   3195
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Frame frReadyRep 
      Caption         =   "frReadyRep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3960
      TabIndex        =   26
      Top             =   5040
      Width           =   3255
      Begin VB.CommandButton btnDailyReport 
         Caption         =   "btnDailyReport"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   36
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton btnReadyExpedition 
         Caption         =   "btnReadyExpedition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton btnDailyProduction 
         Caption         =   "btnDailyProduction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame frOperRep 
      Caption         =   "frOperRep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   360
      TabIndex        =   25
      Top             =   1560
      Width           =   3255
      Begin VB.CommandButton btnOperWork 
         Caption         =   "btnOperWork"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton btnOperAll 
         Caption         =   "btnOperAll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton btnOperDay 
         Caption         =   "btnOperDay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame frSupRep 
      Caption         =   "frSupRep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   7560
      TabIndex        =   24
      Top             =   1560
      Width           =   3255
      Begin VB.CommandButton btnMatRevision 
         Caption         =   "btnMatRevision"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton btnMatSold 
         Caption         =   "btnMatSold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton btnDeliveries 
         Caption         =   "btnDeliveries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame frClntRep 
      Caption         =   "frClntRep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   360
      TabIndex        =   23
      Top             =   5040
      Width           =   3255
      Begin VB.CommandButton btnClntAll 
         Caption         =   "btnClntAll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   37
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CommandButton btnClntOrd 
         Caption         =   "btnClntOrd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton btnClntMix 
         Caption         =   "btnClntMix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton btnClntExpedition 
         Caption         =   "btnClntExpedition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.Frame frDrvRep 
      Caption         =   "frDrvRep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3960
      TabIndex        =   22
      Top             =   1560
      Width           =   3255
      Begin VB.CommandButton btnDrvAll 
         Caption         =   "btnDrvAll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton btnDrvDay 
         Caption         =   "btnDrvDay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton btnDrvExpedition 
         Caption         =   "btnDrvExpedition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Label lblAddver 
      Alignment       =   2  'Center
      Caption         =   "lblAddver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      TabIndex        =   35
      Top             =   8400
      Width           =   12255
   End
   Begin VB.Label lblConn 
      Alignment       =   2  'Center
      Caption         =   "lblConn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   34
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label lblReporter2 
      Alignment       =   2  'Center
      Caption         =   "lblReporter2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   600
      Width           =   7215
   End
   Begin VB.Label lblRepoter 
      Alignment       =   2  'Center
      Caption         =   "lblReporter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   120
      Width           =   12135
   End
End
Attribute VB_Name = "frmStartRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClntAll_Click()
    Me.Hide
    frmClntAllExpedition.Show
End Sub

Public Sub Form_Load()
        
    MousePointer = vbHourglass
    
    Dim MyPath As String
    
    MyPath = strGetCommonAppDataPath() 'път до %AppData% на windows-a

    Set f = New FileSystemObject
    
'създаване на папка на програмата в %AppData%
    If f.FolderExists(MyPath & "\TipReporter") = False Then MkDir MyPath & "\TipReporter"
    
'Файлове за работа на програмата
    PathCore = MyPath & "\TipReporter\"
    DBSetFile = PathCore & "BDstRep.dll"
    ConfirmityFile = PathCore & "ConfirmFile.txt"
    InfoFile = PathCore & "cominf.tpl"
'    LangSetFile = PathCore & "langsetrep.set"
'    LangBgFile = PathCore & "bgrep.lang"
'    LangRusFile = PathCore & "rusrep.lang"
'    LangEnFile = PathCore & "enrep.lang"
'-----------------------------------------------

    Call LoadLang
    
    Me.Caption = frmReportPanelCap
    Me.lblRepoter.Caption = TxtInfoCap
    Me.lblReporter2.Caption = TxtVerCap
    Me.frOperRep.Caption = uniDisp
    Me.frDrvRep.Caption = uniDrv
    Me.frSupRep.Caption = uniSup
    Me.frClntRep.Caption = uniClnt
    Me.frReadyRep.Caption = uniResults
    Me.btnOperDay.Caption = repOperDay
    Me.btnOperAll.Caption = repOperAll
    Me.btnOperWork.Caption = uniLog
    Me.btnDrvExpedition.Caption = repDrvExped
    Me.btnDrvDay.Caption = repDrvDay
    Me.btnDrvAll.Caption = repDrvAll
    Me.btnDeliveries.Caption = uniDlvrs
    Me.btnMatSold.Caption = repMatSold
    Me.btnMatRevision.Caption = uniRevisions
    Me.btnClntMix.Caption = repClntMix
    Me.btnClntExpedition.Caption = repClntExped
    Me.btnClntOrd.Caption = repClntOrd
    Me.btnDailyProduction.Caption = repDailyProd
    Me.btnDailyReport.Caption = repDailyRep
    Me.btnReadyExpedition.Caption = repDailyExped
    Me.btnOrders.Caption = uniOrds
    Me.btnRecs.Caption = uniRecs
    Me.btnClnts.Caption = uniClnts
    Me.btnDrvs.Caption = uniDrvs
    Me.btnSups.Caption = uniSups
    Me.btnMats.Caption = uniMats & " 1"
    Me.btnMats2.Caption = uniMats & " 2"
    Me.btnNotes.Caption = uniNotes & " 1"
    Me.btnNotes2.Caption = uniNotes & " 2"
    Me.btnExit.Caption = UniExit
    Me.btnClntAll.Caption = "Всички клиенти по експедиции"
    
    Me.lblAddver.Caption = "  ТИП-Сервиз ЕООД - гр. Червен бряг - Софтуер, проектиране, изграждане и сервиз на бетонови стопанства"
    
    Me.chMach1.Value = 1
    Me.chMach2.Value = 1
    
    'проверка дали приложението е стартирано вече
    If App.PrevInstance Then
        MousePointer = vbDefault
        MsgBox MsgAnotherRun, vbOKOnly Or vbCritical, MsgErrBx
        Unload Me
        End
    End If

    'проверка дали има файл с IP и парола за достъп до базата данни
    If Dir(DBSetFile) = "" Then
        MousePointer = vbDefault
        frmDBFirst.Show
        GoTo EndSub
    Else
        If Choice = True Then GoTo SkipOpen
        intEmpFileNbr1 = FreeFile
        lCount = 0
        Open DBSetFile For Input As #intEmpFileNbr1
        Do Until EOF(intEmpFileNbr1)
            lCount = lCount + 1
            Input #intEmpFileNbr1, IPConnStr, MachName
        Loop
        Close #intEmpFileNbr1
        
        If lCount = 1 Then
            Open DBSetFile For Input As #intEmpFileNbr1
            Input #intEmpFileNbr1, IPConnStr, MachName
            Close #intEmpFileNbr1
        ElseIf lCount > 1 Then
            Me.Hide
            frmChoose.Show
            GoTo EndSub
        End If
    End If
    
SkipOpen:

    Me.lblConn.Caption = "Свързан с машина: " & MachName
    
    AddMach = False
    
    'connection string PostreSQL
    ConStr = "PROVIDER=PostgreSQL;" _
            & "DATA SOURCE=" & IPConnStr & ";" _
            & "LOCATION=" & DbaseName & ";" _
            & "USER ID=" & DbaseUser & ";" _
            & "PASSWORD=" & PassConnStr & ";"

'-----------------------Start postgreSQL-----------------------------------
    Dim cnR As New ADODB.Connection
    Dim rsR As New Recordset
    Dim ch_clients As String
    Dim ch_worksites As String
    Dim ch_deliveries As String
    Dim ch_drivers As String
    Dim ch_materials As String
    Dim ch_mix_result_bc1 As String
    Dim ch_orders As String
    Dim ch_recepies As String
    Dim ch_suppliers As String
    Dim ch_tempmix_bc1 As String
    Dim ch_admin_data As String
    Dim ch_oper_data As String
    Dim ch_entry_log As String
    Dim ch_other_expen As String
    Dim ch_revision As String
    Dim ch_settings_bc1 As String
    Dim MisTables(0 To 15) As String
    Dim TempString As String
    Dim wrkPerm As String

    For i = 0 To 15
        MisTables(i) = ""
    Next i
    
    TempString = ""
    
    On Error Resume Next
    cnR.ConnectionTimeout = 30
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    'проверка дали има една от стандартните таблици на базата данни
    'ако има значи имаме връзка с базата данни
    Set rsR = cnR.Execute("SELECT * FROM pg_tables WHERE tablename ='pg_statistic';")
    rsR.MoveFirst
    If rsR!tablename <> "pg_statistic" Then 'ако я няма значи базата данни не отговаря
        MousePointer = vbDefault
        MsgBox MsgNoDBConn, vbOKOnly Or vbCritical, MsgErrBx
        'показваме формата за IP за базата данни
        
        rsR.Close
        Set rsR = Nothing
        cnR.Close 'затваряме връзката
        Set cnR = Nothing
        
        frmDBFirst.Show
        Unload Me
        GoTo EndSub
    Else
    End If
    
    'проверка за съществуването на необходимите на програмата таблици в базата данни
    
    'клиенти
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'clients';")
    If Not rsR.BOF And Not rsR.EOF Then
        ch_clients = rsR!table_name
    Else
        ch_clients = "fcku"
        MisTables(0) = "- " & uniClnts
    End If
        
    'обекти
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'worksites';")
    If Not rsR.BOF And Not rsR.EOF Then
        ch_oworksites = rsR!table_name
    Else
        ch_oworksites = "fcku"
        MisTables(1) = "- " & uniObj
    End If
        
    'доставки
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'deliveries';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_deliveries = rsR!table_name
    Else
        ch_deliveries = "fcku"
        MisTables(2) = "- " & uniDlvrs
    End If
    
    'водачи
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'drivers';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_drivers = rsR!table_name
    Else
        ch_drivers = "fcku"
        MisTables(3) = "- " & uniDrvs
    End If
    
    'материали
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'materials_bc1';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_materials = rsR!table_name
    Else
        ch_materials = "fcku"
        MisTables(4) = "- " & uniMats
    End If
    
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'materials_bc2';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_materials = rsR!table_name
    Else
        ch_materials = "fcku"
        MisTables(4) = "- " & uniMats
    End If
    
    'резултати
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'mix_result_bc1';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_mix_result_bc1 = rsR!table_name
    Else
        ch_mix_result_bc1 = "fcku"
        MisTables(5) = "- " & uniResults
    End If
       
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'mix_result_bc2';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_mix_result_bc2 = rsR!table_name
    Else
        ch_mix_result_bc2 = "fcku"
        MisTables(5) = "- " & uniResults
    End If
       
    'заявки
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'orders';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_orders = rsR!table_name
    Else
        ch_orders = "fcku"
        MisTables(6) = "- " & uniOrds
    End If
    
    'рецепти
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'recepies';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_recepies = rsR!table_name
    Else
        ch_recepies = "fcku"
        MisTables(7) = "- " & uniRecs
    End If
    
    'доставчици
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'suppliers';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_suppliers = rsR!table_name
    Else
        ch_suppliers = "fcku"
        MisTables(8) = "- " & uniSups
    End If
    
    'временни резултати
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'tempmix_bc1';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_tempmix_bc1 = rsR!table_name
    Else
        ch_tempmix_bc1 = "fcku"
        MisTables(9) = "- " & uniTempResults
    End If
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'tempmix_bc2';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_tempmix_bc1 = rsR!table_name
    Else
        ch_tempmix_bc1 = "fcku"
        MisTables(9) = "- " & uniTempResults
    End If
    
    'администратор
    Set rs = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'admin_data';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_admin_data = rsR!table_name
    Else
        ch_admin_data = "fcku"
        MisTables(10) = "- " & uniAdmin
    End If
    
    'оператори
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'oper_data';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_oper_data = rsR!table_name
    Else
        ch_oper_data = "fcku"
        MisTables(11) = "- " & uniDisp
    End If
    
    'лог-файл
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'entry_log';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_entry_log = rsR!table_name
    Else
        ch_entry_log = "fcku"
        MisTables(12) = "- " & uniLog
    End If
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'entry_log2';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_entry_log = rsR!table_name
    Else
        ch_entry_log = "fcku"
        MisTables(12) = "- " & uniLog
    End If
    
    'други
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'other_expen';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_other_expen = rsR!table_name
    Else
        ch_other_expen = "fcku"
        MisTables(13) = "- " & uniOther
    End If
        
    'ревизия
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'revision';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_revision = rsR!table_name
    Else
        ch_revision = "fcku"
        MisTables(14) = "- " & uniRevision
    End If
        
    'натройки
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'settings_bc1';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_settings_bc1 = rsR!table_name
    Else
        ch_settings_bc1 = "fcku"
        MisTables(15) = "- " & uniSettings
    End If
    Set rsR = cnR.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'settings_bc1';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        ch_settings_bc1 = rsR!table_name
    Else
        ch_settings_bc1 = "fcku"
        MisTables(15) = "- " & uniSettings
    End If
    
    'ако има липсващи таблици извеждаме съобщение с имената им
    If ch_clients = "fcku" Or ch_deliveries = "fcku" Or ch_drivers = "fcku" Or ch_materials = "fcku" Or _
    ch_mix_result_bc1 = "fcku" Or ch_orders = "fcku" Or ch_recepies = "fcku" Or ch_suppliers = "fcku" Or _
    ch_tempmix_bc1 = "fcku" Or ch_admin_data = "fcku" Or ch_oper_data = "fcku" Or ch_entry_log = "fcku" _
    Or ch_other_expen = "fcku" Or ch_revision = "fcku" Or ch_settings_bc1 = "fcku" Or ch_worksites = "fcku" Then
        For i = 0 To 15
            If MisTables(i) <> "" Then TempString = TempString & MisTables(i) & vbCrLf
        Next i
        MousePointer = vbDefault
        MsgBox MsgTablesNotFound & vbCrLf & vbCrLf & TempString, vbOKOnly Or vbCritical, MsgErrBx
        
        rsR.Close
        Set rsR = Nothing
        cnR.Close
        Set cnR = Nothing
        Unload Me
        End
    Else
    End If
    
    'прочитаме броя на течките на машината от базата данни
    Set rsR = cnR.Execute("SELECT * FROM settings_bc1 WHERE ind = '1';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        n1s1 = Val(rsR!im_num)
        n1s2 = Val(rsR!wat_num)
        n1s3 = Val(rsR!cem_num)
        n1s4 = Val(rsR!chem_num)
    Else
        'ако няма записани бройките извеждаме грешка и спираме програмата
        MousePointer = vbDefault
        MsgBox MsgConfigNotFound, vbOKOnly Or vbCritical, MsgErrBx
        
        rsR.Close
        Set rsR = Nothing
        cnR.Close 'затваряме връзката
        Set cnR = Nothing
        Unload Me
        End
    End If
    
    If n1s1 = 0 Then ns1 = 1
    If n1s2 = 0 Then ns2 = 1
    If n1s3 = 0 Then ns3 = 1
    If n1s4 = 0 Then ns4 = 1
    
    Set rsR = cnR.Execute("SELECT * FROM settings_bc2 WHERE ind = '1';")
    
    If Not rsR.BOF And Not rsR.EOF Then
        n2s1 = Val(rsR!im_num)
        n2s2 = Val(rsR!wat_num)
        n2s3 = Val(rsR!cem_num)
        n2s4 = Val(rsR!chem_num)
    Else
        'ако няма записани бройките извеждаме грешка и спираме програмата
        MousePointer = vbDefault
        n2s1 = 1
        n2s2 = 1
        n2s3 = 1
        n2s4 = 1
        rsR.Close
        Set rsR = Nothing
'        cnR.Close 'затваряме връзката
'        Set cnR = Nothing
    End If
    
    If n2s1 = 0 Then n2s1 = 1
    If n2s2 = 0 Then n2s2 = 1
    If n2s3 = 0 Then n2s3 = 1
    If n2s4 = 0 Then n2s4 = 1
    
    'проверка дали дипечерът е активен или е спрял за неплатен лиценз
    Set rsR = cnR.Execute("SELECT work_permission FROM settings_bc1 ORDER BY ind ASC LIMIT 1;")
    
    If Not rsR.EOF And Not rsR.BOF Then
        rsR.MoveFirst
        If rsR!work_permission <> Null Then
            wrkPerm = rsR!work_permission
        Else
        End If
    Else
        GoTo Skip
    End If
    
    If wrkPerm = "stop" Or Val(rDs(wrkPerm)) < 0 Then
        MsgBox MsgNoPayment & vbCrLf & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBx
        End
    End If
    
Skip:
    MousePointer = vbDefault
    rsR.Close
    Set rsR = Nothing
    cnR.Close 'затваряме връзката
    Set cnR = Nothing
'--------------------------End PostgreSQL-----------------------------------

EndSub:

'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон
    Dim PrevSet As Boolean
    Dim strSubKey As String
    strSubKey = Trim(PlaceProgAllow)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        rQinForm = GetSetting(PlaceProgSettings, PlaceAllow, "RealQinForms", ErrRes)
    Else
        rQinForm = 1
    End If

    MousePointer = vbDefault
End Sub

Private Sub btnOperDay_Click()
    Me.Hide
    frmOperDay.Show
End Sub

Private Sub btnOperAll_Click()
    Me.Hide
    frmOperAll.Show
End Sub

Private Sub btnOperWork_Click()
    Me.Hide
    Call frmOperWork.Show
End Sub

Private Sub btnDrvExpedition_Click()
    Me.Hide
    frmDrvExpedition.Show
End Sub

Private Sub btnDrvDay_Click()
    Me.Hide
    frmDrvDay.Show
End Sub

Private Sub btnDrvAll_Click()
    Me.Hide
    frmDrvAll.Show
End Sub

Private Sub btnClntMix_Click()
    Me.Hide
    frmClntMix.Show
End Sub

Private Sub btnClntExpedition_Click()
    Me.Hide
    frmClntExpedition.Show
End Sub

Private Sub btnClntOrd_Click()
    Me.Hide
    frmClntOrd.Show
End Sub

Private Sub btnDailyProduction_Click()
    Me.Hide
    frmDailyProduction.Show
End Sub

Private Sub btnDailyReport_Click()
    Me.Hide
    frmDailyReport.Show
End Sub

Private Sub btnReadyExpedition_Click()
    Me.Hide
    frmReadyExpedition.Show
End Sub

Private Sub btnDeliveries_Click()
    Me.Hide
    frmDeliveries.Show
End Sub

Private Sub btnMatSold_Click()
    Me.Hide
    frmMatSold.Show
End Sub

Private Sub btnMatRevision_Click()
    Me.Hide
    frmMatRevision.Show
End Sub

Private Sub btnOrders_Click()
    Me.Hide
    frmOrders.Show
End Sub

Private Sub btnRecs_Click()
    Me.Hide
    frmRecs.Show
End Sub

Private Sub btnClnts_Click()
    Me.Hide
    frmClnts.Show
End Sub

Private Sub btnDrvs_Click()
    Me.Hide
    frmDrvs.Show
End Sub

Private Sub btnSups_Click()
    Me.Hide
    frmSups.Show
End Sub

Private Sub btnMats_Click()
    Me.Hide
    MachineNumber = 1
    frmMats.Show
End Sub

Private Sub btnMats2_Click()
    Me.Hide
    MachineNumber = 2
    frmMats.Show
End Sub

Private Sub btnNotes_Click()
    Me.Hide
    MachineNumber = 1
    frmNotes.Show
End Sub

Private Sub btnNotes2_Click()
    Me.Hide
    MachineNumber = 2
    frmNotes.Show
End Sub

Private Sub btnExit_Click()
    Unload Me
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub picReport_Click()
    Me.Hide
    AdminPanel.Show
End Sub
