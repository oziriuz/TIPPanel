VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStart 
   BorderStyle     =   0  'None
   Caption         =   "frmStart"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   ControlBox      =   0   'False
   Icon            =   "frmStart.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmStart.frx":08CA
   ScaleHeight     =   5655
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstOPC 
      Height          =   1335
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.TextBox txtVolNum 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4080
      Width           =   5175
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "btnExit"
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
      TabIndex        =   2
      ToolTipText     =   "Изход"
      Top             =   4680
      Width           =   2460
   End
   Begin VB.CommandButton btnEnter 
      Caption         =   "btnEnter"
      Default         =   -1  'True
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
      TabIndex        =   1
      Top             =   4680
      Width           =   2460
   End
   Begin VB.Label txtInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "txtInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label txtVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "txtVersion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Started As Boolean
Dim WithEvents Getserver As OPCServer
Attribute Getserver.VB_VarHelpID = -1

Public Sub Form_Load()
'стартова форма на програмата
    MousePointer = vbHourglass

    Dim RegKey             As String
    Dim StApp              As Boolean
    Dim Code               As String
    Dim itmX               As MSComctlLib.ListItem
    Dim intEmpFile         As Integer
    Dim f                  As FileSystemObject
    Dim MyPath             As String
    Dim DeskPath           As String
    Dim i                  As Integer
    Dim hw                 As Long
    Dim retval             As Long
    Dim SwWin              As String
    Dim cn                 As ADODB.Connection
    Dim rs                 As Recordset
    Dim ch_clients         As String
    Dim ch_worksites       As String
    Dim ch_deliveries      As String
    Dim ch_drivers         As String
    Dim ch_materials       As String
    Dim ch_mix_result_bc1  As String
    Dim ch_orders          As String
    Dim ch_recepies        As String
    Dim ch_suppliers       As String
    Dim ch_tempmix_bc1     As String
    Dim ch_admin_data      As String
    Dim ch_oper_data       As String
    Dim ch_entry_log       As String
    Dim ch_other_expen     As String
    Dim ch_revision        As String
    Dim ch_settings_bc1    As String
    Dim ch_settings_soft   As String
    Dim ch_daily_expenses  As String
    Dim MisTables(0 To 17) As String
    Dim TempString         As String
    Dim response           As Integer
    Dim comCreate          As String
    Dim ReCheck            As Boolean
    Dim tbName             As String
    Dim comm               As String
    Dim counter            As Long
    Dim comIns             As String
    Dim verApp             As String
        
    Set f = New FileSystemObject
    
    intEmpFile = FreeFile
    
    Call LoadLang 'функция за зареждане на езиковите променливи
    
    MyPath = strGetCommonAppDataPath() 'път до %AppData% на windows-a
    DeskPath = strGetDesktopPath()
    
    'създаване на папка на програмата в %AppData%
    If f.FolderExists(MyPath & "\TipPanel") = False Then MkDir (MyPath & "\TipPanel")
    
    'Файлове за работа на програмата
    PathCore = MyPath & "\TipPanel\"
    BackPath = DeskPath
    OPCSetFile = PathCore & "CPOst.dll"
    DBSetFile = PathCore & "BDst.dll"
    InfoFile = PathCore & "cominf.tpl"
    SilosFile = PathCore & "tipsls.dll"
    If MachineNumber = 2 Then SilosFile = PathCore & "tipsls22.dll"
    ConfirmityFile = PathCore & "ConfirmFile.txt"
'    LangSetFile = PathCore & "langset.set"
'    LangBgFile = PathCore & "bg.lang"
'    LangRusFile = PathCore & "rus.lang"
'    LangEnFile = PathCore & "en.lang"
    
    'проверка за втори старт при две машини
    'бг
    If MachineNumber = 1 Then SwWin = "Машина 2 - ТИП-Панел v" & App.Major & "." & App.Minor
    If MachineNumber = 2 Then SwWin = "Машина 1 - ТИП-Панел v" & App.Major & "." & App.Minor
    hw = FindWindow(vbNullString, SwWin)
    If hw <> 0 Then
        SecondStart = True
        GoTo BreakingBad
    End If
    
    'проверка дали приложението е стартирано вече
    If App.PrevInstance Then
        MousePointer = vbDefault
        MsgBox MsgAnotherRun, vbOKOnly Or vbCritical, MsgErrBx
        End
    End If

BreakingBad:
    'проверка за лицензен ключ
    RegKey = GetSetting(PlaceKey, PlaceKeyAdd, PlcRegLicNum, ErrRes)
    Code = ProdKeyGen()
    If RegKey = ErrRes Then
        MousePointer = vbDefault
        MsgBox MsgNoLic & vbNewLine & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBx
        End
    Else
        If RegKey = Code Then
            txtVolNum.Text = MsgKey & Code
            btnEnter.Default = True
            btnExit.Cancel = True
        Else
            MousePointer = vbDefault
            MsgBox MsgInvalLic & vbNewLine & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBx
            End
        End If
    End If
    
    If hw <> 0 Then
        SecondStart = True
        GoTo BreakingBadAgain
    End If

    'проверка дали работи базата данни
    StApp = isRunning("postgres.exe")
    If StApp = False Then
        MousePointer = vbDefault
        MsgBox MsgNotWorkDB, vbOKOnly Or vbCritical, MsgErrBx
        End
    End If

BreakingBadAgain:
    'проверка дали има файл с IP и парола за достъп до базата данни
    If Dir(DBSetFile) = "" Then
        frmDBFirst.Show
        MousePointer = vbDefault
        GoTo EndSub2
    Else
        Open DBSetFile For Input As #intEmpFile
        Input #intEmpFile, IPConnStr, PassConnStr
        Close #intEmpFile
    End If
    
    'connection string PostreSQL
    ConStr = "PROVIDER=PostgreSQL;" & "DATA SOURCE=" & IPConnStr & ";" & "LOCATION=" & DbaseName & ";" & "USER ID=" & DbaseUser & ";" & "PASSWORD=" & PassConnStr & ";"

    'почистване на променливи
    For i = 0 To 17
        MisTables(i) = ""
    Next i
    TempString = ""
    
    On Error Resume Next
'-----------------------Start postgreSQL-----------------------------------
    Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
    
    'проверка дали има една от стандартните таблици на базата данни
    'ако има значи имаме връзка с базата данни
    Set rs = cn.Execute("SELECT * FROM pg_tables WHERE tablename ='pg_statistic';")
    If Not rs.BOF Or Not rs.EOF Then
        rs.MoveFirst
        tbName = rs!tablename
    End If
    
    If tbName <> "pg_statistic" Then 'ако я няма значи базата данни не отговаря
        MousePointer = vbDefault
        MsgBox MsgNoDBConn, vbOKOnly Or vbCritical, MsgErrBx
        rs.Close
        Set rs = Nothing
        cn.Close 'затваряме връзката
        Set cn = Nothing
        frmDBFirst.Show
        Unload Me
        GoTo EndSub2
    Else
    End If
    
    GoTo CheckTables
    
    ReCheck = False
    
ReCheckTables: 'флаг за проверка след създаването на таблиците
    ReCheck = True
    
CheckTables:
    'проверка за съществуването на необходимите на програмата таблици в базата данни
    
    'клиенти
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'clients';")
    If Not rs.BOF And Not rs.EOF Then
        ch_clients = rs!table_name
    Else
        ch_clients = "fcku"
        MisTables(0) = "- " & uniClnts
    End If
        
    'обекти
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'worksites';")
    If Not rs.BOF And Not rs.EOF Then
        ch_worksites = rs!table_name
    Else
        ch_worksites = "fcku"
        MisTables(1) = "- " & uniObj
    End If
    
    'доставки
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'deliveries';")
    If Not rs.BOF And Not rs.EOF Then
        ch_deliveries = rs!table_name
    Else
        ch_deliveries = "fcku"
        MisTables(2) = "- " & uniDlvrs
    End If
    
    'водачи
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'drivers';")
    If Not rs.BOF And Not rs.EOF Then
        ch_drivers = rs!table_name
    Else
        ch_drivers = "fcku"
        MisTables(3) = "- " & uniDrvs
    End If
    
    'материали
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'materials_bc1';")
    If Not rs.BOF And Not rs.EOF Then
        ch_materials = rs!table_name
    Else
        ch_materials = "fcku"
        MisTables(4) = "- " & uniMats
    End If
    
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'materials_bc2';")
    If Not rs.BOF And Not rs.EOF Then
        ch_materials = rs!table_name
    Else
        ch_materials = "fcku"
        MisTables(4) = "- " & uniMats
    End If
    
    'резултати
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'mix_result_bc1';")
    If Not rs.BOF And Not rs.EOF Then
        ch_mix_result_bc1 = rs!table_name
    Else
        ch_mix_result_bc1 = "fcku"
        MisTables(5) = "- " & uniResults
    End If
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'mix_result_bc2';")
    
    If Not rs.BOF And Not rs.EOF Then
        ch_mix_result_bc1 = rs!table_name
    Else
        ch_mix_result_bc1 = "fcku"
        MisTables(5) = "- " & uniResults
    End If
       
    'заявки
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'orders';")
    If Not rs.BOF And Not rs.EOF Then
        ch_orders = rs!table_name
    Else
        ch_orders = "fcku"
        MisTables(6) = "- " & uniOrds
    End If
    
    'рецепти
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'recepies';")
    If Not rs.BOF And Not rs.EOF Then
        ch_recepies = rs!table_name
    Else
        ch_recepies = "fcku"
        MisTables(7) = "- " & uniRecs
    End If
    
    'доставчици
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'suppliers';")
    If Not rs.BOF And Not rs.EOF Then
        ch_suppliers = rs!table_name
    Else
        ch_suppliers = "fcku"
        MisTables(8) = "- " & uniSups
    End If
    
    'временни резултати
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'tempmix_bc1';")
    If Not rs.BOF And Not rs.EOF Then
        ch_tempmix_bc1 = rs!table_name
    Else
        ch_tempmix_bc1 = "fcku"
        MisTables(9) = "- " & uniTempResults
    End If
    
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'tempmix_bc2';")
    If Not rs.BOF And Not rs.EOF Then
        ch_tempmix_bc1 = rs!table_name
    Else
        ch_tempmix_bc1 = "fcku"
        MisTables(9) = "- " & uniTempResults
    End If
    
    'администратор
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'admin_data';")
    If Not rs.BOF And Not rs.EOF Then
        ch_admin_data = rs!table_name
    Else
        ch_admin_data = "fcku"
        MisTables(10) = "- " & uniAdmin
    End If
    
    'оператори
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'oper_data';")
    If Not rs.BOF And Not rs.EOF Then
        ch_oper_data = rs!table_name
    Else
        ch_oper_data = "fcku"
        MisTables(11) = "- " & uniDisp
    End If
    
    'лог-файл
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'entry_log';")
    If Not rs.BOF And Not rs.EOF Then
        ch_entry_log = rs!table_name
    Else
        ch_entry_log = "fcku"
        MisTables(12) = "- " & uniLog
    End If
    
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'entry_log2';")
    If Not rs.BOF And Not rs.EOF Then
        ch_entry_log = rs!table_name
    Else
        ch_entry_log = "fcku"
        MisTables(12) = "- " & uniLog
    End If
    
    'други
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'other_expen';")
    If Not rs.BOF And Not rs.EOF Then
        ch_other_expen = rs!table_name
    Else
        ch_other_expen = "fcku"
        MisTables(13) = "- " & uniOther
    End If
        
    'ревизия
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'revision';")
    If Not rs.BOF And Not rs.EOF Then
        ch_revision = rs!table_name
    Else
        ch_revision = "fcku"
        MisTables(14) = "- " & uniRevision
    End If
        
    'натройки
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'settings_bc1';")
    If Not rs.BOF And Not rs.EOF Then
        ch_settings_bc1 = rs!table_name
    Else
        ch_settings_bc1 = "fcku"
        MisTables(15) = "- " & uniSettings
    End If
    
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'settings_bc2';")
    If Not rs.BOF And Not rs.EOF Then
        ch_settings_bc1 = rs!table_name
    Else
        ch_settings_bc1 = "fcku"
        MisTables(15) = "- " & uniSettings
    End If
    
    'натройки софтуер
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'settings_soft';")
    If Not rs.BOF And Not rs.EOF Then
        ch_settings_soft = rs!table_name
    Else
        ch_settings_soft = "fcku"
        'бг
        MisTables(16) = "- " & "Настройки Софтуер"
    End If

    'дневен разход
    Set rs = cn.Execute("SELECT table_name FROM information_schema.tables WHERE table_name = 'daily_expenses';")
    If Not rs.BOF And Not rs.EOF Then
        ch_daily_expenses = rs!table_name
    Else
        ch_daily_expenses = "fcku"
        'бг
        MisTables(17) = "- " & "Дневен разход"
    End If
    
    TempString = ""
    
    'ако има липсващи таблици извеждаме съобщение с имената им
    If ch_clients = "fcku" Or ch_deliveries = "fcku" Or ch_drivers = "fcku" Or ch_materials = "fcku" Or _
    ch_mix_result_bc1 = "fcku" Or ch_orders = "fcku" Or ch_recepies = "fcku" Or ch_suppliers = "fcku" Or _
    ch_tempmix_bc1 = "fcku" Or ch_admin_data = "fcku" Or ch_oper_data = "fcku" Or ch_entry_log = "fcku" Or _
    ch_other_expen = "fcku" Or ch_revision = "fcku" Or ch_settings_bc1 = "fcku" Or ch_worksites = "fcku" Or _
    ch_settings_soft = "fcku" Or ch_daily_expenses = "fcku" Then
        For i = 0 To 17
            If MisTables(i) <> "" Then TempString = TempString & MisTables(i) & vbCrLf
        Next i
        MousePointer = vbDefault
        response = MsgBox(MsgTablesNotFound & vbCrLf & vbCrLf & TempString & MsgConfCrTables & vbCrLf & vbCrLf & MsgEndOnCancel, vbYesNo Or vbQuestion, MsgErrBx)
        If response = vbYes Then
            GoTo CreateTables
        Else
            rs.Close
            Set rs = Nothing
            cn.Close
            Set cn = Nothing
            Unload Me
            End
        End If
    Else
        GoTo HaveAllTables
    End If

    'създаваме липсващите таблици
CreateTables:
    MousePointer = vbHourglass

    If ch_clients = "fcku" Then 'таблица клиенти
        comCreate = "CREATE TABLE clients (c_num bigint NOT NULL, c_name text, c_bg text, c_mol text, c_add text, c_tel text, c_show boolean, CONSTRAINT clients_pkey PRIMARY KEY (c_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE clients OWNER TO postgres;")
        cn.Execute ("ALTER TABLE clients SET (autovacuum_enabled = True);")
    End If

    If ch_worksites = "fcku" Then 'таблица обекти
        comCreate = "CREATE TABLE worksites (w_num bigint NOT NULL, w_cnum text, w_name text, w_km text, w_show boolean, CONSTRAINT worksites_pkey PRIMARY KEY (w_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE worksites OWNER TO postgres;")
        cn.Execute ("ALTER TABLE worksites SET (autovacuum_enabled = True);")
    End If
    
    If ch_deliveries = "fcku" Then 'таблица доставки
        comCreate = "CREATE TABLE deliveries (del_num bigint NOT NULL, del_mat text, del_sup_name text, del_sup_bg text, del_doc_type text, del_doc_num text, del_date text, stamp_date date, del_q text, del_op text, CONSTRAINT deliveries_pkey PRIMARY KEY (del_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE deliveries OWNER TO postgres;")
        cn.Execute ("ALTER TABLE deliveries SET (autovacuum_enabled = True);")
    End If

    If ch_drivers = "fcku" Then 'таблица водачи
        comCreate = "CREATE TABLE drivers (d_num integer NOT NULL, d_name text, d_reg text, d_cap text, d_mod text, d_tel text, d_note text, d_show boolean, CONSTRAINT drivers_pkey PRIMARY KEY (d_num), CONSTRAINT drivers_d_name_key UNIQUE (d_name)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE drivers OWNER TO postgres;")
        cn.Execute ("ALTER TABLE drivers SET (autovacuum_enabled = True);")
    End If

    If ch_materials = "fcku" Then 'таблица материали
        comCreate = "CREATE TABLE materials_bc1 (m_num integer NOT NULL, m_name text, m_type text, m_load text, m_del text, m_sold text, CONSTRAINT materials_bc1_pkey PRIMARY KEY (m_num), CONSTRAINT materials_bc1_m_name_key UNIQUE (m_name)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE materials_bc1 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE materials_bc1 ADD COLUMN m_humidity text;")
        cn.Execute ("ALTER TABLE materials_bc1 SET (autovacuum_enabled = True);")
        comCreate = "CREATE TABLE materials_bc2 (m_num integer NOT NULL, m_name text, m_type text, m_load text, m_del text, m_sold text, CONSTRAINT materials_bc2_pkey PRIMARY KEY (m_num), CONSTRAINT materials_bc2_m_name_key UNIQUE (m_name)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE materials_bc2 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE materials_bc2 ADD COLUMN m_humidity text;")
        cn.Execute ("ALTER TABLE materials_bc2 SET (autovacuum_enabled = True);")
    End If

    If ch_mix_result_bc1 = "fcku" Then 'таблица резултати
        comCreate = "CREATE TABLE mix_result_bc1 (mix_num bigint NOT NULL, exp_num bigint, time_exp_start text, time_mix_ready text, stamp_date date, name_op text, ord_num bigint, ord_date text, ord_q text, exp_q text, exp_ord_num integer, mix_ord_num integer,CONSTRAINT mix_result_bc1_pkey PRIMARY KEY (mix_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE mix_result_bc1 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN name_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN bg_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN obj_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN km_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN name_drv text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN reg_drv text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cap_drv text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN name_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN type_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN class_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN classk_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN classv_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN classh_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN classp_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN edm_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im3_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im3z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im3i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im4_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im4z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im4i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im5_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im5z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im5i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im6_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im6z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN im6i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem3_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem3z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem3i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem4_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem4z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN cem4i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN wat1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN wat1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN wat1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN wat2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN wat2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN wat2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem3_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem3z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem3i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem4_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem4z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem4i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem5_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem5z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem5i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem6_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem6z text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN chem6i text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN total_rec_kg text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN total_real_kg text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN total_vol text;")
        cn.Execute ("ALTER TABLE mix_result_bc1 ADD COLUMN avstat boolean;")
        cn.Execute ("ALTER TABLE mix_result_bc1 SET (autovacuum_enabled = True);")
        
        comCreate = "CREATE TABLE mix_result_bc2 (mix_num bigint NOT NULL, exp_num bigint, time_exp_start text, time_mix_ready text, stamp_date date, name_op text, ord_num bigint, ord_date text, ord_q text, exp_q text, exp_ord_num integer, mix_ord_num integer,CONSTRAINT mix_result_bc2_pkey PRIMARY KEY (mix_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE mix_result_bc2 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN name_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN bg_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN obj_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN km_clnt text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN name_drv text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN reg_drv text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cap_drv text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN name_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN type_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN class_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN classk_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN classv_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN classh_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN classp_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN edm_rec text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im3_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im3z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im3i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im4_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im4z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im4i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im5_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im5z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im5i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im6_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im6z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN im6i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem3_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem3z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem3i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem4_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem4z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN cem4i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN wat1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN wat1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN wat1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN wat2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN wat2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN wat2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem1_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem1z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem1i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem2_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem2z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem2i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem3_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem3z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem3i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem4_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem4z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem4i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem5_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem5z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem5i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem6_name text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem6z text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN chem6i text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN total_rec_kg text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN total_real_kg text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN total_vol text;")
        cn.Execute ("ALTER TABLE mix_result_bc2 ADD COLUMN avstat boolean;")
        cn.Execute ("ALTER TABLE mix_result_bc2 SET (autovacuum_enabled = True);")
    End If
    
    If ch_orders = "fcku" Then 'таблица заявки
        comCreate = "CREATE TABLE orders (order_num bigint NOT NULL, order_date text, order_date_que text, stamp_date date, order_q text, order_qmade text, order_rec integer, order_rec_name text, order_rec_class text, order_clnt integer, order_clnt_name text, order_clnt_obj text, CONSTRAINT orders_pkey PRIMARY KEY (order_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE orders OWNER TO postgres;")
        cn.Execute ("ALTER TABLE orders SET (autovacuum_enabled = True);")
    End If
    
    If ch_recepies = "fcku" Then 'таблица рецепти
        comCreate = "CREATE TABLE recepies (r_num integer NOT NULL, r_name text, r_type text, r_class text, r_classk text, r_classv text, r_classh text, r_classp text, r_edm text, r_tpour text, r_tmix text, init_im1 text, kg_im1 text, init_im2 text, kg_im2 text, init_im3 text, kg_im3 text, init_im4 text, kg_im4 text, init_im5 text, kg_im5 text, init_im6 text, kg_im6 text, init_scr1 text, kg_scr1 text, init_scr2 text, kg_scr2 text, init_scr3 text, kg_scr3 text, init_scr4 text, kg_scr4 text, init_wat1 text, kg_wat1 text, init_wat2 text, kg_wat2 text, init_chem1 text, kg_chem1 text, init_chem2 text, kg_chem2 text, init_chem3 text, kg_chem3 text, init_chem4 text, kg_chem4 text, init_chem5 text, kg_chem5 text, init_chem6 text, kg_chem6 text, kg_total text, r_show boolean, CONSTRAINT recepies_pkey PRIMARY KEY (r_num), CONSTRAINT recepies_r_name_key UNIQUE (r_name)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE recepies OWNER TO postgres;")
        cn.Execute ("ALTER TABLE recepies SET (autovacuum_enabled = True);")
    End If
    
    If ch_suppliers = "fcku" Then 'таблица доставчици
        comCreate = "CREATE TABLE suppliers (s_num integer NOT NULL, s_name text, s_bg text, s_mol text, s_add text, s_tel text, s_note text, s_show boolean, CONSTRAINT suppliers_pkey PRIMARY KEY (s_num), CONSTRAINT suppliers_s_name_key UNIQUE (s_name)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE suppliers OWNER TO postgres;")
        cn.Execute ("ALTER TABLE suppliers SET (autovacuum_enabled = True);")
    End If
    
    If ch_tempmix_bc1 = "fcku" Then 'таблица временни резултати
        comCreate = "CREATE TABLE tempmix_bc1 (mix_id bigint NOT NULL, exp_id bigint, ordered_q text, real_q text, total_kg_temp text, CONSTRAINT tempmix_bc1_pkey PRIMARY KEY (mix_id)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE tempmix_bc1 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE tempmix_bc1 SET (autovacuum_enabled = True);")
        comCreate = "CREATE TABLE tempmix_bc2 (mix_id bigint NOT NULL, exp_id bigint, ordered_q text, real_q text, total_kg_temp text, CONSTRAINT tempmix_bc2_pkey PRIMARY KEY (mix_id)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE tempmix_bc2 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE tempmix_bc2 SET (autovacuum_enabled = True);")
    End If
        
    If ch_admin_data = "fcku" Then 'таблица администратор
        comCreate = "CREATE TABLE admin_data (a_num integer NOT NULL, a_name text,  a_pass text, CONSTRAINT admin_data_pkey PRIMARY KEY (a_num), CONSTRAINT admin_data_a_name_key UNIQUE (a_name)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE admin_data OWNER TO postgres;")
        cn.Execute ("ALTER TABLE admin_data SET (autovacuum_enabled = True);")
    End If
        
    If ch_oper_data = "fcku" Then 'таблица оператори
        comCreate = "CREATE TABLE oper_data (o_num integer NOT NULL, o_name text, o_pass text, CONSTRAINT oper_data_pkey PRIMARY KEY (o_num), CONSTRAINT oper_data_o_name_key UNIQUE (o_name)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE oper_data OWNER TO postgres;")
        cn.Execute ("ALTER TABLE oper_data SET (autovacuum_enabled = True);")
    End If
        
    If ch_entry_log = "fcku" Then 'таблица лог-файл
        comCreate = "CREATE TABLE entry_log (log_num bigint NOT NULL, log_name text, log_enter_date date, log_enter text, log_exit_date date, log_exit text, CONSTRAINT entry_log_pkey PRIMARY KEY (log_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE entry_log OWNER TO postgres;")
        cn.Execute ("ALTER TABLE entry_log SET (autovacuum_enabled = True);")
        comCreate = "CREATE TABLE entry_log2 (log_num bigint NOT NULL, log_name text, log_enter_date date, log_enter text, log_exit_date date, log_exit text, CONSTRAINT entry_log2_pkey PRIMARY KEY (log_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE entry_log2 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE entry_log2 SET (autovacuum_enabled = True);")
    End If
        
    If ch_other_expen = "fcku" Then 'таблица други
        comCreate = "CREATE TABLE other_expen (row_num bigint NOT NULL, other_matname text, other_matexp text, other_op text, other_date text, CONSTRAINT other_expen_pkey PRIMARY KEY (row_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE other_expen OWNER TO postgres;")
        cn.Execute ("ALTER TABLE other_expen SET (autovacuum_enabled = True);")
    End If
        
    If ch_revision = "fcku" Then 'таблица ревизия
        comCreate = "CREATE TABLE revision (row_num bigint NOT NULL, rev_num integer, rev_matname text, rev_matqold text, rev_matqnew text, rev_op text, rev_supervisor text, rev_date text, stamp_date date, CONSTRAINT revision_pkey PRIMARY KEY (row_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE revision OWNER TO postgres;")
        cn.Execute ("ALTER TABLE revision SET (autovacuum_enabled = True);")
    End If

    If ch_settings_bc1 = "fcku" Then 'таблица настройки
        comCreate = "CREATE TABLE settings_bc1 (ind integer NOT NULL, im_num text, cem_num text, wat_num text, chem_num text, work_permission text, stamp_date date, CONSTRAINT settings_bc1_pkey PRIMARY KEY (ind)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE settings_bc1 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE settings_bc1 SET (autovacuum_enabled = True);")
        comCreate = "CREATE TABLE settings_bc2 (ind integer NOT NULL, im_num text, cem_num text, wat_num text, chem_num text, work_permission text, stamp_date date, CONSTRAINT settings_bc2_pkey PRIMARY KEY (ind)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE settings_bc2 OWNER TO postgres;")
        cn.Execute ("ALTER TABLE settings_bc2 SET (autovacuum_enabled = True);")
    End If

    If ch_settings_soft = "fcku" Then 'таблица настройки софтуер и запис на данни
        comCreate = "CREATE TABLE settings_soft (ind integer NOT NULL, parameter text, value text, CONSTRAINT settings_soft_pkey PRIMARY KEY (ind)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE settings_soft OWNER TO postgres;")
        cn.Execute ("ALTER TABLE settings_soft SET (autovacuum_enabled = True);")
        
        Set rs = cn.Execute("INSERT INTO settings_soft VALUES(1,'NumSheetsForm1','1')")
        Set rs = cn.Execute("INSERT INTO settings_soft VALUES(2,'NumSheetsForm2','1')")
        Set rs = cn.Execute("INSERT INTO settings_soft VALUES(3,'NumSheetsForm3','1')")
    End If
    
    If ch_daily_expenses = "fcku" Then 'таблица дневен разход
        comCreate = "CREATE TABLE daily_expenses (row_num integer NOT NULL, mat_name text, mat_sold text, date_sold text, stamp_date date, CONSTRAINT daily_expenses_pkey PRIMARY KEY (row_num)) WITH (OIDS = False);"
        cn.Execute (comCreate)
        cn.Execute ("ALTER TABLE daily_expenses OWNER TO postgres;")
        cn.Execute ("ALTER TABLE daily_expenses SET (autovacuum_enabled = True);")
    End If
    
    GoTo CheckTables 'връщаме се на флага за проверка на таблиците
    
    'всички таблици са налични
HaveAllTables:
    MousePointer = vbHourglass
    
    'проверка дали има колона за влажност в таблицата с материалите
    Set rs = cn.Execute("SELECT column_name FROM information_schema.columns WHERE table_name='materials_bc1' AND column_name='m_humidity';")
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
    Else
        cn.Execute ("ALTER TABLE materials_bc1 ADD COLUMN m_humidity text;")
    End If
    
    Set rs = cn.Execute("SELECT column_name FROM information_schema.columns WHERE table_name='materials_bc2' AND column_name='m_humidity';")
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
    Else
        cn.Execute ("ALTER TABLE materials_bc2 ADD COLUMN m_humidity text;")
    End If
    
    'създаваме потребител за четене (user - reporter, pass - reporter)
    Set rs = cn.Execute("SELECT 1 FROM pg_roles WHERE rolname='reporter'")
    If Not rs.BOF Or Not rs.EOF Then
        rs.MoveFirst
    Else
        comCreate = "CREATE ROLE reporter LOGIN ENCRYPTED Password 'md5602ae602f49efdcf7bb2fe925f53d2dc' SUPERUSER INHERIT NOCREATEDB NOCREATEROLE REPLICATION CONNECTION LIMIT 50;"
        cn.Execute (comCreate)
    End If
    
    If ReCheck = True Then
        MousePointer = vbDefault
        MsgBox MsgTablesReady & vbCrLf & TempString, vbOKOnly Or vbInformation, uniSave
    Else
    End If
    
    'прочитаме настройки за брой бележки
    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm1';")
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
    Else
        numSheetsForm1 = 1
    End If
    Do While Not rs.EOF
        numSheetsForm1 = rs!Value
        rs.MoveNext
    Loop
    
    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm2';")
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
    Else
        numSheetsForm2 = 1
    End If
    Do While Not rs.EOF
        numSheetsForm2 = rs!Value
        rs.MoveNext
    Loop
        
    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm3';")
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
    Else
        numSheetsForm3 = 1
    End If
    Do While Not rs.EOF
        numSheetsForm3 = rs!Value
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    cn.Close 'затваряме връзката
    Set cn = Nothing
'--------------------------End PostgreSQL-----------------------------------
    
EndSub:
    MousePointer = vbHourglass
    
    DecSep = GetDecimalSep() 'откриваме десетичния сепаратор от регионалните настройки на компютъра
    
    txtInfo.Caption = TxtInfoCap
    txtInfo.FontSize = 12
    txtInfo.FontBold = True
    txtInfo.ForeColor = &HFF&
    txtInfo.Alignment = 2
    
    txtVersion.Caption = TxtVerCap
    txtVersion.FontSize = 12
    txtVersion.FontBold = True
    txtVersion.ForeColor = &HFF&
    txtVersion.Alignment = 2
    
    'бг
    btnEnter.Caption = UniEnter & " Машина " & MachineNumber
    btnExit.Caption = UniExit & " Машина " & MachineNumber
    
    If Dir(OPCSetFile) <> "" Then
        Open OPCSetFile For Input As intEmpFile
        Do Until EOF(intEmpFile)
            Input #intEmpFile, MyServer
        Loop
        lstOPC.Visible = False
        lstOPC.Enabled = False
        Close #intEmpFile
    Else
        MousePointer = vbHourglass

        'сканира за opc server
        Set Getserver = New OPCServer
        Servers = Getserver.GetOPCServers
        lstOPC.ListItems.Clear
        For i = LBound(Servers) To UBound(Servers)
            Set itmX = lstOPC.ListItems.Add(1, , Servers(i))
        Next i
        Set Getserver = Nothing
    End If

EndSub2:
    MousePointer = vbDefault
    
    If hw <> 0 And Started = False Then
        frmLogin.AdminSuccess = False
        frmLogin.LoginSucceeded = False
        frmLogin.RootUser = False
        OperName = ""
        If SecondStart = True Then
'------------------------------Start PostgreSQL--------------------------------------
            Set cn = New ADODB.Connection
                cn.ConnectionTimeout = 10
                cn.Open ConStr
                
            MousePointer = vbHourglass
    
            If MachineNumber = 1 Then comm = "SELECT * FROM entry_log2 ORDER BY log_num DESC LIMIT 1"
            If MachineNumber = 2 Then comm = "SELECT * FROM entry_log ORDER BY log_num DESC LIMIT 1"
            Set rs = cn.Execute(comm)
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            OperName = rs!log_name
            frmLogin.LoginSucceeded = True
            rs.Close
            Set rs = Nothing
            cn.Close 'затваряме връзката
            Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
Ready:
            If frmLogin.LoginSucceeded = True Then
'------------------------------Start PostgreSQL--------------------------------------
                cn.Open ConStr

                If MachineNumber = 1 Then comm = "SELECT * FROM entry_log ORDER BY log_num DESC LIMIT 1"
                If MachineNumber = 2 Then comm = "SELECT * FROM entry_log2 ORDER BY log_num DESC LIMIT 1"
                Set rs = cn.Execute(comm)
    
                If Not rs.BOF And Not rs.EOF Then
                    counter = Val(rs!log_num) + 1
                Else
                    counter = 1
                End If
                If MachineNumber = 1 Then comIns = "INSERT INTO entry_log (log_num, log_name, log_enter_date, log_enter) VALUES  (" & counter & ",'" & OperName & "','" & Format(Now, "DD-MM-YYYY") & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "')"
                If MachineNumber = 2 Then comIns = "INSERT INTO entry_log2 (log_num, log_name, log_enter_date, log_enter) VALUES  (" & counter & ",'" & OperName & "','" & Format(Now, "DD-MM-YYYY") & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "')"
                Set rs = cn.Execute(comIns)
                rs.Close
                Set rs = Nothing
                cn.Close 'затваряме връзката
                Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
                Unload Me
                DispPanel.Show
            End If
        End If
    End If
End Sub

Private Sub lstOPC_Click()
    MyServer = lstOPC.ListItems(lstOPC.SelectedItem.Index).Text
End Sub

Private Sub btnEnter_Click()

    Dim HasAdmin    As Boolean
    Dim cn          As New ADODB.Connection
    Dim rs          As New Recordset
    
    MousePointer = vbHourglass
    
    Started = True
    HasAdmin = False
    
    If Dir(OPCSetFile) = "" Then
        MyServer = lstOPC.ListItems(lstOPC.SelectedItem.Index).Text
    End If
'------------------------------Start PostgreSQL--------------------------------------
    cn.Open ConStr
    
    Set rs = cn.Execute("SELECT * FROM admin_data")
    If rs.BOF Or rs.EOF Then
        HasAdmin = False
    Else
        HasAdmin = True
    End If
    rs.Close
    Set rs = Nothing
    cn.Close 'затваряме връзката
    Set cn = Nothing
'------------------------------End PostgreSQL----------------------------------------
    Select Case HasAdmin
        Case True
            frmLogin.Show
            Unload Me
        Case Else
            MousePointer = vbDefault
            'бг
            MsgBox MsgMkAdmin, vbOKOnly Or vbInformation, "Администратор"
            frmNwAdmin.Show
            Unload Me
    End Select
    
    MousePointer = vbDefault
End Sub

Private Sub btnExit_Click()

    Dim hw          As Long
    Dim retval      As Long
    Dim SwWin       As String
    Dim response    As Integer
    
    MousePointer = vbDefault
    
    Set Getserver = New OPCServer
    Getserver.Disconnect
    
    'бг
    If MachineNumber = 1 Then SwWin = "Машина 2 - ТИП-Панел v" & App.Major & "." & App.Minor
    If MachineNumber = 2 Then SwWin = "Машина 1 - ТИП-Панел v" & App.Major & "." & App.Minor
    hw = FindWindow(vbNullString, SwWin)
    If hw = 0 Then
        SecondStart = False
        response = MsgBox(MsgArhivQuest, vbYesNo Or vbQuestion, MsgArhBx)
    Else
        response = vbNo
    End If
    If response = vbYes Then
        Unload Me
        frmBackup.Show
    Else
        End
    End If
End Sub

