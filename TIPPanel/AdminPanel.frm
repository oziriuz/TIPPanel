VERSION 5.00
Begin VB.Form AdminPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AdminPanel"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   Icon            =   "AdminPanel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRevision 
      Caption         =   "btnRevision"
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
      Left            =   3840
      TabIndex        =   16
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CommandButton btnClearDB 
      Caption         =   "Изчисти БД"
      Height          =   615
      Left            =   5880
      TabIndex        =   15
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CheckBox chEditor 
      Caption         =   "Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   14
      Top             =   6960
      Width           =   4455
   End
   Begin VB.CheckBox chSilos 
      Caption         =   "chSilos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1440
      TabIndex        =   13
      Top             =   6120
      Width           =   4455
   End
   Begin VB.CommandButton btnRestore 
      Caption         =   "btnRestore"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton btnLabPass 
      Caption         =   "btnLabPass"
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
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton btnForm3 
      Caption         =   "btnForm3"
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
      Left            =   3840
      TabIndex        =   11
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CommandButton btnForm2 
      Caption         =   "btnForm2"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton btnAllow 
      Caption         =   "btnAllow"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton btnForm1 
      Caption         =   "btnForm1"
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
      Left            =   3840
      TabIndex        =   9
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton btnCompanyInfo 
      Caption         =   "btnCompanyInfo"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton btnParam 
      Caption         =   "btnParam"
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
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CommandButton btnNameSilos 
      Caption         =   "btnNameSilos"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton btnAdmin 
      Caption         =   "btnAdmin"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton btnEdOper 
      Caption         =   "btnEdOper"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton btnLogout 
      Cancel          =   -1  'True
      Caption         =   "btnLogout"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   7560
      Width           =   3255
   End
   Begin VB.CommandButton btnNwOper 
      Caption         =   "btnNwOper"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "AdminPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnRevision_Click()

    If frmLogin.AdminSuccess = True Or frmLogin.RootUser = True Then
        frmRevision.Show
    Else
        frmDailyBalance.Show
        If ErrDaily = True Then Unload frmDailyBalance
    End If
End Sub

Private Sub Form_Load()
    'форма администраторски панел / настройки
    
    Me.Caption = frmAdPanel
    Me.btnNwOper.Caption = btnCreateOp
    Me.btnEdOper.Caption = btnEditOp
    Me.btnAdmin.Caption = btnEditAd
    Me.btnNameSilos.Caption = frmNmSilos
    Me.btnParam.Caption = btnParamSys
    Me.btnLogout.Caption = UniExit
    Me.btnCompanyInfo.Caption = uniComInfo
    Me.btnForm1.Caption = uniForm1
    Me.btnForm2.Caption = uniForm2
    Me.btnForm3.Caption = uniForm3
    Me.btnAllow.Caption = uniAllow
    Me.btnLabPass.Caption = frmLabPassCap
    Me.btnRestore.Caption = MsgResBx
    Me.btnRevision.Caption = uniRevision
    'бг
    Me.chSilos.Caption = "Активиране на въпрос при недостатъчно количество в силоза преди изпращане на експедицията."
    
    If frmLogin.AdminSuccess = False Or frmLogin.RootUser = True Then
        Me.btnAdmin.Enabled = False
        Me.btnNwOper.Enabled = False
        Me.btnEdOper.Enabled = False
        Me.btnCompanyInfo.Enabled = False
        Me.btnAllow.Enabled = False
        Me.btnLabPass.Enabled = False
        Me.btnRestore.Enabled = False
        Me.btnClearDB.Enabled = False
        Me.btnClearDB.Visible = False
        Me.btnRevision.Caption = uniDailyBalance

        If rActForm1 = 0 Then
            Me.btnForm1.Enabled = False
        ElseIf rActForm1 = 1 Then
            Me.btnForm1.Enabled = True
        End If

        If rActForm2 = 0 Then
            Me.btnForm2.Enabled = False
        ElseIf rActForm2 = 1 Then
            Me.btnForm2.Enabled = True
        End If

        If rActForm3 = 0 Then
            Me.btnForm3.Enabled = False
        ElseIf rActForm3 = 1 Then
            Me.btnForm3.Enabled = True
        End If
    End If
    
    'проверка в регистъра дали е включен въпроса за недостатъчно количество в силозите
    Dim PrevSetSilos As Boolean
    Dim strSubKeySilos As String

    If MachineNumber = 1 Then
        strSubKeySilos = Trim(Place1SilosQ)
        PrevSetSilos = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeySilos)

        If PrevSetSilos = True Then
            Me.chSilos.Value = GetSetting(PlaceProgSettings, Place1Q, "Quest1Silos", ErrRes)
        Else
            Me.chSilos.Value = 0
        End If
    ElseIf MachineNumber = 2 Then
        strSubKeySilos = Trim(Place2SilosQ)
        PrevSetSilos = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeySilos)

        If PrevSetSilos = True Then
            Me.chSilos.Value = GetSetting(PlaceProgSettings, Place2Q, "Quest2Silos", ErrRes)
        Else
            Me.chSilos.Value = 0
        End If
    End If
    
    QuestSilos = Me.chSilos.Value
    
    'проверка дали е включен редактора на експедиционните бележки
    Dim PrevSetEditor As Boolean
    Dim strSubKeyEditor As String
    
    strSubKeyEditor = Trim(PlaceEditor)
    PrevSetEditor = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeyEditor)
    
    If PrevSetEditor = True Then
        Me.chEditor.Value = GetSetting(PlaceProgSettings, PlaceEd, "NotesEditor", ErrRes)
    Else
        Me.chEditor.Value = 0
    End If
    
    ShowEditor = Me.chEditor.Value
    
End Sub

Private Sub btnAdmin_Click()
    frmEdAdmin.Show
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnNwOper_Click()
    frmNwOper.Show
    Unload frmEdAdmin
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnEdOper_Click()
    frmEdOper.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnNameSilos_Click()
    frmNameSilos.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnParam_Click()
    frmParam.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnCompanyInfo_Click()
    frmComInfo.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnForm1_Click()
    setForm1.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnForm2_Click()
    setForm2.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnForm3_Click()
    setForm3.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload frmRestore
End Sub

Private Sub btnAllow_Click()
    frmAllow.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnLabPass_Click()
    frmLabPass.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmRestore
End Sub

Private Sub btnRestore_Click()
    frmRestore.Show
    Unload frmEdAdmin
    Unload frmNwOper
    Unload frmEdOper
    Unload frmComInfo
    Unload frmParam
    Unload frmLabPass
    Unload frmNameSilos
    Unload frmAllow
    Unload setForm1
    Unload setForm2
    Unload setForm3
End Sub

Private Sub chSilos_Click()
    If MachineNumber = 1 Then
        SaveSetting PlaceProgSettings, Place1Q, "Quest1Silos", Me.chSilos
    ElseIf MachineNumber = 2 Then
        SaveSetting PlaceProgSettings, Place2Q, "Quest2Silos", Me.chSilos
    End If

    QuestSilos = Me.chSilos.Value
End Sub

Private Sub chEditor_Click()
    SaveSetting PlaceProgSettings, PlaceEd, "NotesEditor", Me.chEditor

    ShowEditor = Me.chEditor.Value
End Sub

Private Sub btnLogout_Click()
    Unload setForm1
    Unload setForm2
    Unload setForm3
    Unload frmLabPass
    Unload frmEdAdmin
    Unload frmAllow
    Unload setForm1
    Unload frmComInfo
    Unload frmEdOper
    Unload frmNwOper
    Unload frmParam
    Unload frmNameSilos
    Unload frmRestore
    Unload Me
End Sub

Private Sub btnClearDB_Click()
        '-----------------------Start postgreSQL-----------------------------------
    Dim cn                 As ADODB.Connection

    Dim rs                 As Recordset
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr

    'cn.Execute ("DROP TABLE clients")
    cn.Execute ("DROP TABLE deliveries")
    'cn.Execute ("DROP TABLE drivers")
    cn.Execute ("DROP TABLE mix_result_bc1")
    cn.Execute ("DROP TABLE mix_result_bc2")
    cn.Execute ("DROP TABLE orders")
    cn.Execute ("DROP TABLE revision")
    'cn.Execute ("DROP TABLE recepies")
    cn.Execute ("DROP TABLE suppliers")
    cn.Execute ("DROP TABLE tempmix_bc1")
    cn.Execute ("DROP TABLE tempmix_bc2")
    'cn.Execute ("DROP TABLE worksites")
    
    Set rs = Nothing
    cn.Close 'затваряме връзката
    Set cn = Nothing
    '--------------------------End PostgreSQL-----------------------------------
    
    Unload Me
    Unload DispPanel
    End

End Sub

