VERSION 5.00
Begin VB.Form frmAllow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmAllow"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   Icon            =   "frmAllow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chActiveForm3 
      Caption         =   "chActiveForm3"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   7095
   End
   Begin VB.CheckBox chActiveForm2 
      Caption         =   "chActiveForm2"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   7095
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "btnSave"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox chDeactiveDRPass 
      Caption         =   "chDeactiveDRPass"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   7095
   End
   Begin VB.CheckBox chDeactiveNRPass 
      Caption         =   "chDeactiveNRPass"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   7095
   End
   Begin VB.CheckBox chActiveDelete 
      Caption         =   "chActiveDelete"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   7095
   End
   Begin VB.CheckBox chActiveForm1 
      Caption         =   "chActiveForm1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   7095
   End
End
Attribute VB_Name = "frmAllow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrevSet   As Boolean
Dim strSubKey As String

Private Sub Form_Load()

    Me.Caption = uniAllow
    Me.btnSave.Caption = uniSave
    Me.chActiveForm1.Caption = uniActForm1
    Me.chActiveForm2.Caption = uniActForm2
    Me.chActiveForm3.Caption = uniActForm3
    Me.chActiveDelete.Caption = uniActDel
    Me.chDeactiveNRPass.Caption = uniDeactNRPass
    Me.chDeactiveDRPass.Caption = uniDeactDRPass
    
    strSubKey = Trim(PlaceProgAllow)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        rActForm1 = GetSetting(PlaceProgSettings, PlaceAllow, "ActForm1", ErrRes)
        rActForm2 = GetSetting(PlaceProgSettings, PlaceAllow, "ActForm2", ErrRes)
        rActForm3 = GetSetting(PlaceProgSettings, PlaceAllow, "ActForm3", ErrRes)
        rActDel = GetSetting(PlaceProgSettings, PlaceAllow, "ActDel", ErrRes)
        rDeactNRPass = GetSetting(PlaceProgSettings, PlaceAllow, "DeactNRPass", ErrRes)
        rDeactDRPass = GetSetting(PlaceProgSettings, PlaceAllow, "DeactDRPass", ErrRes)
    Else
        rActForm1 = 0
        rActForm2 = 0
        rActForm3 = 0
        rActDel = 0
        rDeactNRPass = 0
        rDeactDRPass = 0
    End If
    Me.chActiveForm1.Value = rActForm1
    Me.chActiveForm2.Value = rActForm2
    Me.chActiveForm3.Value = rActForm3
    Me.chActiveDelete.Value = rActDel
    Me.chDeactiveNRPass.Value = rDeactNRPass
    Me.chDeactiveDRPass.Value = rDeactDRPass
End Sub

Private Sub btnSave_Click()

    strSubKey = Trim(PlaceProgAllow)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        DeleteSetting PlaceProgSettings, PlaceAllow, "ActForm1"
        DeleteSetting PlaceProgSettings, PlaceAllow, "ActForm2"
        DeleteSetting PlaceProgSettings, PlaceAllow, "ActForm3"
        DeleteSetting PlaceProgSettings, PlaceAllow, "ActDel"
        DeleteSetting PlaceProgSettings, PlaceAllow, "DeactNRPass"
        DeleteSetting PlaceProgSettings, PlaceAllow, "DeactDRPass"
    End If
    SaveSetting PlaceProgSettings, PlaceAllow, "ActForm1", Me.chActiveForm1
    SaveSetting PlaceProgSettings, PlaceAllow, "ActForm2", Me.chActiveForm2
    SaveSetting PlaceProgSettings, PlaceAllow, "ActForm3", Me.chActiveForm3
    SaveSetting PlaceProgSettings, PlaceAllow, "ActDel", Me.chActiveDelete
    SaveSetting PlaceProgSettings, PlaceAllow, "DeactNRPass", Me.chDeactiveNRPass
    SaveSetting PlaceProgSettings, PlaceAllow, "DeactDRPass", Me.chDeactiveDRPass
    Unload Me
End Sub

