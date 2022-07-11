VERSION 5.00
Begin VB.Form frmAllow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmAllow"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   Icon            =   "frmAllow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
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
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CheckBox chRealQinForms 
      Caption         =   "chRealQinForms"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   7095
   End
End
Attribute VB_Name = "frmAllow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim PrevSet As Boolean
    Dim strSubKey As String


Private Sub Form_Load()
    Me.Caption = uniSettings
    Me.btnSave.Caption = uniSave
    strSubKey = Trim(PlaceProgAllow)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    Me.chRealQinForms.Caption = uniRealQinForm
    
    If PrevSet = True Then
        rQinForm = GetSetting(PlaceProgSettings, PlaceAllow, "RealQinForms", ErrRes)
    Else
        rQinForm = 1
    End If
    Me.chRealQinForms.Value = rQinForm
End Sub

Private Sub btnSave_Click()
    strSubKey = Trim(PlaceProgAllow)
    
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        DeleteSetting PlaceProgSettings, PlaceAllow, "RealQinForms"
    End If
    
    SaveSetting PlaceProgSettings, PlaceAllow, "RealQinForms", Me.chRealQinForms
    Unload Me
End Sub

