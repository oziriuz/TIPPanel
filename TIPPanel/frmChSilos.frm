VERSION 5.00
Begin VB.Form frmChSilos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmChSilos"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "btnOK"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtSilosNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSilosNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSilosNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSilosNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSilosNo 
      BackStyle       =   0  'Transparent
      Caption         =   "lblSilosNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblSilosNo 
      BackStyle       =   0  'Transparent
      Caption         =   "lblSilosNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblSilosNo 
      BackStyle       =   0  'Transparent
      Caption         =   "lblSilosNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblSilosNo 
      BackStyle       =   0  'Transparent
      Caption         =   "lblSilosNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmChSilos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrevSet         As Boolean
Dim strSubKey       As String
Dim PlaceSilosSet   As String
Dim PlaceSilos      As String

Private Sub Form_Load()

    Dim i As Integer
    
    Me.Caption = "Смяна на силози - Машина " & MachineNumber
    Me.btnCancel.Caption = "Отказ"
    Me.btnOK.Caption = "ОК"
    
    For i = 1 To ns3
        Me.lblSilosNo(i - 1).Visible = True
        Me.txtSilosNo(i - 1).Visible = True
        Me.txtSilosNo(i - 1).MaxLength = 1
        Me.lblSilosNo(i - 1).Caption = "Силоз " & i
    Next i
    
    If MachineNumber = 1 Then
        PlaceSilosSet = Place1SilosSet
        PlaceSilos = Place1Silos
    End If
    If MachineNumber = 2 Then
        PlaceSilosSet = Place2SilosSet
        PlaceSilos = Place2Silos
    End If
    strSubKey = Trim(PlaceSilosSet)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    On Error Resume Next
    If PrevSet = True Then
        rSilos1 = GetSetting(PlaceProgSettings, PlaceSilos, "Silos1", ErrRes)
        rSilos2 = GetSetting(PlaceProgSettings, PlaceSilos, "Silos2", ErrRes)
        rSilos3 = GetSetting(PlaceProgSettings, PlaceSilos, "Silos3", ErrRes)
        rSilos4 = GetSetting(PlaceProgSettings, PlaceSilos, "Silos4", ErrRes)
    Else
        rSilos1 = 1
        rSilos2 = 2
        rSilos3 = 3
        rSilos4 = 4
    End If

    Me.txtSilosNo(0).Text = rSilos1
    Me.txtSilosNo(1).Text = rSilos2
    Me.txtSilosNo(2).Text = rSilos3
    Me.txtSilosNo(3).Text = rSilos4
End Sub

Private Sub txtSilosNo_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case 49 To (48 + ns3), 8 '1-4 и bksp
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub btnOK_Click()

    Dim i As Integer
    
    On Error Resume Next
    
    For i = 0 To 3
        If CInt(Me.txtSilosNo(i).Text) = 0 Then Me.txtSilosNo(i).Text = i + 1
    Next i
    
    strSubKey = Trim(PlaceSilosSet)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        DeleteSetting PlaceProgSettings, PlaceSilos, "Silos1"
        DeleteSetting PlaceProgSettings, PlaceSilos, "Silos2"
        DeleteSetting PlaceProgSettings, PlaceSilos, "Silos3"
        DeleteSetting PlaceProgSettings, PlaceSilos, "Silos4"
    End If
    SaveSetting PlaceProgSettings, PlaceSilos, "Silos1", Me.txtSilosNo(0).Text
    SaveSetting PlaceProgSettings, PlaceSilos, "Silos2", Me.txtSilosNo(1).Text
    SaveSetting PlaceProgSettings, PlaceSilos, "Silos3", Me.txtSilosNo(2).Text
    SaveSetting PlaceProgSettings, PlaceSilos, "Silos4", Me.txtSilosNo(3).Text
    DispPanel.numSilos(0).Caption = Me.txtSilosNo(0).Text
    DispPanel.numSilos(0).Refresh
    DispPanel.numSilos(1).Caption = Me.txtSilosNo(1).Text
    DispPanel.numSilos(1).Refresh
    DispPanel.numSilos(2).Caption = Me.txtSilosNo(2).Text
    DispPanel.numSilos(2).Refresh
    DispPanel.numSilos(3).Caption = Me.txtSilosNo(3).Text
    DispPanel.numSilos(3).Refresh
    Unload Me
    DispPanel.Show
End Sub

Private Sub btnCancel_Click()

    Unload Me
    DispPanel.Show
End Sub

