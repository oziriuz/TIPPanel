VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrvs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDrvs"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
   Icon            =   "frmDrvs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11670
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
      Left            =   10560
      TabIndex        =   6
      Top             =   7800
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
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   3015
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
      Left            =   8400
      TabIndex        =   3
      Top             =   120
      Width           =   2895
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
      Left            =   6720
      TabIndex        =   2
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
      Left            =   2640
      TabIndex        =   1
      Top             =   7800
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstDrvs 
      Height          =   6735
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11880
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
   Begin VB.Label lblDrv 
      Alignment       =   1  'Right Justify
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
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmDrvs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'зареждане на меню водачи

    Me.Caption = uniDrvs
    Me.btnPrint.Caption = btPrint
    Me.btnExport.Caption = btExport
    Me.btnLoad.Caption = btLoad
    Me.lblDrv.Caption = uniDrv

    Dim itmX As ListItem
    
    Me.lstDrvs.ColumnHeaders.Clear
    Me.lstDrvs.ListItems.Clear
    
    Me.cmbDrv.Clear
    
    Set colx = Me.lstDrvs.ColumnHeaders.Add()
        colx.Text = uniCode
        colx.Width = 700
    
    Set colx = Me.lstDrvs.ColumnHeaders.Add()
        colx.Text = uniNm
        colx.Width = 4000
    
    Set colx = Me.lstDrvs.ColumnHeaders.Add()
        colx.Text = uniDrvReg
        colx.Width = 1200

    Set colx = Me.lstDrvs.ColumnHeaders.Add()
        colx.Text = uniCapacity
        colx.Width = 1000

    Set colx = Me.lstDrvs.ColumnHeaders.Add()
        colx.Text = uniMod
        colx.Width = 1500
    
    Set colx = Me.lstDrvs.ColumnHeaders.Add()
        colx.Text = uniTel
        colx.Width = 1300
    
    Set colx = Me.lstDrvs.ColumnHeaders.Add()
        colx.Text = uniNote
        colx.Width = 3000

'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT * FROM drivers ORDER BY d_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Me.cmbDrv.AddItem rsR!d_name
        
        Set itmX = Me.lstDrvs.ListItems.Add(1, , Format(rsR!d_num, "0000"))
            itmX.SubItems(1) = rsR!d_name
            itmX.SubItems(2) = rsR!d_reg
            itmX.SubItems(3) = rDs(rsR!d_cap)
            itmX.SubItems(4) = rsR!d_mod
            itmX.SubItems(5) = rsR!d_tel
            itmX.SubItems(6) = rsR!d_note
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------
    
    AutoColW Me.lstDrvs
End Sub

Private Sub cmbDrv_KeyPress(KeyAscii As Integer)
    KeyAscii = cmbAutoComplete(cmbDrv, KeyAscii, True)
End Sub

Private Sub btnLoad_Click()

If Me.cmbDrv.Text <> "" Then

    Me.lstDrvs.ListItems.Clear
    
'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    
    Set cnR = New ADODB.Connection
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr
    MousePointer = vbHourglass
    
    Set rsR = cnR.Execute("SELECT * FROM drivers WHERE d_name = '" & Me.cmbDrv.Text & "' ORDER BY d_num ASC;")
    
    If Not rsR.EOF And Not rsR.BOF Then rsR.MoveFirst
    
    Do While Not rsR.EOF
        Set itmX = Me.lstDrvs.ListItems.Add(1, , Format(rsR!d_num, "0000"))
            itmX.SubItems(1) = rsR!d_name
            itmX.SubItems(2) = rsR!d_reg
            itmX.SubItems(3) = rDs(rsR!d_cap)
            itmX.SubItems(4) = rsR!d_mod
            itmX.SubItems(5) = rsR!d_tel
            itmX.SubItems(6) = rsR!d_note
        rsR.MoveNext
    Loop
    
    rsR.Close
    Set rsR = Nothing
    cnR.Close
    MousePointer = vbDefault
    Set cnR = Nothing
'--------------------------End PostgreSQL------------------------------------------

    AutoColW Me.lstDrvs
Else
    Call Form_Load
End If
End Sub

Private Sub btnPrint_Click()
    Call PrintLVPic(Me.lstDrvs, 1, True, True, True, uniDrvs)
End Sub

Private Sub btnExport_Click()
    Call ExportToExcel(Me.lstDrvs)
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

