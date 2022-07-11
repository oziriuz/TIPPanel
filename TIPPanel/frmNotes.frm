VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmNotes"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCarData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1560
   End
   Begin VB.TextBox txtObjData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3000
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "uniClose"
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
      Left            =   11400
      TabIndex        =   4
      Top             =   7440
      Width           =   2175
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
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   7440
      Width           =   2175
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
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   7440
      Width           =   2175
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
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox txtClassH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1320
   End
   Begin VB.TextBox txtClassP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1320
   End
   Begin VB.TextBox txtDrvData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   3000
   End
   Begin VB.TextBox txtClassK 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1320
   End
   Begin VB.TextBox txtClassV 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1320
   End
   Begin VB.TextBox txtOperData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   3000
   End
   Begin VB.TextBox txtDateReady 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2040
   End
   Begin VB.TextBox txtClntData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   3000
   End
   Begin VB.TextBox txtNotes 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1440
   End
   Begin MSComctlLib.ListView lstMixOld 
      Height          =   4455
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   7858
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
   Begin ComCtl2.UpDown udNotes 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   1
      OrigLeft        =   6720
      OrigTop         =   1800
      OrigRight       =   6975
      OrigBottom      =   2175
      Max             =   1000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblClassP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblClassP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   27
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblClassH 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblClassH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   26
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblClassV 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblClassV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblClassK 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblClassK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblCarData 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCarData"
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
      Left            =   6720
      TabIndex        =   23
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblDateReady 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDateReady"
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
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblExp 
      BackStyle       =   0  'Transparent
      Caption         =   "lblExp"
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
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblOperData 
      BackStyle       =   0  'Transparent
      Caption         =   "lblOperData"
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
      TabIndex        =   20
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblObjData 
      BackStyle       =   0  'Transparent
      Caption         =   "lblObjData"
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
      Left            =   3000
      TabIndex        =   19
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblClntData 
      BackStyle       =   0  'Transparent
      Caption         =   "lblClntData"
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
      Left            =   3000
      TabIndex        =   18
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblDrvData 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDrvData"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim CountMixd       As Integer
    Dim TotalRecKGsd    As Integer
    Dim TotalKGsd       As Integer
    Dim cnnew           As ADODB.Connection
    Dim rsnew           As Recordset
    Dim i               As Integer

    'нулиране на брояча на замесите и сумарните тегла
    CountMixd = 0
    TotalRecKGsd = 0
    TotalKGsd = 0
    
    Me.txtNotes.MaxLength = 7
    Me.Caption = uniNotes
    Me.lblExp.Caption = uniExped
    Me.lblDateReady.Caption = uniDate
    Me.lblClntData.Caption = uniClnt
    Me.lblObjData.Caption = uniObj
    Me.lblDrvData.Caption = uniDrv
    Me.lblCarData.Caption = uniDrvReg
    Me.lblOperData.Caption = uniDisp
    Me.lblClassK.Caption = uniClassK
    Me.lblClassV.Caption = uniClassV
    Me.lblClassH.Caption = uniClassH
    Me.lblClassP.Caption = uniClassP
    Me.btnForm1.Caption = uniForm1
    Me.btnForm2.Caption = uniForm2
    Me.btnForm3.Caption = uniForm3
    Me.btnClose.Caption = UniCancel
    
'------------------------------Start PostgreSQL----------------------------------
    Set cnnew = New ADODB.Connection 'връзка с база данни
        cnnew.ConnectionTimeout = 10
        cnnew.Open ConStr 'отваряме връзката
        
    MousePointer = vbHourglass
    
    'визуализация на замесите от последната експедиция
    Set rsnew = cnnew.Execute("SELECT exp_num FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1") 'маркираме последния замес
    If Not rsnew.EOF And Not rsnew.BOF Then
        i = rsnew!exp_num 'маркираме номера на експедиция от последния замес
    Else 'ако няма замеси
        GoTo EndLoad
    End If
    Me.udNotes.BuddyControl = Me.txtNotes
    Me.udNotes.Max = i
    Me.udNotes.Increment = 1
    Me.txtNotes.Text = Me.udNotes.Max
EndLoad:
    rsnew.Close 'затваряме записите
    Set rsnew = Nothing
    cnnew.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnnew = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------
    If Me.lstMixOld.ListItems.count < 1 Then
        Me.btnForm1.Enabled = False
        Me.btnForm2.Enabled = False
        Me.btnForm3.Enabled = False
    Else
        Me.btnForm1.Enabled = True
        Me.btnForm2.Enabled = True
        Me.btnForm3.Enabled = True
    End If
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtNotes_Change()

    Dim i                       As Integer
    Dim TotalIMKGz(0 To 5)      As Single
    Dim TotalIMKGi(0 To 5)      As Single
    Dim TotalCemKGz(0 To 3)     As Single
    Dim TotalCemKGi(0 To 3)     As Single
    Dim TotalWatKGz(0 To 1)     As Single
    Dim TotalWatKGi(0 To 1)     As Single
    Dim TotalChemKGz(0 To 5)    As Single
    Dim TotalChemKGi(0 To 5)    As Single
    Dim TotalKGsd               As Single
    Dim TotalRecKGsd            As Single
    Dim TotalVold               As Single
    Dim cnnew                   As ADODB.Connection
    Dim rsnew                   As Recordset
    Dim comm                    As String
    Dim imz(1 To 6)             As Integer
    Dim imi(1 To 6)             As Integer
    Dim cemz(1 To 4)            As Integer
    Dim cemi(1 To 4)            As Integer
    Dim watz(1 To 2)            As Integer
    Dim wati(1 To 2)            As Integer
    Dim chemz(1 To 6)           As Single
    Dim chemi(1 To 6)           As Single
    Dim count                   As Integer
    Dim CountMixd               As Integer
    Dim IMold(1 To 6)           As String
    Dim ScrOld(1 To 4)          As String
    Dim Watold(1 To 2)          As String
    Dim Chemold(1 To 6)         As String
    Dim colx                    As MSComctlLib.ColumnHeader
    Dim itmX                    As MSComctlLib.ListItem
    Dim colw                    As Integer
    Dim colwn                   As Integer
    Dim e                       As Integer
    
    TotalRecKGsd = 0 'нулираме променливата за кг по рецепта
    TotalKGsd = 0 'нулираме променливата за кг по изпълнение
    TotalVold = 0 'нулираме променливата за обем по изпълнение
    colw = 10
    colwn = 20
    
    For i = 0 To 5
        TotalIMKGz(i) = 0
        TotalIMKGi(i) = 0
    Next i
    For i = 0 To 3
        TotalCemKGz(i) = 0
        TotalCemKGi(i) = 0
    Next i
    For i = 0 To 1
        TotalWatKGz(i) = 0
        TotalWatKGi(i) = 0
    Next i
    For i = 0 To 5
        TotalChemKGz(i) = 0
        TotalChemKGi(i) = 0
    Next i
    Me.txtNotes.Refresh
    Me.txtCarData.Refresh
    Me.txtClassH.Refresh
    Me.txtClassK.Refresh
    Me.txtClassP.Refresh
    Me.txtClassV.Refresh
    Me.txtClntData.Refresh
    Me.txtDateReady.Refresh
    Me.txtDrvData.Refresh
    Me.txtObjData.Refresh
    Me.txtOperData.Refresh
    Me.lstMixOld.ColumnHeaders.Clear
    Me.lstMixOld.ListItems.Clear
'------------------------------Start PostgreSQL----------------------------------
    Set cnnew = New ADODB.Connection 'връзка с база данни
        cnnew.ConnectionTimeout = 10
        cnnew.Open ConStr 'отваряме връзката
        
    MousePointer = vbHourglass
    
    comm = "SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & Val(Me.txtNotes) & " ORDER BY mix_num ASC;"
    Set rsnew = cnnew.Execute(comm) 'маркираме всички замеси от намерената експедиция
    If Not rsnew.BOF And Not rsnew.EOF Then
        rsnew.MoveFirst 'отиваме в началото на маркираните замеси
        IMold(1) = rsnew!im1_name
        IMold(2) = rsnew!im2_name
        IMold(3) = rsnew!im3_name
        IMold(4) = rsnew!im4_name
        IMold(5) = rsnew!im5_name
        IMold(6) = rsnew!im6_name
        ScrOld(1) = rsnew!cem1_name
        ScrOld(2) = rsnew!cem2_name
        ScrOld(3) = rsnew!cem3_name
        ScrOld(4) = rsnew!cem4_name
        Watold(1) = rsnew!wat1_name
        Watold(2) = rsnew!wat2_name
        Chemold(1) = rsnew!chem1_name
        Chemold(2) = rsnew!chem2_name
        Chemold(3) = rsnew!chem3_name
        Chemold(4) = rsnew!chem4_name
        Chemold(5) = rsnew!chem5_name
        Chemold(6) = rsnew!chem6_name
        Me.txtDateReady.Text = rsnew!time_mix_ready
        Me.txtClntData.Text = rsnew!name_clnt
        Me.txtObjData.Text = rsnew!obj_clnt
        Me.txtDrvData.Text = rsnew!name_drv
        Me.txtCarData.Text = rsnew!reg_drv
        Me.txtOperData.Text = rsnew!name_op
        Me.txtClassK.Text = rsnew!classk_rec
        Me.txtClassV.Text = rsnew!classv_rec
        Me.txtClassH.Text = rsnew!classh_rec
        Me.txtClassP.Text = rsnew!classp_rec
        'зареждане на заглавките на експедициите от първия замес на експедицията в базата данни
        Set colx = Me.lstMixOld.ColumnHeaders.Add()
            colx.Text = uniNr
            colx.Width = 500
        Set colx = Me.lstMixOld.ColumnHeaders.Add()
            colx.Text = uniOrdCode
            colx.Width = 1150
        Set colx = Me.lstMixOld.ColumnHeaders.Add()
            colx.Text = uniRec & " " & uniNm
            colx.Width = 1200
        Set colx = Me.lstMixOld.ColumnHeaders.Add()
            colx.Text = uniClass
            colx.Width = 1300
        For count = 1 To ns1
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = IMold(count) 'имена на течките на им
                colx.Width = colwn
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        For count = 1 To ns3
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = ScrOld(count) 'имена на течки цимент
                colx.Width = colwn
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        For count = 1 To ns2
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = Watold(count) 'име на течка вода
                colx.Width = colw
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        For count = 1 To ns4
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = Chemold(count) 'имена на течки хд
                colx.Width = colwn
            Set colx = Me.lstMixOld.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        Set colx = Me.lstMixOld.ColumnHeaders.Add()
            colx.Text = "тегло заявено" 'тегло по заявено
            colx.Width = 1200
        Set colx = Me.lstMixOld.ColumnHeaders.Add()
            colx.Text = "тегло измерено" 'тегло по измерено
            colx.Width = 1200
        Set colx = Me.lstMixOld.ColumnHeaders.Add()
            colx.Text = "обем измерено" 'обем по измерено
            colx.Width = 1200
    Else 'ако няма замеси
        GoTo EndLoad
    End If
    Do While Not rsnew.EOF
        If rsnew!exp_num <> Val(Me.txtNotes) Then Exit Do 'излизаме ако номера на експедицията не отговаря на търсенето
        'запис на данните от базата дани в масиви от променливи
        imz(1) = rsnew!im1z
        imz(2) = rsnew!im2z
        imz(3) = rsnew!im3z
        imz(4) = rsnew!im4z
        imz(5) = rsnew!im5z
        imz(6) = rsnew!im6z
        imi(1) = rsnew!im1i
        imi(2) = rsnew!im2i
        imi(3) = rsnew!im3i
        imi(4) = rsnew!im4i
        imi(5) = rsnew!im5i
        imi(6) = rsnew!im6i
        cemz(1) = rsnew!cem1z
        cemz(2) = rsnew!cem2z
        cemz(3) = rsnew!cem3z
        cemz(4) = rsnew!cem4z
        cemi(1) = rsnew!cem1i
        cemi(2) = rsnew!cem2i
        cemi(3) = rsnew!cem3i
        cemi(4) = rsnew!cem4i
        watz(1) = rsnew!wat1z
        wati(1) = rsnew!wat1i
        watz(2) = rsnew!wat2z
        wati(2) = rsnew!wat2i
        chemz(1) = rDs(rsnew!chem1z)
        chemz(2) = rDs(rsnew!chem2z)
        chemz(3) = rDs(rsnew!chem3z)
        chemz(4) = rDs(rsnew!chem4z)
        chemz(5) = rDs(rsnew!chem5z)
        chemz(6) = rDs(rsnew!chem6z)
        chemi(1) = rDs(rsnew!chem1i)
        chemi(2) = rDs(rsnew!chem2i)
        chemi(3) = rDs(rsnew!chem3i)
        chemi(4) = rDs(rsnew!chem4i)
        chemi(5) = rDs(rsnew!chem5i)
        chemi(6) = rDs(rsnew!chem6i)
        TotalRecKGsd = TotalRecKGsd + CSng(rDs(rsnew!total_rec_kg)) 'сумираме теглата по заявено от всеки замес
        TotalKGsd = TotalKGsd + CSng(rDs(rsnew!total_real_kg)) 'сумираме теглата по измерено от всеки замес
        TotalVold = TotalVold + CSng(rDs(rsnew!total_vol)) 'сумираме количества от всеки замес

        'сума на отделните ИМ по зададено
        TotalIMKGz(0) = TotalIMKGz(0) + Val(rsnew!im1z)
        TotalIMKGz(1) = TotalIMKGz(1) + Val(rsnew!im2z)
        TotalIMKGz(2) = TotalIMKGz(2) + Val(rsnew!im3z)
        TotalIMKGz(3) = TotalIMKGz(3) + Val(rsnew!im4z)
        TotalIMKGz(4) = TotalIMKGz(4) + Val(rsnew!im5z)
        TotalIMKGz(5) = TotalIMKGz(5) + Val(rsnew!im6z)
        
        'сума на отделните ИМ по изпълнено
        TotalIMKGi(0) = TotalIMKGi(0) + Val(rsnew!im1i)
        TotalIMKGi(1) = TotalIMKGi(1) + Val(rsnew!im2i)
        TotalIMKGi(2) = TotalIMKGi(2) + Val(rsnew!im3i)
        TotalIMKGi(3) = TotalIMKGi(3) + Val(rsnew!im4i)
        TotalIMKGi(4) = TotalIMKGi(4) + Val(rsnew!im5i)
        TotalIMKGi(5) = TotalIMKGi(5) + Val(rsnew!im6i)
        
        'сума на отделните цименти по зададено
        TotalCemKGz(0) = TotalCemKGz(0) + Val(rsnew!cem1z)
        TotalCemKGz(1) = TotalCemKGz(1) + Val(rsnew!cem2z)
        TotalCemKGz(2) = TotalCemKGz(2) + Val(rsnew!cem3z)
        TotalCemKGz(3) = TotalCemKGz(3) + Val(rsnew!cem4z)
        
        'сума на отделните цименти по изпълнено
        TotalCemKGi(0) = TotalCemKGi(0) + Val(rsnew!cem1i)
        TotalCemKGi(1) = TotalCemKGi(1) + Val(rsnew!cem2i)
        TotalCemKGi(2) = TotalCemKGi(2) + Val(rsnew!cem3i)
        TotalCemKGi(3) = TotalCemKGi(3) + Val(rsnew!cem4i)
        
        'сума на вода по зададено
        TotalWatKGz(0) = TotalWatKGz(0) + Val(rsnew!wat1z)
        TotalWatKGz(1) = TotalWatKGz(1) + Val(rsnew!wat2z)
        
        'сума на вода по изпълнено
        TotalWatKGi(0) = TotalWatKGi(0) + Val(rsnew!wat1i)
        TotalWatKGi(1) = TotalWatKGi(1) + Val(rsnew!wat2i)
    
        'сума на отделните хд по зададено
        TotalChemKGz(0) = TotalChemKGz(0) + CSng(rDs(rsnew!chem1z))
        TotalChemKGz(1) = TotalChemKGz(1) + CSng(rDs(rsnew!chem2z))
        TotalChemKGz(2) = TotalChemKGz(2) + CSng(rDs(rsnew!chem3z))
        TotalChemKGz(3) = TotalChemKGz(3) + CSng(rDs(rsnew!chem4z))
        TotalChemKGz(4) = TotalChemKGz(4) + CSng(rDs(rsnew!chem5z))
        TotalChemKGz(5) = TotalChemKGz(5) + CSng(rDs(rsnew!chem6z))
        
        'сума на отделните хд по изпълнено
        TotalChemKGi(0) = TotalChemKGi(0) + CSng(rDs(rsnew!chem1i))
        TotalChemKGi(1) = TotalChemKGi(1) + CSng(rDs(rsnew!chem2i))
        TotalChemKGi(2) = TotalChemKGi(2) + CSng(rDs(rsnew!chem3i))
        TotalChemKGi(3) = TotalChemKGi(3) + CSng(rDs(rsnew!chem4i))
        TotalChemKGi(4) = TotalChemKGi(4) + CSng(rDs(rsnew!chem5i))
        TotalChemKGi(5) = TotalChemKGi(5) + CSng(rDs(rsnew!chem6i))
        
        CountMixd = CountMixd + 1 'брояч на замесите
        Set itmX = Me.lstMixOld.ListItems.Add(1, , Format(CountMixd, "00")) 'запис в ListView
            itmX.SubItems(1) = Format(rsnew!ord_num, "0000000")
            itmX.SubItems(2) = rsnew!name_rec
            itmX.SubItems(3) = rsnew!class_rec
        For e = 1 To ns1
            itmX.SubItems(2 * e + 2) = imz(e)
            itmX.SubItems(2 * e + 3) = imi(e)
        Next e
        For e = 1 To ns3
            itmX.SubItems(2 * (e + ns1) + 2) = cemz(e)
            itmX.SubItems(2 * (e + ns1) + 3) = cemi(e)
        Next e
        For e = 1 To ns2
            itmX.SubItems(2 * (e + ns1 + ns3) + 2) = watz(e)
            itmX.SubItems(2 * (e + ns1 + ns3) + 3) = wati(e)
        Next e
        For e = 1 To ns4
            itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 2) = chemz(e)
            itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 3) = chemi(e)
        Next e
            itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 2) = rDs(rsnew!total_rec_kg)
            itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 3) = rDs(rsnew!total_real_kg)
            itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 4) = rDs(rsnew!total_vol)
        rsnew.MoveNext
    Loop
    
    'запис в ListView на празен ред
    Set itmX = Me.lstMixOld.ListItems.Add(1, , "X")
        itmX.SubItems(1) = "-------------"
        itmX.SubItems(2) = "-------------"
        itmX.SubItems(3) = "-------------"
    For e = 1 To ns1
        itmX.SubItems(2 * e + 2) = "-------------"
        itmX.SubItems(2 * e + 3) = "-------------"
    Next e
    For e = 1 To ns3
        itmX.SubItems(2 * (e + ns1) + 2) = "-------------"
        itmX.SubItems(2 * (e + ns1) + 3) = "-------------"
    Next e
    For e = 1 To ns2
        itmX.SubItems(2 * (e + ns1 + ns3) + 2) = "-------------"
        itmX.SubItems(2 * (e + ns1 + ns3) + 3) = "-------------"
    Next e
    For e = 1 To ns4
        itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 2) = "-------------"
        itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 3) = "-------------"
    Next e
        itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 2) = "-------------"
        itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 3) = "-------------"
        itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 4) = "-------------"
    
    'запис в ListView на тоталите
    Set itmX = Me.lstMixOld.ListItems.Add(1, , "XX")
        itmX.SubItems(1) = "-ТОТАЛ-"

    For e = 1 To ns1
        itmX.SubItems(2 * e + 2) = TotalIMKGz(e - 1)
        itmX.SubItems(2 * e + 3) = TotalIMKGi(e - 1)
    Next e
    For e = 1 To ns3
        itmX.SubItems(2 * (e + ns1) + 2) = TotalCemKGz(e - 1)
        itmX.SubItems(2 * (e + ns1) + 3) = TotalCemKGi(e - 1)
    Next e
    For e = 1 To ns2
        itmX.SubItems(2 * (ns1 + ns3 + ns2) + 2) = TotalWatKGz(e - 1)
        itmX.SubItems(2 * (ns1 + ns3 + ns2) + 3) = TotalWatKGi(e - 1)
    Next e
    For e = 1 To ns4
        itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 2) = TotalChemKGz(e - 1)
        itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 3) = TotalChemKGi(e - 1)
    Next e
        itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 2) = TotalRecKGsd
        itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 3) = TotalKGsd
        itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 4) = TotalVold
    
EndLoad:
    rsnew.Close 'затваряме записите
    Set rsnew = Nothing
    cnnew.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnnew = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------

    'автонастройка на ListView
    If Me.lstMixOld.ListItems.count > 0 Then AutoColW Me.lstMixOld
    
    'активиране и деактивиране на бутони за печат
    If Me.lstMixOld.ListItems.count < 1 Then
        Me.btnForm1.Enabled = False
        Me.btnForm2.Enabled = False
        Me.btnForm3.Enabled = False
    Else
        Me.btnForm1.Enabled = True
        Me.btnForm2.Enabled = True
        Me.btnForm3.Enabled = True
    End If
End Sub

Private Sub btnForm1_Click()

    Call BtnFillForm1
End Sub

Private Sub btnForm2_Click()

    Call BtnFillForm2
End Sub

Private Sub btnForm3_Click()

    Call BtnFillForm3
End Sub

Private Sub btnClose_Click()

    Unload Me
End Sub

