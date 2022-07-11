VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmNotes"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
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
      Left            =   12720
      TabIndex        =   27
      Top             =   7440
      Width           =   735
   End
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
      TabIndex        =   15
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3000
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
      Left            =   8640
      TabIndex        =   3
      Top             =   7440
      Width           =   2055
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
      Left            =   5880
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
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
      Left            =   3120
      TabIndex        =   1
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox txtClassH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
   Begin MSComctlLib.ListView lstNotes 
      Height          =   4455
      Left            =   240
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
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
    

'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    Dim i As Integer
    
    Set cnR = New ADODB.Connection 'връзка с база данни
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr 'отваряме връзката
    MousePointer = vbHourglass
    
'визуализация на замесите от последната експедиция
    Set rsR = cnR.Execute("SELECT exp_num FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1") 'маркираме последния замес
    
    If Not rsR.EOF And Not rsR.BOF Then
        i = rsR!exp_num 'маркираме номера на експедиция от последния замес
    Else 'ако няма замеси
        GoTo EndLoad
    End If
    
    Me.udNotes.BuddyControl = Me.txtNotes
    Me.udNotes.Max = i
    Me.udNotes.Increment = 1
    Me.txtNotes.Text = Me.udNotes.Max
    
EndLoad:
    rsR.Close 'затваряме записите
    Set rsR = Nothing
    cnR.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnR = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------
    
    If Me.lstNotes.ListItems.count < 1 Then
        Me.btnForm1.Enabled = False
        Me.btnForm2.Enabled = False
        Me.btnForm3.Enabled = False
    Else
        Me.btnForm1.Enabled = True
        Me.btnForm2.Enabled = True
        Me.btnForm3.Enabled = True
    End If
    
End Sub

Private Sub txtNotes_Change()
    
    Dim TotalIMKGz(0 To 5) As Single
    Dim TotalIMKGi(0 To 5) As Single
    Dim TotalCemKGz(0 To 3) As Single
    Dim TotalCemKGi(0 To 3) As Single
    Dim TotalWatKGz(0 To 1) As Single
    Dim TotalWatKGi(0 To 1) As Single
    Dim TotalChemKGz(0 To 5) As Single
    Dim TotalChemKGi(0 To 5) As Single
    Dim TotalKG As Single
    Dim TotalRecKG As Single
    Dim TotalVol As Single
    
    If MachineNumber = 1 Then
        ns1 = n1s1
        ns2 = n1s2
        ns3 = n1s3
        ns4 = n1s4
    ElseIf MachineNumber = 2 Then
        ns1 = n2s1
        ns2 = n2s2
        ns3 = n2s3
        ns4 = n2s4
    End If
    
    TotalRecKG = 0 'нулираме променливата за кг по рецепта
    TotalKG = 0 'нулираме променливата за кг по изпълнение
    TotalVol = 0 'нулираме променливата за обем по изпълнение
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
    Me.lstNotes.ColumnHeaders.Clear
    Me.lstNotes.ListItems.Clear
    
'------------------------------Start PostgreSQL----------------------------------
    Dim cnR As ADODB.Connection
    Dim rsR As Recordset
    Dim comm As String
    Dim imz(1 To 6) As Integer
    Dim imi(1 To 6) As Integer
    Dim cemz(1 To 4) As Integer
    Dim cemi(1 To 4) As Integer
    Dim watz(1 To 2) As Integer
    Dim wati(1 To 2) As Integer
    Dim chemz(1 To 6) As Single
    Dim chemi(1 To 6) As Single
    Dim expq As Single
    Dim count As Integer
    Dim IMOld(1 To 6) As String
    Dim ScrOld(1 To 4) As String
    Dim WatOld(1 To 2) As String
    Dim ChemOld(1 To 6) As String
    
    Set cnR = New ADODB.Connection 'връзка с база данни
    cnR.ConnectionTimeout = 10
    cnR.Open ConStr 'отваряме връзката
    MousePointer = vbHourglass
    
    comm = "SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & Val(Me.txtNotes) & ";"
    
    Set rsR = cnR.Execute(comm) 'маркираме всички замеси от намерената експедиция
    
    If Not rsR.BOF And Not rsR.EOF Then
        rsR.MoveFirst 'отиваме в началото на маркираните замеси
        
        IMOld(1) = rsR!im1_name
        IMOld(2) = rsR!im2_name
        IMOld(3) = rsR!im3_name
        IMOld(4) = rsR!im4_name
        IMOld(5) = rsR!im5_name
        IMOld(6) = rsR!im6_name
        ScrOld(1) = rsR!cem1_name
        ScrOld(2) = rsR!cem2_name
        ScrOld(3) = rsR!cem3_name
        ScrOld(4) = rsR!cem4_name
        WatOld(1) = rsR!wat1_name
        WatOld(2) = rsR!wat2_name
        ChemOld(1) = rsR!chem1_name
        ChemOld(2) = rsR!chem2_name
        ChemOld(3) = rsR!chem3_name
        ChemOld(4) = rsR!chem4_name
        ChemOld(5) = rsR!chem5_name
        ChemOld(6) = rsR!chem6_name
        
        Me.txtDateReady.Text = rsR!time_mix_ready
        Me.txtClntData.Text = rsR!name_clnt
        Me.txtObjData.Text = rsR!obj_clnt
        Me.txtDrvData.Text = rsR!name_drv
        Me.txtCarData.Text = rsR!reg_drv
        Me.txtOperData.Text = rsR!name_op
        Me.txtClassK.Text = rsR!classk_rec
        Me.txtClassV.Text = rsR!classv_rec
        Me.txtClassH.Text = rsR!classh_rec
        Me.txtClassP.Text = rsR!classp_rec
        
        'зареждане на заглавките на експедициите от първия замес на експедицията в базата данни
        Set colx = Me.lstNotes.ColumnHeaders.Add()
            colx.Text = uniNr
            colx.Width = 500
        Set colx = Me.lstNotes.ColumnHeaders.Add()
            colx.Text = uniOrdCode
            colx.Width = 1150
        Set colx = Me.lstNotes.ColumnHeaders.Add()
            colx.Text = uniRec & " " & uniNm
            colx.Width = 1200
        Set colx = Me.lstNotes.ColumnHeaders.Add()
            colx.Text = uniClass
            colx.Width = 1300
        For count = 1 To ns1
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = IMOld(count) 'имена на течките на им
                colx.Width = colwn
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        For count = 1 To ns3
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = ScrOld(count) 'имена на течки цимент
                colx.Width = colwn
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        For count = 1 To ns2
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = WatOld(count) 'име на течка вода
                colx.Width = colw
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        For count = 1 To ns4
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = ChemOld(count) 'имена на течки хд
                colx.Width = colwn
            Set colx = Me.lstNotes.ColumnHeaders.Add()
                colx.Text = uniMeasured
                colx.Width = colw
        Next count
        Set colx = Me.lstNotes.ColumnHeaders.Add()
            colx.Text = "тегло заявено" 'тегло по заявено
            colx.Width = 1200
        Set colx = Me.lstNotes.ColumnHeaders.Add()
            colx.Text = "тегло измерено" 'тегло по измерено
            colx.Width = 1200
        Set colx = Me.lstNotes.ColumnHeaders.Add()
            colx.Text = "обем измерено" 'обем по измерено
            colx.Width = 1200
    Else 'ако няма замеси
        GoTo EndLoad
    End If
        
    Do While Not rsR.EOF
        If rsR!exp_num <> Val(Me.txtNotes) Then Exit Do 'излизаме ако номера на експедицията не отговаря на търсенето
        
        expq = rDs(rsR!exp_q) 'запис на данните от базата дани в масиви от променливи
        imz(1) = rsR!im1z
        imz(2) = rsR!im2z
        imz(3) = rsR!im3z
        imz(4) = rsR!im4z
        imz(5) = rsR!im5z
        imz(6) = rsR!im6z
        imi(1) = rsR!im1i
        imi(2) = rsR!im2i
        imi(3) = rsR!im3i
        imi(4) = rsR!im4i
        imi(5) = rsR!im5i
        imi(6) = rsR!im6i
        cemz(1) = rsR!cem1z
        cemz(2) = rsR!cem2z
        cemz(3) = rsR!cem3z
        cemz(4) = rsR!cem4z
        cemi(1) = rsR!cem1i
        cemi(2) = rsR!cem2i
        cemi(3) = rsR!cem3i
        cemi(4) = rsR!cem4i
        watz(1) = rsR!wat1z
        wati(1) = rsR!wat1i
        watz(2) = rsR!wat2z
        wati(2) = rsR!wat2i
        chemz(1) = rDs(rsR!chem1z)
        chemz(2) = rDs(rsR!chem2z)
        chemz(3) = rDs(rsR!chem3z)
        chemz(4) = rDs(rsR!chem4z)
        chemz(5) = rDs(rsR!chem5z)
        chemz(6) = rDs(rsR!chem6z)
        chemi(1) = rDs(rsR!chem1i)
        chemi(2) = rDs(rsR!chem2i)
        chemi(3) = rDs(rsR!chem3i)
        chemi(4) = rDs(rsR!chem4i)
        chemi(5) = rDs(rsR!chem5i)
        chemi(6) = rDs(rsR!chem6i)
        TotalRecKGsd = TotalRecKGsd + CSng(rDs(rsR!total_rec_kg)) 'сумираме теглата по заявено от всеки замес
        TotalKGsd = TotalKGsd + CSng(rDs(rsR!total_real_kg)) 'сумираме теглата по измерено от всеки замес
        TotalVold = TotalVold + CSng(rDs(rsR!total_vol)) 'сумираме количества от всеки замес
        
        'сума на отделните ИМ по зададено
        TotalIMKGz(0) = TotalIMKGz(0) + Val(rsR!im1z)
        TotalIMKGz(1) = TotalIMKGz(1) + Val(rsR!im2z)
        TotalIMKGz(2) = TotalIMKGz(2) + Val(rsR!im3z)
        TotalIMKGz(3) = TotalIMKGz(3) + Val(rsR!im4z)
        TotalIMKGz(4) = TotalIMKGz(4) + Val(rsR!im5z)
        TotalIMKGz(5) = TotalIMKGz(5) + Val(rsR!im6z)
        
        'сума на отделните ИМ по изпълнено
        TotalIMKGi(0) = TotalIMKGi(0) + Val(rsR!im1i)
        TotalIMKGi(1) = TotalIMKGi(1) + Val(rsR!im2i)
        TotalIMKGi(2) = TotalIMKGi(2) + Val(rsR!im3i)
        TotalIMKGi(3) = TotalIMKGi(3) + Val(rsR!im4i)
        TotalIMKGi(4) = TotalIMKGi(4) + Val(rsR!im5i)
        TotalIMKGi(5) = TotalIMKGi(5) + Val(rsR!im6i)
        
        'сума на отделните цименти по зададено
        TotalCemKGz(0) = TotalCemKGz(0) + Val(rsR!cem1z)
        TotalCemKGz(1) = TotalCemKGz(1) + Val(rsR!cem2z)
        TotalCemKGz(2) = TotalCemKGz(2) + Val(rsR!cem3z)
        TotalCemKGz(3) = TotalCemKGz(3) + Val(rsR!cem4z)
        
        'сума на отделните цименти по изпълнено
        TotalCemKGi(0) = TotalCemKGi(0) + Val(rsR!cem1i)
        TotalCemKGi(1) = TotalCemKGi(1) + Val(rsR!cem2i)
        TotalCemKGi(2) = TotalCemKGi(2) + Val(rsR!cem3i)
        TotalCemKGi(3) = TotalCemKGi(3) + Val(rsR!cem4i)
        
        'сума на вода по зададено
        TotalWatKGz(0) = TotalWatKGz(0) + Val(rsR!wat1z)
        TotalWatKGz(1) = TotalWatKGz(1) + Val(rsR!wat2z)
        
        'сума на вода по изпълнено
        TotalWatKGi(0) = TotalWatKGi(0) + Val(rsR!wat1i)
        TotalWatKGi(1) = TotalWatKGi(1) + Val(rsR!wat2i)
        
        'сума на отделните хд по зададено
        TotalChemKGz(0) = TotalChemKGz(0) + CSng(rDs(rsR!chem1z))
        TotalChemKGz(1) = TotalChemKGz(1) + CSng(rDs(rsR!chem2z))
        TotalChemKGz(2) = TotalChemKGz(2) + CSng(rDs(rsR!chem3z))
        TotalChemKGz(3) = TotalChemKGz(3) + CSng(rDs(rsR!chem4z))
        TotalChemKGz(4) = TotalChemKGz(4) + CSng(rDs(rsR!chem5z))
        TotalChemKGz(5) = TotalChemKGz(5) + CSng(rDs(rsR!chem6z))
        
        'сума на отделните хд по изпълнено
        TotalChemKGi(0) = TotalChemKGi(0) + CSng(rDs(rsR!chem1i))
        TotalChemKGi(1) = TotalChemKGi(1) + CSng(rDs(rsR!chem2i))
        TotalChemKGi(2) = TotalChemKGi(2) + CSng(rDs(rsR!chem3i))
        TotalChemKGi(3) = TotalChemKGi(3) + CSng(rDs(rsR!chem4i))
        TotalChemKGi(4) = TotalChemKGi(4) + CSng(rDs(rsR!chem5i))
        TotalChemKGi(5) = TotalChemKGi(5) + CSng(rDs(rsR!chem6i))
        
        CountMixd = CountMixd + 1 'брояч на замесите
        Set itmX = Me.lstNotes.ListItems.Add(1, , Format(CountMixd, "00")) 'запис в ListView
            itmX.SubItems(1) = Format(rsR!ord_num, "0000000")
            itmX.SubItems(2) = rsR!name_rec
            itmX.SubItems(3) = rsR!class_rec
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
            itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 2) = rDs(rsR!total_rec_kg)
            itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 3) = rDs(rsR!total_real_kg)
            itmX.SubItems(2 * (ns1 + ns3 + ns4 + ns2 + 1) + 4) = rDs(rsR!total_vol)
        rsR.MoveNext
    Loop
    
    Set itmX = Me.lstNotes.ListItems.Add(1, , "X") 'запис в ListView на празен ред
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
    
    Set itmX = Me.lstNotes.ListItems.Add(1, , "XX") 'запис в ListView на тоталите
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
    rsR.Close 'затваряме записите
    Set rsR = Nothing
    cnR.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnR = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------

'автонастройка на ListView
    If Me.lstNotes.ListItems.count > 0 Then AutoColW Me.lstNotes
    
    If Me.lstNotes.ListItems.count < 1 Then
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

Private Sub btnForm1_Click()
    Call BtnFillForm1
End Sub

Private Sub btnForm2_Click()
    Call BtnFillForm2
End Sub

Private Sub btnForm3_Click()
    Call BtnFillForm3
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStartRep.Show
End Sub

