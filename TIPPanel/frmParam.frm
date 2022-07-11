VERSION 5.00
Begin VB.Form frmParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmParam"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2940
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   2940
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumChemSilos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtNumWaterSilos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtNumCementSilos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtNumIMSilos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtTimePourDefault 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtTimeMixDefault 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtMixCap 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblTimePourDefault 
      Alignment       =   1  'Right Justify
      Caption         =   "lblTimePourDefault"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblTimeMixDefault 
      Alignment       =   1  'Right Justify
      Caption         =   "lblTimeMixDefault"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblNumChemSilos 
      Alignment       =   1  'Right Justify
      Caption         =   "lblNumChemSilos"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblNumWaterSilos 
      Alignment       =   1  'Right Justify
      Caption         =   "lblNumWaterSilos"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblNumCementSilos 
      Alignment       =   1  'Right Justify
      Caption         =   "lblNumCementSilos"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblNumIMSilos 
      Alignment       =   1  'Right Justify
      Caption         =   "lblNumIMSilos"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblMixCap 
      Alignment       =   1  'Right Justify
      Caption         =   "lblMixCap"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim cn    As ADODB.Connection
    Dim rs    As Recordset
    Dim MCap  As Single
    Dim Tmix  As Integer
    Dim Tpour As Integer
    Dim Nim   As Integer
    Dim Nsil  As Integer
    Dim Nwat  As Integer
    Dim Nchem As Integer
    
    frmParam.Caption = btnParamSys
    lblMixCap.Caption = uniMixCap
    lblTimeMixDefault.Caption = uniTimeMix
    lblTimePourDefault.Caption = uniOpenMix
    lblNumIMSilos.Caption = uniNumIM
    lblNumCementSilos.Caption = uniNumCem
    lblNumWaterSilos.Caption = uniNumWat
    lblNumChemSilos.Caption = uniNumChem
    
    MCap = 1
    Tmix = 10
    Tpour = 10
    Nim = 1
    Nsil = 1
    Nwat = 1
    Nchem = 1
    '-----------------------Start postgreSQL-----------------------------------
    'изчитане на параметри от база данни
    Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
    
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT * FROM settings_bc" & MachineNumber & " WHERE ind = 1;")
    If Not rs.BOF And Not rs.EOF Then
        Nim = rs!im_num
        Nsil = rs!cem_num
        Nwat = rs!wat_num
        Nchem = rs!chem_num
    Else
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
'--------------------------End PostgreSQL-----------------------------------

    txtMixCap.Text = ARound(IEEE754(frmOPC.MixCap.ItemValue), 2)
    txtTimeMixDefault.Text = Hex(frmOPC.TimeMixDefault.ItemValue) / 10
    txtTimePourDefault.Text = Hex(frmOPC.TimePourDefault.ItemValue) / 10
    txtNumIMSilos.Text = frmOPC.NumIMSilos.ItemValue
    txtNumCementSilos.Text = frmOPC.NumCementSilos.ItemValue
    txtNumWaterSilos.Text = frmOPC.NumWaterSilos.ItemValue
    txtNumChemSilos.Text = frmOPC.NumChemSilos.ItemValue
    
    If Val(txtMixCap.Text) <= 0 Then txtMixCap.Text = MCap
    If Val(txtTimeMixDefault.Text) <= 0 Then txtTimeMixDefault.Text = Tmix
    If Val(txtTimePourDefault.Text) <= 0 Then txtTimePourDefault.Text = Tpour
    If Val(txtNumIMSilos.Text) <= 0 Then txtNumIMSilos.Text = Nim
    If Val(txtNumCementSilos.Text) <= 0 Then txtNumCementSilos.Text = Nsil
    If Val(txtNumWaterSilos.Text) <= 0 Then txtNumWaterSilos.Text = Nwat
    If Val(txtNumChemSilos.Text) <= 0 Then txtNumChemSilos.Text = Nchem
    If Val(txtNumIMSilos.Text) > 5 Then txtNumIMSilos.Text = 5
    If Val(txtNumCementSilos.Text) > 4 Then txtNumCementSilos.Text = 4
    If Val(txtNumWaterSilos.Text) > 1 Then txtNumWaterSilos.Text = 1
    If Val(txtNumChemSilos.Text) > 6 Then txtNumChemSilos.Text = 6
End Sub

