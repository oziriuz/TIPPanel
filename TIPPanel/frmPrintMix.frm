VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintMix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmPrintMix"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnKillPrint 
      Caption         =   "btnKillPrint"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar barPrint 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   500
      Scrolling       =   1
   End
   Begin VB.PictureBox pbPrint 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmPrintMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnKillPrint_Click()

    Printer.KillDoc
    Me.Hide
End Sub

Private Sub Form_Load()

    Me.Caption = uniSendingPrinter
    Me.btnKillPrint.Caption = UniCancel
End Sub
