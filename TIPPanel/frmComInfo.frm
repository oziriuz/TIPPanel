VERSION 5.00
Begin VB.Form frmComInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmComInfo"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   Icon            =   "frmComInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFax 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtTel 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtConcretePlant 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "btnCancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
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
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtTown 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtCompany 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblFax 
      Alignment       =   1  'Right Justify
      Caption         =   "lblFax"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblTel 
      Alignment       =   1  'Right Justify
      Caption         =   "lblTel"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblConcretePlant 
      Alignment       =   1  'Right Justify
      Caption         =   "lblConcretePlant"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblTown 
      Alignment       =   1  'Right Justify
      Caption         =   "lblTown"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      Caption         =   "lblCompany"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmComInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intEmpFile As Integer
Dim Comp           As String
Dim Town           As String
Dim ConcP          As String
Dim Tel            As String
Dim Fax            As String

Private Sub Form_Load()

    Me.Caption = uniComInfo
    Me.lblCompany = uniFirm
    Me.lblTown = uniTown
    Me.lblConcretePlant = uniConcPlant
    '·„
    Me.lblTel = "»ÌÙÓ 1:"
    Me.lblFax = "»ÌÙÓ 2:"
    Me.btnSave.Caption = uniSave
    Me.btnCancel.Caption = UniCancel
    
    intEmpFile = FreeFile

    If Dir(InfoFile) <> "" Then
        Open InfoFile For Input As intEmpFile
        Input #intEmpFile, Comp, Town, ConcP, Tel, Fax
        Close
    End If
    
    Me.txtCompany.Text = Comp
    Me.txtTown.Text = Town
    Me.txtConcretePlant.Text = ConcP
    Me.txtTel.Text = Tel
    Me.txtFax.Text = Fax
End Sub

Private Sub btnSave_Click()

    intEmpFile = FreeFile
    Comp = Me.txtCompany
    Town = Me.txtTown
    ConcP = Me.txtConcretePlant
    Tel = Me.txtTel
    Fax = Me.txtFax
    
    Open InfoFile For Output As intEmpFile
    Write #intEmpFile, Comp, Town, ConcP, Tel, Fax
    Close
    Unload Me
End Sub

Private Sub btnCancel_Click()

    Unload Me
End Sub

