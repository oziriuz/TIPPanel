VERSION 5.00
Begin VB.Form setForm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������� ����� 2"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12210
   Icon            =   "setForm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOper 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2655
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
      Height          =   495
      Left            =   9840
      TabIndex        =   82
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CheckBox chRealVol 
      Caption         =   "������������ �� �������� ���������� ���������� ����� � ����� ""����"""
      Height          =   195
      Left            =   2040
      TabIndex        =   81
      Top             =   6720
      Width           =   6975
   End
   Begin VB.TextBox txtClassP 
      Appearance      =   0  'Flat
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtChem1 
      Appearance      =   0  'Flat
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtMixTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtChem2 
      Appearance      =   0  'Flat
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtCem2 
      Appearance      =   0  'Flat
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtExpTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtEDM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3840
      Width           =   210
   End
   Begin VB.TextBox txtChem3 
      Appearance      =   0  'Flat
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtCem3 
      Appearance      =   0  'Flat
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtCem1 
      Appearance      =   0  'Flat
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtClassH 
      Appearance      =   0  'Flat
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
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtClass 
      Appearance      =   0  'Flat
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtOrdVol 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtRecType 
      Appearance      =   0  'Flat
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtDist 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtExpNote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtClassV 
      Appearance      =   0  'Flat
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtClassK 
      Appearance      =   0  'Flat
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtVol 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtW 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtDrvNo 
      Appearance      =   0  'Flat
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtDrv 
      Appearance      =   0  'Flat
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txtObj 
      Appearance      =   0  'Flat
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox txtClnt 
      Appearance      =   0  'Flat
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtOrd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.CheckBox chDist 
      Height          =   195
      Left            =   11040
      TabIndex        =   18
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chRecType 
      Height          =   195
      Left            =   3840
      TabIndex        =   17
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chVol 
      Height          =   195
      Left            =   5520
      TabIndex        =   16
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chW 
      Height          =   195
      Left            =   8040
      TabIndex        =   15
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chOrdVol 
      Height          =   195
      Left            =   11160
      TabIndex        =   14
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chClass 
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chClassK 
      Height          =   195
      Left            =   5520
      TabIndex        =   12
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chClassV 
      Height          =   195
      Left            =   8040
      TabIndex        =   11
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chClassH 
      Height          =   195
      Left            =   11520
      TabIndex        =   10
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chCem3 
      Height          =   195
      Left            =   11520
      TabIndex        =   9
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chCem2 
      Height          =   195
      Left            =   8520
      TabIndex        =   8
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chCem1 
      Height          =   195
      Left            =   5520
      TabIndex        =   7
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chChem1 
      Height          =   195
      Left            =   5520
      TabIndex        =   6
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chChem2 
      Height          =   195
      Left            =   8520
      TabIndex        =   5
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chChem3 
      Height          =   195
      Left            =   11520
      TabIndex        =   4
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chEDM 
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chClassP 
      Height          =   195
      Left            =   5520
      TabIndex        =   2
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chMixTime 
      Height          =   195
      Left            =   8640
      TabIndex        =   1
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chExpTime 
      Height          =   195
      Left            =   11520
      TabIndex        =   0
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblOperSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/..................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4800
      TabIndex        =   95
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"setForm2.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   840
      TabIndex        =   94
      Top             =   5760
      Width           =   11070
   End
   Begin VB.Label lblAddSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������������: ........................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8160
      TabIndex        =   93
      Top             =   6120
      Width           =   3780
   End
   Begin VB.Label lblAddType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���: ..................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   92
      Top             =   6120
      Width           =   2580
   End
   Begin VB.Label lblAddPlace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����: ................................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   91
      Top             =   6120
      Width           =   3555
   End
   Begin VB.Label lblTimeArr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� �� ����������: ......................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   90
      Top             =   4440
      Width           =   4065
   End
   Begin VB.Label lblPlaceMix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������������: .............................................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   89
      Top             =   4440
      Width           =   5205
   End
   Begin VB.Label lblSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����: ....................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   88
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label lblDrvSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����: ........................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   87
      Top             =   5280
      Width           =   2190
   End
   Begin VB.Label lblOper 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   86
      Top             =   5280
      Width           =   1020
   End
   Begin VB.Label lblPlaceEnd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� ��������: ........................................"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   85
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label lblPlaceStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ �� ����������: ...................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   84
      Top             =   4800
      Width           =   4260
   End
   Begin VB.Label lblClassP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   80
      Top             =   3840
      Width           =   1320
   End
   Begin VB.Label lblCem2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   6120
      TabIndex        =   79
      Top             =   3120
      Width           =   285
   End
   Begin VB.Label lblCem1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3000
      TabIndex        =   78
      Top             =   3120
      Width           =   285
   End
   Begin VB.Label lblMM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   77
      Top             =   3840
      Width           =   330
   End
   Begin VB.Label lblChem3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   9360
      TabIndex        =   76
      Top             =   3480
      Width           =   285
   End
   Begin VB.Label lblChem2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   6120
      TabIndex        =   75
      Top             =   3480
      Width           =   285
   End
   Begin VB.Label lblChem1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   3000
      TabIndex        =   74
      Top             =   3480
      Width           =   285
   End
   Begin VB.Label lblCem3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9360
      TabIndex        =   73
      Top             =   3120
      Width           =   285
   End
   Begin VB.Label lblM34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11760
      TabIndex        =   72
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label lblM33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11520
      TabIndex        =   71
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblKG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   70
      Top             =   2400
      Width           =   210
   End
   Begin VB.Label lblM32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6120
      TabIndex        =   69
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label lblM31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      TabIndex        =   68
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label lblOrdVol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� ��������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   67
      Top             =   2400
      Width           =   1710
   End
   Begin VB.Label lblRecType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� �������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   66
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label lblKM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "km"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11520
      TabIndex        =   65
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label lblDist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9960
      TabIndex        =   64
      Top             =   2040
      Width           =   570
   End
   Begin VB.Label lblChem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������� �������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   63
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblEDM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D max �� ���:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   62
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblCem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� ������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   61
      Top             =   3120
      Width           =   1185
   End
   Begin VB.Label lblClassH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� �-��� �� �������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   60
      Top             =   2760
      Width           =   2385
   End
   Begin VB.Label lblClassV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� ����.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6000
      TabIndex        =   59
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Label lblClassK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� �������.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   58
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �� �����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   57
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6720
      TabIndex        =   56
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label lvlVol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4440
      TabIndex        =   55
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label lblExpTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� �� ����������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9240
      TabIndex        =   54
      Top             =   3840
      Width           =   1785
   End
   Begin VB.Label lblMixTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ �� ���������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   53
      Top             =   3840
      Width           =   2010
   End
   Begin VB.Label lblDrvNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6720
      TabIndex        =   52
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblDrv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   51
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lblObj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6720
      TabIndex        =   50
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label lblClnt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   49
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label lblOrd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� ������ No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9240
      TabIndex        =   48
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7080
      TabIndex        =   47
      Top             =   840
      Width           =   105
   End
   Begin VB.Label lblExpNote 
      BackStyle       =   0  'Transparent
      Caption         =   "������������� ������� No."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   46
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������� ���-����� v1.0b/2013"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10320
      TabIndex        =   45
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "setForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim PrevSet As Boolean
    Dim strSubKey As String

Private Sub Form_Load()
    Me.btnSave.Caption = uniSave
    strSubKey = Trim(PlaceProgSet2)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    On Error Resume Next
    
    If PrevSet = True Then
        rDist = GetSetting(PlaceProgSettings, PlaceForm2, "Dist", ErrRes)
        rRecType = GetSetting(PlaceProgSettings, PlaceForm2, "RecType", ErrRes)
        rVol = GetSetting(PlaceProgSettings, PlaceForm2, "Vol", ErrRes)
        rW = GetSetting(PlaceProgSettings, PlaceForm2, "W", ErrRes)
        rOrdVol = GetSetting(PlaceProgSettings, PlaceForm2, "OrdVol", ErrRes)
        rClass = GetSetting(PlaceProgSettings, PlaceForm2, "Class", ErrRes)
        rClassK = GetSetting(PlaceProgSettings, PlaceForm2, "ClassK", ErrRes)
        rClassV = GetSetting(PlaceProgSettings, PlaceForm2, "ClassV", ErrRes)
        rClassH = GetSetting(PlaceProgSettings, PlaceForm2, "ClassH", ErrRes)
        rClassP = GetSetting(PlaceProgSettings, PlaceForm2, "ClassP", ErrRes)
        rCem1 = GetSetting(PlaceProgSettings, PlaceForm2, "Cem1", ErrRes)
        rCem2 = GetSetting(PlaceProgSettings, PlaceForm2, "Cem2", ErrRes)
        rCem3 = GetSetting(PlaceProgSettings, PlaceForm2, "Cem3", ErrRes)
        rChem1 = GetSetting(PlaceProgSettings, PlaceForm2, "Chem1", ErrRes)
        rChem2 = GetSetting(PlaceProgSettings, PlaceForm2, "Chem2", ErrRes)
        rChem3 = GetSetting(PlaceProgSettings, PlaceForm2, "Chem3", ErrRes)
        rEDM = GetSetting(PlaceProgSettings, PlaceForm2, "EDM", ErrRes)
        rMixTime = GetSetting(PlaceProgSettings, PlaceForm2, "MixTime", ErrRes)
        rExpTime = GetSetting(PlaceProgSettings, PlaceForm2, "ExpTime", ErrRes)
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm2, "RealVol", ErrRes)
    Else
        rDist = 1
        rRecType = 1
        rVol = 1
        rW = 1
        rOrdVol = 1
        rClass = 1
        rClassK = 1
        rClassV = 1
        rClassH = 1
        rClassP = 1
        rCem1 = 1
        rCem2 = 1
        rCem3 = 1
        rChem1 = 1
        rChem2 = 1
        rChem3 = 1
        rEDM = 1
        rMixTime = 1
        rExpTime = 1
        rRealVol = 1
    End If
    Me.chDist.Value = rDist
    Me.chRecType.Value = rRecType
    Me.chVol.Value = rVol
    Me.chW.Value = rW
    Me.chOrdVol.Value = rOrdVol
    Me.chClass.Value = rClass
    Me.chClassK.Value = rClassK
    Me.chClassV.Value = rClassV
    Me.chClassH.Value = rClassH
    Me.chClassP.Value = rClassP
    Me.chCem1.Value = rCem1
    Me.chCem2.Value = rCem2
    Me.chCem3.Value = rCem3
    Me.chChem1.Value = rChem1
    Me.chChem2.Value = rChem2
    Me.chChem3.Value = rChem3
    Me.chEDM.Value = rEDM
    Me.chMixTime.Value = rMixTime
    Me.chExpTime.Value = rExpTime
    Me.chRealVol.Value = rRealVol
End Sub

Private Sub btnSave_Click()
    
    On Error Resume Next
    
    strSubKey = Trim(PlaceProgSet2)
    
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        DeleteSetting PlaceProgSettings, PlaceForm2, "Dist"
        DeleteSetting PlaceProgSettings, PlaceForm2, "RecType"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Vol"
        DeleteSetting PlaceProgSettings, PlaceForm2, "W"
        DeleteSetting PlaceProgSettings, PlaceForm2, "OrdVol"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Class"
        DeleteSetting PlaceProgSettings, PlaceForm2, "ClassK"
        DeleteSetting PlaceProgSettings, PlaceForm2, "ClassV"
        DeleteSetting PlaceProgSettings, PlaceForm2, "ClassH"
        DeleteSetting PlaceProgSettings, PlaceForm2, "ClassP"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Cem1"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Cem2"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Cem3"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Chem1"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Chem2"
        DeleteSetting PlaceProgSettings, PlaceForm2, "Chem3"
        DeleteSetting PlaceProgSettings, PlaceForm2, "EDM"
        DeleteSetting PlaceProgSettings, PlaceForm2, "MixTime"
        DeleteSetting PlaceProgSettings, PlaceForm2, "ExpTime"
        DeleteSetting PlaceProgSettings, PlaceForm2, "RealVol"
    End If
    
    SaveSetting PlaceProgSettings, PlaceForm2, "Dist", Me.chDist
    SaveSetting PlaceProgSettings, PlaceForm2, "RecType", Me.chRecType
    SaveSetting PlaceProgSettings, PlaceForm2, "Vol", Me.chVol
    SaveSetting PlaceProgSettings, PlaceForm2, "W", Me.chW
    SaveSetting PlaceProgSettings, PlaceForm2, "OrdVol", Me.chOrdVol
    SaveSetting PlaceProgSettings, PlaceForm2, "Class", Me.chClass
    SaveSetting PlaceProgSettings, PlaceForm2, "ClassK", Me.chClassK
    SaveSetting PlaceProgSettings, PlaceForm2, "ClassV", Me.chClassV
    SaveSetting PlaceProgSettings, PlaceForm2, "ClassH", Me.chClassH
    SaveSetting PlaceProgSettings, PlaceForm2, "ClassP", Me.chClassP
    SaveSetting PlaceProgSettings, PlaceForm2, "Cem1", Me.chCem1
    SaveSetting PlaceProgSettings, PlaceForm2, "Cem2", Me.chCem2
    SaveSetting PlaceProgSettings, PlaceForm2, "Cem3", Me.chCem3
    SaveSetting PlaceProgSettings, PlaceForm2, "Chem1", Me.chChem1
    SaveSetting PlaceProgSettings, PlaceForm2, "Chem2", Me.chChem2
    SaveSetting PlaceProgSettings, PlaceForm2, "Chem3", Me.chChem3
    SaveSetting PlaceProgSettings, PlaceForm2, "EDM", Me.chEDM
    SaveSetting PlaceProgSettings, PlaceForm2, "MixTime", Me.chMixTime
    SaveSetting PlaceProgSettings, PlaceForm2, "ExpTime", Me.chExpTime
    SaveSetting PlaceProgSettings, PlaceForm2, "RealVol", Me.chRealVol
    Unload Me
End Sub

