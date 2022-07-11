VERSION 5.00
Begin VB.Form prntForm1btn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Форма 1"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrint 
      Caption         =   "OK"
      Height          =   375
      Left            =   10080
      TabIndex        =   89
      Top             =   8160
      Width           =   1475
   End
   Begin VB.TextBox txtExpNote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   4680
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   780
      Width           =   2655
   End
   Begin VB.TextBox txtOrd 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9480
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtDate 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtClnt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox txtObj 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4335
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Width           =   4695
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6960
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1695
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtVol 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2280
      Width           =   735
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtRecType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1695
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10320
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtClass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtEDM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3720
      Width           =   495
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtOper 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2655
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3360
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3720
      Width           =   615
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblAtt1 
      BackStyle       =   0  'Transparent
      Caption         =   "Избягвайте допир на кожата с бетонова смес."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   88
      Top             =   7560
      Width           =   11025
   End
   Begin VB.Label lblAtt2 
      BackStyle       =   0  'Transparent
      Caption         =   "и потърсете медицинска помощ."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   87
      Top             =   8040
      Width           =   4905
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " (атмосферни, химични, трептения и вибрации)."
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   86
      Top             =   6840
      Width           =   11100
   End
   Begin VB.Label lblInstr1 
      BackStyle       =   0  'Transparent
      Caption         =   "монолитността на бетона в конструкциите. Сместа трябва да бъде положена не по-късно от 90 минути от часа на разбъркването."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   85
      Top             =   6120
      Width           =   10560
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Диспечер ТИП-Панел v1.2/2014"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8880
      TabIndex        =   83
      Top             =   120
      Width           =   2625
   End
   Begin VB.Label lblInstr1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Полагането и уплътняването на сместа трябва да се извърши ръчно или машинно по технология, осигуряваща еднородността и"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   82
      Top             =   5880
      Width           =   11280
   End
   Begin VB.Label lblExpNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Експедиционна бележка No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   81
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   7440
      TabIndex        =   80
      Top             =   840
      Width           =   105
   End
   Begin VB.Label lblOrd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "по заявка No. / дата"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9360
      TabIndex        =   79
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label lblClnt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Клиент:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   78
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblObj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Обект:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6240
      TabIndex        =   77
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblDrv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Водач:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   76
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label lblDrvNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Кола:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6240
      TabIndex        =   75
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lblMixTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Начало на миксиране:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5520
      TabIndex        =   74
      Top             =   3720
      Width           =   2190
   End
   Begin VB.Label lblExpTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Час на експедиция:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8760
      TabIndex        =   73
      Top             =   3720
      Width           =   1890
   End
   Begin VB.Label lblVol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Обем:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3840
      TabIndex        =   72
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Тегло:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6240
      TabIndex        =   71
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label lblPlaceMix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Местополагане: .............................................................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   70
      Top             =   4080
      Width           =   4425
   End
   Begin VB.Label lblPlaceStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Начало на полагането: ...................................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   69
      Top             =   4440
      Width           =   3945
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Клас по якост:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   68
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label lblClassK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Клас по конс.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3120
      TabIndex        =   67
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblClassV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Клас по възд.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5400
      TabIndex        =   66
      Top             =   2640
      Width           =   1380
   End
   Begin VB.Label lblClassH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Клас по с-ние на хлориди:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7920
      TabIndex        =   65
      Top             =   2640
      Width           =   2505
   End
   Begin VB.Label lblCem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Клас цимент:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   64
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label lblEDM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D max на ЕДМ:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   63
      Top             =   3720
      Width           =   1605
   End
   Begin VB.Label lblChem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Химически добавки:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   62
      Top             =   3360
      Width           =   1980
   End
   Begin VB.Label lblTimeArr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Час на пристигане: ......................................"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6120
      TabIndex        =   61
      Top             =   4080
      Width           =   3630
   End
   Begin VB.Label lblPlaceEnd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Край на полагане: ........................................"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6120
      TabIndex        =   60
      Top             =   4440
      Width           =   3630
   End
   Begin VB.Label lblOper 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Диспечер:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   59
      Top             =   4800
      Width           =   1005
   End
   Begin VB.Label lblDrvSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Водач: ........................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6120
      TabIndex        =   58
      Top             =   4800
      Width           =   1830
   End
   Begin VB.Label lblSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Приел: ....................................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8520
      TabIndex        =   57
      Top             =   4800
      Width           =   2385
   End
   Begin VB.Label lblAddPlace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Място: ................................................"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   56
      Top             =   5400
      Width           =   2850
   End
   Begin VB.Label lblAddType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Вид: ..................................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4200
      TabIndex        =   55
      Top             =   5400
      Width           =   2040
   End
   Begin VB.Label lblAddSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Разпоредител: ........................................"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7680
      TabIndex        =   54
      Top             =   5400
      Width           =   3330
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Допълнително третиране на разтвора (добавяне на вода или друго) на отговорност на Разпоредителя:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   53
      Top             =   5160
      Width           =   10260
   End
   Begin VB.Label lblInstr2 
      BackStyle       =   0  'Transparent
      Caption         =   "Не се допуска добавяне на материали от Разпоредителя."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   52
      Top             =   6360
      Width           =   5100
   End
   Begin VB.Label lblInstr4 
      BackStyle       =   0  'Transparent
      Caption         =   "2. До набиране на необходимата якост трябва да бъдат полагани грижи към пресния бетон за предпазване от едри въздействия"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   51
      Top             =   6600
      Width           =   11100
   End
   Begin VB.Label lblInstr 
      BackStyle       =   0  'Transparent
      Caption         =   "Указание за ползване и грижи при съсъхването на бетонната смес:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   360
      TabIndex        =   50
      Top             =   5640
      Width           =   5970
   End
   Begin VB.Label lblAtt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ВНИМАНИЕ!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   49
      Top             =   7080
      Width           =   1110
   End
   Begin VB.Label lblAtt1 
      BackStyle       =   0  'Transparent
      Caption         =   $"prntForm1btn.frx":0000
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   48
      Top             =   7320
      Width           =   11025
   End
   Begin VB.Label lblAtt2 
      BackStyle       =   0  'Transparent
      Caption         =   "Измийте незабавно бетонната смес попаднала по кожата. При попадане на бетонна смес в очите незабавно изплакнете обилно с вода"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   47
      Top             =   7800
      Width           =   11265
   End
   Begin VB.Label lblDist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Разст.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9480
      TabIndex        =   46
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label lblKM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "km"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   11040
      TabIndex        =   45
      Top             =   1920
      Width           =   270
   End
   Begin VB.Label lblRecType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Вид разтвор:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   44
      Top             =   2280
      Width           =   1470
   End
   Begin VB.Label lblOrdVol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Общо по заявката:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8400
      TabIndex        =   43
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblM3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   5400
      TabIndex        =   42
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label lblKG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7920
      TabIndex        =   41
      Top             =   2280
      Width           =   225
   End
   Begin VB.Label lblM3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   11040
      TabIndex        =   40
      Top             =   2280
      Width           =   270
   End
   Begin VB.Label lblCem3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8880
      TabIndex        =   39
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblChem1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2520
      TabIndex        =   38
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblChem2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5640
      TabIndex        =   37
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblChem3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8880
      TabIndex        =   36
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblMM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2640
      TabIndex        =   35
      Top             =   3720
      Width           =   330
   End
   Begin VB.Label lblOperSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/..................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4320
      TabIndex        =   34
      Top             =   4800
      Width           =   915
   End
   Begin VB.Label lblCem1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2520
      TabIndex        =   33
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblCem2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5640
      TabIndex        =   32
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblClassP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Водоплътност:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3120
      TabIndex        =   31
      Top             =   3720
      Width           =   1470
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCompany"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   600
      TabIndex        =   30
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblTown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   600
      TabIndex        =   29
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblTel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6240
      TabIndex        =   28
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lblFax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblFax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6240
      TabIndex        =   27
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblConcP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblConcP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3720
      TabIndex        =   26
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "prntForm1btn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim PrevSet         As Boolean
    Dim strSubKey       As String
    Dim Comp            As String
    Dim Town            As String
    Dim ConcP           As String
    Dim Tel             As String
    Dim Fax             As String
    Dim intEmpFile      As Integer

    intEmpFile = FreeFile
    
    'зареждане на настройките на форма 1 от регистъра
    strSubKey = Trim(PlaceProgSet1)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    If PrevSet = True Then
        rDist = GetSetting(PlaceProgSettings, PlaceForm1, "Dist", ErrRes)
        rRecType = GetSetting(PlaceProgSettings, PlaceForm1, "RecType", ErrRes)
        rVol = GetSetting(PlaceProgSettings, PlaceForm1, "Vol", ErrRes)
        rW = GetSetting(PlaceProgSettings, PlaceForm1, "W", ErrRes)
        rOrdVol = GetSetting(PlaceProgSettings, PlaceForm1, "OrdVol", ErrRes)
        rClass = GetSetting(PlaceProgSettings, PlaceForm1, "Class", ErrRes)
        rClassK = GetSetting(PlaceProgSettings, PlaceForm1, "ClassK", ErrRes)
        rClassV = GetSetting(PlaceProgSettings, PlaceForm1, "ClassV", ErrRes)
        rClassH = GetSetting(PlaceProgSettings, PlaceForm1, "ClassH", ErrRes)
        rClassP = GetSetting(PlaceProgSettings, PlaceForm1, "ClassP", ErrRes)
        rCem1 = GetSetting(PlaceProgSettings, PlaceForm1, "Cem1", ErrRes)
        rCem2 = GetSetting(PlaceProgSettings, PlaceForm1, "Cem2", ErrRes)
        rCem3 = GetSetting(PlaceProgSettings, PlaceForm1, "Cem3", ErrRes)
        rChem1 = GetSetting(PlaceProgSettings, PlaceForm1, "Chem1", ErrRes)
        rChem2 = GetSetting(PlaceProgSettings, PlaceForm1, "Chem2", ErrRes)
        rChem3 = GetSetting(PlaceProgSettings, PlaceForm1, "Chem3", ErrRes)
        rEDM = GetSetting(PlaceProgSettings, PlaceForm1, "EDM", ErrRes)
        rMixTime = GetSetting(PlaceProgSettings, PlaceForm1, "MixTime", ErrRes)
        rExpTime = GetSetting(PlaceProgSettings, PlaceForm1, "ExpTime", ErrRes)
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm1, "RealVol", ErrRes)
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
    
    If Dir(InfoFile) <> "" Then
        Open InfoFile For Input As intEmpFile
        Input #intEmpFile, Comp, Town, ConcP, Tel, Fax
        Close
        If Len(Comp) Then
            Me.lblCompany.Caption = Comp
        Else
            Me.lblCompany.Caption = ""
        End If
        If Len(Town) Then
            Me.lblTown.Caption = Town
        Else
            Me.lblTown.Caption = ""
        End If
        If Len(ConcP) Then
            Me.lblConcP.Caption = ConcP
        Else
            Me.lblConcP.Caption = ""
        End If
        If Len(Tel) Then
            Me.lblTel.Caption = Tel
        Else
            Me.lblTel.Caption = ""
        End If
        If Len(Fax) Then
            Me.lblFax.Caption = Fax
        Else
            Me.lblFax.Caption = ""
        End If
    Else
        Me.lblCompany.Caption = ""
        Me.lblTown.Caption = ""
        Me.lblConcP.Caption = ""
        Me.lblTel.Caption = ""
        Me.lblFax.Caption = ""
    End If
    If rDist = 0 Then
        Me.lblDist.Visible = False
        Me.txtDist.Enabled = False
        Me.lblKM.Visible = False
    ElseIf rDist = 1 Then
        Me.lblDist.Visible = True
        Me.txtDist.Enabled = True
        Me.lblKM.Visible = True
    End If
    If rRecType = 0 Then
        Me.lblRecType.Visible = False
        Me.txtRecType.Enabled = False
    ElseIf rRecType = 1 Then
        Me.lblRecType.Visible = True
        Me.txtRecType.Enabled = True
    End If
    If rVol = 0 Then
        Me.lblVol.Visible = False
        Me.txtVol.Enabled = False
        Me.lblM3(0).Visible = False
    ElseIf rVol = 1 Then
        Me.lblVol.Visible = True
        Me.txtVol.Enabled = True
        Me.lblM3(0).Visible = True
    End If
    If rW = 0 Then
        Me.lblW.Visible = False
        Me.txtW.Enabled = False
        Me.lblKG.Visible = False
    ElseIf rW = 1 Then
        Me.lblW.Visible = True
        Me.txtW.Enabled = True
        Me.lblKG.Visible = True
    End If
    If rOrdVol = 0 Then
        Me.lblOrdVol.Visible = False
        Me.txtOrdVol.Enabled = False
        Me.lblM3(1).Visible = False
    ElseIf rOrdVol = 1 Then
        Me.lblOrdVol.Visible = True
        Me.txtOrdVol.Enabled = True
        Me.lblM3(1).Visible = True
    End If
    If rClass = 0 Then
        Me.lblClass.Visible = False
        Me.txtClass.Enabled = False
    ElseIf rClass = 1 Then
        Me.lblClass.Visible = True
        Me.txtClass.Enabled = True
    End If
    If rClassK = 0 Then
        Me.lblClassK.Visible = False
        Me.txtClassK.Enabled = False
    ElseIf rClassK = 1 Then
        Me.lblClassK.Visible = True
        Me.txtClassK.Enabled = True
    End If
    If rClassV = 0 Then
        Me.lblClassV.Visible = False
        Me.txtClassV.Enabled = False
    ElseIf rClassV = 1 Then
        Me.lblClassV.Visible = True
        Me.txtClassV.Enabled = True
    End If
    If rClassH = 0 Then
        Me.lblClassH.Visible = False
        Me.txtClassH.Enabled = False
    ElseIf rClassH = 1 Then
        Me.lblClassH.Visible = True
        Me.txtClassH.Enabled = True
    End If
    If rClassP = 0 Then
        Me.lblClassP.Visible = False
        Me.txtClassP.Enabled = False
    ElseIf rClassP = 1 Then
        Me.lblClassP.Visible = True
        Me.txtClassP.Enabled = True
    End If
    If rCem1 = 0 Then
        Me.lblCem1.Visible = False
        Me.txtCem1.Enabled = False
    ElseIf rCem1 = 1 Then
        Me.lblCem1.Visible = True
        Me.txtCem1.Enabled = True
    End If
    If rCem2 = 0 Then
        Me.lblCem2.Visible = False
        Me.txtCem2.Enabled = False
    ElseIf rCem2 = 1 Then
        Me.lblCem2.Visible = True
        Me.txtCem2.Enabled = True
    End If
    If rCem3 = 0 Then
        Me.lblCem3.Visible = False
        Me.txtCem3.Enabled = False
    ElseIf rCem3 = 1 Then
        Me.lblCem3.Visible = True
        Me.txtCem3.Enabled = True
    End If
    If rCem1 = 0 And rCem2 = 0 And rCem3 = 0 Then
        lblCem.Visible = False
    Else
        lblCem.Visible = True
    End If
    If rChem1 = 0 Then
        Me.lblChem1.Visible = False
        Me.txtChem1.Enabled = False
    ElseIf rChem1 = 1 Then
        Me.lblChem1.Visible = True
        Me.txtChem1.Enabled = True
    End If
    If rChem2 = 0 Then
        Me.lblChem2.Visible = False
        Me.txtChem2.Enabled = False
    ElseIf rChem2 = 1 Then
        Me.lblChem2.Visible = True
        Me.txtChem2.Enabled = True
    End If
    If rChem3 = 0 Then
        Me.lblChem3.Visible = False
        Me.txtChem3.Enabled = False
    ElseIf rChem3 = 1 Then
        Me.lblChem3.Visible = True
        Me.txtChem3.Enabled = True
    End If
    If rChem1 = 0 And rChem2 = 0 And rChem3 = 0 Then
        lblChem.Visible = False
    Else
        lblChem.Visible = True
    End If
    If rEDM = 0 Then
        Me.lblEDM.Visible = False
        Me.txtEDM.Enabled = False
        Me.lblMM.Visible = False
    ElseIf rEDM = 1 Then
        Me.lblEDM.Visible = True
        Me.txtEDM.Enabled = True
        Me.lblMM.Visible = True
    End If
    If rMixTime = 0 Then
        Me.lblMixTime.Visible = False
        Me.txtMixTime.Enabled = False
    ElseIf rMixTime = 1 Then
        Me.lblMixTime.Visible = True
        Me.txtMixTime.Enabled = True
    End If
    If rExpTime = 0 Then
        Me.lblExpTime.Visible = False
        Me.txtExpTime.Enabled = False
    ElseIf rExpTime = 1 Then
        Me.lblExpTime.Visible = True
        Me.txtExpTime.Enabled = True
    End If
End Sub

Private Sub btnPrint_Click()

    Call PrintBtnForm1(prntForm1btn)
End Sub

