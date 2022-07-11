VERSION 5.00
Begin VB.Form frmOPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmOPCVIPA"
   ClientHeight    =   10335
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Ready 
      Height          =   285
      Index           =   2
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox Ready 
      Height          =   285
      Index           =   1
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Stat 
      Height          =   285
      Index           =   4
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   2520
      Width           =   615
   End
   Begin VB.Frame frmOPC 
      Caption         =   "OPC Items"
      Height          =   9975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   38
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox Ready 
         Height          =   285
         Index           =   3
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   8280
         Width           =   615
      End
      Begin VB.TextBox Config 
         Height          =   285
         Index           =   4
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   16
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   9360
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   15
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   9000
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   34
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   9360
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   18
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   9360
         Width           =   1335
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Close"
         Height          =   375
         Left            =   6360
         TabIndex        =   73
         Top             =   9360
         Width           =   855
      End
      Begin VB.TextBox Ready 
         Height          =   285
         Index           =   0
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   14
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   8640
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   13
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   8280
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   12
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   7800
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   11
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   10
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   6960
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   9
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   6600
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   8
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   6240
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   7
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   6
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   5
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   4
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   3
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   2
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   1
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Result 
         Height          =   285
         Index           =   0
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   37
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   36
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   35
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   20
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   33
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   9000
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   32
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   8640
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   31
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   8280
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   30
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   29
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   28
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   6960
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   27
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   6600
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   26
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   25
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   24
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   23
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   22
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   21
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   19
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   17
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   9000
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   16
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   8640
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   15
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   8280
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   14
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   7800
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   13
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   12
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   6960
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   11
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   6600
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   10
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   6240
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   9
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   8
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   7
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   5
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   2
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox RecInput 
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox Stat 
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Stat 
         Height          =   285
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Stat 
         Height          =   285
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Stat 
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Config 
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Config 
         Height          =   285
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Config 
         Height          =   285
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Config 
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Panel 
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Cons 
         Height          =   285
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Cons 
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Panel 
         Height          =   285
         Index           =   3
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Panel 
         Height          =   285
         Index           =   2
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Panel 
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Results"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   56
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "RecInput"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Status"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Config"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Connection Tests"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Weghts"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmOPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    Me.Hide
End Sub

