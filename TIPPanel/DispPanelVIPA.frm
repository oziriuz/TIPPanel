VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form DispPanel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDispPanelCap"
   ClientHeight    =   68580
   ClientLeft      =   3240
   ClientTop       =   -3990
   ClientWidth     =   16215
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   68580
   ScaleWidth      =   16215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnMixCap 
      Caption         =   "btnMixCap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   273
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton btnChSilos 
      Caption         =   "btnChSilos"
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
      Left            =   12720
      TabIndex        =   263
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton btnChMach 
      Caption         =   "btnChMach"
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
      Left            =   14400
      TabIndex        =   262
      Top             =   480
      Width           =   1695
   End
   Begin VB.Timer AVTimer 
      Interval        =   77
      Left            =   13800
      Top             =   720
   End
   Begin VB.Timer FormT 
      Enabled         =   0   'False
      Interval        =   77
      Left            =   13440
      Top             =   720
   End
   Begin VB.Timer TimerStartReq 
      Enabled         =   0   'False
      Interval        =   777
      Left            =   13800
      Top             =   0
   End
   Begin VB.Timer TimerRes 
      Enabled         =   0   'False
      Interval        =   777
      Left            =   13800
      Top             =   360
   End
   Begin VB.PictureBox imgLogo 
      Height          =   495
      Left            =   120
      Picture         =   "DispPanelVIPA.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   245
      Top             =   1320
      Width           =   530
   End
   Begin VB.CheckBox chPrintConf 
      Caption         =   "chPrintConf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   11160
      TabIndex        =   235
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Frame frAbout 
      Caption         =   "frAbout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   204
      Top             =   60240
      Width           =   14175
      Begin RichTextLib.RichTextBox rtxtLicAgr 
         Height          =   6015
         Left            =   360
         TabIndex        =   243
         Top             =   600
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   10610
         _Version        =   393217
         BackColor       =   -2147483633
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         FileName        =   "D:\Dispatcher\license agreement.rtf"
         TextRTF         =   $"DispPanelVIPA.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frStatus 
      Caption         =   "frStatus"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   10
      Top             =   1080
      Width           =   9975
      Begin VB.Label numSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "numSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   271
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label numSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "numSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   270
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label numSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "numSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   269
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label numSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "numSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   268
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   267
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   266
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   265
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSilos 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSilos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   264
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label stExp 
         BackStyle       =   0  'Transparent
         Caption         =   "stExp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7560
         TabIndex        =   260
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label stClnt 
         BackStyle       =   0  'Transparent
         Caption         =   "stClnt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7560
         TabIndex        =   259
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label stOrd 
         BackStyle       =   0  'Transparent
         Caption         =   "stOrd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7560
         TabIndex        =   258
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label indValveMix 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "indValveMix"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label indReq 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "indReq"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label indAvaria 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "indAvaria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label indMode 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "indMode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame frMaterials 
      Caption         =   "frMaterials"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   184
      Top             =   53280
      Width           =   14175
      Begin VB.CommandButton btnRevision 
         Caption         =   "btnRevision"
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
         Left            =   3360
         TabIndex        =   242
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton btnSvExp 
         Caption         =   "btnSvExp"
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
         Left            =   12000
         TabIndex        =   192
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox s1 
         Caption         =   "s1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   12480
         TabIndex        =   202
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox s1 
         Caption         =   "s1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   11160
         TabIndex        =   201
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox s1 
         Caption         =   "s1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   9840
         TabIndex        =   200
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox s1 
         Caption         =   "s1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8520
         TabIndex        =   199
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox s1 
         Caption         =   "s1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   198
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox s1 
         Caption         =   "s1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   197
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton btnAddMatDlvr 
         Caption         =   "btnAddMatDlvr"
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
         Left            =   12000
         TabIndex        =   188
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton btnClearMat 
         Caption         =   "btnClearMat"
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
         Left            =   240
         TabIndex        =   194
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton btnSvNwMat 
         Caption         =   "btnSvNwMat"
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
         Left            =   2040
         TabIndex        =   195
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton btnDelMat 
         Caption         =   "btnDelMat"
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
         TabIndex        =   196
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cmbMatType 
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
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   187
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtMatName 
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
         Left            =   1800
         TabIndex        =   190
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtMat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin MSComctlLib.ListView lstMat 
         Height          =   4455
         Left            =   240
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   2280
         Width           =   13695
         _ExtentX        =   24156
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
      Begin VB.Label lblMatLoad 
         BackStyle       =   0  'Transparent
         Caption         =   "lblMatLoad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   193
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label lblMatName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblMatName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   191
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblMatType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblMatType"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   189
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblMatIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblMatIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   186
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame frDisp 
      Caption         =   "frDisp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   25
      Top             =   11520
      Width           =   14175
      Begin VB.TextBox txtDispClntObj 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   255
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtDispOrdDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   14640
         Locked          =   -1  'True
         TabIndex        =   244
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDispDrvCap 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtDispOrdQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox cmbDispDrvName 
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
         Left            =   8040
         TabIndex        =   34
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDispClnt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton btnDispStart 
         Caption         =   "btnDispStart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11640
         TabIndex        =   43
         Top             =   1440
         Width           =   2175
      End
      Begin ComCtl2.UpDown updownDispWat 
         Height          =   375
         Left            =   13320
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "txtDispWat"
         BuddyDispid     =   196652
         OrigLeft        =   6720
         OrigTop         =   1800
         OrigRight       =   6975
         OrigBottom      =   2175
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtDispWat 
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
         Left            =   12720
         TabIndex        =   46
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtDispQuant 
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
         Left            =   12720
         TabIndex        =   45
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtDispClntName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtDispDrvReg 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtDispRecClass 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtDispRecName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtDispRec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cmbDispDrv 
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
         Left            =   8640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbDispOrd 
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
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstOrdWait 
         Height          =   2175
         Left            =   600
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2280
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
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
      Begin MSComctlLib.ListView lstMixReady 
         Height          =   2295
         Left            =   600
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   4440
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
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
      Begin VB.Label lblM3 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   5
         Left            =   11040
         TabIndex        =   238
         Top             =   1920
         Width           =   270
      End
      Begin VB.Label lblDispDrvCap 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispDrvCap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   236
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblDispOrdQuant 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispOrdQuant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   232
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblM3 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   3
         Left            =   6960
         TabIndex        =   231
         Top             =   1920
         Width           =   270
      End
      Begin VB.Label lblKg 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   4
         Left            =   13680
         TabIndex        =   229
         Top             =   480
         Width           =   225
      End
      Begin VB.Label lblM3 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   2
         Left            =   13440
         TabIndex        =   227
         Top             =   960
         Width           =   270
      End
      Begin VB.Label lblReady2 
         Alignment       =   2  'Center
         Caption         =   "lblReady2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   360
         TabIndex        =   54
         Top             =   4440
         Width           =   135
      End
      Begin VB.Label lblReady 
         Alignment       =   2  'Center
         Caption         =   "lblReady"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   53
         Top             =   4440
         Width           =   135
      End
      Begin VB.Label lblWait 
         Alignment       =   2  'Center
         Caption         =   "lblWait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   51
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label lblDispClnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispClnt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDispDrvReg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispDrvReg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   42
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblDispDrvName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispDrvName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblDispWat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispWat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   50
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblDispQuant 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispQuant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   49
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblDispClntObj 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispClntObj"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblDispClntName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispClntName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblDispDrv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispDrv"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblDispRecClass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispRecClass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblDispRecName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispRecName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblDispRec 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispRec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblDispOrd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDispOrd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame frOrders 
      Caption         =   "frOrders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   56
      Top             =   18480
      Width           =   14175
      Begin VB.ComboBox cmbOrdClntName 
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
         Left            =   6360
         TabIndex        =   257
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox cmbOrdClntObj 
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
         Left            =   6360
         TabIndex        =   254
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton btnDelOrd 
         Caption         =   "btnDelOrd"
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
         Left            =   12600
         TabIndex        =   77
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton btnSvNwOrd 
         Caption         =   "btnSvNwOrd"
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
         Left            =   11160
         TabIndex        =   76
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton btnClearOrd 
         Caption         =   "btnClearOrd"
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
         Left            =   9720
         TabIndex        =   75
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtOrdQuant 
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
         Left            =   12960
         TabIndex        =   60
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtOrd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbOrdClnt 
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
         Left            =   6360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtOrdRecClass 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtOrdRecName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2655
      End
      Begin VB.ComboBox cmbOrdRec 
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
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   480
         Width           =   975
      End
      Begin MSComCtl2.DTPicker nowOrdDate 
         Height          =   375
         Left            =   14280
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   -2147483639
         CustomFormat    =   "dd.MM.yyy"
         Format          =   113639427
         CurrentDate     =   41426.3333333333
         MaxDate         =   44196
         MinDate         =   41426
      End
      Begin MSComctlLib.ListView lstOrd 
         Height          =   4455
         Left            =   240
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2280
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
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
      Begin MSComCtl2.DTPicker queOrdDate 
         Height          =   375
         Left            =   10080
         TabIndex        =   70
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   -2147483639
         CustomFormat    =   "dd.MM.yyy"
         Format          =   113639427
         CurrentDate     =   41487.3333333333
         MaxDate         =   45291
         MinDate         =   41487
      End
      Begin MSComCtl2.DTPicker queOrdTime 
         Height          =   375
         Left            =   12120
         TabIndex        =   71
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   -2147483639
         CustomFormat    =   "dd.MM.yyy"
         Format          =   113639426
         CurrentDate     =   41426.3333333333
         MaxDate         =   44196
         MinDate         =   41426
      End
      Begin VB.Label lblM3 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   1
         Left            =   13680
         TabIndex        =   228
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblOrdQuant 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdQuant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11400
         TabIndex        =   65
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblOrd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   64
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblOrdClntObj 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdClntObj"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   74
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblOrdClntName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdClntName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   69
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblOrdClntIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdClntIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   63
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblOrdRecClass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdRecClass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblOrdRecName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdRecName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblOrdDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdDate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   66
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblOrdRecIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOrdRecIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame frSuppliers 
      Caption         =   "frSuppliers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   165
      Top             =   46320
      Width           =   14175
      Begin VB.CommandButton btnShowSup 
         Caption         =   "btnShowSup"
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
         Left            =   12240
         TabIndex        =   241
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtNoteSup 
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
         Left            =   7680
         TabIndex        =   179
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox txtTelSup 
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
         Left            =   7680
         TabIndex        =   175
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtMOLSup 
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
         Left            =   1560
         TabIndex        =   178
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox txtBGSup 
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
         Left            =   1560
         TabIndex        =   174
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton btnDelSup 
         Caption         =   "btnDelSup"
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
         Left            =   12240
         TabIndex        =   182
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btnSvNwSup 
         Caption         =   "btnSvNwSup"
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
         Left            =   12240
         TabIndex        =   173
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton btnClearSup 
         Caption         =   "btnClearSup"
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
         Left            =   12240
         TabIndex        =   168
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtSup 
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
         Left            =   1560
         TabIndex        =   166
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtNameSup 
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
         Left            =   1560
         TabIndex        =   169
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtAddSup 
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
         Left            =   7680
         TabIndex        =   170
         Top             =   840
         Width           =   4095
      End
      Begin MSComctlLib.ListView lstSup 
         Height          =   4575
         Left            =   240
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   2160
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   8070
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
      Begin VB.Label lblSupNote 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSupNote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   181
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblSupIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSupIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   167
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblSupName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSupName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   171
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSupBG 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSupBG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   176
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblSupMOL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSupMOL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   180
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblSupAdd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSupAdd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   172
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSupTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSupTel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   177
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame frDrivers 
      Caption         =   "frDrivers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   146
      Top             =   39360
      Width           =   14175
      Begin VB.CommandButton btnShowDrv 
         Caption         =   "btnShowDrv"
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
         Left            =   12240
         TabIndex        =   240
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtTelDrv 
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
         Left            =   9600
         MaxLength       =   15
         TabIndex        =   156
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtCapDrv 
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
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   155
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtNoteDrv 
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
         Left            =   6600
         MaxLength       =   15
         TabIndex        =   161
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox txtNameDrv 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   154
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtRegDrv 
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
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   150
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtModDrv 
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
         Left            =   9600
         MaxLength       =   15
         TabIndex        =   151
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton btnSvNwDrv 
         Caption         =   "btnSvNwDrv"
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
         Left            =   12240
         TabIndex        =   160
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton btnClearDrv 
         Caption         =   "btnClearDrv"
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
         Left            =   12240
         TabIndex        =   147
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnDelDrv 
         Caption         =   "btnDelDrv"
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
         Left            =   12240
         TabIndex        =   163
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtDrv 
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
         Left            =   1560
         TabIndex        =   148
         Top             =   480
         Width           =   735
      End
      Begin MSComctlLib.ListView lstDrv 
         Height          =   4575
         Left            =   240
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   2160
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   8070
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
      Begin VB.Label lblM3 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   0
         Left            =   7440
         TabIndex        =   226
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label lblDrvNote 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDrvNote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   162
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblNameDrv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblNameDrv"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblDrvReg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDrvReg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   152
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDrvCap 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDrvCap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   158
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblDrvMod 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDrvMod"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   153
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblDrvTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDrvTel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   159
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblDrvIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblDrvIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   149
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame frClients 
      Caption         =   "frClients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   129
      Top             =   32400
      Width           =   14175
      Begin VB.CommandButton btnShowObj 
         Caption         =   "btnShowObj"
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
         Left            =   12480
         TabIndex        =   256
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton btnDelObj 
         Caption         =   "btnDelObj"
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
         Left            =   11160
         TabIndex        =   253
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton btnObjects 
         Caption         =   "btnObjects"
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
         Left            =   9840
         TabIndex        =   252
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txtTelClnt 
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
         Left            =   1920
         TabIndex        =   136
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtAddClnt 
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
         Left            =   1920
         TabIndex        =   133
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton btnShowClnt 
         Caption         =   "btnShowClnt"
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
         Left            =   8040
         TabIndex        =   239
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtMOLClnt 
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
         Left            =   1920
         TabIndex        =   143
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txtBGClnt 
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
         Left            =   1920
         TabIndex        =   140
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtClnt 
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
         Left            =   1920
         TabIndex        =   130
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtNameClnt 
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
         Left            =   1920
         TabIndex        =   135
         Top             =   840
         Width           =   3855
      End
      Begin VB.CommandButton btnDelClnt 
         Caption         =   "btnDelClnt"
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
         Left            =   8040
         TabIndex        =   142
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton btnClearClnt 
         Caption         =   "btnClearClnt"
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
         Left            =   8040
         TabIndex        =   132
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton btnSvNwClnt 
         Caption         =   "btnSvNwClnt"
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
         Left            =   8040
         TabIndex        =   139
         Top             =   960
         Width           =   1455
      End
      Begin MSComctlLib.ListView lstClnt 
         Height          =   3975
         Left            =   240
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   2760
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7011
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
      Begin MSComctlLib.ListView lstObj 
         Height          =   5655
         Left            =   9840
         TabIndex        =   251
         TabStop         =   0   'False
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9975
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
      Begin VB.Label lblClntIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblClntIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   131
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblClntName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblClntName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   137
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblClntBG 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblClntBG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   141
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblClntMOL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblClntMOL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   144
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblClntAdd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblClntAdd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   134
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblClntTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblClntTel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   138
         Top             =   2400
         Width           =   1455
      End
   End
   Begin VB.CommandButton btnMaterials 
      Caption         =   "btnMaterials"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   20
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Frame frRecepies 
      Caption         =   "frRecepies"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   120
      TabIndex        =   79
      Top             =   25440
      Visible         =   0   'False
      Width           =   14175
      Begin VB.ComboBox cmbRec2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   250
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtRec2 
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
         Height          =   285
         Index           =   1
         Left            =   9720
         TabIndex        =   249
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton btnShowRec 
         Caption         =   "btnShowRec"
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
         Left            =   12480
         TabIndex        =   248
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtRec1 
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
         Height          =   285
         Index           =   5
         Left            =   6360
         TabIndex        =   247
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbRec1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   246
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtEDMRec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   108
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtClassRecP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   218
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtClassRecH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   206
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtClassRecV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   207
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtTotalKg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12600
         Locked          =   -1  'True
         TabIndex        =   234
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtTimePourRec 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   214
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtTimeMixRec 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   213
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtTypeRec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   212
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtClassRecK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   208
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtRec2 
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
         Height          =   285
         Index           =   0
         Left            =   9720
         TabIndex        =   123
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbRec2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtRec4 
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
         Height          =   285
         Index           =   0
         Left            =   13080
         TabIndex        =   90
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbRec4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtRec4 
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
         Height          =   285
         Index           =   1
         Left            =   13080
         TabIndex        =   98
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtRec4 
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
         Height          =   285
         Index           =   2
         Left            =   13080
         TabIndex        =   106
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtRec4 
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
         Height          =   285
         Index           =   3
         Left            =   13080
         TabIndex        =   114
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtRec4 
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
         Height          =   285
         Index           =   4
         Left            =   13080
         TabIndex        =   118
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtRec4 
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
         Height          =   285
         Index           =   5
         Left            =   13080
         TabIndex        =   121
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtRec3 
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
         Height          =   285
         Index           =   0
         Left            =   9720
         TabIndex        =   88
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbRec3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtRec3 
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
         Height          =   285
         Index           =   1
         Left            =   9720
         TabIndex        =   96
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtRec3 
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
         Height          =   285
         Index           =   2
         Left            =   9720
         TabIndex        =   104
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtRec3 
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
         Height          =   285
         Index           =   3
         Left            =   9720
         TabIndex        =   112
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtRec1 
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
         Height          =   315
         Index           =   0
         Left            =   6360
         TabIndex        =   86
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbRec1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cmbRec1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtRec1 
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
         Height          =   285
         Index           =   1
         Left            =   6360
         TabIndex        =   94
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtRec1 
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
         Height          =   285
         Index           =   2
         Left            =   6360
         TabIndex        =   102
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtRec1 
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
         Height          =   285
         Index           =   3
         Left            =   6360
         TabIndex        =   110
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtRec1 
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
         Height          =   285
         Index           =   4
         Left            =   6360
         TabIndex        =   116
         Top             =   2160
         Width           =   735
      End
      Begin VB.Frame frChem 
         Caption         =   "frChem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   82
         Top             =   360
         Width           =   2535
      End
      Begin VB.Frame frWat 
         Caption         =   "frWat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   119
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Frame frScr 
         Caption         =   "frScr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   81
         Top             =   360
         Width           =   2535
      End
      Begin VB.Frame frIM 
         Caption         =   "frIM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   80
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtRec 
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
         Height          =   285
         Left            =   1200
         TabIndex        =   84
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtNameRec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   92
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtClassRec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   100
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton btnDelRec 
         Caption         =   "btnDelRec"
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
         Left            =   11160
         TabIndex        =   126
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton btnClearRec 
         Caption         =   "btnClearRec"
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
         Left            =   8520
         TabIndex        =   124
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton btnSvNwRec 
         Caption         =   "btnSvNwRec"
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
         Left            =   9840
         TabIndex        =   125
         Top             =   3360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstRec 
         Height          =   2895
         Left            =   240
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   3840
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   5106
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
      Begin VB.Label lblTotalKg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblTotalKg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   233
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3480
         TabIndex        =   225
         Top             =   3480
         Width           =   105
      End
      Begin VB.Label lblS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3480
         TabIndex        =   224
         Top             =   3120
         Width           =   105
      End
      Begin VB.Label lblKg 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   3
         Left            =   9840
         TabIndex        =   223
         Top             =   2160
         Width           =   225
      End
      Begin VB.Label lblKg 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   2
         Left            =   13200
         TabIndex        =   222
         Top             =   360
         Width           =   225
      End
      Begin VB.Label lblKg 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   1
         Left            =   9840
         TabIndex        =   221
         Top             =   360
         Width           =   225
      End
      Begin VB.Label lblKg 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Index           =   0
         Left            =   6480
         TabIndex        =   220
         Top             =   360
         Width           =   225
      End
      Begin VB.Label lblRecClassP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecClassP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   219
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblRecType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecType"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   217
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblTimePour 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblTimePour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   216
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblTimeMix 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblTimeMix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   215
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblRecClassH 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecClassH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   211
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblRecClassV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecClassV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   210
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblRecClassK 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecClassK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   209
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblRecNote 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecNote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   3840
         TabIndex        =   127
         Top             =   2880
         Width           =   3375
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRecIndex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecIndex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRecName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblRecClass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecClass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   99
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblRecEDM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecEDM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   107
         Top             =   2760
         Width           =   1935
      End
   End
   Begin VB.CommandButton btnAdminPanel 
      Caption         =   "btnAdminPanel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   22
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox kgChm 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "CHEM"
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox kgWt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "WAT"
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox kgCem 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "CEM"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer ScalesT 
      Left            =   13440
      Top             =   0
   End
   Begin VB.TextBox Clock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "Clock"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Timer ClockT 
      Left            =   13440
      Top             =   360
   End
   Begin VB.CommandButton btnNotes 
      Caption         =   "btnNotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   21
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton btnSuppliers 
      Caption         =   "btnSuppliers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   19
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton btnRecepies 
      Caption         =   "btnRecepies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   16
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton btnClients 
      Caption         =   "btnClients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   17
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton btnDrivers 
      Caption         =   "btnDrivers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   18
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton btnOrders 
      Caption         =   "btnOrders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton btnDisp 
      Caption         =   "btnDisp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "btnExit"
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
      Left            =   14400
      TabIndex        =   24
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox kgAggr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "AGGR"
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   205
      Top             =   68205
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   661
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAddver 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblAddver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   272
      Top             =   0
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMach 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblMach"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   14400
      TabIndex        =   261
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblLoading"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   14400
      TabIndex        =   4
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Label lblPanel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblChem"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   10680
      TabIndex        =   8
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblPanel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblWat"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   7680
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblPanel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblCem"
      Enabled         =   0   'False
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
      Index           =   2
      Left            =   4680
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblPanel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblAggr"
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.Menu rcMnuPrnt 
      Caption         =   "rcMnuPrnt"
      Visible         =   0   'False
      Begin VB.Menu rcForm1 
         Caption         =   "Форма 1"
      End
      Begin VB.Menu rcForm2 
         Caption         =   "Форма 2"
      End
      Begin VB.Menu rcForm3 
         Caption         =   "Форма 3"
      End
   End
   Begin VB.Menu rcMnuPrntOrdNow 
      Caption         =   "rcMnuPrntOrdNow"
      Visible         =   0   'False
      Begin VB.Menu rcFormOrdNow 
         Caption         =   "Принтирай текущи"
      End
      Begin VB.Menu rcFormOrdNowExp 
         Caption         =   "Експорт в MSExcel"
      End
   End
   Begin VB.Menu rcMnuPrntOrd 
      Caption         =   "rcMnuPrntOrd"
      Visible         =   0   'False
      Begin VB.Menu rcFormOrd 
         Caption         =   "Принтирай всички"
      End
      Begin VB.Menu rcFormOrdExp 
         Caption         =   "Експорт в MSExcel"
      End
   End
   Begin VB.Menu rcMnuPrntRec 
      Caption         =   "rcMnuPrntRec"
      Visible         =   0   'False
      Begin VB.Menu PrintRec 
         Caption         =   "Принтирай"
      End
      Begin VB.Menu ExpRec 
         Caption         =   "Експорт в MSExcel"
      End
   End
   Begin VB.Menu rcMnuPrntClnt 
      Caption         =   "rcMnuPrntClnt"
      Visible         =   0   'False
      Begin VB.Menu PrintClnt 
         Caption         =   "Принтирай"
      End
      Begin VB.Menu ExpClnt 
         Caption         =   "Експорт в MSExcel"
      End
   End
   Begin VB.Menu rcMnuPrntDrv 
      Caption         =   "rcMnuPrntDrv"
      Visible         =   0   'False
      Begin VB.Menu PrintDrv 
         Caption         =   "Принтирай"
      End
      Begin VB.Menu ExpDrv 
         Caption         =   "Експорт в MSExcel"
      End
   End
   Begin VB.Menu rcMnuPrntSup 
      Caption         =   "rcMnuPrntSup"
      Visible         =   0   'False
      Begin VB.Menu PrintSup 
         Caption         =   "Принтирай"
      End
      Begin VB.Menu ExpSup 
         Caption         =   "Експорт в MSExcel"
      End
   End
   Begin VB.Menu rcMnuPrntMat 
      Caption         =   "rcMnuPrntMat"
      Visible         =   0   'False
      Begin VB.Menu PrintMat 
         Caption         =   "Принтирай"
      End
      Begin VB.Menu ExpMat 
         Caption         =   "Експорт в MSExcel"
      End
   End
End
Attribute VB_Name = "DispPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Base 1

Public WithEvents ConOPCServer As OPCServer
Attribute ConOPCServer.VB_VarHelpID = -1

Dim ConSrvGroup As OPCGroups

Public WithEvents ConGroupPanels As OPCGroup
Attribute ConGroupPanels.VB_VarHelpID = -1
Dim OPCItemCollPanels As OPCItems
Dim ItemCountPanels As Long
Dim ItemSrvHandlesPanels() As Long
Dim OPCItemIDPanels(4) As String
Dim ClientHandlesPanels(4) As Long
Dim ItemSrvErrPanels() As Long

Public WithEvents ConGroupConfig As OPCGroup
Attribute ConGroupConfig.VB_VarHelpID = -1
Dim OPCItemCollConfig As OPCItems
Public ItemCountConfig As Long
Dim ItemSrvHandlesConfig() As Long
Dim OPCItemIDConfig(1) As String
Dim ClientHandlesConfig(1) As Long
Dim ItemSrvErrConfig() As Long

Public WithEvents ConGroupStatus As OPCGroup
Attribute ConGroupStatus.VB_VarHelpID = -1
Dim OPCItemCollStatus As OPCItems
Dim ItemCountStatus As Long
Dim ItemSrvHandlesStatus() As Long
Dim OPCItemIDStatus(5) As String
Dim ClientHandlesStatus(5) As Long
Dim ItemSrvErrStatus() As Long

Public WithEvents ConGroupRec As OPCGroup
Attribute ConGroupRec.VB_VarHelpID = -1
Dim OPCItemCollRec As OPCItems
Public ItemCountRec As Long
Dim ItemSrvHandlesRec() As Long
Dim OPCItemIDRec(39) As String
Dim ClientHandlesRec(39) As Long
Dim ItemSrvErrRec() As Long

Public WithEvents ConGroupResults As OPCGroup
Attribute ConGroupResults.VB_VarHelpID = -1
Dim OPCItemCollResults As OPCItems
Dim ItemCountResults As Long
Dim ItemSrvHandlesResults() As Long
Dim OPCItemIDResults(17) As String
Dim ClientHandlesResults(17) As Long
Dim ItemSrvErrResults() As Long

Public WithEvents ConGroupReady As OPCGroup
Attribute ConGroupReady.VB_VarHelpID = -1
Dim OPCItemCollReady As OPCItems
Public ItemCountReady As Long
Dim ItemSrvHandlesReady() As Long
Dim OPCItemIDReady(4) As String
Dim ClientHandlesReady(4) As Long
Dim ItemSrvErrReady() As Long

Dim PointLook1 As Boolean
Dim PointLook2 As Boolean
Dim PointLook3 As Boolean
Dim PointLook4 As Boolean
Dim ShiftTest As Integer

Private Sub btnChMach_Click()
    Dim hw As Long
    Dim retval As Long
    Dim SwWin As String
    If MachineNumber = 1 Then SwWin = "Машина 2 - ТИП-Панел v1.2"
    If MachineNumber = 2 Then SwWin = "Машина 1 - ТИП-Панел v1.2"
    hw = FindWindow(vbNullString, SwWin)
'    retval = ShowWindow(hw, 0)
    If hw <> 0 Then retval = ShowWindow(hw, 9)
    
    If hw <> 0 And Me.WindowState <> vbMinimized Then
        Me.WindowState = vbMinimized
    Else
        Me.WindowState = vbNormal
    End If
    
    If MachineNumber = 1 Then
        If hw = 0 And Dir(App.Path & "\TIP2Panelv12.exe") <> "" Then
            Shell App.Path & "\TIP2Panelv12.exe"
        End If
    ElseIf MachineNumber = 2 Then
        If hw = 0 And Dir(App.Path & "\TIP1Panelv12.exe") <> "" Then
            Shell App.Path & "\TIP1Panelv12.exe"
        End If
    End If
End Sub

Private Sub btnChSilos_Click()
    frmChSilos.Show
End Sub

Private Sub btnMixCap_Click()
    frmMixCap.Show
End Sub

Private Sub Form_Load()

    MousePointer = vbHourglass
    
    Dim ConNodeName As String

    dispatcher = True
    
    MousePointer = vbHourglass
    
    ExpeditionStarted = False
    
    EmptyData = False
    
    Dim NotMachineNumber As Integer
    
    Dim oper           As Panel

    Dim today          As Panel

    Dim mixes          As Panel

    Dim exps           As Panel

    Dim voexps         As Panel

    Dim vexps          As Panel

    Dim kgexps         As Panel

    Dim OrderedQuantl  As Single

    Dim RealQuantl     As Single

    Dim TotalKGsl      As Single

    Dim Index          As Integer

    Dim intEmpFileNbr1 As Integer

    Dim lblW()         As String

    Dim lblR()         As String

    Dim lblR2()        As String

    Dim X              As Integer

    Dim i              As Integer
    
    FlagButRec = -1 'флаг за бутон от рамка рецепти дали е "запис" или "изтрий"
    
    frDisp.Top = 21880
    frOrders.Top = 21880
    frRecepies.Top = 21880
    frClients.Top = 21880
    frDrivers.Top = 21880
    frSuppliers.Top = 21880
    frMaterials.Top = 21880
    frAbout.Top = 21880

    DispPanel.Height = 10000
    
    DecSep = GetDecimalSep()
    
    Me.lblAddver.Caption = "  ТИП-Сервиз   ЕООД    гр. Червен бряг    Софтуер, проектиране, изграждане и сервиз   на бетонови стопанства"
    
    Me.btnDisp.Enabled = False
    Me.btnOrders.Enabled = False
    Me.btnRecepies.Enabled = False
    Me.btnClients.Enabled = False
    Me.btnDrivers.Enabled = False
    Me.btnSuppliers.Enabled = False
    Me.btnMaterials.Enabled = False
    Me.btnNotes.Enabled = False
    Me.btnAdminPanel.Enabled = False
    Me.BtnExit.Enabled = False
    Me.chPrintConf.Enabled = False
    
    DispPanel.Caption = frmDispPanelCap
    lblPanel(1).Caption = uniIM
    lblPanel(2).Caption = uniCement
    lblPanel(3).Caption = uniWat
    lblPanel(4).Caption = uniChem
    chPrintConf.Caption = uniAutoPrint
    btnDisp.Caption = uniDisp
    btnOrders.Caption = uniOrds
    btnRecepies.Caption = uniRecs
    btnClients.Caption = uniClnts
    btnDrivers.Caption = uniDrvs
    btnSuppliers.Caption = uniSups
    btnMaterials.Caption = uniMats
    btnNotes.Caption = uniNotes
    btnAdminPanel.Caption = uniSettings
    BtnExit.Caption = UniExit
    frStatus.Caption = uniStatus
    frDisp.Caption = uniDisp
    frOrders.Caption = uniOrds
    frRecepies.Caption = uniRecs
    frClients.Caption = uniClnts
    frDrivers.Caption = uniDrvs
    frSuppliers.Caption = uniSups
    frMaterials.Caption = uniMats
    frAbout.Caption = uniAbout
    btnDispStart.Caption = uniSTART
    lblDispOrd.Caption = uniOrdCode
    lblDispRec.Caption = uniRecCode
    lblDispRecName.Caption = uniNm
    lblDispRecClass.Caption = uniClass
    lblDispClnt.Caption = uniClntCode
    lblDispClntName.Caption = uniFirm
    lblDispClntObj.Caption = uniObj
    lblDispQuant.Caption = uniQuant
    lblDispOrdQuant.Caption = uniOrdered
    lblDispWat.Caption = uniWat1cb
    lblDispDrv.Caption = uniDrvCode
    lblDispDrvName.Caption = uniNm
    lblDispDrvReg.Caption = uniDrvReg
    lblDispDrvCap.Caption = uniCapacity
    lblOrdRecIndex.Caption = uniRecCode
    lblOrdRecName.Caption = uniNm
    lblOrdRecClass.Caption = uniClass
    lblOrdClntIndex.Caption = uniClntCode
    lblOrdClntName.Caption = uniFirm
    lblOrdClntObj.Caption = uniObj
    lblOrd.Caption = uniCode
    lblOrdQuant.Caption = uniQuant
    lblOrdDate.Caption = uniDateReady
    btnClearOrd.Caption = uniNewa
    btnSvNwOrd.Caption = uniSave
    btnDelOrd.Caption = uniDel
    lblRecIndex.Caption = uniCode
    lblRecName.Caption = uniNm
    lblRecType.Caption = uniRecType
    lblRecClass.Caption = uniClass
    lblRecClassK.Caption = uniClassK
    lblRecClassV.Caption = uniClassV
    lblRecClassH.Caption = uniClassH
    lblRecClassP.Caption = uniClassP
    lblRecEDM.Caption = uniEDM
    lblTotalKg.Caption = uniTotalKg
    lblTimeMix.Caption = uniTimeMix
    lblTimePour.Caption = uniTimePour
    lblRecNote.Caption = uniRecNote
    frIM.Caption = uniIM
    frScr.Caption = uniCem
    frWat.Caption = uniWat
    frChem.Caption = uniChem
    btnClearRec.Caption = uniNewa
    btnSvNwRec.Caption = uniSave
    btnDelRec.Caption = uniDel
    btnShowRec.Caption = uniSettings
    lblClntIndex.Caption = uniCode
    lblClntName.Caption = uniFirm
    lblClntBG.Caption = uniBG
    lblClntMOL.Caption = uniMOL
    lblClntAdd.Caption = uniAdd
    lblClntTel.Caption = uniTel
    btnClearClnt.Caption = uniNew
    btnSvNwClnt.Caption = uniSave
    btnDelClnt.Caption = uniDel
    btnShowClnt.Caption = uniSettings
    btnObjects.Caption = uniNew
    btnDelObj.Caption = uniDel
    btnShowObj.Caption = uniSettings
    lblDrvIndex.Caption = uniCode
    lblNameDrv.Caption = uniNm
    lblDrvReg.Caption = uniDrvReg
    lblDrvCap.Caption = uniCapacity
    lblDrvMod.Caption = uniMod
    lblDrvTel.Caption = uniTel
    lblDrvNote.Caption = uniNote
    btnClearDrv.Caption = uniNew
    btnSvNwDrv.Caption = uniSave
    btnDelDrv.Caption = uniDel
    btnShowDrv.Caption = uniSettings
    lblSupIndex.Caption = uniCode
    lblSupName.Caption = uniFirm
    lblSupBG.Caption = uniBG
    lblSupMOL.Caption = uniMOL
    lblSupAdd.Caption = uniAdd
    lblSupTel.Caption = uniTel
    lblSupNote.Caption = uniNote
    btnClearSup.Caption = uniNew
    btnSvNwSup.Caption = uniSave
    btnDelSup.Caption = uniDel
    btnShowSup.Caption = uniSettings
    lblMatIndex.Caption = uniCode
    lblMatName.Caption = uniNm
    lblMatType.Caption = uniType
    lblMatLoad.Caption = uniLoad
    btnClearMat.Caption = uniNew
    btnSvNwMat.Caption = uniSave
    btnDelMat.Caption = uniDel
    btnAddMatDlvr.Caption = uniDlvr
    btnSvExp.Caption = uniEnterExp
    btnRevision.Caption = uniRevision
    
    stOrd.Caption = ""
    stClnt.Caption = ""
    stExp.Caption = ""
    lblMach.Caption = "Машина " & MachineNumber
    If MachineNumber = 1 Then NotMachineNumber = 2
    If MachineNumber = 2 Then NotMachineNumber = 1
    Me.btnChMach.Caption = "Към Машина " & NotMachineNumber
    Me.btnChSilos.Caption = "Смяна силози"
    Me.lblSilos(0).Caption = "Силоз 1 -->"
    Me.lblSilos(1).Caption = "Силоз 2 -->"
    Me.lblSilos(2).Caption = "Силоз 3 -->"
    Me.lblSilos(3).Caption = "Силоз 4 -->"
    
    
    'етикети на смяна на силозите
Dim PrevSetSilos   As Boolean

Dim strSubKeySilos As String

Dim PlaceSilosSet As String

Dim PlaceSilos As String
    
    If MachineNumber = 1 Then
        PlaceSilosSet = Place1SilosSet
        PlaceSilos = Place1Silos
    End If
    If MachineNumber = 2 Then
        PlaceSilosSet = Place2SilosSet
        PlaceSilos = Place2Silos
    End If
    
    strSubKeySilos = Trim(PlaceSilosSet)
    PrevSetSilos = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeySilos)
    
    On Error Resume Next
    
    If PrevSetSilos = True Then
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

    Me.numSilos(0).Caption = rSilos1
    Me.numSilos(1).Caption = rSilos2
    Me.numSilos(2).Caption = rSilos3
    Me.numSilos(3).Caption = rSilos4
    
'--------------------------------------------------------------
    
    btnSvExp.Visible = False

    lblW = Split(uniOrdsVert)
    lblWait.Caption = ""

    For X = 0 To UBound(lblW)
        lblWait.Caption = lblWait.Caption & lblW(X) & vbCrLf
    Next
    
    lblR = Split(uniReadyVert)
    lblReady.Caption = ""

    For X = 0 To UBound(lblR)
        lblReady.Caption = lblReady.Caption & lblR(X) & vbCrLf
    Next
    
    lblR2 = Split(uniReady2Vert)
    lblReady2.Caption = ""

    For X = 0 To UBound(lblR2)
        lblReady2.Caption = lblReady2.Caption & lblR2(X) & vbCrLf
    Next
    
    For i = 0 To 5
        s1(i).Caption = uniFlow & i + 1
    Next i
    
    If InStr(txtDispQuant.Text, DecSep) = 0 Then PointLook1 = False
    If InStr(txtOrdQuant.Text, DecSep) = 0 Then PointLook2 = False
    If InStr(txtRec4(Index).Text, DecSep) = 0 Then PointLook3 = False
    If InStr(txtCapDrv.Text, DecSep) = 0 Then PointLook4 = False
    
    '-----------------------Start postgreSQL-----------------------------------
    Dim cn        As ADODB.Connection

    Dim rs        As Recordset

    Dim rsLog     As Recordset

    Dim WrkPerm   As String

    Dim LogErr    As Boolean

    Dim chD       As Date

    Dim chD1      As Date

    Dim chDnow    As Date

    Dim chDnowStr As String

    Dim MixID     As Long

    Dim ExpID     As Long
    
    chDnow = Format(Now, "DD-MM-YYYY")
    chDnowStr = Format(Now, "DD-MM-YYYY")
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT * FROM tempmix_bc" & MachineNumber & " ORDER BY mix_id ASC LIMIT 1;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        MixID = rs!mix_id
        ExpID = rs!exp_id
        OrderedQuantl = rDs(rs!ordered_q)
        RealQuantl = rDs(rs!real_q)
        TotalKGsl = rDs(rs!total_kg_temp)
        rs.MoveNext
    Loop
    
    'проверка за разрешение за старт на програмата
    Set rsLog = cn.Execute("SELECT log_enter_date FROM entry_log ORDER BY log_num DESC LIMIT 2;")

    If Not rsLog.EOF And Not rsLog.BOF Then
        rsLog.MoveFirst
        rsLog.MoveNext

        If Not rsLog.EOF Then
            chD1 = rsLog!log_enter_date
        Else
            GoTo Skip
        End If
    End If

    LogErr = False

    If (chD1 - chDnow) > 0 Then LogErr = True
    
    Set rs = cn.Execute("SELECT work_permission, stamp_date FROM settings_bc1 ORDER BY ind ASC LIMIT 1;")
    
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst

        If rs!work_permission <> "Null" Then
            WrkPerm = rs!work_permission
        Else
            GoTo Skip
        End If

    Else
        GoTo Skip
    End If
    
    If WrkPerm = "stop" Or Val(rDs(WrkPerm)) < 0 Then
        MsgBox MsgNoPayment & vbCrLf & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBx

        End

    ElseIf WrkPerm = "delay" Then
        MsgBox MsgNoPaymentDays & vbCrLf & MsgDaysLeft & " - " & "10" & vbCrLf & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBx
        cn.Execute ("UPDATE settings_bc1 SET work_permission = '10', stamp_date = '" & chDnowStr & "'")
    ElseIf WrkPerm = "longdelay" Then
        MsgBox MsgNoPaymentDays & vbCrLf & MsgDaysLeft & " - " & "30" & vbCrLf & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBx
        cn.Execute ("UPDATE settings_bc1 SET work_permission = '30', stamp_date = '" & chDnowStr & "'")
    ElseIf Val(rDs(WrkPerm)) >= 0 Then

        Dim calcDay  As Integer

        Dim DaysLeft As Integer

        chD = rs!stamp_date
        calcDay = (chDnow - chD)

        If calcDay <= 0 And LogErr = True Then
            MsgBox MsgNoPayment & vbCrLf & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBxFatal
            cn.Execute ("UPDATE settings_bc1 SET work_permission = 'stop', stamp_date = '" & chDnowStr & "';")

            End

        End If

        If DaysLeft >= 0 Then
            DaysLeft = Val(rDs(WrkPerm)) - calcDay
            MsgBox MsgNoPaymentDays & vbCrLf & MsgDaysLeft & " - " & DaysLeft & vbCrLf & MsgCallTIP, vbOKOnly Or vbCritical, MsgErrBx
            cn.Execute ("UPDATE settings_bc1 SET work_permission = '" & DaysLeft & "', stamp_date = '" & chDnowStr & "';")
        End If

        If DaysLeft < 0 Then
            cn.Execute ("UPDATE settings_bc1 SET work_permission = 'stop', stamp_date = '" & chDnowStr & "';")
        End If
    End If

Skip:
    
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    '--------------------------End PostgreSQL-----------------------------------
    
    MousePointer = vbHourglass
    
    StatusBar.Style = sbrNormal
    
    Set today = StatusBar.Panels.Add()
    today.Style = sbrDate
    today.AutoSize = sbrContents
        
    Set oper = StatusBar.Panels.Add()
    oper.Style = sbrText
    oper.AutoSize = sbrContents
    oper.Text = OperName
    
    Set mixes = StatusBar.Panels.Add()
    mixes.Style = sbrText
    mixes.AutoSize = sbrContents
    mixes.Text = "Замеси: " & MixID
        
    Set exps = StatusBar.Panels.Add()
    exps.Style = sbrText
    exps.AutoSize = sbrContents
    exps.Text = "Експедиции: " & ExpID

    Set voexps = StatusBar.Panels.Add()
    voexps.Style = sbrText
    voexps.AutoSize = sbrContents
    voexps.Text = "Заявена експедиция: " & OrderedQuantl & " m3"

    Set vexps = StatusBar.Panels.Add()
    vexps.Style = sbrText
    vexps.AutoSize = sbrContents
    vexps.Text = "Обем експедиция: " & RealQuantl & " m3"
    
    Set kgexps = StatusBar.Panels.Add()
    kgexps.Style = sbrText
    kgexps.AutoSize = sbrContents
    kgexps.Text = "Тегло експедиция: " & TotalKGsl & " kg"
    
    intEmpFileNbr1 = FreeFile
    DispPanel.Hide
    DispPanel.Show
    
    MousePointer = vbHourglass
    
'----------------------------------------Start OPC Initialize---------------------------------
'настройки на opc server
    VipaActive = True
    
    OffMode = False
    
    MousePointer = vbHourglass

    If VipaActive Then
        RecMin = 11
    Else
        RecMin = 1
    End If
'осъществяваме връзката
    Set ConOPCServer = New OPCServer
    ConNodeName = ""
    ConOPCServer.Connect MyServer, ConNodeName

'подготвяме формата за OPC за VIPA S-7 CPU-313
    Load frmOPC
    frmOPC.Hide
    
    Const RateUp = 100
    
    Dim OPCch As String
    Dim OPCdev As String
    Dim OPCpref As String
    Dim grNamePanels As String
    Dim grNameConfig As String
    Dim grNameStatus As String
    Dim grNameRec As String
    Dim grNameResult As String
    Dim grNameReady As String
    Dim itemIDPanels(0 To 3) As String
    Dim itemIDConfig(0 To 0) As String
    Dim itemIDStatus(0 To 4) As String
    Dim itemIDRec(0 To 38) As String
    Dim itemIDResults(0 To 16) As String
    Dim itemIDReady(0 To 3) As String
    Dim ActiveItemErrPanels() As Long
    Dim ActiveItemErrTest() As Long
    Dim ActiveItemErrConfig() As Long
    Dim ActiveItemErrStatus() As Long
    Dim ActiveItemErrRec() As Long
    Dim ActiveItemErrResults() As Long
    Dim ActiveItemErrReady() As Long
    Dim AnItemIsGood As Boolean
    Dim SyncItemValuesConfig() As Variant
    Dim SyncItemSrvErrConfig() As Long
    
'клас за групите
    Set ConSrvGroup = ConOPCServer.OPCGroups
        ConSrvGroup.DefaultGroupIsActive = True
        ConSrvGroup.DefaultGroupDeadband = 0
    
'имена на групите
    grNamePanels = "Panels"
    grNameConfig = "Config"
    grNameStatus = "Status"
    grNameRec = "RecInput"
    grNameResult = "Result"
    grNameReady = "Ready"

'име на канал и устройство
    OPCch = "TIPService."
    OPCdev = "VIPA."
    OPCpref = OPCch & OPCdev
    
'създаваме групите
    Set ConGroupPanels = ConSrvGroup.Add(grNamePanels)
        ConGroupPanels.UpdateRate = RateUp
        ConGroupPanels.IsSubscribed = True
        ConGroupPanels.IsActive = True
        
    Set ConGroupConfig = ConSrvGroup.Add(grNameConfig)
        ConGroupConfig.UpdateRate = RateUp
        ConGroupConfig.IsSubscribed = True
        ConGroupConfig.IsActive = True
        
    Set ConGroupStatus = ConSrvGroup.Add(grNameStatus)
        ConGroupStatus.UpdateRate = RateUp
        ConGroupStatus.IsSubscribed = True
        ConGroupStatus.IsActive = True
        
    Set ConGroupRec = ConSrvGroup.Add(grNameRec)
        ConGroupRec.UpdateRate = RateUp
        ConGroupRec.IsSubscribed = True
        ConGroupRec.IsActive = True
    
    Set ConGroupResults = ConSrvGroup.Add(grNameResult)
        ConGroupResults.UpdateRate = RateUp
        ConGroupResults.IsSubscribed = True
        ConGroupResults.IsActive = True

    Set ConGroupReady = ConSrvGroup.Add(grNameReady)
        ConGroupReady.UpdateRate = RateUp
        ConGroupReady.IsSubscribed = True
        ConGroupReady.IsActive = True

'задаваме имената на клетките от паметта
    ItemCountPanels = 4 'приборите
    itemIDPanels(0) = "Panels.IMPanel"
    itemIDPanels(1) = "Panels.CemPanel"
    itemIDPanels(2) = "Panels.WatPanel"
    itemIDPanels(3) = "Panels.HDPanel"
      
    ItemCountConfig = 1 'параметрите на машината
    itemIDConfig(0) = "Config.MixCap" 'MixCap
    
    ItemCountStatus = 5 'статусите на машината
    itemIDStatus(0) = "Status.Auto"
    itemIDStatus(1) = "Status.Emergency"
    itemIDStatus(2) = "Status.MixerClosed"
    itemIDStatus(3) = "Status.MixerOpened"
    itemIDStatus(4) = "_System._Error"
    
    ItemCountRec = 39 'адреси за въвеждане на заявка към контролера
    itemIDRec(0) = "RecInput.RecNum"
    itemIDRec(1) = "RecInput.RecName"
    itemIDRec(2) = "RecInput.Quant"
    itemIDRec(3) = "RecInput.IM1"
    itemIDRec(4) = "RecInput.IM2"
    itemIDRec(5) = "RecInput.IM3"
    itemIDRec(6) = "RecInput.IM4"
    itemIDRec(7) = "RecInput.IM5"
    itemIDRec(8) = "RecInput.IM6"
    itemIDRec(9) = "RecInput.Cem1"
    itemIDRec(10) = "RecInput.Cem2"
    itemIDRec(11) = "RecInput.Cem3"
    itemIDRec(12) = "RecInput.Cem4"
    itemIDRec(13) = "RecInput.Wat1"
    itemIDRec(14) = "RecInput.Wat2"
    itemIDRec(15) = "RecInput.HD1"
    itemIDRec(16) = "RecInput.HD2"
    itemIDRec(17) = "RecInput.HD3"
    itemIDRec(18) = "RecInput.HD4"
    itemIDRec(19) = "RecInput.OrdIM1"
    itemIDRec(20) = "RecInput.OrdIM2"
    itemIDRec(21) = "RecInput.OrdIM3"
    itemIDRec(22) = "RecInput.OrdIM4"
    itemIDRec(23) = "RecInput.OrdIM5"
    itemIDRec(24) = "RecInput.OrdIM6"
    itemIDRec(25) = "RecInput.OrdCem1"
    itemIDRec(26) = "RecInput.OrdCem2"
    itemIDRec(27) = "RecInput.OrdCem3"
    itemIDRec(28) = "RecInput.OrdCem4"
    itemIDRec(29) = "RecInput.OrdWat1"
    itemIDRec(30) = "RecInput.OrdWat2"
    itemIDRec(31) = "RecInput.OrdHD1"
    itemIDRec(32) = "RecInput.OrdHD2"
    itemIDRec(33) = "RecInput.OrdHD3"
    itemIDRec(34) = "RecInput.OrdHD4"
    itemIDRec(35) = "RecInput.Mtime"
    itemIDRec(36) = "RecInput.Mpour"
    itemIDRec(37) = "RecInput.RecReady"
    itemIDRec(38) = "RecInput.RecName1"
    
    ItemCountResults = 17 'адреси за резултатите
    itemIDResults(0) = "Results.CycCount"
    itemIDResults(1) = "Results.im1"
    itemIDResults(2) = "Results.im2"
    itemIDResults(3) = "Results.im3"
    itemIDResults(4) = "Results.im4"
    itemIDResults(5) = "Results.im5"
    itemIDResults(6) = "Results.im6"
    itemIDResults(7) = "Results.cem1"
    itemIDResults(8) = "Results.cem2"
    itemIDResults(9) = "Results.cem3"
    itemIDResults(10) = "Results.cem4"
    itemIDResults(11) = "Results.wat1"
    itemIDResults(12) = "Results.wat2"
    itemIDResults(13) = "Results.hd1"
    itemIDResults(14) = "Results.hd2"
    itemIDResults(15) = "Results.hd3"
    itemIDResults(16) = "Results.hd4"

    ItemCountReady = 4 'адреси за резултатите
    itemIDReady(0) = "Results.IMReady"
    itemIDReady(1) = "Results.CemReady"
    itemIDReady(2) = "Results.WatReady"
    itemIDReady(3) = "Results.HDReady"

'създаваме връзка на клетките
    For i = 0 To ItemCountPanels - 1
        OPCItemIDPanels(i + 1) = OPCpref & itemIDPanels(i)
        ClientHandlesPanels(i + 1) = i
    Next i
    Set OPCItemCollPanels = ConGroupPanels.OPCItems
        OPCItemCollPanels.DefaultIsActive = True
        OPCItemCollPanels.AddItems ItemCountPanels, OPCItemIDPanels, _
        ClientHandlesPanels, ItemSrvHandlesPanels, ItemSrvErrPanels
        OPCItemCollPanels.SetActive ItemCountPanels, ItemSrvHandlesPanels, _
        True, ActiveItemErrPanels
        
    For i = 0 To ItemCountConfig - 1
        OPCItemIDConfig(i + 1) = OPCpref & itemIDConfig(i)
        ClientHandlesConfig(i + 1) = i
    Next i
    Set OPCItemCollConfig = ConGroupConfig.OPCItems
        OPCItemCollConfig.DefaultIsActive = True
        OPCItemCollConfig.AddItems ItemCountConfig, OPCItemIDConfig, _
        ClientHandlesConfig, ItemSrvHandlesConfig, ItemSrvErrConfig
        OPCItemCollConfig.SetActive ItemCountConfig, ItemSrvHandlesConfig, _
        True, ActiveItemErrConfig
    For i = 1 To ItemCountConfig
        handyConfig(i) = ItemSrvHandlesConfig(i)
    Next i
    
    For i = 0 To ItemCountStatus - 1
        OPCItemIDStatus(i + 1) = OPCpref & itemIDStatus(i)
        ClientHandlesStatus(i + 1) = i
    Next i
    Set OPCItemCollStatus = ConGroupStatus.OPCItems
        OPCItemCollStatus.DefaultIsActive = True
        OPCItemCollStatus.AddItems ItemCountStatus, OPCItemIDStatus, _
        ClientHandlesStatus, ItemSrvHandlesStatus, ItemSrvErrStatus
        OPCItemCollStatus.SetActive ItemCountStatus, ItemSrvHandlesStatus, _
        True, ActiveItemErrStatus

    For i = 0 To ItemCountRec - 1
        OPCItemIDRec(i + 1) = OPCpref & itemIDRec(i)
        ClientHandlesRec(i + 1) = i
    Next i
    Set OPCItemCollRec = ConGroupRec.OPCItems
        OPCItemCollRec.DefaultIsActive = True
        OPCItemCollRec.AddItems ItemCountRec, OPCItemIDRec, _
        ClientHandlesRec, ItemSrvHandlesRec, ItemSrvErrRec
        OPCItemCollRec.SetActive ItemCountRec, ItemSrvHandlesRec, _
        True, ActiveItemErrRec
    For i = 1 To ItemCountRec
        handyRec(i) = ItemSrvHandlesRec(i)
    Next i

    For i = 0 To ItemCountResults - 1
        OPCItemIDResults(i + 1) = OPCpref & itemIDResults(i)
        ClientHandlesResults(i + 1) = i
    Next i
    Set OPCItemCollResults = ConGroupResults.OPCItems
        OPCItemCollResults.DefaultIsActive = True
        OPCItemCollResults.AddItems ItemCountResults, OPCItemIDResults, _
        ClientHandlesResults, ItemSrvHandlesResults, ItemSrvErrResults
        OPCItemCollResults.SetActive ItemCountResults, ItemSrvHandlesResults, _
        True, ActiveItemErrResults

    For i = 0 To ItemCountReady - 1
        OPCItemIDReady(i + 1) = OPCpref & itemIDReady(i)
        ClientHandlesReady(i + 1) = i
    Next i
    Set OPCItemCollReady = ConGroupReady.OPCItems
        OPCItemCollReady.DefaultIsActive = True
        OPCItemCollReady.AddItems ItemCountReady, OPCItemIDReady, _
        ClientHandlesReady, ItemSrvHandlesReady, ItemSrvErrReady
        OPCItemCollReady.SetActive ItemCountReady, ItemSrvHandlesReady, _
        True, ActiveItemErrReady
    For i = 1 To ItemCountReady
        handyReady(i) = ItemSrvHandlesReady(i)
    Next i

    lblLoading.ForeColor = &HC000&
    lblLoading.Caption = "  " & uniLoading & "  "
    lblLoading.Refresh
    Sleep 77
    
    For i = 0 To 77
        lblLoading.Caption = "\ " & uniLoading & " /"
        lblLoading.Refresh
        Sleep 21
        lblLoading.Caption = "- " & uniLoading & " -"
        lblLoading.Refresh
        Sleep 21
        lblLoading.Caption = "/ " & uniLoading & " \"
        lblLoading.Refresh
        Sleep 21
        lblLoading.Caption = "| " & uniLoading & " |"
        lblLoading.Refresh
        Sleep 21
        If frmOPC.Stat(4).Text = "False" Then
            lblLoading.Caption = uniLoaded
            lblLoading.Refresh
            Exit For
        Else
            lblLoading.ForeColor = &HFF&
            lblLoading.Caption = MsgNotRespOPC
            lblLoading.Refresh

            Dim response As Integer

            MousePointer = vbDefault
            response = MsgBox(MsgOffline, vbQuestion Or vbYesNo, MsgNotRespOPC)

            If response = vbYes Then
                Me.btnDispStart.Enabled = False
                OffMode = True
                GoTo OfflineMode
            Else
                DontAskExit = True
                Unload Me

            End If
        End If
        
    Next i
    
    MousePointer = vbHourglass
    
    Sleep 7

    Open OPCSetFile For Output As intEmpFileNbr1
    Write #intEmpFileNbr1, MyServer
    Close
    
    Sleep 7
    
    MousePointer = vbHourglass
    
'    ConGroupConfig.SyncRead OPCDevice, ItemCountConfig, ItemSrvHandlesConfig, SyncItemValuesConfig, SyncItemSrvErrConfig
'
'    For I = 1 To ItemCountConfig
'        If SyncItemSrvErrConfig(I) = 0 Then frmOPC.Config(I - 1).Text = SyncItemValuesConfig(I)
'    Next I
'----------------------------------------END OPC-----------------------------------
    
    MousePointer = vbHourglass
    
    'зареждане на параметрите на машината
    Load frmParam
    ns1 = Val(frmParam.txtNumIMSilos.Text)
    ns3 = Val(frmParam.txtNumCementSilos.Text)
    ns2 = Val(frmParam.txtNumWaterSilos.Text)
    ns4 = Val(frmParam.txtNumChemSilos.Text)
    MixCap = CSng(rDs(frmOPC.Config(0)))
    TMd = Val(frmParam.txtTimeMixDefault)
    TPd = Val(frmParam.txtTimePourDefault)
    Unload frmParam
    
    For i = 1 To ns3
        Me.lblSilos(i - 1).Visible = True
        Me.numSilos(i - 1).Visible = True
    Next i
    
    'запис на параметрите на машината

    '-----------------------Start postgreSQL-----------------------------------
    Dim comIns  As String

    Dim comEdit As String
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    
    MousePointer = vbHourglass
    
    comIns = "INSERT INTO settings_bc" & MachineNumber & " VALUES(1,'" & ns1 & "','" & ns3 & "','" & ns2 & "','" & ns4 & "')"
    
    comEdit = "UPDATE settings_bc" & MachineNumber & " SET im_num = '" & ns1 & "',cem_num = '" & ns3 & "',wat_num = '" & ns2 & "',chem_num = '" & ns4 & "' WHERE ind = 1"
                    
    Set rs = cn.Execute("SELECT * FROM settings_bc" & MachineNumber & " WHERE ind = 1;")
    
    If Not rs.BOF And Not rs.EOF Then
        Set rs = cn.Execute(comEdit)
    Else
        Set rs = cn.Execute(comIns)
    End If
    
    'зареждане на променливите за брой на експедициони бележки за печат
    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm1';")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        numSheetsForm1 = rs!Value
        rs.MoveNext
    Loop

    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm2';")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        numSheetsForm2 = rs!Value
        rs.MoveNext
    Loop
    
    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm3';")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        numSheetsForm3 = rs!Value
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    '--------------------------End PostgreSQL-----------------------------------

    GoTo SkipOffline
OfflineMode:

    '-----------------------Start postgreSQL-----------------------------------
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT * FROM settings_bc" & MachineNumber & " WHERE ind = 1;")
    
    If Not rs.BOF And Not rs.EOF Then
        ns1 = rs!im_num
        ns3 = rs!cem_num
        ns2 = rs!wat_num
        ns4 = rs!chem_num
    Else
    End If
    
    'зареждане на променливите за брой на експедициони бележки за печат
    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm1';")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        numSheetsForm1 = rs!Value
        rs.MoveNext
    Loop

    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm2';")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        numSheetsForm2 = rs!Value
        rs.MoveNext
    Loop
    
    Set rs = cn.Execute("SELECT * FROM settings_soft WHERE parameter = 'NumSheetsForm3';")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        numSheetsForm3 = rs!Value
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    '--------------------------End PostgreSQL-----------------------------------

SkipOffline:

    'запис на имената на течките в случай на промяна в данните
    Load frmNameSilos
    frmNameSilos.Hide
    Call frmNameSilos.btnSaveSilos_Click
    Unload frmNameSilos
    
    MousePointer = vbHourglass
    
    ClockT.Interval = 100
    ScalesT.Interval = 100
    
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = False
    frOrders.Visible = False
    
    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = True
    frAbout.Visible = True
    
    Call OpenAbout
    
    MousePointer = vbHourglass
    
    Dim PrevSet   As Boolean

    Dim strSubKey As String

    strSubKey = Trim(PlaceProgAllow)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)

    If PrevSet = True Then
        rActForm1 = GetSetting(PlaceProgSettings, PlaceAllow, "ActForm1", ErrRes)
        rActForm2 = GetSetting(PlaceProgSettings, PlaceAllow, "ActForm2", ErrRes)
        rActForm3 = GetSetting(PlaceProgSettings, PlaceAllow, "ActForm3", ErrRes)
        rActDel = GetSetting(PlaceProgSettings, PlaceAllow, "ActDel", ErrRes)
        rDeactNRPass = GetSetting(PlaceProgSettings, PlaceAllow, "DeactNRPass", ErrRes)
        rDeactDRPass = GetSetting(PlaceProgSettings, PlaceAllow, "DeactDRPass", ErrRes)
    Else
        rActForm1 = 0
        rActForm2 = 0
        rActForm3 = 0
        rActDel = 0
        rDeactNRPass = 0
        rDeactDRPass = 0
        Me.chPrintConf.Value = 0
    End If
    
    'проверка в регистъра дали е включен автопечата
    Dim PrevSet1   As Boolean

    Dim strSubKey1 As String

    strSubKey1 = Trim(PlaceProgPrint)
    PrevSet1 = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey1)

    If PrevSet1 = True Then
        Me.chPrintConf.Value = GetSetting(PlaceProgSettings, PlacePrint, "AutoPrForm", ErrRes)
    Else
        Me.chPrintConf.Value = 0
    End If
    
    If frmLogin.AdminSuccess = True And frmLogin.RootUser = False Then
        Me.btnDispStart.Enabled = False 'забрана на бутона "старт" в режим администратор
    End If

    'проверка в регистъра дали е включен въпроса за недостатъчно количество в силозите
    Dim PrevSetQSilos As Boolean
    Dim strSubKeyQSilos As String

    If MachineNumber = 1 Then
        strSubKeyQSilos = Trim(Place1SilosQ)
        PrevSetQSilos = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeyQSilos)

        If PrevSetQSilos = True Then
            QuestSilos = GetSetting(PlaceProgSettings, Place1Q, "Quest1Silos", ErrRes)
        Else
            QuestSilos = 0
        End If
    ElseIf MachineNumber = 2 Then
        strSubKeyQSilos = Trim(Place2SilosQ)
        PrevSetQSilos = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeyQSilos)

        If PrevSetQSilos = True Then
            QuestSilos = GetSetting(PlaceProgSettings, Place2Q, "Quest2Silos", ErrRes)
        Else
            QuestSilos = 0
        End If
    End If

    'проверка в регистъра дали е включена редакцията на бележките
    Dim PrevSetEditor As Boolean
    Dim strSubKeyEditor As String
    
    strSubKeyEditor = Trim(PlaceEditor)
    PrevSetEditor = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeyEditor)
    
    If PrevSetEditor = True Then
        ShowEditor = GetSetting(PlaceProgSettings, PlaceEd, "NotesEditor", ErrRes)
    Else
        ShowEditor = 0
    End If

    Me.btnDisp.Enabled = True
    Me.btnOrders.Enabled = True
    Me.btnRecepies.Enabled = True
    Me.btnClients.Enabled = True
    Me.btnDrivers.Enabled = True
    Me.btnSuppliers.Enabled = True
    Me.btnMaterials.Enabled = True
    Me.btnNotes.Enabled = True
    Me.btnAdminPanel.Enabled = True
    Me.BtnExit.Enabled = True
    Me.chPrintConf.Enabled = True
    
    If frmLogin.AdminSuccess = True Or frmLogin.RootUser = True Then
        If frmLogin.RootUser = False Then Me.btnDelOrd.Visible = False
        Me.btnRevision.Visible = True
        Me.btnDelDrv.Enabled = True
        Me.btnDelClnt.Enabled = True
        Me.btnDelRec.Enabled = True
        Me.btnDelSup.Enabled = True
        Me.btnDelMat.Enabled = True
    ElseIf frmLogin.AdminSuccess = False Or frmLogin.RootUser = False Then
        Me.btnRevision.Visible = True
        Me.btnRevision.Caption = "Дневен отчет"
        If rActDel = 0 Then
            Me.btnDelOrd.Visible = False
            Me.btnDelDrv.Enabled = False
            Me.btnDelClnt.Enabled = False
            Me.btnDelRec.Enabled = False
            Me.btnDelSup.Enabled = False
            Me.btnDelMat.Enabled = False
        Else
            Me.btnDelOrd.Visible = False

            If frmLogin.RootUser = True Then Me.btnDelOrd.Visible = True
            Me.btnDelDrv.Enabled = True
            Me.btnDelClnt.Enabled = True
            Me.btnDelRec.Enabled = True
            Me.btnDelSup.Enabled = True
            Me.btnDelMat.Enabled = True
        End If
    End If
    
    PrintRightBut = False
    
    MousePointer = vbDefault
    
            'проверка дали е сговнясана програмата
            Dim PrevSetShit As Boolean
            Dim strSubKeyShit As String
    
            strSubKeyShit = Trim(PlaceShit)
            PrevSetShit = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeyShit)
    
            If PrevSetShit = True Then
                ShitEnabled = GetSetting(PlaceProgSettings, Shit, "Shit", ErrRes)
            Else
                ShitEnabled = 0
            End If
            
            If ShitEnabled = True Then
                ConStr = ""
            Else
                ConStr = "PROVIDER=PostgreSQL;" & "DATA SOURCE=" & IPConnStr & ";" & "LOCATION=" & DbaseName & ";" & "USER ID=" & DbaseUser & ";" & "PASSWORD=" & PassConnStr & ";"
            End If

End Sub

Sub ConGroupPanels_DataChange(ByVal TransactionIDPanels As Long, ByVal NumItemsPanels As Long, _
ClientHandlesPanels() As Long, ItemValuesPanels() As Variant, QualitiesPanels() As Long, _
TimeStampsPanels() As Date)
'събития за обновяване на стойностите на клетките от OPC
    Dim ip As Integer
    For ip = 1 To NumItemsPanels
        frmOPC.Panel(ClientHandlesPanels(ip)).Text = ItemValuesPanels(ip)
    Next ip
End Sub

Sub ConGroupConfig_DataChange(ByVal TransactionIDConfig As Long, ByVal NumItemsConfig As Long, _
ClientHandlesConfig() As Long, ItemValuesConfig() As Variant, QualitiesConfig() As Long, _
TimeStampsConfig() As Date)
'събития за обновяване на стойностите на клетките от OPC
    Dim ic As Integer
    For ic = 1 To NumItemsConfig
        frmOPC.Config(ClientHandlesConfig(ic)).Text = ItemValuesConfig(ic)
        Me.btnMixCap.Caption = "Обем Миксер: " & ItemValuesConfig(ic)
    Next ic
End Sub

Sub ConGroupStatus_DataChange(ByVal TransactionIDStatus As Long, ByVal NumItemsStatus As Long, _
ClientHandlesStatus() As Long, ItemValuesStatus() As Variant, QualitiesStatus() As Long, _
TimeStampsStatus() As Date)
'събития за обновяване на стойностите на клетките от OPC
    Dim ist As Integer
    For ist = 1 To NumItemsStatus
        frmOPC.Stat(ClientHandlesStatus(ist)).Text = ItemValuesStatus(ist)
        frmOPC.Stat(ClientHandlesStatus(ist)).Refresh
    Next ist
End Sub

Sub ConGroupRec_DataChange(ByVal TransactionIDRec As Long, ByVal NumItemsRec As Long, _
ClientHandlesRec() As Long, ItemValuesRec() As Variant, QualitiesRec() As Long, _
TimeStampsRec() As Date)
'събития за обновяване на стойностите на клетките от OPC
    Dim ir As Integer
    For ir = 1 To NumItemsRec
        frmOPC.RecInput(ClientHandlesRec(ir)).Text = ItemValuesRec(ir)
    Next ir
End Sub

Sub ConGroupResults_DataChange(ByVal TransactionIDResults As Long, ByVal NumItemsResults As Long, _
ClientHandlesResults() As Long, ItemValuesResults() As Variant, QualitiesResults() As Long, _
TimeStampsResults() As Date)
'събития за обновяване на стойностите на клетките от OPC
    Dim irs As Integer
    For irs = 1 To NumItemsResults
        frmOPC.Result(ClientHandlesResults(irs)).Text = ItemValuesResults(irs)
    Next irs
End Sub

Sub ConGroupReady_DataChange(ByVal TransactionIDReady As Long, ByVal NumItemsReady As Long, _
ClientHandlesReady() As Long, ItemValuesReady() As Variant, QualitiesReady() As Long, _
TimeStampsReady() As Date)
'събития за обновяване на стойностите на клетките от OPC
    Dim ird As Integer
    For ird = 1 To NumItemsReady
        frmOPC.Ready(ClientHandlesReady(ird)).Text = ItemValuesReady(ird)
    Next ird
End Sub

Private Sub ClockT_Timer()

    Dim MyTime As String

    Call GetStatVIPA
    MyTime = Format$(Now, "hh:mm:ss")
    Clock.Text = Left$(MyTime, 2) & ":" & Mid$(MyTime, 4, 2) & ":" & Right$(MyTime, 2)
    DayToday = Format(Now, "DD-MM-YYYY")
    Call MaintConn

    If MaintConn = True And OffMode = True Then
        MousePointer = vbDefault
        MsgBox MsgConnEst, vbInformation, uniLoaded
        Me.btnDispStart.Enabled = True
        OffMode = False
    End If

    If OffMode = False And MaintConn = False Then

        Dim response As Integer

        MousePointer = vbDefault
        response = MsgBox(MsgOffline, vbQuestion Or vbYesNo, MsgNotRespOPC)

        If response = vbYes Then
            Me.btnDispStart.Enabled = False
            OffMode = True
        Else

            End

        End If
    End If

    If DispPanel.indReq.Caption = statReqStarted And Me.indAvaria.Caption <> statAvaria Then
        DispConfirm.btnSendToController.Enabled = False
    Else
        DispConfirm.btnSendToController.Enabled = True
    End If

End Sub

Private Sub ScalesT_Timer()
    If OffMode = False Then
        kgAggr.Text = ARound(CSng(rDs(frmOPC.Panel(0))), 0)
        kgCem.Text = ARound(CSng(rDs(frmOPC.Panel(1))), 0)
        kgWt.Text = ARound(CSng(rDs(frmOPC.Panel(2))), 0)
        kgChm.Text = Format(ARound(CSng(rDs(frmOPC.Panel(3))), 2), "0.00")
    Else
        kgAggr.Text = "NoConn"
        kgCem.Text = "NoConn"
        kgWt.Text = "NoConn"
        kgChm.Text = "NoConn"
    End If
End Sub

Private Sub AVTimer_Timer()
    If Me.indAvaria.Caption = statAvaria And Me.TimerRes.Enabled = True Then
        Dim response As Integer
        
        response = MsgBox(MsgAvaria & vbNewLine & MsgContinue, vbYesNo Or vbQuestion, statAvaria)
        If response = vbNo Then
            Me.TimerRes.Enabled = False
            Me.TimerStartReq.Enabled = False
            Me.FormT.Enabled = False
            ExpeditionStarted = False
            DispConfirm.btnSendToController.Enabled = True
        Else
        End If
    Else
    End If
End Sub

Private Sub TimerStartReq_Timer()

    If Me.indReq.Caption = statReqStarted Then
        ExpeditionStarted = True
        ReqTime = Format(Now, "DD.MM.YYYY - HH:MM:SS")
        Me.TimerStartReq.Enabled = False
    End If

End Sub

Private Sub TimerRes_Timer()
    If Me.indMode.Caption = statAuto Then
    
        WasAuto = True
        
        Call GetAggrVIPA
        Call GetCemVIPA
        Call GetWatVIPA
        Call GetHDVIPA
    
        If Val(frmOPC.Result(0).Text) > 0 And Me.indValveMix.Caption = statMixClosed And _
        Val(frmOPC.Result(0).Text) = HelpRes And ExpeditionStarted = True Then
            Call GetResultVIPA
        End If
    End If
    
    If Me.indMode.Caption = statMan And WasAuto = True Then
        
        Call GetAggrVIPA
        Call GetCemVIPA
        Call GetWatVIPA
        Call GetHDVIPA
        
        If okAggr And okCem And okWat And okHD Then
            Call GetResultVIPA
        Else
        End If
        
    Else
    End If

End Sub

Private Sub FormT_Timer()
    If Me.indValveMix.Caption = statMixClosed Then
        Dim response As Integer

        If DispPanel.chPrintConf.Value = 1 Then 'изключваме въпроса за печат при маркиран Авто-печат
            MousePointer = vbDefault
            response = vbYes
        Else
            MousePointer = vbDefault
            response = MsgBox(uniQuestPrint, vbQuestion Or vbYesNo, uniPrint)
        End If

        If response = vbYes Then
            PrintAnyForm = True
            Call FillForm1 'извикваме функция за попълване на форма ако е необходимо
        Else
            DispPanel.FormT.Enabled = False
        End If
    Else
    End If
End Sub

Private Sub chPrintConf_Click()
    SaveSetting PlaceProgSettings, PlacePrint, "AutoPrForm", Me.chPrintConf
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    ShiftTest = Shift And 7

    Select Case ShiftTest

        Case 6
            'проверка дали е сговнясана програмата
            Dim PrevSetShit As Boolean
            Dim strSubKeyShit As String
    
            strSubKeyShit = Trim(PlaceShit)
            PrevSetShit = CheckRegistryKey(HKEY_CURRENT_USER, strSubKeyShit)
    
            If PrevSetShit = True Then
                ShitEnabled = GetSetting(PlaceProgSettings, Shit, "Shit", ErrRes)
            Else
                ShitEnabled = 0
            End If
            
            If ShitEnabled = False Then
                ShitEnabled = True
                ConStr = ""
            Else
                ShitEnabled = False
                ConStr = "PROVIDER=PostgreSQL;" & "DATA SOURCE=" & IPConnStr & ";" & "LOCATION=" & DbaseName & ";" & "USER ID=" & DbaseUser & ";" & "PASSWORD=" & PassConnStr & ";"
            End If

            SaveSetting PlaceProgSettings, Shit, "Shit", ShitEnabled
            
            frDisp.Enabled = False
            frDisp.Visible = False
    
            frOrders.Enabled = False
            frOrders.Visible = False
    
            frRecepies.Enabled = False
            frRecepies.Visible = False
    
            frClients.Enabled = False
            frClients.Visible = False
    
            frDrivers.Enabled = False
            frDrivers.Visible = False
    
            frSuppliers.Enabled = False
            frSuppliers.Visible = False
    
            frMaterials.Enabled = False
            frMaterials.Visible = False
    
            frAbout.Enabled = True
            frAbout.Visible = True
    
    Call OpenAbout

        Case 7
            frmOPC.Show
    End Select

End Sub

Private Sub imgLogo_Click()
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = False
    frOrders.Visible = False
    
    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = True
    frAbout.Visible = True
    
    Call OpenAbout
End Sub

Private Sub btnDisp_Click()
    frDisp.Enabled = True
    frDisp.Visible = True
    
    frOrders.Enabled = False
    frOrders.Visible = False
    
    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = False
    frAbout.Visible = False
    
    Call OpenDisp
End Sub

Private Sub btnDispStart_Click()
    PrintAnyForm = False

    If cmbDispOrd = "" Or cmbDispDrv = "" Then
        MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
    Else
        DispConfirm.Show
        Call DispConfSend
    End If

End Sub

Private Sub txtDispClnt_Change()

    '------------------------------Start PostgreSQL----------------------------------
    Dim cn        As ADODB.Connection

    Dim rs        As Recordset

    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Me.cmbOrdClntObj.Clear
    
    Set rs = cn.Execute("SELECT w_name, w_km FROM worksites WHERE w_cnum = '" & Val(Me.txtDispClnt.Text) & "' ORDER BY w_name;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Me.cmbOrdClntObj.AddItem rs!w_name
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
    '--------------------------End PostgreSQL------------------------------------------

End Sub

Private Sub txtDispQuant_GotFocus()
    txtDispQuant.SelStart = 0
    txtDispQuant.SelLength = Len(txtDispQuant.Text)

    If InStr(txtDispQuant.Text, DecSep) <> 0 Then
        PointLook1 = True
    Else
        PointLook1 = False
    End If

End Sub

Private Sub txtDispQuant_Change()

    If InStr(txtDispQuant.Text, DecSep) <> 0 Then
        PointLook1 = True
    Else
        PointLook1 = False
    End If

End Sub

Private Sub txtDispQuant_KeyPress(KeyAscii As Integer)

    If InStr(txtDispQuant.Text, DecSep) <> 0 Then
        PointLook1 = True
    Else
        PointLook1 = False
    End If

    If txtDispQuant.SelLength = Len(txtDispQuant.Text) Then
        PointLook1 = False
    Else
    End If

    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If

    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook1 = True Then
        KeyAscii = 0
    Else

        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook1 = True
        Else
        End If
    End If

End Sub

Private Sub txtDispWat_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtDispWat_GotFocus()
    txtDispWat.SelStart = 0
    txtDispWat.SelLength = Len(txtDispWat.Text)
End Sub

Private Sub cmbDispOrd_Click()
    Call ChangeDispOrd
End Sub

Private Sub cmbDispDrv_Click()
    Call ChangeDispDrv
End Sub

Private Sub cmbDispDrvName_Click()
    Call ChangeDispDrvName
End Sub

Private Sub cmbDispDrvName_Change()
    Call ChangeDispDrvName
End Sub

Private Sub cmbDispDrvName_KeyPress(KeyAscii As Integer)

    Dim CB         As Long

    Dim FindString As String
    
    If KeyAscii < 32 Then GoTo EndSub
    
    If cmbDispDrvName.SelLength = 0 Then
        FindString = cmbDispDrvName.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(cmbDispDrvName.Text, cmbDispDrvName.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage cmbDispDrvName.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(cmbDispDrvName.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        cmbDispDrvName.ListIndex = CB
        cmbDispDrvName.SelStart = Len(FindString)
        cmbDispDrvName.SelLength = Len(cmbDispDrvName.Text) - cmbDispDrvName.SelStart
    End If

    KeyAscii = 0
    Call ChangeDispDrvName

EndSub:

End Sub

Private Sub lstOrdWait_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If Button = 2 And Me.lstOrdWait.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrntOrdNow)
End Sub

Private Sub rcFormOrdNow_Click()
    Call PrintLVPic(Me.lstOrdWait, 2, True, True, True, uniOrds)
End Sub

Private Sub rcFormOrdNowExp_Click()
    Call ExportToExcel(lstOrdWait)
End Sub

Private Sub lstOrdWait_Click()
    Call ListOrdWaitClick
End Sub

Private Sub lstMixReady_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If Button = 2 And Me.lstMixReady.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrnt)
End Sub

Private Sub rcForm1_Click()
    PrintRightBut = True
    Call FillForm1
End Sub

Private Sub rcForm2_Click()
    PrintRightBut = True
    Call FillForm2
End Sub

Private Sub rcForm3_Click()
    PrintRightBut = True
    Call FillForm3
End Sub
'------------------------------------------------------

Private Sub btnOrders_Click()
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = True
    frOrders.Visible = True
    
    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = False
    frAbout.Visible = False
    
    Call OpenOrders
End Sub

Private Sub btnClearOrd_Click()
    Call ClearOrdBut
End Sub

Private Sub btnSvNwOrd_Click()
    Call SvNwOrdBut
End Sub

Private Sub btnDelOrd_Click()

    If frmLogin.RootUser = True Then Call DelOrdBut
End Sub

Private Sub txtOrdQuant_GotFocus()
    txtOrdQuant.SelStart = 0
    txtOrdQuant.SelLength = Len(txtOrdQuant.Text)

    If InStr(txtOrdQuant.Text, DecSep) <> 0 Then
        PointLook2 = True
    Else
        PointLook2 = False
    End If

End Sub

Private Sub txtOrdQuant_Change()

    If InStr(txtOrdQuant.Text, DecSep) <> 0 Then
        PointLook2 = True
    Else
        PointLook2 = False
    End If

End Sub

Private Sub txtOrdQuant_KeyPress(KeyAscii As Integer)

    If InStr(txtOrdQuant.Text, DecSep) <> 0 Then
        PointLook2 = True
    Else
        PointLook2 = False
    End If

    If txtOrdQuant.SelLength = Len(txtOrdQuant.Text) Then
        PointLook2 = False
    Else
    End If

    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If

    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook2 = True Then
        KeyAscii = 0
    Else

        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook2 = True
        Else
        End If
    End If

End Sub

Private Sub cmbOrdRec_Change()
    Call ChangeOrdRec
End Sub

Private Sub cmbOrdRec_Click()
    Call ChangeOrdRec
End Sub

Private Sub cmbOrdClnt_Change()
    Call ChangeOrdClnt
End Sub

Private Sub cmbOrdClnt_Click()
    Call ChangeOrdClnt
End Sub

Private Sub cmbOrdClntObj_KeyPress(KeyAscii As Integer)

    Dim CB         As Long

    Dim FindString As String
    
    If KeyAscii < 32 Then GoTo EndSub
    
    If cmbOrdClntObj.SelLength = 0 Then
        FindString = cmbOrdClntObj.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(cmbOrdClntObj.Text, cmbOrdClntObj.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage cmbOrdClntObj.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(cmbOrdClntObj.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        cmbOrdClntObj.ListIndex = CB
        cmbOrdClntObj.SelStart = Len(FindString)
        cmbOrdClntObj.SelLength = Len(cmbOrdClntObj.Text) - cmbOrdClntObj.SelStart
    End If

    KeyAscii = 0

EndSub:

End Sub

Private Sub cmbOrdClntName_Change()
    Call ChangeOrdClntName
End Sub

Private Sub cmbOrdClntName_Click()
    Call ChangeOrdClntName
End Sub

Private Sub cmbOrdClntName_KeyPress(KeyAscii As Integer)

    Dim CB         As Long

    Dim FindString As String
    
    If KeyAscii < 32 Then GoTo EndSub
    
    If cmbOrdClntName.SelLength = 0 Then
        FindString = cmbOrdClntName.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(cmbOrdClntName.Text, cmbOrdClntName.SelStart) & Chr$(KeyAscii)
    End If
    
    SendMessage cmbOrdClntName.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&

    CB = SendMessage(cmbOrdClntName.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If CB <> CB_ERR Then
        cmbOrdClntName.ListIndex = CB
        cmbOrdClntName.SelStart = Len(FindString)
        cmbOrdClntName.SelLength = Len(cmbOrdClntName.Text) - cmbOrdClntName.SelStart
    End If

    KeyAscii = 0

EndSub:

End Sub

Private Sub lstOrd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And Me.lstOrd.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrntOrd)
End Sub

Private Sub rcFormOrd_Click()
    Call PrintLVPic(Me.lstOrd, 2, True, True, True, uniOrds)
End Sub

Private Sub rcFormOrdExp_Click()
    Call ExportToExcel(lstOrd)
End Sub

Private Sub lstOrd_Click()
    Call ListOrdClick
End Sub
'---------------------------------------------------------

Private Sub btnRecepies_Click()
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = False
    frOrders.Visible = False
    
    frRecepies.Enabled = True
    frRecepies.Visible = True
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = False
    frAbout.Visible = False
    
    Call OpenRecepies
End Sub

Private Sub btnClearRec_Click()
    Call ClearRecBut
End Sub

Private Sub btnSvNwRec_Click()
    FlagButRec = 1

    If rDeactNRPass = 0 And (frmLogin.AdminSuccess = False Or frmLogin.RootUser = False) Then

        Dim PassCheck As String

        PassCheck = GetSetting(PlacePass, PlacePassAdd, "Lab", ErrRes)

        If PassCheck <> ErrRes Then
            frmPass.Show
        Else
            Call SvNwRecBut
        End If

    Else
        Call SvNwRecBut
    End If

End Sub

Private Sub btnDelRec_Click()
    FlagButRec = 2

    If rDeactDRPass = 0 And (frmLogin.AdminSuccess = False Or frmLogin.RootUser = False) Then

        Dim PassCheck As String

        PassCheck = GetSetting(PlacePass, PlacePassAdd, "Lab", ErrRes)

        If PassCheck <> ErrRes Then
            frmPass.Show
        Else
        End If

    Else
        Call DelRecBut
    End If

End Sub

Private Sub btnShowRec_Click()
    setRecepies.Show
End Sub

Private Sub txtRec_GotFocus()
    txtRec.SelStart = 0
    txtRec.SelLength = Len(txtRec.Text)
End Sub

Private Sub txtRec_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtNameRec_GotFocus()
    txtNameRec.SelStart = 0
    txtNameRec.SelLength = Len(txtNameRec.Text)
End Sub

Private Sub txtNameRec_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+
        
            'пропуска кода
        Case 46 '.
        
            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtTypeRec_GotFocus()
    txtNameRec.SelStart = 0
    txtNameRec.SelLength = Len(txtTypeRec.Text)
End Sub

Private Sub txtTypeRec_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtClassRec_GotFocus()
    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRec.Text)
End Sub

Private Sub txtClassRec_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+
            'пропуска кода
        Case 46 '.
            'пропуска кода
            
        Case 92
        
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtClassRecK_GotFocus()
    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecK.Text)
End Sub

Private Sub txtClassRecK_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtClassRecV_GotFocus()
    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecV.Text)
End Sub

Private Sub txtClassRecV_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtClassRecH_GotFocus()
    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecH.Text)
End Sub

Private Sub txtClassRecH_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtClassRecP_GotFocus()
    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecP.Text)
End Sub

Private Sub txtClassRecP_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+
        
        Case 46
            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtEDMRec_GotFocus()
    txtEDMRec.SelStart = 0
    txtEDMRec.SelLength = Len(txtEDMRec.Text)
End Sub

Private Sub txtEDMRec_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 '-

            'пропуска кода
        Case 45 '+

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtTimePourRec_GotFocus()
    txtTimePourRec.SelStart = 0
    txtTimePourRec.SelLength = Len(txtTimePourRec.Text)
End Sub

Private Sub txtTimePourRec_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtTimeMixRec_GotFocus()
    txtTimeMixRec.SelStart = 0
    txtTimeMixRec.SelLength = Len(txtTimeMixRec.Text)
End Sub

Private Sub txtTimeMixRec_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtRec1_GotFocus(Index As Integer)
    txtRec1(Index).SelStart = 0
    txtRec1(Index).SelLength = Len(txtRec1(Index).Text)
End Sub

Private Sub txtRec1_KeyPress(Index As Integer, KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtRec2_GotFocus(Index As Integer)
    txtRec2(Index).SelStart = 0
    txtRec2(Index).SelLength = Len(txtRec2(Index).Text)
End Sub

Private Sub txtRec2_KeyPress(Index As Integer, KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtRec3_GotFocus(Index As Integer)
    txtRec3(Index).SelStart = 0
    txtRec3(Index).SelLength = Len(txtRec3(Index).Text)
End Sub

Private Sub txtRec3_KeyPress(Index As Integer, KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtRec4_GotFocus(Index As Integer)
    txtRec4(Index).SelStart = 0
    txtRec4(Index).SelLength = Len(txtRec4(Index).Text)

    If InStr(txtRec4(Index).Text, DecSep) <> 0 Then
        PointLook3 = True
    Else
        PointLook3 = False
    End If

End Sub

Private Sub txtRec4_Change(Index As Integer)

    If InStr(txtRec4(Index).Text, DecSep) <> 0 Then
        PointLook3 = True
    Else
        PointLook3 = False
    End If

End Sub

Private Sub txtRec4_KeyPress(Index As Integer, KeyAscii As Integer)

    If InStr(txtRec4(Index).Text, DecSep) <> 0 Then
        PointLook3 = True
    Else
        PointLook3 = False
    End If

    If txtRec4(Index).SelLength = Len(txtRec4(Index).Text) Then
        PointLook3 = False
    Else
    End If

    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If

    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook3 = True Then
        KeyAscii = 0
    Else

        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook3 = True
        Else
        End If
    End If

End Sub

Private Sub lstRec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And Me.lstRec.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrntRec)
End Sub

Private Sub PrintRec_Click()
    Call PrintLVPic(Me.lstRec, 2, True, True, True, uniRecs)
End Sub

Private Sub ExpRec_Click()
    Call ExportToExcel(lstRec)
End Sub

Private Sub lstRec_Click()
    Call ListRecClick
End Sub
'-----------------------------------------------------

Private Sub btnClients_Click()
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = False
    frOrders.Visible = False
    
    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = True
    frClients.Visible = True
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = False
    frAbout.Visible = False
    
    Call OpenClients
End Sub

Private Sub btnClearClnt_Click()
    Call ClearClntBut
End Sub

Private Sub btnSvNwClnt_Click()
    Call SvNwClntBut
End Sub

Private Sub btnDelClnt_Click()
    Call DelClntBut
End Sub

Private Sub btnShowClnt_Click()
    setClients.Show
End Sub

Private Sub btnDelObj_Click()
    Call DelObjBut
End Sub

Private Sub btnObjects_Click()
    frmAddObj.Show
End Sub

Private Sub btnShowObj_Click()
    setObjects.Show
End Sub

Private Sub txtClnt_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtClnt_GotFocus()
    txtClnt.SelStart = 0
    txtClnt.SelLength = Len(txtClnt.Text)
End Sub

Private Sub txtNameClnt_GotFocus()
    txtNameClnt.SelStart = 0
    txtNameClnt.SelLength = Len(txtNameClnt.Text)
End Sub

Private Sub txtNameClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtBGClnt_GotFocus()
    txtBGClnt.SelStart = 0
    txtBGClnt.SelLength = Len(txtBGClnt.Text)
End Sub

Private Sub txtBGClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква

        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtMOLClnt_GotFocus()
    txtMOLClnt.SelStart = 0
    txtMOLClnt.SelLength = Len(txtMOLClnt.Text)
End Sub

Private Sub txtMOLClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 45 To 46 '- .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtAddClnt_GotFocus()
    txtAddClnt.SelStart = 0
    txtAddClnt.SelLength = Len(txtAddClnt.Text)
End Sub

Private Sub txtAddClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtTelClnt_GotFocus()
    txtTelClnt.SelStart = 0
    txtTelClnt.SelLength = Len(txtTelClnt.Text)
End Sub

Private Sub txtTelClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 48 To 57, 8 ' 0-9 и bksp

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub lstClnt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And Me.lstClnt.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrntClnt)
End Sub

Private Sub PrintClnt_Click()
    Call PrintLVPic(Me.lstClnt, 2, True, True, True, uniClnts)
End Sub

Private Sub ExpClnt_Click()
    Call ExportToExcel(lstClnt)
End Sub

Private Sub lstClnt_Click()
    Call ListClntClick
End Sub
'------------------------------------------------------

Private Sub btnDrivers_Click()
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = False
    frOrders.Visible = False

    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = True
    frDrivers.Visible = True
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = False
    frAbout.Visible = False
    
    Call OpenDrivers
End Sub

Private Sub btnClearDrv_Click()
    Call ClearDrvBut
End Sub

Private Sub btnSvNwDrv_Click()
    Call SvNwDrvBut
End Sub

Private Sub btnDelDrv_Click()
    Call DelDrvBut
End Sub

Private Sub btnShowDrv_Click()
    setDrivers.Show
End Sub

Private Sub txtDrv_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtDrv_GotFocus()
    txtDrv.SelStart = 0
    txtDrv.SelLength = Len(txtDrv.Text)
End Sub

Private Sub txtNameDrv_GotFocus()
    txtNameDrv.SelStart = 0
    txtNameDrv.SelLength = Len(txtNameDrv.Text)
End Sub

Private Sub txtNameDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtRegDrv_GotFocus()
    txtRegDrv.SelStart = 0
    txtRegDrv.SelLength = Len(txtRegDrv.Text)
End Sub

Private Sub txtRegDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква

        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква

        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtCapDrv_GotFocus()
    txtCapDrv.SelStart = 0
    txtCapDrv.SelLength = Len(txtCapDrv.Text)

    If InStr(txtCapDrv.Text, DecSep) <> 0 Then
        PointLook4 = True
    Else
        PointLook4 = False
    End If

End Sub

Private Sub txtCapDrv_Change()

    If InStr(txtCapDrv.Text, DecSep) <> 0 Then
        PointLook4 = True
    Else
        PointLook4 = False
    End If

End Sub

Private Sub txtCapDrv_KeyPress(KeyAscii As Integer)

    If InStr(txtCapDrv.Text, DecSep) <> 0 Then
        PointLook4 = True
    Else
        PointLook4 = False
    End If

    If txtCapDrv.SelLength = Len(txtCapDrv.Text) Then
        PointLook4 = False
    Else
    End If

    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If

    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook4 = True Then
        KeyAscii = 0
    Else

        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook4 = True
        Else
        End If
    End If

End Sub

Private Sub txtModDrv_GotFocus()
    txtModDrv.SelStart = 0
    txtModDrv.SelLength = Len(txtModDrv.Text)
End Sub

Private Sub txtModDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtTelDrv_GotFocus()
    txtTelDrv.SelStart = 0
    txtTelDrv.SelLength = Len(txtTelDrv.Text)
End Sub

Private Sub txtTelDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 48 To 57, 8 ' 0-9 и bksp

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtNoteDrv_GotFocus()
    txtNoteDrv.SelStart = 0
    txtNoteDrv.SelLength = Len(txtNoteDrv.Text)
End Sub

Private Sub txtNoteDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub lstDrv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And Me.lstDrv.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrntDrv)
End Sub

Private Sub PrintDrv_Click()
    Call PrintLVPic(Me.lstDrv, 1, True, True, True, uniDrvs)
End Sub

Private Sub ExpDrv_Click()
    Call ExportToExcel(lstDrv)
End Sub

Private Sub lstDrv_Click()
    Call ListDrvClick
End Sub
'-------------------------------------------------

Private Sub btnSuppliers_Click()
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = False
    frOrders.Visible = False
    
    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = True
    frSuppliers.Visible = True
    
    frMaterials.Enabled = False
    frMaterials.Visible = False
    
    frAbout.Enabled = False
    frAbout.Visible = False
    
    Call OpenSuppliers
End Sub

Private Sub btnClearSup_Click()
    Call ClearSupBut
End Sub

Private Sub btnSvNwSup_Click()
    Call SvNwSupBut
End Sub

Private Sub btnDelSup_Click()
    Call DelSupBut
End Sub

Private Sub btnShowSup_Click()
    setSuppliers.Show
End Sub

Private Sub txtSup_KeyPress(KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtSup_GotFocus()
    txtSup.SelStart = 0
    txtSup.SelLength = Len(txtSup.Text)
End Sub

Private Sub txtNameSup_GotFocus()
    txtNameSup.SelStart = 0
    txtNameSup.SelLength = Len(txtNameSup.Text)
End Sub

Private Sub txtNameSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtBGSup_GotFocus()
    txtBGSup.SelStart = 0
    txtBGSup.SelLength = Len(txtBGSup.Text)
End Sub

Private Sub txtBGSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква

        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtMOLSup_GotFocus()
    txtMOLSup.SelStart = 0
    txtMOLSup.SelLength = Len(txtMOLSup.Text)
End Sub

Private Sub txtMOLSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 45 To 46 '- .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtAddSup_GotFocus()
    txtAddSup.SelStart = 0
    txtAddSup.SelLength = Len(txtAddSup.Text)
End Sub

Private Sub txtAddSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtTelSup_GotFocus()
    txtTelSup.SelStart = 0
    txtTelSup.SelLength = Len(txtTelSup.Text)
End Sub

Private Sub txtTelSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 48 To 57, 8 ' 0-9 и bksp

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub txtNoteSup_GotFocus()
    txtNoteSup.SelStart = 0
    txtNoteSup.SelLength = Len(txtNoteSup.Text)
End Sub

Private Sub txtNoteSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub lstSup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And Me.lstSup.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrntSup)
End Sub

Private Sub PrintSup_Click()
    Call PrintLVPic(Me.lstSup, 2, True, True, True, uniSups)
End Sub

Private Sub ExpSup_Click()
    Call ExportToExcel(lstSup)
End Sub

Private Sub lstSup_Click()
    Call ListSupClick
End Sub
'-----------------------------------------------------

Private Sub btnRevision_Click()
    If frmLogin.AdminSuccess = True Or frmLogin.RootUser = True Then
        frmRevision.Show
    Else
        frmDailyBalance.Show
        If ErrDaily = True Then Unload frmDailyBalance
    End If
End Sub

Private Sub btnMaterials_Click()
    frDisp.Enabled = False
    frDisp.Visible = False
    
    frOrders.Enabled = False
    frOrders.Visible = False
    
    frRecepies.Enabled = False
    frRecepies.Visible = False
    
    frClients.Enabled = False
    frClients.Visible = False
    
    frDrivers.Enabled = False
    frDrivers.Visible = False
    
    frSuppliers.Enabled = False
    frSuppliers.Visible = False
    
    frMaterials.Enabled = True
    frMaterials.Visible = True
    
    frAbout.Enabled = False
    frAbout.Visible = False
    
    Call OpenMaterials
End Sub

Private Sub btnClearMat_Click()
    Call ClearMatBut
End Sub

Private Sub btnSvNwMat_Click()
    Call SvNwMatBut
End Sub

Private Sub btnDelMat_Click()
    Call DelMatBut
End Sub

Private Sub btnSvExp_Click()
    frmSvExpen.Show
End Sub

Private Sub txtMatName_GotFocus()
    txtMatName.SelStart = 0
    txtMatName.SelLength = Len(txtMatName.Text)
End Sub

Private Sub txtMatName_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 32 'интервал

            'пропуска кода
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp

            'пропуска кода
        Case 97 To 122 'латиница a-z

            'пропуска кода
        Case 192 To 223 'кирилица А-Я

            'пропуска кода
        Case 224 To 255 'кирилица а-я

            'пропуска кода
        Case 43 To 46 '+ , - .

            'пропуска кода
        Case Else
            'всички останали
            KeyAscii = 0 ' код ascii = 0
    End Select

End Sub

Private Sub btnAddMatDlvr_Click()
    frmAddDlvr.Show
End Sub

Private Sub lstMat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And Me.lstMat.ListItems.count > 0 Then Call Me.PopupMenu(rcMnuPrntMat)
End Sub

Private Sub PrintMat_Click()
    Call PrintLVPic(Me.lstMat, 2, True, True, True, uniMats)
End Sub

Private Sub Expmat_Click()
    Call ExportToExcel(lstMat)
End Sub

Private Sub lstMat_Click()
    Call ListMatClick
End Sub

Private Sub cmbMatType_Change()
    Call LoadMat
End Sub

Private Sub cmbMatType_Click()
    Call LoadMat
End Sub

Private Sub btnNotes_Click()
    frmNotes.Show
End Sub

Private Sub btnAdminPanel_Click()
    AdminPanel.Show
End Sub

Private Sub btnExit_Click()

    If Me.indReq.Caption <> statReqStarted Or Me.indAvaria.Caption = statAvaria Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim msgNow   As String

    Dim response As Integer

    If Me.indReq.Caption <> statReqStarted Or Me.indAvaria.Caption = statAvaria Then
        If Val(frmOPC.RecInput(2).Text) > 0 And frmOPC.Stat(0).Text = "True" Then
            msgNow = MsgExpWait & vbCrLf & MsgNoResOnExit & vbCrLf & MsgClose
        Else
            msgNow = MsgClose
        End If
    
        If DontAskExit = True Then
            response = vbYes
        Else
            response = MsgBox(msgNow, vbQuestion Or vbYesNo, UniExit)
        End If
        
        If response = vbYes Then
    
            frDisp.Enabled = False
            frDisp.Visible = False
    
            frOrders.Enabled = False
            frOrders.Visible = False
        
            frRecepies.Enabled = False
            frRecepies.Visible = False
        
            frClients.Enabled = False
            frClients.Visible = False
    
            frDrivers.Enabled = False
            frDrivers.Visible = False
    
            frSuppliers.Enabled = False
            frSuppliers.Visible = False
    
            frMaterials.Enabled = False
            frMaterials.Visible = False
    
            frAbout.Enabled = False
            frAbout.Visible = False
    
            Unload AdminPanel
            Unload DispConfirm
    
            '------------------------------Start PostgreSQL--------------------------------------
            Dim cn      As New ADODB.Connection

            Dim rs      As New Recordset

            Dim counter As Long

            Dim comm    As String

            Dim comEdit As String

            cn.Open ConStr
            If MachineNumber = 1 Then comm = "SELECT * FROM entry_log ORDER BY log_num DESC LIMIT 1"
            If MachineNumber = 2 Then comm = "SELECT * FROM entry_log2 ORDER BY log_num DESC LIMIT 1"
            Set rs = cn.Execute(comm)
    
            If Not rs.BOF And Not rs.EOF Then
                counter = Val(rs!log_num)
            Else
                counter = 1
            End If
    
            If MachineNumber = 1 Then comEdit = "UPDATE entry_log SET log_exit_date ='" & Format(Now, "DD-MM-YYYY") & "', log_exit ='" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE log_num = " & counter & ""
            If MachineNumber = 2 Then comEdit = "UPDATE entry_log2 SET log_exit_date ='" & Format(Now, "DD-MM-YYYY") & "', log_exit ='" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE log_num = " & counter & ""
            Set rs = cn.Execute(comEdit)
    
            rs.Close
            Set rs = Nothing
            cn.Close 'затваряме връзката
            Set cn = Nothing
    
            '------------------------------End PostgreSQL----------------------------------------
    
    'прекъсване на връзките със OPC Server
        Dim ip As Integer
        Dim ist As Integer
        Dim ic As Integer
        Dim ir As Integer
        Dim irs As Integer
        Dim RemItemSrvErrPanels() As Long
        Dim RemItemSrvErrConfig() As Long
        Dim RemItemSrvErrStatus() As Long
        Dim RemItemSrvErrRec() As Long
        Dim RemItemSrvErrResults() As Long
        Dim RemItemSrvHandlesPanels(4) As Long
        Dim RemItemSrvHandlesConfig(1) As Long
        Dim RemItemSrvHandlesStatus(5) As Long
        Dim RemItemSrvHandlesRec(39) As Long
        Dim RemItemSrvHandlesResults(21) As Long
    
        If Not OPCItemCollPanels Is Nothing Then
            For ip = 1 To ItemCountPanels
                If ItemSrvHandlesPanels(ip) <> 0 Then
                    RemItemSrvHandlesPanels(ip) = ItemSrvHandlesPanels(ip)
                End If
                ItemSrvHandlesPanels(ip) = 0
            Next ip
            OPCItemCollPanels.Remove ItemCountPanels, _
            RemItemSrvHandlesPanels, RemItemSrvErrPanels
            Set OPCItemCollPanels = Nothing
        End If

        If Not OPCItemCollConfig Is Nothing Then
            For ic = 1 To ItemCountConfig
                If ItemSrvHandlesConfig(ic) <> 0 Then
                    RemItemSrvHandlesConfig(ic) = ItemSrvHandlesConfig(ic)
                End If
                ItemSrvHandlesConfig(ic) = 0
            Next ic
            OPCItemCollConfig.Remove ItemCountConfig, _
            RemItemSrvHandlesConfig, RemItemSrvErrConfig
            Set OPCItemCollConfig = Nothing
        End If

        If Not OPCItemCollStatus Is Nothing Then
            For ist = 1 To ItemCountStatus
                If ItemSrvHandlesStatus(ist) <> 0 Then
                    RemItemSrvHandlesStatus(ist) = ItemSrvHandlesStatus(ist)
                End If
                ItemSrvHandlesStatus(ist) = 0
            Next ist
            OPCItemCollStatus.Remove ItemCountStatus, _
            RemItemSrvHandlesStatus, RemItemSrvErrStatus
            Set OPCItemCollStatus = Nothing
        End If

        If Not OPCItemCollRec Is Nothing Then
            For ir = 1 To ItemCountRec
                If ItemSrvHandlesRec(ir) <> 0 Then
                    RemItemSrvHandlesRec(ir) = ItemSrvHandlesRec(ir)
                End If
                ItemSrvHandlesRec(ir) = 0
            Next ir
            OPCItemCollRec.Remove ItemCountRec, _
            RemItemSrvHandlesRec, RemItemSrvErrRec
            Set OPCItemCollRec = Nothing
        End If

        If Not OPCItemCollResults Is Nothing Then
            For irs = 1 To ItemCountResults
                If ItemSrvHandlesResults(irs) <> 0 Then
                    RemItemSrvHandlesResults(irs) = ItemSrvHandlesResults(irs)
                End If
                ItemSrvHandlesResults(irs) = 0
            Next irs
            OPCItemCollResults.Remove ItemCountResults, _
            RemItemSrvHandlesResults, RemItemSrvErrResults
            Set OPCItemCollResults = Nothing
        End If

        If Not ConGroupPanels Is Nothing Then
            Set ConGroupPanels = Nothing
        End If
        If Not ConGroupConfig Is Nothing Then
            Set ConGroupConfig = Nothing
        End If
        If Not ConGroupStatus Is Nothing Then
            Set ConGroupStatus = Nothing
        End If
        If Not ConGroupRec Is Nothing Then
            Set ConGroupRec = Nothing
        End If
        If Not ConGroupResults Is Nothing Then
            Set ConGroupResults = Nothing
        End If

        Set ConSrvGroup = Nothing
    
        If Not ConOPCServer Is Nothing Then
            ConOPCServer.Disconnect
            Set ConOPCServer = Nothing
        End If
    
    
            Dim hw As Long
            Dim retval As Long
            Dim SwWin As String
            If MachineNumber = 1 Then SwWin = "Машина 2 - ТИП-Панел v1.1"
            If MachineNumber = 2 Then SwWin = "Машина 1 - ТИП-Панел v1.1"
            hw = FindWindow(vbNullString, SwWin)
    
            If DontAskExit = True Then
                End
            Else
            End If
            
            Unload Me
            
            frmStart.Started = True
            frmStart.Show
            
            If hw <> 0 Then
                retval = ShowWindow(hw, 9)
                frmStart.BtnExit = True
            End If
        Else
            msgNow = MsgClose
            Cancel = 1
                        
            Exit Sub

        End If

    Else
        Cancel = 1
    End If

End Sub

