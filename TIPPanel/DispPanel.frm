VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DispPanel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDispPanelCap"
   ClientHeight    =   12375
   ClientLeft      =   150
   ClientTop       =   -7185
   ClientWidth     =   16245
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
   Icon            =   "DispPanel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12375
   ScaleMode       =   0  'User
   ScaleWidth      =   16245
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
      TabIndex        =   271
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
      TabIndex        =   261
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
      TabIndex        =   260
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
      Picture         =   "DispPanel.frx":08CA
      ScaleHeight     =   435
      ScaleWidth      =   465
      TabIndex        =   243
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
      TabIndex        =   234
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
      TabIndex        =   203
      Top             =   60240
      Width           =   14175
      Begin RichTextLib.RichTextBox rtxtLicAgr 
         Height          =   6015
         Left            =   360
         TabIndex        =   241
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
         TextRTF         =   $"DispPanel.frx":1194
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
         TabIndex        =   269
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
         TabIndex        =   268
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
         TabIndex        =   267
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
         TabIndex        =   266
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
         TabIndex        =   265
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
         TabIndex        =   264
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
         TabIndex        =   263
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
         TabIndex        =   262
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
         TabIndex        =   258
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
         TabIndex        =   257
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
         TabIndex        =   256
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
      TabIndex        =   183
      Top             =   53280
      Width           =   14175
      Begin VB.TextBox txtMatHum 
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
         Left            =   10200
         TabIndex        =   272
         Top             =   960
         Width           =   975
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
         TabIndex        =   191
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
         Index           =   4
         Left            =   11160
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
         Index           =   3
         Left            =   9840
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
         Index           =   2
         Left            =   8520
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
         Index           =   1
         Left            =   7200
         TabIndex        =   197
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
         TabIndex        =   196
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
         TabIndex        =   187
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
         TabIndex        =   193
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
         TabIndex        =   194
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
         TabIndex        =   195
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
         TabIndex        =   186
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
         TabIndex        =   189
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
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin MSComctlLib.ListView lstMat 
         Height          =   4455
         Left            =   240
         TabIndex        =   202
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
      Begin VB.Label lblMatHum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblMatHum"
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
         Left            =   7560
         TabIndex        =   273
         Top             =   1080
         Width           =   2535
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
         TabIndex        =   192
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
         TabIndex        =   190
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
         TabIndex        =   188
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
         TabIndex        =   185
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   253
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
         TabIndex        =   242
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
         TabIndex        =   236
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   229
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
         ItemData        =   "DispPanel.frx":D224
         Left            =   8040
         List            =   "DispPanel.frx":D226
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
         Left            =   4920
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
         BuddyDispid     =   196653
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
         Left            =   4920
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
         Left            =   1680
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
         Left            =   1680
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
         Left            =   1680
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
         Left            =   1680
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
         TabIndex        =   54
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
         TabIndex        =   237
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
         TabIndex        =   235
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
         Left            =   4800
         TabIndex        =   231
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
         Left            =   7200
         TabIndex        =   230
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
         TabIndex        =   228
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
         TabIndex        =   226
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
         Left            =   240
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
         Left            =   3600
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
         Left            =   3840
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
         Left            =   3840
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
         Width           =   1455
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
         Width           =   1455
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
         Left            =   240
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
         Left            =   240
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
      TabIndex        =   55
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
         ItemData        =   "DispPanel.frx":D228
         Left            =   6360
         List            =   "DispPanel.frx":D22A
         TabIndex        =   255
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
         ItemData        =   "DispPanel.frx":D22C
         Left            =   6360
         List            =   "DispPanel.frx":D22E
         TabIndex        =   252
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   71
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
         TabIndex        =   66
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
         TabIndex        =   56
         Top             =   480
         Width           =   975
      End
      Begin MSComCtl2.DTPicker nowOrdDate 
         Height          =   375
         Left            =   14280
         TabIndex        =   60
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
         Format          =   122224643
         CurrentDate     =   41426.3333333333
         MaxDate         =   44196
         MinDate         =   41426
      End
      Begin MSComctlLib.ListView lstOrd 
         Height          =   4455
         Left            =   240
         TabIndex        =   77
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
         TabIndex        =   69
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
         Format          =   122224643
         CurrentDate     =   41487.3333333333
         MaxDate         =   45291
         MinDate         =   41487
      End
      Begin MSComCtl2.DTPicker queOrdTime 
         Height          =   375
         Left            =   12120
         TabIndex        =   70
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
         Format          =   122224642
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
         TabIndex        =   227
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   73
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
         TabIndex        =   68
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
         TabIndex        =   62
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
         TabIndex        =   72
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
         TabIndex        =   67
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
         TabIndex        =   65
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
         TabIndex        =   61
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
      TabIndex        =   164
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
         TabIndex        =   240
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
         TabIndex        =   178
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
         TabIndex        =   174
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
         TabIndex        =   177
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
         TabIndex        =   173
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
         TabIndex        =   181
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
         TabIndex        =   172
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
         TabIndex        =   167
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
         TabIndex        =   165
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
         TabIndex        =   168
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
         TabIndex        =   169
         Top             =   840
         Width           =   4095
      End
      Begin MSComctlLib.ListView lstSup 
         Height          =   4575
         Left            =   240
         TabIndex        =   182
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
         TabIndex        =   180
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
         TabIndex        =   166
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
         TabIndex        =   170
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
         TabIndex        =   175
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
         TabIndex        =   179
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
         TabIndex        =   171
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
         TabIndex        =   176
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
      TabIndex        =   145
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
         TabIndex        =   239
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   160
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
         TabIndex        =   153
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
         TabIndex        =   149
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
         TabIndex        =   150
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
         TabIndex        =   159
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
         TabIndex        =   146
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
         TabIndex        =   162
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
         TabIndex        =   147
         Top             =   480
         Width           =   735
      End
      Begin MSComctlLib.ListView lstDrv 
         Height          =   4575
         Left            =   240
         TabIndex        =   163
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
         TabIndex        =   225
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
         TabIndex        =   161
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
         TabIndex        =   156
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
         TabIndex        =   151
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
         TabIndex        =   157
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
         TabIndex        =   152
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
         TabIndex        =   158
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
         TabIndex        =   148
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
      TabIndex        =   128
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
         TabIndex        =   254
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
         TabIndex        =   251
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
         TabIndex        =   250
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
         TabIndex        =   135
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
         TabIndex        =   132
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
         TabIndex        =   238
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
         TabIndex        =   142
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
         TabIndex        =   139
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
         TabIndex        =   129
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
         TabIndex        =   134
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
         TabIndex        =   141
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
         TabIndex        =   131
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
         TabIndex        =   138
         Top             =   960
         Width           =   1455
      End
      Begin MSComctlLib.ListView lstClnt 
         Height          =   3975
         Left            =   240
         TabIndex        =   144
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
         TabIndex        =   249
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
         TabIndex        =   130
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
         TabIndex        =   136
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
         TabIndex        =   140
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
         TabIndex        =   143
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
         TabIndex        =   133
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
         TabIndex        =   137
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
      TabIndex        =   78
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
         ItemData        =   "DispPanel.frx":D230
         Left            =   7320
         List            =   "DispPanel.frx":D232
         Style           =   2  'Dropdown List
         TabIndex        =   248
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
         TabIndex        =   247
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
         TabIndex        =   246
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
         TabIndex        =   245
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
         ItemData        =   "DispPanel.frx":D234
         Left            =   3960
         List            =   "DispPanel.frx":D236
         Style           =   2  'Dropdown List
         TabIndex        =   244
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
         TabIndex        =   107
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
         TabIndex        =   217
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
         TabIndex        =   205
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
         TabIndex        =   206
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
         TabIndex        =   233
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
         TabIndex        =   213
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
         TabIndex        =   212
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
         TabIndex        =   211
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
         TabIndex        =   207
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
         TabIndex        =   122
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
         ItemData        =   "DispPanel.frx":D238
         Left            =   7320
         List            =   "DispPanel.frx":D23A
         Style           =   2  'Dropdown List
         TabIndex        =   121
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
         ItemData        =   "DispPanel.frx":D23C
         Left            =   10680
         List            =   "DispPanel.frx":D23E
         Style           =   2  'Dropdown List
         TabIndex        =   88
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
         TabIndex        =   89
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
         ItemData        =   "DispPanel.frx":D240
         Left            =   10680
         List            =   "DispPanel.frx":D242
         Style           =   2  'Dropdown List
         TabIndex        =   96
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
         ItemData        =   "DispPanel.frx":D244
         Left            =   10680
         List            =   "DispPanel.frx":D246
         Style           =   2  'Dropdown List
         TabIndex        =   104
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
         ItemData        =   "DispPanel.frx":D248
         Left            =   10680
         List            =   "DispPanel.frx":D24A
         Style           =   2  'Dropdown List
         TabIndex        =   112
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
         ItemData        =   "DispPanel.frx":D24C
         Left            =   10680
         List            =   "DispPanel.frx":D24E
         Style           =   2  'Dropdown List
         TabIndex        =   116
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
         ItemData        =   "DispPanel.frx":D250
         Left            =   10680
         List            =   "DispPanel.frx":D252
         Style           =   2  'Dropdown List
         TabIndex        =   119
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
         TabIndex        =   97
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
         TabIndex        =   105
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
         TabIndex        =   113
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
         TabIndex        =   117
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
         TabIndex        =   120
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
         TabIndex        =   87
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
         ItemData        =   "DispPanel.frx":D254
         Left            =   7320
         List            =   "DispPanel.frx":D256
         Style           =   2  'Dropdown List
         TabIndex        =   86
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
         ItemData        =   "DispPanel.frx":D258
         Left            =   7320
         List            =   "DispPanel.frx":D25A
         Style           =   2  'Dropdown List
         TabIndex        =   94
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
         ItemData        =   "DispPanel.frx":D25C
         Left            =   7320
         List            =   "DispPanel.frx":D25E
         Style           =   2  'Dropdown List
         TabIndex        =   102
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
         ItemData        =   "DispPanel.frx":D260
         Left            =   7320
         List            =   "DispPanel.frx":D262
         Style           =   2  'Dropdown List
         TabIndex        =   110
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
         TabIndex        =   95
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
         TabIndex        =   103
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
         TabIndex        =   111
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
         TabIndex        =   85
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
         ItemData        =   "DispPanel.frx":D264
         Left            =   3960
         List            =   "DispPanel.frx":D266
         Style           =   2  'Dropdown List
         TabIndex        =   84
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
         ItemData        =   "DispPanel.frx":D268
         Left            =   3960
         List            =   "DispPanel.frx":D26A
         Style           =   2  'Dropdown List
         TabIndex        =   92
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
         ItemData        =   "DispPanel.frx":D26C
         Left            =   3960
         List            =   "DispPanel.frx":D26E
         Style           =   2  'Dropdown List
         TabIndex        =   100
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
         ItemData        =   "DispPanel.frx":D270
         Left            =   3960
         List            =   "DispPanel.frx":D272
         Style           =   2  'Dropdown List
         TabIndex        =   108
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
         ItemData        =   "DispPanel.frx":D274
         Left            =   3960
         List            =   "DispPanel.frx":D276
         Style           =   2  'Dropdown List
         TabIndex        =   114
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
         TabIndex        =   93
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
         TabIndex        =   101
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
         TabIndex        =   109
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
         TabIndex        =   115
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
         TabIndex        =   81
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
         TabIndex        =   118
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   83
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
         TabIndex        =   91
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
         TabIndex        =   99
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
         TabIndex        =   125
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
         TabIndex        =   123
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
         TabIndex        =   124
         Top             =   3360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstRec 
         Height          =   2895
         Left            =   240
         TabIndex        =   127
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
         TabIndex        =   232
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
         TabIndex        =   224
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
         TabIndex        =   223
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
         TabIndex        =   222
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
         Index           =   1
         Left            =   9840
         TabIndex        =   220
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
         TabIndex        =   219
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
         TabIndex        =   218
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
         TabIndex        =   216
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
         TabIndex        =   215
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
         TabIndex        =   214
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
         TabIndex        =   210
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
         TabIndex        =   209
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
         TabIndex        =   208
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
         TabIndex        =   126
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
         TabIndex        =   82
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
         TabIndex        =   90
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
         TabIndex        =   98
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
         TabIndex        =   106
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
      TabIndex        =   204
      Top             =   12000
      Width           =   16245
      _ExtentX        =   28654
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
      TabIndex        =   270
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
      TabIndex        =   259
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
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Dim WithEvents Getserver As OPCServer
Attribute Getserver.VB_VarHelpID = -1
Dim PointLook1           As Boolean
Dim PointLook2           As Boolean
Dim PointLook3           As Boolean
Dim PointLook4           As Boolean
Dim PointLook5           As Boolean
Dim ShiftTest            As Integer

Private Sub Form_Load()
'главната форма на диспечерския панел

    Const RateUp            As Integer = 10

    Dim NotMachineNumber    As Integer
    Dim oper                As Panel
    Dim today               As Panel
    Dim mixes               As Panel
    Dim exps                As Panel
    Dim voexps              As Panel
    Dim vexps               As Panel
    Dim kgexps              As Panel
    Dim OrderedQuantl       As Single
    Dim RealQuantl          As Single
    Dim TotalKGsl           As Single
    Dim Index               As Integer
    Dim intEmpFile          As Integer
    Dim lblW()              As String
    Dim lblR()              As String
    Dim lblR2()             As String
    Dim X                   As Integer
    Dim i                   As Integer
    Dim PrevSetSilos        As Boolean
    Dim strSubKeySilos      As String
    Dim PlaceSilosSet       As String
    Dim PlaceSilos          As String
    Dim cn                  As ADODB.Connection
    Dim rs                  As Recordset
    Dim rsLog               As Recordset
    Dim WrkPerm             As String
    Dim LogErr              As Boolean
    Dim chD                 As Date
    Dim chD1                As Date
    Dim chDnow              As Date
    Dim chDnowStr           As String
    Dim MixID               As Long
    Dim ExpID               As Long
    Dim calcDay             As Integer
    Dim DaysLeft            As Integer
    Dim StrDev              As String
    Dim response            As Integer
    Dim comIns              As String
    Dim comEdit             As String
    Dim PrevSet             As Boolean
    Dim strSubKey           As String

    MousePointer = vbHourglass
    
    intEmpFile = FreeFile

    'височина на панела
    DispPanel.Height = 10000
    
    'настройка на рамките
    frDisp.Top = 21880
    frOrders.Top = 21880
    frRecepies.Top = 21880
    frClients.Top = 21880
    frDrivers.Top = 21880
    frSuppliers.Top = 21880
    frMaterials.Top = 21880
    frAbout.Top = 21880
        
    'изключване на бутоните
    Me.btnDisp.Enabled = False
    Me.btnOrders.Enabled = False
    Me.btnRecepies.Enabled = False
    Me.btnClients.Enabled = False
    Me.btnDrivers.Enabled = False
    Me.btnSuppliers.Enabled = False
    Me.btnMaterials.Enabled = False
    Me.btnNotes.Enabled = False
    Me.btnAdminPanel.Enabled = False
    Me.btnExit.Enabled = False
    Me.chPrintConf.Enabled = False
    Me.btnSvExp.Visible = False
    
    'рекламен надпис
    'бг
    Me.lblAddver.Caption = "  ТИП-Сервиз   ЕООД    гр. Червен бряг    Софтуер, проектиране, изграждане и сервиз   на бетонови стопанства"
    
    'надписи
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
    btnExit.Caption = UniExit
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
    lblMatHum = uniHumidity
    btnClearMat.Caption = uniNew
    btnSvNwMat.Caption = uniSave
    btnDelMat.Caption = uniDel
    btnAddMatDlvr.Caption = uniDlvr
    btnSvExp.Caption = uniEnterExp
    
    stOrd.Caption = ""
    stClnt.Caption = ""
    stExp.Caption = ""
    'бг
    lblMach.Caption = "Машина " & MachineNumber
    If MachineNumber = 1 Then NotMachineNumber = 2
    If MachineNumber = 2 Then NotMachineNumber = 1
    Me.btnChMach.Caption = "Към Машина " & NotMachineNumber
    Me.btnChSilos.Caption = "Смяна силози"
    Me.lblSilos(0).Caption = "Силоз 1 -->"
    Me.lblSilos(1).Caption = "Силоз 2 -->"
    Me.lblSilos(2).Caption = "Силоз 3 -->"
    Me.lblSilos(3).Caption = "Силоз 4 -->"

    'флагове
    PrintRightBut = False
    Dispatcher = True
    ExpeditionStarted = False
    EmptyData = False
    OffMode = False
    FlagButRec = -1 'флаг за бутон от рамка рецепти дали е "запис" или "изтрий"
    
    DecSep = GetDecimalSep()
    
    'проверка за размяна на силозите
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
    
    'вертикални надписи за меню диспечер
    lblW = Split(uniOrdsVert)
    lblWait.Caption = ""
    For X = 0 To UBound(lblW)
        lblWait.Caption = lblWait.Caption & lblW(X) & vbCrLf
    Next
    lblR = Split(uniReadyVert)
    lblR2 = Split(uniReady2Vert)
    lblReady2.Caption = ""
    For X = 0 To UBound(lblR2)
        lblReady2.Caption = lblReady2.Caption & lblR2(X) & vbCrLf
    Next
    
    'надпис на течки в меню материали
    For i = 0 To 5
        Me.s1(i).Caption = uniFlow & i + 1
    Next i
    
    'проверка за десетичен знак в текстовите полета
    If InStr(txtDispQuant.Text, DecSep) = 0 Then PointLook1 = False
    If InStr(txtOrdQuant.Text, DecSep) = 0 Then PointLook2 = False
    If InStr(txtRec4(Index).Text, DecSep) = 0 Then PointLook3 = False
    If InStr(txtCapDrv.Text, DecSep) = 0 Then PointLook4 = False
    If InStr(txtMatHum.Text, DecSep) = 0 Then PointLook5 = False
    
    'днешна дата използвана при разрешение за старт на програмата
    chDnow = Format(Now, "DD-MM-YYYY")
    chDnowStr = Format(Now, "DD-MM-YYYY")

'-----------------------Start postgreSQL-----------------------------------
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
    
    'настройка на статус-бар
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
    
    DispPanel.StatusBar.Refresh
    
    'рефреш на панела ???
    DispPanel.Hide
    DispPanel.Show
    
    MousePointer = vbHourglass
    
    'надпис свързване...
    lblLoading.ForeColor = &HC000&
    lblLoading.Caption = "  " & uniLoading & "  "
    lblLoading.Refresh

    Load frmOPC
    'frmOPC.Hide
'----------------------------------------------------------------start OPC-----------------------
    'настройки на opc server
    If MachineNumber = 1 Then StrDev = "ConcreteNodePLC."
    If MachineNumber = 2 Then StrDev = "ConcreteNodePLC2."
    
    MousePointer = vbHourglass
    
    'адреси за конфигурацията
    frmOPC.my12.ServerName = MyServer
    frmOPC.my12.AccessPath = StrDev
    frmOPC.my12.ItemName = "Test"
    frmOPC.my12.Attach
    frmOPC.my12.ItemValue = 3333
    frmOPC.my12.UpdateRate = RateUp
    
    frmOPC.cio1000.ServerName = MyServer
    frmOPC.cio1000.AccessPath = StrDev & "Mixer.Online."
    frmOPC.cio1000.ItemName = "OnlineStatus"
    frmOPC.cio1000.DataType = WordType
    frmOPC.cio1000.Attach
    frmOPC.cio1000.UpdateRate = RateUp
        
    Sleep 777
        
    For i = 0 To 77
        DoEvents
        If frmOPC.my12.ItemValue = 3333 And CInt(frmOPC.cio1000.ItemValue) <> 0 Then Exit For
        lblLoading.Caption = uniLoading
        If frmOPC.my12.ItemError = 1 Or frmOPC.cio1000.ItemError = 1 Then GoTo BreakingBad
        frmOPC.my12.ItemValue = 3333
    Next i
    
    MousePointer = vbHourglass
    
    Sleep 7
    
BreakingBad:
    If frmOPC.my12.ItemValue <> 3333 Or Val(frmOPC.cio1000.ItemValue) = 0 Then
        lblLoading.ForeColor = &HFF&
        lblLoading.Caption = MsgNotRespOPC
        lblLoading.Refresh
        MyServer = ""
        frmOPC.my12.ServerName = MyServer
        Set Getserver = New OPCServer
        Getserver.Disconnect

        MousePointer = vbDefault
        response = MsgBox(MsgOffline, vbQuestion Or vbYesNo, MsgNotRespOPC)

        If response = vbYes Then
            Me.btnDispStart.Enabled = False
            OffMode = True
            GoTo OfflineMode
        Else

            End

        End If

    Else
        Open OPCSetFile For Output As intEmpFile
        Write #intEmpFile, MyServer
        Close
    End If
    
    Sleep 7
    
    'адреси за конфигурацията
    frmOPC.cio1001.ServerName = MyServer
    frmOPC.cio1001.AccessPath = StrDev & "Scales.Scale_IM.Online."
    frmOPC.cio1001.ItemName = "OnlineStatus"
    frmOPC.cio1001.DataType = IntegerType
    frmOPC.cio1001.Attach
    frmOPC.cio1001.UpdateRate = RateUp
    
    frmOPC.cio1002.ServerName = MyServer
    frmOPC.cio1002.AccessPath = StrDev & "Scales.Scale_H2O.Online."
    frmOPC.cio1002.ItemName = "OnlineStatus"
    frmOPC.cio1002.DataType = IntegerType
    frmOPC.cio1002.Attach
    frmOPC.cio1002.UpdateRate = RateUp
    
    frmOPC.cio1003.ServerName = MyServer
    frmOPC.cio1003.AccessPath = StrDev & "Scales.Scale_Cement.Online."
    frmOPC.cio1003.ItemName = "OnlineStatus"
    frmOPC.cio1003.DataType = IntegerType
    frmOPC.cio1003.Attach
    frmOPC.cio1003.UpdateRate = RateUp
    
    frmOPC.cio1004.ServerName = MyServer
    frmOPC.cio1004.AccessPath = StrDev & "Scales.Scale_Chemicals.Online."
    frmOPC.cio1004.ItemName = "OnlineStatus"
    frmOPC.cio1004.DataType = IntegerType
    frmOPC.cio1004.Attach
    frmOPC.cio1004.UpdateRate = RateUp
    
    frmOPC.cio1005.ServerName = MyServer
    frmOPC.cio1005.AccessPath = StrDev & "Mixer.Online."
    frmOPC.cio1005.ItemName = "PCCommands"
    frmOPC.cio1005.DataType = IntegerType
    frmOPC.cio1005.Attach
    frmOPC.cio1005.UpdateRate = RateUp
    
    frmOPC.MixCap.ServerName = MyServer
    frmOPC.MixCap.AccessPath = StrDev & "MachineSettings."
    frmOPC.MixCap.ItemName = "MixerCapacity"
    frmOPC.MixCap.DataType = DWordType
    frmOPC.MixCap.Attach
    frmOPC.MixCap.UpdateRate = RateUp
    
    frmOPC.TimeMixDefault.ServerName = MyServer
    frmOPC.TimeMixDefault.AccessPath = StrDev & "MachineSettings."
    frmOPC.TimeMixDefault.ItemName = "TimeMixDefault"
    frmOPC.TimeMixDefault.DataType = WordType
    frmOPC.TimeMixDefault.Attach
    frmOPC.TimeMixDefault.UpdateRate = RateUp
    
    frmOPC.TimePourDefault.ServerName = MyServer
    frmOPC.TimePourDefault.AccessPath = StrDev & "MachineSettings."
    frmOPC.TimePourDefault.ItemName = "TimePourDefault"
    frmOPC.TimePourDefault.DataType = WordType
    frmOPC.TimePourDefault.Attach
    frmOPC.TimePourDefault.UpdateRate = RateUp
    
    frmOPC.NumIMSilos.ServerName = MyServer
    frmOPC.NumIMSilos.AccessPath = StrDev & "MachineSettings."
    frmOPC.NumIMSilos.ItemName = "NumIMSilos"
    frmOPC.NumIMSilos.DataType = IntegerType
    frmOPC.NumIMSilos.Attach
    frmOPC.NumIMSilos.UpdateRate = RateUp
    
    frmOPC.NumCementSilos.ServerName = MyServer
    frmOPC.NumCementSilos.AccessPath = StrDev & "MachineSettings."
    frmOPC.NumCementSilos.ItemName = "NumCementSilos"
    frmOPC.NumCementSilos.DataType = IntegerType
    frmOPC.NumCementSilos.Attach
    frmOPC.NumCementSilos.UpdateRate = RateUp
    
    frmOPC.NumWaterSilos.ServerName = MyServer
    frmOPC.NumWaterSilos.AccessPath = StrDev & "MachineSettings."
    frmOPC.NumWaterSilos.ItemName = "NumWaterSilos"
    frmOPC.NumWaterSilos.DataType = IntegerType
    frmOPC.NumWaterSilos.Attach
    frmOPC.NumWaterSilos.UpdateRate = RateUp
    
    frmOPC.NumChemSilos.ServerName = MyServer
    frmOPC.NumChemSilos.AccessPath = StrDev & "MachineSettings."
    frmOPC.NumChemSilos.ItemName = "NumChemSilos"
    frmOPC.NumChemSilos.DataType = IntegerType
    frmOPC.NumChemSilos.Attach
    frmOPC.NumChemSilos.UpdateRate = RateUp

    'адреси за инициализация на течките
    frmOPC.dm1.ServerName = MyServer
    frmOPC.dm1.AccessPath = StrDev & "Scales.Scale_IM."
    frmOPC.dm1.ItemName = "Settings"
    frmOPC.dm1.DataType = WordType
    frmOPC.dm1.Attach
    frmOPC.dm1.UpdateRate = RateUp
    
    frmOPC.dm11.ServerName = MyServer
    frmOPC.dm11.AccessPath = StrDev & "Scales.Scale_IM."
    frmOPC.dm11.ItemName = "Silo_1"
    frmOPC.dm11.DataType = WordType
    frmOPC.dm11.Attach
    frmOPC.dm11.UpdateRate = RateUp
    
    frmOPC.dm12.ServerName = MyServer
    frmOPC.dm12.AccessPath = StrDev & "Scales.Scale_IM."
    frmOPC.dm12.ItemName = "Silo_2"
    frmOPC.dm12.DataType = WordType
    frmOPC.dm12.Attach
    frmOPC.dm12.UpdateRate = RateUp
    
    frmOPC.dm13.ServerName = MyServer
    frmOPC.dm13.AccessPath = StrDev & "Scales.Scale_IM."
    frmOPC.dm13.ItemName = "Silo_3"
    frmOPC.dm13.DataType = WordType
    frmOPC.dm13.Attach
    frmOPC.dm13.UpdateRate = RateUp
    
    frmOPC.dm14.ServerName = MyServer
    frmOPC.dm14.AccessPath = StrDev & "Scales.Scale_IM."
    frmOPC.dm14.ItemName = "Silo_4"
    frmOPC.dm14.DataType = WordType
    frmOPC.dm14.Attach
    frmOPC.dm14.UpdateRate = RateUp
    
    frmOPC.dm15.ServerName = MyServer
    frmOPC.dm15.AccessPath = StrDev & "Scales.Scale_IM."
    frmOPC.dm15.ItemName = "Silo_5"
    frmOPC.dm15.DataType = WordType
    frmOPC.dm15.Attach
    frmOPC.dm15.UpdateRate = RateUp
    
    frmOPC.dm2.ServerName = MyServer
    frmOPC.dm2.AccessPath = StrDev & "Scales.Scale_H2O."
    frmOPC.dm2.ItemName = "Settings"
    frmOPC.dm2.DataType = WordType
    frmOPC.dm2.Attach
    frmOPC.dm2.UpdateRate = RateUp
    
    frmOPC.dm21.ServerName = MyServer
    frmOPC.dm21.AccessPath = StrDev & "Scales.Scale_H2O."
    frmOPC.dm21.ItemName = "Silo_1"
    frmOPC.dm21.DataType = WordType
    frmOPC.dm21.Attach
    frmOPC.dm21.UpdateRate = RateUp
    
    frmOPC.dm3.ServerName = MyServer
    frmOPC.dm3.AccessPath = StrDev & "Scales.Scale_Cement."
    frmOPC.dm3.ItemName = "Settings"
    frmOPC.dm3.DataType = WordType
    frmOPC.dm3.Attach
    frmOPC.dm3.UpdateRate = RateUp
    
    frmOPC.dm31.ServerName = MyServer
    frmOPC.dm31.AccessPath = StrDev & "Scales.Scale_Cement."
    frmOPC.dm31.ItemName = "Silo_1"
    frmOPC.dm31.DataType = WordType
    frmOPC.dm31.Attach
    frmOPC.dm31.UpdateRate = RateUp
    
    frmOPC.dm32.ServerName = MyServer
    frmOPC.dm32.AccessPath = StrDev & "Scales.Scale_Cement."
    frmOPC.dm32.ItemName = "Silo_2"
    frmOPC.dm32.DataType = WordType
    frmOPC.dm32.Attach
    frmOPC.dm32.UpdateRate = RateUp
    
    frmOPC.dm33.ServerName = MyServer
    frmOPC.dm33.AccessPath = StrDev & "Scales.Scale_Cement."
    frmOPC.dm33.ItemName = "Silo_3"
    frmOPC.dm33.DataType = WordType
    frmOPC.dm33.Attach
    frmOPC.dm33.UpdateRate = RateUp
    
    frmOPC.dm34.ServerName = MyServer
    frmOPC.dm34.AccessPath = StrDev & "Scales.Scale_Cement."
    frmOPC.dm34.ItemName = "Silo_4"
    frmOPC.dm34.DataType = WordType
    frmOPC.dm34.Attach
    frmOPC.dm34.UpdateRate = RateUp
    
    frmOPC.dm4.ServerName = MyServer
    frmOPC.dm4.AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmOPC.dm4.ItemName = "Settings"
    frmOPC.dm4.DataType = WordType
    frmOPC.dm4.Attach
    frmOPC.dm4.UpdateRate = RateUp
    
    frmOPC.dm41.ServerName = MyServer
    frmOPC.dm41.AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmOPC.dm41.ItemName = "Silo_1"
    frmOPC.dm41.DataType = WordType
    frmOPC.dm41.Attach
    frmOPC.dm41.UpdateRate = RateUp
    
    frmOPC.dm42.ServerName = MyServer
    frmOPC.dm42.AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmOPC.dm42.ItemName = "Silo_2"
    frmOPC.dm42.DataType = WordType
    frmOPC.dm42.Attach
    frmOPC.dm42.UpdateRate = RateUp
    
    frmOPC.dm43.ServerName = MyServer
    frmOPC.dm43.AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmOPC.dm43.ItemName = "Silo_3"
    frmOPC.dm43.DataType = WordType
    frmOPC.dm43.Attach
    frmOPC.dm43.UpdateRate = RateUp
    
    frmOPC.dm44.ServerName = MyServer
    frmOPC.dm44.AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmOPC.dm44.ItemName = "Silo_4"
    frmOPC.dm44.DataType = WordType
    frmOPC.dm44.Attach
    frmOPC.dm44.UpdateRate = RateUp
    
    frmOPC.dm45.ServerName = MyServer
    frmOPC.dm45.AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmOPC.dm45.ItemName = "Silo_5"
    frmOPC.dm45.DataType = WordType
    frmOPC.dm45.Attach
    frmOPC.dm45.UpdateRate = RateUp
    
    frmOPC.dm46.ServerName = MyServer
    frmOPC.dm46.AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmOPC.dm46.ItemName = "Silo_6"
    frmOPC.dm46.DataType = WordType
    frmOPC.dm46.Attach
    frmOPC.dm46.UpdateRate = RateUp
    
    'адреси на дозаторите
    frmOPC.BCDAggr.ServerName = MyServer
    frmOPC.BCDAggr.AccessPath = StrDev & "Scales.Scale_IM.Online."
    frmOPC.BCDAggr.ItemName = "OnlineValue"
    frmOPC.BCDAggr.DataType = DWordType
    frmOPC.BCDAggr.Attach
    frmOPC.BCDAggr.UpdateRate = RateUp
    
    frmOPC.BCDCem.ServerName = MyServer
    frmOPC.BCDCem.AccessPath = StrDev & "Scales.Scale_Cement.Online."
    frmOPC.BCDCem.ItemName = "OnlineValue"
    frmOPC.BCDCem.DataType = DWordType
    frmOPC.BCDCem.Attach
    frmOPC.BCDCem.UpdateRate = RateUp
    
    frmOPC.BCDWat.ServerName = MyServer
    frmOPC.BCDWat.AccessPath = StrDev & "Scales.Scale_H2O.Online."
    frmOPC.BCDWat.ItemName = "OnlineValue"
    frmOPC.BCDWat.DataType = DWordType
    frmOPC.BCDWat.Attach
    frmOPC.BCDWat.UpdateRate = RateUp
    
    frmOPC.BCDChem.ServerName = MyServer
    frmOPC.BCDChem.AccessPath = StrDev & "Scales.Scale_Chemicals.Online."
    frmOPC.BCDChem.ItemName = "OnlineValue"
    frmOPC.BCDChem.DataType = DWordType
    frmOPC.BCDChem.Attach
    frmOPC.BCDChem.UpdateRate = RateUp
    
    MousePointer = vbHourglass
    
    'адреси на рецептата
    For i = 0 To 63
        frmOPC.dm1000(i).ServerName = MyServer
        frmOPC.dm1000(i).AccessPath = StrDev & "Commissioning.Receipt.Buffer."
        If i < 10 Then
            frmOPC.dm1000(i).ItemName = "Word_100" & i
        Else
            frmOPC.dm1000(i).ItemName = "Word_10" & i
        End If
        frmOPC.dm1000(i).DataType = WordType
        frmOPC.dm1000(i).Attach
        frmOPC.dm1000(i).UpdateRate = RateUp
    Next i
    
    'адреси за четене на последна рецепта
    For i = 0 To 51
        frmOPC.dm1100(i).ServerName = MyServer
        frmOPC.dm1100(i).AccessPath = StrDev & "Buffer_FinishedMixes."
        If i < 10 Then
            frmOPC.dm1100(i).ItemName = "Word_110" & i
        Else
            frmOPC.dm1100(i).ItemName = "Word_11" & i
        End If
        frmOPC.dm1100(i).DataType = WordType
        frmOPC.dm1100(i).Attach
        frmOPC.dm1100(i).UpdateRate = RateUp
    Next i
    
    'адреси за брой бъркала
    frmOPC.dm500.ServerName = MyServer
    frmOPC.dm500.AccessPath = StrDev & "Mixer."
    frmOPC.dm500.ItemName = "ReadyCycles"
    frmOPC.dm500.DataType = IntegerType
    frmOPC.dm500.Attach
    frmOPC.dm500.UpdateRate = RateUp

    frmOPC.dm501.ServerName = MyServer
    frmOPC.dm501.AccessPath = StrDev & "Mixer."
    frmOPC.dm501.ItemName = "ReadyAutoCycles"
    frmOPC.dm501.DataType = IntegerType
    frmOPC.dm501.Attach
    frmOPC.dm501.UpdateRate = RateUp
    
    'адрес за име на рецепта
    frmOPC.dm1070.ServerName = MyServer
    frmOPC.dm1070.AccessPath = StrDev
    frmOPC.dm1070.ItemName = "name1"
    frmOPC.dm1070.DataType = StringType
    frmOPC.dm1070.Attach
    frmOPC.dm1070.UpdateRate = RateUp
'----------------------------------------------------------------end OPC------------------------------

    'надпис свързан
    lblLoading.Caption = uniLoaded
    lblLoading.Refresh
    
    MousePointer = vbHourglass
    
    'зареждане на параметрите на машината
    Load frmParam
    ns1 = Val(frmParam.txtNumIMSilos.Text)
    ns3 = Val(frmParam.txtNumCementSilos.Text)
    ns2 = Val(frmParam.txtNumWaterSilos.Text)
    ns4 = Val(frmParam.txtNumChemSilos.Text)
    MixCap = ARound(IEEE754(frmOPC.MixCap.ItemValue), 2)
    TMd = Val(frmParam.txtTimeMixDefault)
    TPd = Val(frmParam.txtTimePourDefault)
    Unload frmParam
    
    'настройка на брой видими силози
    For i = 1 To ns3
        Me.lblSilos(i - 1).Visible = True
        Me.numSilos(i - 1).Visible = True
    Next i
    
    'запис на параметрите на машината ако няма реални данни за броя на течките от контролера (18.11.2015)
    If ns1 <= 0 Then GoTo OfflineMode 'четене на брой течки от базата данни
    If ns2 <= 0 Then GoTo OfflineMode
    If ns3 <= 0 Then GoTo OfflineMode
    If ns4 <= 0 Then GoTo OfflineMode
    
'-----------------------Start postgreSQL-----------------------------------
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
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
'--------------------------End PostgreSQL-----------------------------------

    GoTo SkipOffline
OfflineMode:
'-----------------------Start postgreSQL-----------------------------------
    'изчитане на параметри в офлайн режим
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
    
    'интервали на таймерите
    ClockT.Interval = 100
    ScalesT.Interval = 100
    
    'настройка на рамките
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
    
    'проверка в регистъра кои форми да се печатат
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
    strSubKey = Trim(PlaceProgPrint)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        Me.chPrintConf.Value = GetSetting(PlaceProgSettings, PlacePrint, "AutoPrForm", ErrRes)
    Else
        Me.chPrintConf.Value = 0
    End If
    If frmLogin.AdminSuccess = True And frmLogin.RootUser = False Then
        Me.btnDispStart.Enabled = False 'забрана на бутона "старт" в режим администратор
    End If

    'проверка в регистъра дали е включен въпроса за недостатъчно количество в силозите
    If MachineNumber = 1 Then
        strSubKey = Trim(Place1SilosQ)
        PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
        If PrevSet = True Then
            QuestSilos = GetSetting(PlaceProgSettings, Place1Q, "Quest1Silos", ErrRes)
        Else
            QuestSilos = 0
        End If
    ElseIf MachineNumber = 2 Then
        strSubKey = Trim(Place2SilosQ)
        PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
        If PrevSet = True Then
            QuestSilos = GetSetting(PlaceProgSettings, Place2Q, "Quest2Silos", ErrRes)
        Else
            QuestSilos = 0
        End If
    End If

    'проверка в регистъра дали е включена редакцията на бележките
    strSubKey = Trim(PlaceEditor)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    If PrevSet = True Then
        ShowEditor = GetSetting(PlaceProgSettings, PlaceEd, "NotesEditor", ErrRes)
    Else
        ShowEditor = 0
    End If

    'включване на бутоните
    Me.btnDisp.Enabled = True
    Me.btnOrders.Enabled = True
    Me.btnRecepies.Enabled = True
    Me.btnClients.Enabled = True
    Me.btnDrivers.Enabled = True
    Me.btnSuppliers.Enabled = True
    Me.btnMaterials.Enabled = True
    Me.btnNotes.Enabled = True
    Me.btnAdminPanel.Enabled = True
    Me.btnExit.Enabled = True
    Me.chPrintConf.Enabled = True
    
    If frmLogin.AdminSuccess = True Or frmLogin.RootUser = True Then
        If frmLogin.RootUser = False Then Me.btnDelOrd.Visible = False
        Me.btnDelDrv.Enabled = True
        Me.btnDelClnt.Enabled = True
        Me.btnDelRec.Enabled = True
        Me.btnDelSup.Enabled = True
        Me.btnDelMat.Enabled = True
    ElseIf frmLogin.AdminSuccess = False Or frmLogin.RootUser = False Then
        'бг
        If rActDel = 0 Then
            Me.btnDelOrd.Visible = False
            Me.btnDelDrv.Enabled = False
            Me.btnDelClnt.Enabled = False
            Me.btnDelRec.Enabled = False
            Me.btnDelSup.Enabled = False
            Me.btnDelMat.Enabled = False
        Else
            If frmLogin.RootUser = True Then Me.btnDelOrd.Visible = True
            Me.btnDelOrd.Visible = False
            Me.btnDelDrv.Enabled = True
            Me.btnDelClnt.Enabled = True
            Me.btnDelRec.Enabled = True
            Me.btnDelSup.Enabled = True
            Me.btnDelMat.Enabled = True
        End If
    End If
    
    MousePointer = vbDefault
    
    'проверка дали е сговнясана програмата (ctrl+alt+rightmouse)- симулиране на авария в софтуера
    strSubKey = Trim(PlaceShit)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    If PrevSet = True Then
        ShitEnabled = GetSetting(PlaceProgSettings, Shit, "Shit", ErrRes)
    Else
        ShitEnabled = 0
    End If
    If ShitEnabled = True Then
        ConStr = "" 'изтриваме стринга за връзка с базата данни
    Else
        ConStr = "PROVIDER=PostgreSQL;" & "DATA SOURCE=" & IPConnStr & ";" & "LOCATION=" & DbaseName & ";" & "USER ID=" & DbaseUser & ";" & "PASSWORD=" & PassConnStr & ";"
    End If
End Sub

Private Sub ClockT_Timer()
'таймер на часовника
    Dim MyTime      As String
    Dim response    As Integer
    Dim mConn       As Boolean
    

    MyTime = Format$(Now, "hh:mm:ss")
    Clock.Text = Left$(MyTime, 2) & ":" & Mid$(MyTime, 4, 2) & ":" & Right$(MyTime, 2)
    DayToday = Format(Now, "DD-MM-YYYY")
    
    'следене на обема на миксера
    Me.btnMixCap.Caption = "Обем Миксер: " & ARound(IEEE754(frmOPC.MixCap.ItemValue), 2)
    
    'функции за статус и следене на връзката
    Call GetStat
    mConn = MaintConn()
    
    'следене за офлайн режим
    If mConn = True And OffMode = True Then
        MousePointer = vbDefault
        MsgBox MsgConnEst, vbInformation, uniLoaded
        OffMode = False
    End If
    If OffMode = False And mConn = False Then
        MousePointer = vbDefault
        response = MsgBox(MsgOffline, vbQuestion Or vbYesNo, MsgNotRespOPC)
        If response = vbYes Then
            Me.btnDispStart.Enabled = False
            OffMode = True
        Else
            End
        End If
    End If

    'вклюване и изключване на бутона за изпращане към контролера
    If DispPanel.indReq.Caption = statReqStarted And Me.indAvaria.Caption <> statAvaria Then
        DispConfirm.btnSendToController.Enabled = False
    Else
        DispConfirm.btnSendToController.Enabled = True
    End If
End Sub

Private Sub ScalesT_Timer()
'таймер на дозаторите

    If OffMode = False Then
        kgAggr.Text = ARound(IEEE754(frmOPC.BCDAggr.ItemValue), 0)
        kgCem.Text = ARound(IEEE754(frmOPC.BCDCem.ItemValue), 0)
        kgWt.Text = ARound(IEEE754(frmOPC.BCDWat.ItemValue), 0)
        kgChm.Text = Format(ARound(IEEE754(frmOPC.BCDChem.ItemValue), 2), "0.00")
    Else
        kgAggr.Text = "NoCon"
        kgCem.Text = "NoCon"
        kgWt.Text = "NoCon"
        kgChm.Text = "NoCon"
    End If
End Sub

Private Sub AVTimer_Timer()
'таймер за следене за авария

    Dim response As Integer
    
    If Me.indAvaria.Caption = statAvaria And Me.TimerRes.Enabled = True Then
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
'таймер за следене на стартирана заявка

    If Me.indReq.Caption = statReqStarted Then
        ExpeditionStarted = True
        ReqTime = Format(Now, "DD.MM.YYYY - HH:MM:SS")
        Me.TimerStartReq.Enabled = False
    End If
End Sub

Private Sub TimerRes_Timer()
'таймер за следене на резулати от замесите

    If Me.indValveMix.Caption = statMixOpened And CInt(frmOPC.dm1100(0).ItemValue) = HelpRes And ExpeditionStarted = True Then
        Call GetResult
    End If
End Sub

Private Sub FormT_Timer()
'таймер за следене на шибъра на миксера

    Dim response As Integer
    
    If Me.indValveMix.Caption = statMixClosed Then
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

Private Sub btnChMach_Click()
'бутон за смяна на машините

    Dim hw          As Long
    Dim retval      As Long
    Dim SwWin       As String
    
    If MachineNumber = 1 Then SwWin = "Машина 2 - ТИП-Панел v" & App.Major & "." & App.Minor
    If MachineNumber = 2 Then SwWin = "Машина 1 - ТИП-Панел v" & App.Major & "." & App.Minor
    hw = FindWindow(vbNullString, SwWin)
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

Private Sub btnMixCap_Click()
'бутон за обема на миксера

    frmMixCap.Show
End Sub

Private Sub btnChSilos_Click()
'бутон за смяна на силозите
    frmChSilos.Show
End Sub

Private Sub chPrintConf_Click()
'авто-печат
    SaveSetting PlaceProgSettings, PlacePrint, "AutoPrForm", Me.chPrintConf
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'натискане на логото

    Dim PrevSet As Boolean
    Dim strSubKey As String
    
    ShiftTest = Shift And 7

    Select Case ShiftTest
        Case 6
            'проверка дали е сговнясана програмата
            strSubKey = Trim(PlaceShit)
            PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
            If PrevSet = True Then
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
'клик на логото

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
'бутон диспечер

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
'бутон старт

    PrintAnyForm = False

    If cmbDispOrd = "" Or cmbDispDrv = "" Then
        MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
    Else
        DispConfirm.Show
        Call DispConfSend
    End If
End Sub

Private Sub txtDispClnt_Change()
'при промяна на клиента от меню диспечер

    Dim cn        As ADODB.Connection
    Dim rs        As Recordset
    
'------------------------------Start PostgreSQL----------------------------------
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

Private Sub lstOrdWait_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

Private Sub lstMixReady_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

Private Sub btnOrders_Click()
'бутон заявки

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

Private Sub txtMatHum_GotFocus()

    txtMatHum.SelStart = 0
    txtMatHum.SelLength = Len(txtMatHum.Text)
    If InStr(txtMatHum.Text, DecSep) <> 0 Then
        PointLook5 = True
    Else
        PointLook5 = False
    End If
End Sub


Private Sub txtMatHum_KeyPress(KeyAscii As Integer)

    If InStr(txtMatHum.Text, DecSep) <> 0 Then
        PointLook5 = True
    Else
        PointLook5 = False
    End If
    If txtMatHum.SelLength = Len(txtMatHum.Text) Then
        PointLook5 = False
    Else
    End If
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If
    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook5 = True Then
        KeyAscii = 0
    Else
        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook5 = True
        Else
        End If
    End If

End Sub

Private Sub txtMatHum_Change()

    If InStr(txtMatHum.Text, DecSep) <> 0 Then
        PointLook5 = True
    Else
        PointLook5 = False
    End If
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

Private Sub btnRecepies_Click()
'бутон рецепти

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

    Dim PassCheck As String
    
    FlagButRec = 1

    If rDeactNRPass = 0 And (frmLogin.AdminSuccess = False Or frmLogin.RootUser = False) Then
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

    Dim PassCheck As String
    
    FlagButRec = 2

    If rDeactDRPass = 0 And (frmLogin.AdminSuccess = False Or frmLogin.RootUser = False) Then
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
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 46 '.
        Case 47 '/
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtTypeRec_GotFocus()

    txtNameRec.SelStart = 0
    txtNameRec.SelLength = Len(txtTypeRec.Text)
End Sub

Private Sub txtTypeRec_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 47 '/
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtClassRec_GotFocus()

    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRec.Text)
End Sub

Private Sub txtClassRec_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 46 '.
        Case 47 '/
        Case 92
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtClassRecK_GotFocus()

    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecK.Text)
End Sub

Private Sub txtClassRecK_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 47 '/
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtClassRecV_GotFocus()

    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecV.Text)
End Sub

Private Sub txtClassRecV_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 47 '/
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtClassRecH_GotFocus()

    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecH.Text)
End Sub

Private Sub txtClassRecH_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 47 '/
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtClassRecP_GotFocus()

    txtClassRec.SelStart = 0
    txtClassRec.SelLength = Len(txtClassRecP.Text)
End Sub

Private Sub txtClassRecP_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 46
        Case 47 '/
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtEDMRec_GotFocus()

    txtEDMRec.SelStart = 0
    txtEDMRec.SelLength = Len(txtEDMRec.Text)
End Sub

Private Sub txtEDMRec_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 '-
        Case 45 '+
        Case 47 '/
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
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

Private Sub btnClients_Click()
'бутон клиенти

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
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtBGClnt_GotFocus()

    txtBGClnt.SelStart = 0
    txtBGClnt.SelLength = Len(txtBGClnt.Text)
End Sub

Private Sub txtBGClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtMOLClnt_GotFocus()

    txtMOLClnt.SelStart = 0
    txtMOLClnt.SelLength = Len(txtMOLClnt.Text)
End Sub

Private Sub txtMOLClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 45 To 46 '- .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtAddClnt_GotFocus()

    txtAddClnt.SelStart = 0
    txtAddClnt.SelLength = Len(txtAddClnt.Text)
End Sub

Private Sub txtAddClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtTelClnt_GotFocus()

    txtTelClnt.SelStart = 0
    txtTelClnt.SelLength = Len(txtTelClnt.Text)
End Sub

Private Sub txtTelClnt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 48 To 57, 8 ' 0-9 и bksp
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
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

Private Sub btnDrivers_Click()
'бутон водачи

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
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtRegDrv_GotFocus()

    txtRegDrv.SelStart = 0
    txtRegDrv.SelLength = Len(txtRegDrv.Text)
End Sub

Private Sub txtRegDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
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
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtTelDrv_GotFocus()

    txtTelDrv.SelStart = 0
    txtTelDrv.SelLength = Len(txtTelDrv.Text)
End Sub

Private Sub txtTelDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 48 To 57, 8 ' 0-9 и bksp
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtNoteDrv_GotFocus()

    txtNoteDrv.SelStart = 0
    txtNoteDrv.SelLength = Len(txtNoteDrv.Text)
End Sub

Private Sub txtNoteDrv_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
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
'бутон доставчици

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
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtBGSup_GotFocus()

    txtBGSup.SelStart = 0
    txtBGSup.SelLength = Len(txtBGSup.Text)
End Sub

Private Sub txtBGSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
            KeyAscii = KeyAscii - 32 'пропуска кода към главна буква
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtMOLSup_GotFocus()

    txtMOLSup.SelStart = 0
    txtMOLSup.SelLength = Len(txtMOLSup.Text)
End Sub

Private Sub txtMOLSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 45 To 46 '- .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtAddSup_GotFocus()

    txtAddSup.SelStart = 0
    txtAddSup.SelLength = Len(txtAddSup.Text)
End Sub

Private Sub txtAddSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtTelSup_GotFocus()

    txtTelSup.SelStart = 0
    txtTelSup.SelLength = Len(txtTelSup.Text)
End Sub

Private Sub txtTelSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 48 To 57, 8 ' 0-9 и bksp
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
    End Select
End Sub

Private Sub txtNoteSup_GotFocus()

    txtNoteSup.SelStart = 0
    txtNoteSup.SelLength = Len(txtNoteSup.Text)
End Sub

Private Sub txtNoteSup_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 32 'интервал
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
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

Private Sub btnMaterials_Click()
'бутон материали

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
        Case 65 To 90, 48 To 57, 8 'латиница A-Z, 0-9 и bksp
        Case 97 To 122 'латиница a-z
        Case 192 To 223 'кирилица А-Я
        Case 224 To 255 'кирилица а-я
        Case 43 To 46 '+ , - .
        Case Else
            KeyAscii = 0 'всички останали код ascii = 0
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
'при изход от формата

    Dim msgNow      As String
    Dim response    As Integer
    Dim cn          As New ADODB.Connection
    Dim rs          As New Recordset
    Dim counter     As Long
    Dim comm        As String
    Dim comEdit     As String
    Dim hw          As Long
    Dim retval      As Long
    Dim SwWin       As String

    If Me.indReq.Caption <> statReqStarted Or Me.indAvaria.Caption = statAvaria Then
        If frmOPC.dm1000(0).ItemValue > 0 Then
            msgNow = MsgExpWait & vbCrLf & MsgNoResOnExit & vbCrLf & MsgClose
        Else
            msgNow = MsgClose
        End If
        response = MsgBox(msgNow, vbQuestion Or vbYesNo, UniExit)
        If response = vbYes Then
            Unload AdminPanel
            Unload DispConfirm
'------------------------------Start PostgreSQL--------------------------------------
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
'бг
            If MachineNumber = 1 Then SwWin = "Машина 2 - ТИП-Панел v" & App.Major & "." & App.Minor
            If MachineNumber = 2 Then SwWin = "Машина 1 - ТИП-Панел v" & App.Major & "." & App.Minor
            hw = FindWindow(vbNullString, SwWin)
            Unload Me
            frmStart.Started = True
            frmStart.Show
            If hw <> 0 Then
                retval = ShowWindow(hw, 9)
                frmStart.btnExit = True
                
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

