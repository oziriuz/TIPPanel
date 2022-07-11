VERSION 5.00
Object = "{68254760-8F65-4BB1-9AA4-5F9F4C53FEFD}#1.5#0"; "iDAXCE.ocx"
Begin VB.Form frmSim 
   Caption         =   "frmSim"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   5520
   End
   Begin VB.TextBox resreadymix 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   219
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton dorec 
      Caption         =   "изпълни рецепта"
      Height          =   735
      Left            =   4800
      TabIndex        =   218
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame frResult 
      Caption         =   "Резултат"
      Height          =   4215
      Left            =   120
      TabIndex        =   153
      Top             =   6240
      Width           =   6255
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   23
         Left            =   2160
         TabIndex        =   177
         Top             =   1440
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   13
         Left            =   1200
         TabIndex        =   167
         Top             =   1440
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   157
         Top             =   1440
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm501 
         Height          =   375
         Left            =   5160
         TabIndex        =   226
         Top             =   1560
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "dm501"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm500 
         Height          =   375
         Left            =   5160
         TabIndex        =   225
         Top             =   1200
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "dm500"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   42
         Left            =   4080
         TabIndex        =   196
         Top             =   1080
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   32
         Left            =   3120
         TabIndex        =   186
         Top             =   1080
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   22
         Left            =   2160
         TabIndex        =   176
         Top             =   1080
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   12
         Left            =   1200
         TabIndex        =   166
         Top             =   1080
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   156
         Top             =   1080
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   155
         Top             =   720
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   165
         Top             =   720
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   21
         Left            =   2160
         TabIndex        =   175
         Top             =   720
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   31
         Left            =   3120
         TabIndex        =   185
         Top             =   720
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   41
         Left            =   4080
         TabIndex        =   195
         Top             =   720
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   51
         Left            =   5040
         TabIndex        =   205
         Top             =   720
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   50
         Left            =   5040
         TabIndex        =   204
         Top             =   360
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   40
         Left            =   4080
         TabIndex        =   194
         Top             =   360
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   30
         Left            =   3120
         TabIndex        =   184
         Top             =   360
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   20
         Left            =   2160
         TabIndex        =   174
         Top             =   360
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   10
         Left            =   1200
         TabIndex        =   164
         Top             =   360
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   154
         Top             =   360
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   158
         Top             =   1800
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   159
         Top             =   2160
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   160
         Top             =   2520
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   161
         Top             =   2880
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   162
         Top             =   3240
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   163
         Top             =   3600
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   14
         Left            =   1200
         TabIndex        =   168
         Top             =   1800
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   15
         Left            =   1200
         TabIndex        =   169
         Top             =   2160
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   16
         Left            =   1200
         TabIndex        =   170
         Top             =   2520
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   17
         Left            =   1200
         TabIndex        =   171
         Top             =   2880
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   18
         Left            =   1200
         TabIndex        =   172
         Top             =   3240
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   19
         Left            =   1200
         TabIndex        =   173
         Top             =   3600
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   24
         Left            =   2160
         TabIndex        =   178
         Top             =   1800
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   25
         Left            =   2160
         TabIndex        =   179
         Top             =   2160
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   26
         Left            =   2160
         TabIndex        =   180
         Top             =   2520
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   27
         Left            =   2160
         TabIndex        =   181
         Top             =   2880
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   28
         Left            =   2160
         TabIndex        =   182
         Top             =   3240
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   29
         Left            =   2160
         TabIndex        =   183
         Top             =   3600
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   33
         Left            =   3120
         TabIndex        =   187
         Top             =   1440
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   34
         Left            =   3120
         TabIndex        =   188
         Top             =   1800
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   35
         Left            =   3120
         TabIndex        =   189
         Top             =   2160
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   36
         Left            =   3120
         TabIndex        =   190
         Top             =   2520
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   37
         Left            =   3120
         TabIndex        =   191
         Top             =   2880
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   38
         Left            =   3120
         TabIndex        =   192
         Top             =   3240
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   39
         Left            =   3120
         TabIndex        =   193
         Top             =   3600
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   43
         Left            =   4080
         TabIndex        =   197
         Top             =   1440
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   44
         Left            =   4080
         TabIndex        =   198
         Top             =   1800
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   45
         Left            =   4080
         TabIndex        =   199
         Top             =   2160
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   46
         Left            =   4080
         TabIndex        =   200
         Top             =   2520
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   47
         Left            =   4080
         TabIndex        =   201
         Top             =   2880
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   48
         Left            =   4080
         TabIndex        =   202
         Top             =   3240
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   49
         Left            =   4080
         TabIndex        =   203
         Top             =   3600
         Width           =   975
         _Version        =   65541
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
   End
   Begin VB.Frame frstatus 
      Caption         =   "статуси"
      Height          =   5775
      Left            =   10800
      TabIndex        =   127
      Top             =   120
      Width           =   1815
      Begin IDAXCELib.IDAXCE status 
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   360
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "status"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE watcc 
         Height          =   255
         Left            =   240
         TabIndex        =   222
         Top             =   5160
         Width           =   1335
         _Version        =   65541
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "watcc"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE watP 
         Height          =   255
         Left            =   1200
         TabIndex        =   223
         Top             =   5400
         Width           =   375
         _Version        =   65541
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "watP"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE watM 
         Height          =   255
         Left            =   240
         TabIndex        =   224
         Top             =   5400
         Width           =   375
         _Version        =   65541
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "watM"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin VB.Label indskipfull 
         Caption         =   "количка пълна"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   151
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label indmixopened 
         Caption         =   "клапа отворена"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label skipwait 
         Caption         =   "скип пауза"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label indskipup 
         Caption         =   "скип горе"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   140
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label indskipwait 
         Caption         =   "скип чака"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label indskipdown 
         Caption         =   "скип долу"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label indemgstop 
         Caption         =   "авариен стоп"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   137
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label indavaria 
         Caption         =   "авария"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label indavariamode 
         Caption         =   "авариен режим"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label indmanualmode 
         Caption         =   "ръчен режим"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label indautomode 
         Caption         =   "автоматичен режим"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label mixermix 
         Caption         =   "миксер изключен"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label finishedreq 
         Caption         =   "заявка завършена"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label okreadnewreq 
         Caption         =   "заявка стартирана"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label readyfornew 
         Caption         =   "готов за заявка"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.Frame frControl 
      Caption         =   "контролен блок"
      Height          =   1095
      Left            =   4680
      TabIndex        =   125
      Top             =   4800
      Width           =   6135
      Begin VB.CommandButton swav 
         Caption         =   "авария"
         Height          =   375
         Left            =   120
         TabIndex        =   144
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton swskd 
         Caption         =   "скип долу"
         Height          =   375
         Left            =   840
         TabIndex        =   146
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton swskf 
         Caption         =   "количка пълна"
         Height          =   375
         Left            =   1800
         TabIndex        =   152
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton swkpw 
         Caption         =   "скип чака"
         Height          =   375
         Left            =   3120
         TabIndex        =   147
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton swsku 
         Caption         =   "скип горе"
         Height          =   375
         Left            =   4080
         TabIndex        =   148
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton btnpause 
         Caption         =   "скип пауза"
         Height          =   375
         Left            =   5040
         TabIndex        =   149
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton btnklapamix 
         Caption         =   "клапа миксер"
         Height          =   375
         Left            =   4680
         TabIndex        =   150
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton turnmix 
         Caption         =   "миксер"
         Height          =   375
         Left            =   2520
         TabIndex        =   126
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton btnavstop 
         Caption         =   "авариен стоп"
         Height          =   375
         Left            =   1320
         TabIndex        =   145
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnchange 
         Caption         =   "смяна режим"
         Height          =   375
         Left            =   120
         TabIndex        =   143
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frDosing 
      Caption         =   "дозатори"
      Height          =   2535
      Left            =   120
      TabIndex        =   75
      Top             =   3000
      Width           =   4575
      Begin VB.TextBox reschem1 
         Height          =   285
         Index           =   5
         Left            =   3360
         TabIndex        =   96
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox reschem1 
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   231
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox reschemall 
         Height          =   285
         Left            =   3360
         TabIndex        =   98
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox reswatall 
         Height          =   285
         Left            =   2280
         TabIndex        =   97
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox reschem1 
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   95
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox reschem1 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   94
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox reschem1 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   93
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox reschem1 
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   92
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox reswat 
         Height          =   285
         Left            =   2280
         TabIndex        =   91
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox rescemall 
         Height          =   285
         Left            =   1200
         TabIndex        =   90
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox rescem1 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   89
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox rescem1 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   88
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox rescem1 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   87
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox rescem1 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   86
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox resaggrall 
         Height          =   285
         Left            =   120
         TabIndex        =   85
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox resaggr1 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   84
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox resaggr1 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   83
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox resaggr1 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox resaggr1 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox resaggr1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   855
      End
      Begin IDAXCELib.IDAXCE BCDCem 
         Height          =   375
         Left            =   1200
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   12648384
         BorderStyle     =   1
         MouseAction     =   3
         Caption         =   "BCDCEM"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE BCDWat 
         Height          =   375
         Left            =   2280
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   16777152
         BorderStyle     =   1
         MouseAction     =   3
         Caption         =   "BCDWAT"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE BCDAggr 
         Height          =   375
         Left            =   120
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   14737632
         BorderStyle     =   1
         MouseAction     =   3
         Caption         =   "BCDAGGR"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         ButtonAction    =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE BCDChem 
         Height          =   375
         Left            =   3360
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   12648447
         BorderStyle     =   1
         MouseAction     =   3
         Caption         =   "BCDCHEM"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
   End
   Begin VB.Frame frData 
      Caption         =   "данни рецепта"
      Height          =   4695
      Left            =   5760
      TabIndex        =   23
      Top             =   120
      Width           =   5055
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   12
         Left            =   1320
         TabIndex        =   36
         Top             =   840
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1012"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   35
         Top             =   600
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1011"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   10
         Left            =   1320
         TabIndex        =   34
         Top             =   360
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1010"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   33
         Top             =   2520
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1009"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   32
         Top             =   2280
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1008"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   31
         Top             =   2040
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1007"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   30
         Top             =   1800
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1006"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   29
         Top             =   1560
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1005"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   28
         Top             =   1320
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1004"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   27
         Top             =   1080
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1003"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   26
         Top             =   840
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1002"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   25
         Top             =   600
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1001"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1000"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   49
         Left            =   3840
         TabIndex        =   73
         Top             =   2520
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1049"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   39
         Left            =   3000
         TabIndex        =   63
         Top             =   2520
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1039"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   29
         Left            =   2160
         TabIndex        =   53
         Top             =   2520
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1029"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   43
         Top             =   2520
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1019"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   48
         Left            =   3840
         TabIndex        =   72
         Top             =   2280
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1048"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   38
         Left            =   3000
         TabIndex        =   62
         Top             =   2280
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1038"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   28
         Left            =   2160
         TabIndex        =   52
         Top             =   2280
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1028"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   18
         Left            =   1320
         TabIndex        =   42
         Top             =   2280
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1018"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   17
         Left            =   1320
         TabIndex        =   41
         Top             =   2040
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1017"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   27
         Left            =   2160
         TabIndex        =   51
         Top             =   2040
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1027"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   37
         Left            =   3000
         TabIndex        =   61
         Top             =   2040
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1037"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   47
         Left            =   3840
         TabIndex        =   71
         Top             =   2040
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1047"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   46
         Left            =   3840
         TabIndex        =   70
         Top             =   1800
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1046"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   36
         Left            =   3000
         TabIndex        =   60
         Top             =   1800
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1036"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   26
         Left            =   2160
         TabIndex        =   50
         Top             =   1800
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1026"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   16
         Left            =   1320
         TabIndex        =   40
         Top             =   1800
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1016"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   45
         Left            =   3840
         TabIndex        =   69
         Top             =   1560
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1045"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   35
         Left            =   3000
         TabIndex        =   59
         Top             =   1560
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1035"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   25
         Left            =   2160
         TabIndex        =   49
         Top             =   1560
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1025"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   15
         Left            =   1320
         TabIndex        =   39
         Top             =   1560
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1015"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   44
         Left            =   3840
         TabIndex        =   68
         Top             =   1320
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1044"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   34
         Left            =   3000
         TabIndex        =   58
         Top             =   1320
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1034"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   24
         Left            =   2160
         TabIndex        =   48
         Top             =   1320
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1024"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   14
         Left            =   1320
         TabIndex        =   38
         Top             =   1320
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1014"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   13
         Left            =   1320
         TabIndex        =   37
         Top             =   1080
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1013"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   23
         Left            =   2160
         TabIndex        =   47
         Top             =   1080
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1023"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   33
         Left            =   3000
         TabIndex        =   57
         Top             =   1080
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1033"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   43
         Left            =   3840
         TabIndex        =   67
         Top             =   1080
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1043"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   42
         Left            =   3840
         TabIndex        =   66
         Top             =   840
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1042"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   32
         Left            =   3000
         TabIndex        =   56
         Top             =   840
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1032"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   22
         Left            =   2160
         TabIndex        =   46
         Top             =   840
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1022"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   21
         Left            =   2160
         TabIndex        =   45
         Top             =   600
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1021"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   31
         Left            =   3000
         TabIndex        =   55
         Top             =   600
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1031"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   41
         Left            =   3840
         TabIndex        =   65
         Top             =   600
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1041"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   40
         Left            =   3840
         TabIndex        =   64
         Top             =   360
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1040"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   30
         Left            =   3000
         TabIndex        =   54
         Top             =   360
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1030"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   255
         Index           =   20
         Left            =   2160
         TabIndex        =   44
         Top             =   360
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "dm1020"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin VB.TextBox rectpour 
         Height          =   285
         Left            =   120
         TabIndex        =   117
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox rectmix 
         Height          =   285
         Left            =   120
         TabIndex        =   116
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox recchem6 
         Height          =   285
         Left            =   3960
         TabIndex        =   115
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox recchem5 
         Height          =   285
         Left            =   3960
         TabIndex        =   114
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox recchem4 
         Height          =   285
         Left            =   3960
         TabIndex        =   113
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox recchem3 
         Height          =   285
         Left            =   3960
         TabIndex        =   112
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox recchem2 
         Height          =   285
         Left            =   3960
         TabIndex        =   111
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox recchem1 
         Height          =   285
         Left            =   3960
         TabIndex        =   110
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox recwat 
         Height          =   285
         Left            =   3000
         TabIndex        =   109
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox reccem4 
         Height          =   285
         Left            =   2040
         TabIndex        =   108
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox reccem3 
         Height          =   285
         Left            =   2040
         TabIndex        =   107
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox reccem2 
         Height          =   285
         Left            =   2040
         TabIndex        =   106
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox reccem1 
         Height          =   285
         Left            =   2040
         TabIndex        =   105
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox recim5 
         Height          =   285
         Left            =   1080
         TabIndex        =   104
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox recim4 
         Height          =   285
         Left            =   1080
         TabIndex        =   103
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox recim3 
         Height          =   285
         Left            =   1080
         TabIndex        =   102
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox recim2 
         Height          =   285
         Left            =   1080
         TabIndex        =   101
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox recim1 
         Height          =   285
         Left            =   1080
         TabIndex        =   100
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox recmix 
         Height          =   285
         Left            =   120
         TabIndex        =   99
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lbltpour 
         Alignment       =   2  'Center
         Caption         =   "изсипване"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lbltmix 
         Alignment       =   2  'Center
         Caption         =   "бъркане"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblMix 
         Alignment       =   2  'Center
         Caption         =   "замеси"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   2880
         Width           =   735
      End
   End
   Begin VB.Frame frSetting 
      Caption         =   "настройки"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin IDAXCELib.IDAXCE NumChemSilos 
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumChemSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE NumWaterSilos 
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumWaterSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE NumCementSilos 
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumCementSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE NumIMSilos 
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumIMSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE TimePourDefault 
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   840
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "TimePourDefault"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE TimeMixDefault 
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "TimeMixDefault"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE MixCap 
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "MixCap"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idchem 
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   217
         Top             =   1800
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idchem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idchem 
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   216
         Top             =   1560
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idchem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idchem 
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   215
         Top             =   1320
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idchem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idchem 
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   214
         Top             =   1080
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idchem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idchem 
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   213
         Top             =   840
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idchem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idchem 
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   124
         Top             =   600
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idchem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idwat 
         Height          =   255
         Left            =   4800
         TabIndex        =   123
         Top             =   360
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idwat"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idcem 
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   212
         Top             =   2280
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idcem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idcem 
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   211
         Top             =   2040
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idcem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idcem 
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   210
         Top             =   1800
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idcem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idcem 
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   122
         Top             =   1560
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idcem"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idim 
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   209
         Top             =   1320
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idim"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idim 
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   208
         Top             =   1080
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idim"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idim 
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   207
         Top             =   840
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idim"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idim 
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   206
         Top             =   600
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idim"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE idim 
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   121
         Top             =   360
         Width           =   615
         _Version        =   65541
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   13
         BorderStyle     =   3
         Caption         =   "idim"
         AnimateColor1   =   0
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.25
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin VB.CommandButton btnSaveConfig 
         Caption         =   "Запиши настройки"
         Height          =   495
         Left            =   1560
         TabIndex        =   22
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox visMixCap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox visTimeMixDefault 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox visTimePourDefault 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox visNumIMSilos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox visNumCementSilos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox visNumWaterSilos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox visNumChemSilos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1800
         Width           =   495
      End
      Begin IDAXCELib.IDAXCE my 
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   2400
         Width           =   1335
         _Version        =   65541
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "my"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   12
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin VB.Label lblMixCap 
         Alignment       =   1  'Right Justify
         Caption         =   "миксер"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblNumIMSilos 
         Alignment       =   1  'Right Justify
         Caption         =   "течки им"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblNumCementSilos 
         Alignment       =   1  'Right Justify
         Caption         =   "течки цимент"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblNumWaterSilos 
         Alignment       =   1  'Right Justify
         Caption         =   "течки вода"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblNumChemSilos 
         Alignment       =   1  'Right Justify
         Caption         =   "течки хд"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblTimeMixDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "време бъркане"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblTimePourDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "време изсипване"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
   End
   Begin IDAXCELib.IDAXCE cio1001 
      Height          =   375
      Left            =   6480
      TabIndex        =   227
      Top             =   6600
      Width           =   975
      _Version        =   65541
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   13
      BorderStyle     =   3
      MouseAction     =   3
      Caption         =   "cio1001"
      AnimateColor1   =   1917560804
      AnimateColor2   =   1917560804
      AnimateColor3   =   1917560804
      AnimateColor4   =   16777215
      AutoConnect     =   0   'False
      UpdateRate      =   10
      OLECE_Signature =   -22662
      OLECE_Name      =   "MS Sans Serif"
      OLECE_Size      =   8.25
      OLECE_Bold      =   0   'False
      OLECE_Italic    =   0   'False
      OLECE_Underline =   0   'False
      OLECE_Strikethrough=   0   'False
      OLECE_Weight    =   400
      OLECE_Charset   =   204
   End
   Begin IDAXCELib.IDAXCE cio1002 
      Height          =   375
      Left            =   7440
      TabIndex        =   228
      Top             =   6600
      Width           =   975
      _Version        =   65541
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   13
      BorderStyle     =   3
      MouseAction     =   3
      Caption         =   "cio1002"
      AnimateColor1   =   1917560804
      AnimateColor2   =   1917560804
      AnimateColor3   =   1917560804
      AnimateColor4   =   16777215
      AutoConnect     =   0   'False
      UpdateRate      =   10
      OLECE_Signature =   -22662
      OLECE_Name      =   "MS Sans Serif"
      OLECE_Size      =   8.25
      OLECE_Bold      =   0   'False
      OLECE_Italic    =   0   'False
      OLECE_Underline =   0   'False
      OLECE_Strikethrough=   0   'False
      OLECE_Weight    =   400
      OLECE_Charset   =   204
   End
   Begin IDAXCELib.IDAXCE cio1003 
      Height          =   375
      Left            =   8400
      TabIndex        =   229
      Top             =   6600
      Width           =   975
      _Version        =   65541
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   13
      BorderStyle     =   3
      MouseAction     =   3
      Caption         =   "cio1003"
      AnimateColor1   =   1917560804
      AnimateColor2   =   1917560804
      AnimateColor3   =   1917560804
      AnimateColor4   =   16777215
      AutoConnect     =   0   'False
      UpdateRate      =   10
      OLECE_Signature =   -22662
      OLECE_Name      =   "MS Sans Serif"
      OLECE_Size      =   8.25
      OLECE_Bold      =   0   'False
      OLECE_Italic    =   0   'False
      OLECE_Underline =   0   'False
      OLECE_Strikethrough=   0   'False
      OLECE_Weight    =   400
      OLECE_Charset   =   204
   End
   Begin IDAXCELib.IDAXCE cio1004 
      Height          =   375
      Left            =   9360
      TabIndex        =   230
      Top             =   6600
      Width           =   975
      _Version        =   65541
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   13
      BorderStyle     =   3
      MouseAction     =   3
      Caption         =   "cio1004"
      AnimateColor1   =   1917560804
      AnimateColor2   =   1917560804
      AnimateColor3   =   1917560804
      AnimateColor4   =   16777215
      AutoConnect     =   0   'False
      UpdateRate      =   10
      OLECE_Signature =   -22662
      OLECE_Name      =   "MS Sans Serif"
      OLECE_Size      =   8.25
      OLECE_Bold      =   0   'False
      OLECE_Italic    =   0   'False
      OLECE_Underline =   0   'False
      OLECE_Strikethrough=   0   'False
      OLECE_Weight    =   400
      OLECE_Charset   =   204
   End
   Begin VB.Label finished 
      Alignment       =   2  'Center
      Caption         =   "експедиция завършена"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   221
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "готови замеси"
      Height          =   375
      Left            =   4920
      TabIndex        =   220
      Top             =   3960
      Width           =   615
   End
End
Attribute VB_Name = "frmSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Free1 As Integer
    Dim Free2 As Integer
    Dim MixC As Single
    Dim Nim As Integer
    Dim Nsil As Integer
    Dim Nwat As Integer
    Dim Nchem As Integer
    Dim TPour As Integer
    Dim TMix As Integer

    Const myfile = "d:\simfile.txt"

Private Sub dorec_Click()
    Call AutoStart
End Sub

Private Sub Form_Load()
    
    Me.Caption = "Simulator - Machine " & MachineNumber
    
    Free2 = FreeFile + 1
    
    If Dir(myfile) <> "" Then
        Open myfile For Input As Free2
        Input #Free2, MixC, TMix, TPour, Nim, Nsil, Nwat, Nchem
        Close
    Else
        MixC = 1
        TMix = 11
        TPour = 7
        Nim = 4
        Nsil = 2
        Nwat = 1
        Nchem = 2
    End If
    finished.Caption = ""
    '----------------------------------------------------------------start OPC-----------------------
    'настройки на opc server
    
    Dim StrDev As String
    Const MyServer = "CimQuest.IGOMOPC"
    
    If MachineNumber = 1 Then StrDev = "ConcreteNodePLC."
    If MachineNumber = 2 Then StrDev = "ConcreteNodePLC2."
    
'    MousePointer = vbHourglass
    
    'адреси за конфигурацията

    Const RateUp As Integer = 10

    frmSim.my.ServerName = MyServer
    frmSim.my.AccessPath = StrDev
    frmSim.my.ItemName = "Test"
    frmSim.my.Attach
    frmSim.my.UpdateRate = RateUp
    
    frmSim.status.ServerName = MyServer
    frmSim.status.AccessPath = StrDev & "Mixer.Online."
    frmSim.status.ItemName = "OnlineStatus"
    frmSim.status.DataType = WordType
    frmSim.status.Attach
    frmSim.status.UpdateRate = RateUp
    
    Sleep 7
    
    'адреси за конфигурацията
    frmSim.cio1001.ServerName = MyServer
    frmSim.cio1001.AccessPath = StrDev & "Scales.Scale_IM.Online."
    frmSim.cio1001.ItemName = "OnlineStatus"
    frmSim.cio1001.DataType = IntegerType
    frmSim.cio1001.Attach
    frmSim.cio1001.UpdateRate = RateUp
    
    frmSim.cio1002.ServerName = MyServer
    frmSim.cio1002.AccessPath = StrDev & "Scales.Scale_H2O.Online."
    frmSim.cio1002.ItemName = "OnlineStatus"
    frmSim.cio1002.DataType = IntegerType
    frmSim.cio1002.Attach
    frmSim.cio1002.UpdateRate = RateUp
    
    frmSim.cio1003.ServerName = MyServer
    frmSim.cio1003.AccessPath = StrDev & "Scales.Scale_Cement.Online."
    frmSim.cio1003.ItemName = "OnlineStatus"
    frmSim.cio1003.DataType = IntegerType
    frmSim.cio1003.Attach
    frmSim.cio1003.UpdateRate = RateUp
    
    frmSim.cio1004.ServerName = MyServer
    frmSim.cio1004.AccessPath = StrDev & "Scales.Scale_Chemicals.Online."
    frmSim.cio1004.ItemName = "OnlineStatus"
    frmSim.cio1004.DataType = IntegerType
    frmSim.cio1004.Attach
    frmSim.cio1004.UpdateRate = RateUp
    
'    frmSim.cio1005.ServerName = MyServer
'    frmSim.cio1005.AccessPath = StrDev & "Mixer.Online."
'    frmSim.cio1005.ItemName = "PCCommands"
'    frmSim.cio1005.DataType = IntegerType
'    frmSim.cio1005.Attach
'    frmSim.cio1005.UpdateRate = RateUp
    
    frmSim.MixCap.ServerName = MyServer
    frmSim.MixCap.AccessPath = StrDev & "MachineSettings."
    frmSim.MixCap.ItemName = "MixerCapacity"
    frmSim.MixCap.DataType = DWordType
    frmSim.MixCap.Attach
    frmSim.MixCap.UpdateRate = RateUp
    
    frmSim.TimeMixDefault.ServerName = MyServer
    frmSim.TimeMixDefault.AccessPath = StrDev & "MachineSettings."
    frmSim.TimeMixDefault.ItemName = "TimeMixDefault"
    frmSim.TimeMixDefault.DataType = WordType
    frmSim.TimeMixDefault.Attach
    frmSim.TimeMixDefault.UpdateRate = RateUp
    
    frmSim.TimePourDefault.ServerName = MyServer
    frmSim.TimePourDefault.AccessPath = StrDev & "MachineSettings."
    frmSim.TimePourDefault.ItemName = "TimePourDefault"
    frmSim.TimePourDefault.DataType = WordType
    frmSim.TimePourDefault.Attach
    frmSim.TimePourDefault.UpdateRate = RateUp
    
    frmSim.NumIMSilos.ServerName = MyServer
    frmSim.NumIMSilos.AccessPath = StrDev & "MachineSettings."
    frmSim.NumIMSilos.ItemName = "NumIMSilos"
    frmSim.NumIMSilos.DataType = IntegerType
    frmSim.NumIMSilos.Attach
    frmSim.NumIMSilos.UpdateRate = RateUp
    
    frmSim.NumCementSilos.ServerName = MyServer
    frmSim.NumCementSilos.AccessPath = StrDev & "MachineSettings."
    frmSim.NumCementSilos.ItemName = "NumCementSilos"
    frmSim.NumCementSilos.DataType = IntegerType
    frmSim.NumCementSilos.Attach
    frmSim.NumCementSilos.UpdateRate = RateUp
    
    frmSim.NumWaterSilos.ServerName = MyServer
    frmSim.NumWaterSilos.AccessPath = StrDev & "MachineSettings."
    frmSim.NumWaterSilos.ItemName = "NumWaterSilos"
    frmSim.NumWaterSilos.DataType = IntegerType
    frmSim.NumWaterSilos.Attach
    frmSim.NumWaterSilos.UpdateRate = RateUp
    
    frmSim.NumChemSilos.ServerName = MyServer
    frmSim.NumChemSilos.AccessPath = StrDev & "MachineSettings."
    frmSim.NumChemSilos.ItemName = "NumChemSilos"
    frmSim.NumChemSilos.DataType = IntegerType
    frmSim.NumChemSilos.Attach
    frmSim.NumChemSilos.UpdateRate = RateUp

    'адреси за инициализация на течките
'    frmSim.dm1.ServerName = MyServer
'    frmSim.dm1.AccessPath = StrDev & "Scales.Scale_IM."
'    frmSim.dm1.ItemName = "Settings"
'    frmSim.dm1.DataType = WordType
'    frmSim.dm1.Attach
'    frmSim.dm1.UpdateRate = RateUp
    
    frmSim.idim(0).ServerName = MyServer
    frmSim.idim(0).AccessPath = StrDev & "Scales.Scale_IM."
    frmSim.idim(0).ItemName = "Silo_1"
    frmSim.idim(0).DataType = WordType
    frmSim.idim(0).Attach
    frmSim.idim(0).UpdateRate = RateUp
    
    frmSim.idim(1).ServerName = MyServer
    frmSim.idim(1).AccessPath = StrDev & "Scales.Scale_IM."
    frmSim.idim(1).ItemName = "Silo_2"
    frmSim.idim(1).DataType = WordType
    frmSim.idim(1).Attach
    frmSim.idim(1).UpdateRate = RateUp
    
    frmSim.idim(2).ServerName = MyServer
    frmSim.idim(2).AccessPath = StrDev & "Scales.Scale_IM."
    frmSim.idim(2).ItemName = "Silo_3"
    frmSim.idim(2).DataType = WordType
    frmSim.idim(2).Attach
    frmSim.idim(2).UpdateRate = RateUp
    
    frmSim.idim(3).ServerName = MyServer
    frmSim.idim(3).AccessPath = StrDev & "Scales.Scale_IM."
    frmSim.idim(3).ItemName = "Silo_4"
    frmSim.idim(3).DataType = WordType
    frmSim.idim(3).Attach
    frmSim.idim(3).UpdateRate = RateUp
    
    frmSim.idim(4).ServerName = MyServer
    frmSim.idim(4).AccessPath = StrDev & "Scales.Scale_IM."
    frmSim.idim(4).ItemName = "Silo_5"
    frmSim.idim(4).DataType = WordType
    frmSim.idim(4).Attach
    frmSim.idim(4).UpdateRate = RateUp
    
'    frmSim.dm2.ServerName = MyServer
'    frmSim.dm2.AccessPath = StrDev & "Scales.Scale_H2O."
'    frmSim.dm2.ItemName = "Settings"
'    frmSim.dm2.DataType = WordType
'    frmSim.dm2.Attach
'    frmSim.dm2.UpdateRate = RateUp
    
    frmSim.idwat.ServerName = MyServer
    frmSim.idwat.AccessPath = StrDev & "Scales.Scale_H2O."
    frmSim.idwat.ItemName = "Silo_1"
    frmSim.idwat.DataType = WordType
    frmSim.idwat.Attach
    frmSim.idwat.UpdateRate = RateUp
    
'    frmSim.dm3.ServerName = MyServer
'    frmSim.dm3.AccessPath = StrDev & "Scales.Scale_Cement."
'    frmSim.dm3.ItemName = "Settings"
'    frmSim.dm3.DataType = WordType
'    frmSim.dm3.Attach
'    frmSim.dm3.UpdateRate = RateUp
    
    frmSim.idcem(0).ServerName = MyServer
    frmSim.idcem(0).AccessPath = StrDev & "Scales.Scale_Cement."
    frmSim.idcem(0).ItemName = "Silo_1"
    frmSim.idcem(0).DataType = WordType
    frmSim.idcem(0).Attach
    frmSim.idcem(0).UpdateRate = RateUp
    
    frmSim.idcem(1).ServerName = MyServer
    frmSim.idcem(1).AccessPath = StrDev & "Scales.Scale_Cement."
    frmSim.idcem(1).ItemName = "Silo_2"
    frmSim.idcem(1).DataType = WordType
    frmSim.idcem(1).Attach
    frmSim.idcem(1).UpdateRate = RateUp
    
    frmSim.idcem(2).ServerName = MyServer
    frmSim.idcem(2).AccessPath = StrDev & "Scales.Scale_Cement."
    frmSim.idcem(2).ItemName = "Silo_3"
    frmSim.idcem(2).DataType = WordType
    frmSim.idcem(2).Attach
    frmSim.idcem(2).UpdateRate = RateUp
    
    frmSim.idcem(3).ServerName = MyServer
    frmSim.idcem(3).AccessPath = StrDev & "Scales.Scale_Cement."
    frmSim.idcem(3).ItemName = "Silo_4"
    frmSim.idcem(3).DataType = WordType
    frmSim.idcem(3).Attach
    frmSim.idcem(3).UpdateRate = RateUp
    
'    frmSim.dm4.ServerName = MyServer
'    frmSim.dm4.AccessPath = StrDev & "Scales.Scale_Chemicals."
'    frmSim.dm4.ItemName = "Settings"
'    frmSim.dm4.DataType = WordType
'    frmSim.dm4.Attach
'    frmSim.dm4.UpdateRate = RateUp
    
    frmSim.idchem(0).ServerName = MyServer
    frmSim.idchem(0).AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmSim.idchem(0).ItemName = "Silo_1"
    frmSim.idchem(0).DataType = WordType
    frmSim.idchem(0).Attach
    frmSim.idchem(0).UpdateRate = RateUp
    
    frmSim.idchem(1).ServerName = MyServer
    frmSim.idchem(1).AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmSim.idchem(1).ItemName = "Silo_2"
    frmSim.idchem(1).DataType = WordType
    frmSim.idchem(1).Attach
    frmSim.idchem(1).UpdateRate = RateUp
    
    frmSim.idchem(2).ServerName = MyServer
    frmSim.idchem(2).AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmSim.idchem(2).ItemName = "Silo_3"
    frmSim.idchem(2).DataType = WordType
    frmSim.idchem(2).Attach
    frmSim.idchem(2).UpdateRate = RateUp
    
    frmSim.idchem(3).ServerName = MyServer
    frmSim.idchem(3).AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmSim.idchem(3).ItemName = "Silo_4"
    frmSim.idchem(3).DataType = WordType
    frmSim.idchem(3).Attach
    frmSim.idchem(3).UpdateRate = RateUp
    
    frmSim.idchem(4).ServerName = MyServer
    frmSim.idchem(4).AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmSim.idchem(4).ItemName = "Silo_5"
    frmSim.idchem(4).DataType = WordType
    frmSim.idchem(4).Attach
    frmSim.idchem(4).UpdateRate = RateUp
    
    frmSim.idchem(5).ServerName = MyServer
    frmSim.idchem(5).AccessPath = StrDev & "Scales.Scale_Chemicals."
    frmSim.idchem(5).ItemName = "Silo_6"
    frmSim.idchem(5).DataType = WordType
    frmSim.idchem(5).Attach
    frmSim.idchem(5).UpdateRate = RateUp
    
    'адреси на дозаторите
    frmSim.BCDAggr.ServerName = MyServer
    frmSim.BCDAggr.AccessPath = StrDev & "Scales.Scale_IM.Online."
    frmSim.BCDAggr.ItemName = "OnlineValue"
    frmSim.BCDAggr.DataType = DWordType
    frmSim.BCDAggr.Attach
    frmSim.BCDAggr.UpdateRate = RateUp
    
    frmSim.BCDCem.ServerName = MyServer
    frmSim.BCDCem.AccessPath = StrDev & "Scales.Scale_Cement.Online."
    frmSim.BCDCem.ItemName = "OnlineValue"
    frmSim.BCDCem.DataType = DWordType
    frmSim.BCDCem.Attach
    frmSim.BCDCem.UpdateRate = RateUp
    
    frmSim.BCDWat.ServerName = MyServer
    frmSim.BCDWat.AccessPath = StrDev & "Scales.Scale_H2O.Online."
    frmSim.BCDWat.ItemName = "OnlineValue"
    frmSim.BCDWat.DataType = DWordType
    frmSim.BCDWat.Attach
    frmSim.BCDWat.UpdateRate = RateUp
    
    frmSim.BCDChem.ServerName = MyServer
    frmSim.BCDChem.AccessPath = StrDev & "Scales.Scale_Chemicals.Online."
    frmSim.BCDChem.ItemName = "OnlineValue"
    frmSim.BCDChem.DataType = DWordType
    frmSim.BCDChem.Attach
    frmSim.BCDChem.UpdateRate = RateUp
    
'    MousePointer = vbHourglass
    
    'адреси на рецептата
    For I = 0 To 49
        frmSim.dm1000(I).ServerName = MyServer
        frmSim.dm1000(I).AccessPath = StrDev & "Commissioning.Receipt.Buffer."

        If I < 10 Then
            frmSim.dm1000(I).ItemName = "Word_100" & I
        Else
            frmSim.dm1000(I).ItemName = "Word_10" & I
        End If

        frmSim.dm1000(I).DataType = WordType
        frmSim.dm1000(I).Attach
        frmSim.dm1000(I).UpdateRate = RateUp
    Next I
    
    'адреси за четене на последна рецепта
    For I = 0 To 51
        frmSim.dm1100(I).ServerName = MyServer
        frmSim.dm1100(I).AccessPath = StrDev & "Buffer_FinishedMixes."

        If I < 10 Then
            frmSim.dm1100(I).ItemName = "Word_110" & I
        Else
            frmSim.dm1100(I).ItemName = "Word_11" & I
        End If

        frmSim.dm1100(I).DataType = WordType
        frmSim.dm1100(I).Attach
        frmSim.dm1100(I).UpdateRate = RateUp
    Next I
    
    'адреси за брой бъркала
    frmSim.dm500.ServerName = MyServer
    frmSim.dm500.AccessPath = StrDev & "Mixer."
    frmSim.dm500.ItemName = "ReadyCycles"
    frmSim.dm500.DataType = IntegerType
    frmSim.dm500.Attach
    frmSim.dm500.UpdateRate = RateUp

    frmSim.dm501.ServerName = MyServer
    frmSim.dm501.AccessPath = StrDev & "Mixer."
    frmSim.dm501.ItemName = "ReadyAutoCycles"
    frmSim.dm501.DataType = IntegerType
    frmSim.dm501.Attach
    frmSim.dm501.UpdateRate = RateUp
    
'    frmSim.dm1070.ServerName = MyServer
'    frmSim.dm1070.AccessPath = StrDev
'    frmSim.dm1070.ItemName = "name1"
'    frmSim.dm1070.DataType = StringType
'    frmSim.dm1070.Attach
'    frmSim.dm1070.UpdateRate = RateUp
    
'    frmSim.cio00014.ServerName = MyServer
'    frmSim.cio00014.AccessPath = StrDev & "Mixer.Online."
'    frmSim.cio00014.ItemName = "mix"
'    frmSim.cio00014.DataType = DefaultType
'    frmSim.cio00014.Attach
'    frmSim.cio00014.UpdateRate = RateUp
    
    frmSim.watcc.ServerName = MyServer
    frmSim.watcc.AccessPath = StrDev & "Scales.Scale_H2O.Online."
    frmSim.watcc.ItemName = "WatReal"
    frmSim.watcc.DataType = DWordType
    frmSim.watcc.Attach
    frmSim.watcc.UpdateRate = RateUp
    
    frmSim.watM.ServerName = MyServer
    frmSim.watM.AccessPath = StrDev & "Scales.Scale_H2O.Online."
    frmSim.watM.ItemName = "WatMinus"
    frmSim.watM.DataType = ByteType
    frmSim.watM.Attach
    frmSim.watM.UpdateRate = RateUp
    
    frmSim.watP.ServerName = MyServer
    frmSim.watP.AccessPath = StrDev & "Scales.Scale_H2O.Online."
    frmSim.watP.ItemName = "WatPlus"
    frmSim.watP.DataType = ByteType
    frmSim.watP.Attach
    frmSim.watP.UpdateRate = RateUp
    '----------------------------------------------------------------end OPC------------------------------
    
    MixCap.ItemValue = ToIEEE754(MixC)
    TimeMixDefault.ItemValue = CLng("&H" & TMix * 10)
    TimePourDefault.ItemValue = CLng("&H" & TPour * 10)
    NumIMSilos.ItemValue = Nim
    NumCementSilos.ItemValue = Nsil
    NumWaterSilos.ItemValue = Nwat
    NumChemSilos.ItemValue = Nchem
    
    visMixCap.Text = MixC
    visTimeMixDefault.Text = TMix
    visTimePourDefault.Text = TPour
    visNumIMSilos.Text = Nim
    visNumCementSilos.Text = Nsil
    visNumWaterSilos.Text = Nwat
    visNumChemSilos.Text = Nchem
    
    idim(0).ItemValue = 51
    idim(1).ItemValue = 52
    idim(2).ItemValue = 53
    idim(3).ItemValue = 54
    idim(4).ItemValue = 55
    
    idcem(0).ItemValue = 56
    idcem(1).ItemValue = 57
    idcem(2).ItemValue = 58
    idcem(3).ItemValue = 59
    
    idwat.ItemValue = 60
    
    idchem(0).ItemValue = 61
    idchem(1).ItemValue = 62
    idchem(2).ItemValue = 63
    idchem(3).ItemValue = 64
    idchem(4).ItemValue = 65
    idchem(5).ItemValue = 66
    
    rfornreq = 0
    streq = 0
    finreq = 0
    mixmix = 0
    btnopen = 0
    iauto = 0
    iman = 0
    iavar = 1
    iskd = 1
    iskw = 0
    isku = 0
    btnskpause = 0
    iopened = 0
    iskf = 0
    iavaria = 0
    avstop = 0

    stLentaa = 0
    stVoda = 1
    stCiment = 1
    stHimiq = 1
    
    Call GetStat
    
    NowMixing = False
End Sub

Private Sub btnSaveConfig_Click()
    Free1 = FreeFile

    Open myfile For Output As Free1
    Write #Free1, MixC, TMix, TPour, Nim, Nsil, Nwat, Nchem
    Close
End Sub

Private Sub btnchange_Click()
    If iauto = 1 Then
        iauto = 0
        iman = 1
        iavar = 0
        Call GetStat
        Exit Sub
    Else
    End If
    
    If iman = 1 Then
        iauto = 0
        iman = 0
        iavar = 1
        Call GetStat
        Exit Sub
    Else
    End If
    
    If iavar = 1 Then
        iauto = 1
        iman = 0
        iavar = 0
        Call GetStat
        Exit Sub
    Else
    End If
End Sub

Private Sub btnavstop_Click()
    Call GetStat
    If avstop = 0 Then
        avstop = 1
    Else
        avstop = 0
    End If
    Call GetStat
End Sub

Public Sub Timer1_Timer()
    If Me.dorec.Visible = True And Val(Me.dm1000(0).ItemValue) > 0 And NowMixing = False Then
        frmSim.dm500.ItemValue = 0
        frmSim.dm501.ItemValue = 0
        Sleep 500
        Call AutoStart
        NowMixing = False
    End If
End Sub

Private Sub turnmix_Click()
    Call GetStat
    If mixmix = 0 Then
        mixmix = 1
    Else
        mixmix = 0
    End If
    Call GetStat
End Sub

Private Sub btnpause_Click()
    Call GetStat
    If btnskpause = 0 Then
        btnskpause = 1
    Else
        btnskpause = 0
    End If
    Call GetStat
End Sub

Private Sub btnklapamix_Click()
    Call OpenMix
End Sub

Private Sub swav_Click()
    Call GetStat
    If iavaria = 0 Then
        iavaria = 1
    Else
        iavaria = 0
    End If
    Call GetStat
End Sub

Private Sub swskd_Click()
    Call GetStat
    If iskd = 0 And iskw = 1 Then
        iskd = 1
        iskw = 0
        isku = 0
    Else
    End If
    Call GetStat
End Sub

Private Sub swskf_Click()
    Call GetStat
    If iskf = 0 And iskd = 1 Then
        iskf = 1
    Else
    End If
    Call GetStat
End Sub

Private Sub swkpw_Click()
    Call GetStat
    If iskw = 0 Then
        iskw = 1
        iskd = 0
        isku = 0
    Else
    End If
    Call GetStat
End Sub

Private Sub swsku_Click()
    Call GetStat
    If isku = 0 And iskw = 1 Then
        isku = 1
        iskw = 0
        iskd = 0
        If iskf = 1 Then
            Call GetStat
            Sleep 3000
            iskf = 0
        Else
        End If
    Else
    End If
    Call GetStat
End Sub

Private Sub Form_Unload(Cancel As Integer)
    status.ItemValue = 0
End Sub

