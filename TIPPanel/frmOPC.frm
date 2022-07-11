VERSION 5.00
Object = "{68254760-8F65-4BB1-9AA4-5F9F4C53FEFD}#1.5#0"; "iDAXCE.ocx"
Begin VB.Form frmOPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmOPC"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   15495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frOPC 
      Caption         =   "frOPC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   13800
         TabIndex        =   166
         Top             =   6840
         Width           =   1215
      End
      Begin IDAXCELib.IDAXCE my12 
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _Version        =   65541
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         Caption         =   "my12"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   100
         CaptionBad      =   "Bad"
         CaptionError    =   "Error"
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1000"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1001"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1002"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1003"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1004"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   5
         Left            =   480
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1005"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   6
         Left            =   480
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1006"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1007"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   8
         Left            =   480
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1008"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   9
         Left            =   480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1009"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   10
         Left            =   1560
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1010"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   11
         Left            =   1560
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1011"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   12
         Left            =   1560
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1012"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   13
         Left            =   1560
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1013"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   14
         Left            =   1560
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1014"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   15
         Left            =   1560
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1015"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   16
         Left            =   1560
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1016"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   17
         Left            =   1560
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1017"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   18
         Left            =   1560
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1018"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   19
         Left            =   1560
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1019"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   20
         Left            =   2640
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1020"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   21
         Left            =   2640
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1021"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   22
         Left            =   2640
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1022"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   23
         Left            =   2640
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1023"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   24
         Left            =   2640
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1024"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   25
         Left            =   2640
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1025"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   26
         Left            =   2640
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1026"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   27
         Left            =   2640
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1027"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   28
         Left            =   2640
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1028"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   29
         Left            =   2640
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1029"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   30
         Left            =   3720
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1030"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   31
         Left            =   3720
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1031"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   32
         Left            =   3720
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1032"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   33
         Left            =   3720
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1033"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   34
         Left            =   3720
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1034"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   35
         Left            =   3720
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1035"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   36
         Left            =   3720
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1036"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   37
         Left            =   3720
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1037"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   38
         Left            =   3720
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1038"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   39
         Left            =   3720
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1039"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   40
         Left            =   4800
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1040"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   41
         Left            =   4800
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1041"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   42
         Left            =   4800
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1042"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   43
         Left            =   4800
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1043"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   44
         Left            =   4800
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1044"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   45
         Left            =   4800
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1045"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   46
         Left            =   4800
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1046"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   47
         Left            =   4800
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1047"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   48
         Left            =   4800
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1048"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   49
         Left            =   4800
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1049"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE BCDCem 
         Height          =   375
         Left            =   1680
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.74
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   5
         BorderStyle     =   2
         MouseAction     =   3
         Caption         =   "BCDCEM"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         DataType        =   6
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "System"
         OLECE_Size      =   9.74
         OLECE_Bold      =   -1  'True
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   700
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE BCDWat 
         Height          =   375
         Left            =   3240
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.74
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   5
         BorderStyle     =   2
         MouseAction     =   3
         Caption         =   "BCDWAT"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         DataType        =   6
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "System"
         OLECE_Size      =   9.74
         OLECE_Bold      =   -1  'True
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   700
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE BCDChem 
         Height          =   375
         Left            =   4800
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.74
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   5
         BorderStyle     =   2
         MouseAction     =   3
         Caption         =   "BCDCHEM"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         DataType        =   6
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "System"
         OLECE_Size      =   9.74
         OLECE_Bold      =   -1  'True
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   700
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE BCDAggr 
         Height          =   375
         Left            =   120
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
         _Version        =   65541
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   13
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.74
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   5
         BorderStyle     =   2
         MouseAction     =   3
         Caption         =   "BCDAGGR"
         AnimateColor1   =   255
         AnimateColor2   =   0
         AnimateColor3   =   0
         AnimateColor4   =   0
         ButtonAction    =   4
         DataType        =   6
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "System"
         OLECE_Size      =   9.74
         OLECE_Bold      =   -1  'True
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   700
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE MixCap 
         Height          =   375
         Left            =   5400
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1455
         _Version        =   65541
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "MixCap"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         DataType        =   6
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.24
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE TimeMixDefault 
         Height          =   375
         Left            =   5400
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   5160
         Width           =   735
         _Version        =   65541
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "TimeMixDefault"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.24
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE TimePourDefault 
         Height          =   375
         Left            =   6120
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   5160
         Width           =   735
         _Version        =   65541
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "TimePourDefault"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   8.24
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE NumIMSilos 
         Height          =   375
         Left            =   5400
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   5520
         Width           =   735
         _Version        =   65541
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumIMSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         DataType        =   3
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
      Begin IDAXCELib.IDAXCE NumCementSilos 
         Height          =   375
         Left            =   5400
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   5880
         Width           =   735
         _Version        =   65541
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumCementSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         DataType        =   3
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
      Begin IDAXCELib.IDAXCE NumWaterSilos 
         Height          =   375
         Left            =   6120
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   5520
         Width           =   735
         _Version        =   65541
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumWaterSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         DataType        =   3
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
      Begin IDAXCELib.IDAXCE NumChemSilos 
         Height          =   375
         Left            =   6120
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   5880
         Width           =   735
         _Version        =   65541
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         MouseAction     =   3
         Caption         =   "NumChemSilos"
         AnimateColor1   =   1917560804
         AnimateColor2   =   1917560804
         AnimateColor3   =   1917560804
         AnimateColor4   =   16777215
         DataType        =   3
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
      Begin IDAXCELib.IDAXCE dm32 
         Height          =   375
         Left            =   1560
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   5160
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm32"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm12 
         Height          =   375
         Left            =   480
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   5160
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm12"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm13 
         Height          =   375
         Left            =   480
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm13"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm14 
         Height          =   375
         Left            =   480
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm14"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm15 
         Height          =   375
         Left            =   480
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm15"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm11 
         Height          =   375
         Left            =   480
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm11"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm21 
         Height          =   375
         Left            =   1560
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   6600
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm21"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm31 
         Height          =   375
         Left            =   1560
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm31"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm33 
         Height          =   375
         Left            =   1560
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm33"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm34 
         Height          =   375
         Left            =   1560
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm34"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm41 
         Height          =   375
         Left            =   2640
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm41"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm42 
         Height          =   375
         Left            =   2640
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   5160
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm42"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm43 
         Height          =   375
         Left            =   2640
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm43"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm44 
         Height          =   375
         Left            =   2640
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm44"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm45 
         Height          =   375
         Left            =   2640
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm45"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm46 
         Height          =   375
         Left            =   2640
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   6600
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm46"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm1 
         Height          =   375
         Left            =   3840
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm3 
         Height          =   375
         Left            =   3840
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm3"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE dm2 
         Height          =   375
         Left            =   3840
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   5160
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm2"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm4 
         Height          =   375
         Left            =   3840
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
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
         BorderStyle     =   3
         Caption         =   "dm4"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
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
      Begin IDAXCELib.IDAXCE cio1000 
         Height          =   375
         Left            =   7800
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _Version        =   65541
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "cio1000"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE cio1001 
         Height          =   375
         Left            =   9000
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _Version        =   65541
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "cio1001"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE cio1002 
         Height          =   375
         Left            =   10200
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _Version        =   65541
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "cio1002"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE cio1003 
         Height          =   375
         Left            =   11400
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _Version        =   65541
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "cio1003"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE cio1004 
         Height          =   375
         Left            =   12600
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _Version        =   65541
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "cio1004"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   100
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE cio1005 
         Height          =   375
         Left            =   13800
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _Version        =   65541
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "cio1005"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1100"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1101"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1102"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1103"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1104"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1105"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1106"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1107"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1108"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   8400
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1109"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1110"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1111"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1112"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1113"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1114"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1115"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1116"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1117"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1118"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   9480
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1119"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1120"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1121"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1122"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1100 
         Height          =   375
         Index           =   23
         Left            =   10560
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1123"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1124"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1125"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1126"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1127"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1128"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   10560
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1129"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1130"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1131"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1133"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1134"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1135"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1136"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1137"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1138"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1139"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1140"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1141"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1142"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1143"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1144"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1145"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1146"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1147"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1148"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   12720
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1149"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   50
         Left            =   5880
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1050"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   51
         Left            =   5880
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1051"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   52
         Left            =   5880
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1052"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   53
         Left            =   5880
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1053"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   54
         Left            =   5880
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1054"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   55
         Left            =   5880
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1055"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   56
         Left            =   5880
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1056"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   57
         Left            =   5880
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1057"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   58
         Left            =   5880
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1058"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   59
         Left            =   5880
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1059"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   60
         Left            =   6960
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1060"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   61
         Left            =   6960
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1061"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   62
         Left            =   6960
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1062"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1000 
         Height          =   375
         Index           =   63
         Left            =   6960
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Caption         =   "dm1063"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   13800
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1150"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   13800
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1151"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm500 
         Height          =   375
         Left            =   14040
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   3840
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm500"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm501 
         Height          =   375
         Left            =   14040
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   4200
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm501"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   4
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
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
         Left            =   11640
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
         _Version        =   65541
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1132"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   3
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin IDAXCELib.IDAXCE dm1070 
         Height          =   375
         Left            =   7200
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   5160
         Width           =   855
         _Version        =   65541
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   11.99
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         Caption         =   "dm1070"
         AnimateColor1   =   1791338468
         AnimateColor2   =   1791338468
         AnimateColor3   =   1791338468
         AnimateColor4   =   16777215
         DataType        =   9
         AutoConnect     =   0   'False
         UpdateRate      =   10
         OLECE_Signature =   -22662
         OLECE_Name      =   "MS Sans Serif"
         OLECE_Size      =   11.99
         OLECE_Bold      =   0   'False
         OLECE_Italic    =   0   'False
         OLECE_Underline =   0   'False
         OLECE_Strikethrough=   0   'False
         OLECE_Weight    =   400
         OLECE_Charset   =   204
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "5"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "6"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "7"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "8"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "9"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   4320
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmOPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Exit_Click()

    Me.Hide
End Sub

