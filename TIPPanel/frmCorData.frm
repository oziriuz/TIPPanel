VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmCorData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmCorData"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSvData 
      Caption         =   "btnSvData"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   8760
      Width           =   2055
   End
   Begin VB.TextBox txtChemm 
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
      Index           =   5
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemm 
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
      Index           =   4
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemm 
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
      Index           =   3
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemm 
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
      Index           =   2
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemm 
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
      Index           =   1
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemm 
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
      Index           =   0
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemst 
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
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemst 
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
      Index           =   4
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemst 
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
      Index           =   3
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemst 
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
      Index           =   2
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemst 
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
      Index           =   1
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtChemst 
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
      Index           =   0
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtWatm 
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
      Index           =   1
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtWatm 
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
      Index           =   0
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtWatst 
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
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtWatst 
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
      Index           =   0
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemm 
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
      Index           =   3
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemm 
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
      Index           =   2
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemm 
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
      Index           =   1
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemm 
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
      Index           =   0
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemst 
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
      Index           =   3
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemst 
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
      Index           =   2
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemst 
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
      Index           =   1
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtCemst 
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
      Index           =   0
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMm 
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
      Index           =   5
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMm 
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
      Index           =   4
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMm 
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
      Index           =   3
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMm 
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
      Index           =   2
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMm 
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
      Index           =   1
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMm 
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
      Index           =   0
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMst 
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
      Index           =   5
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMst 
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
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMst 
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
      Index           =   3
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMst 
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
      Index           =   2
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMst 
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
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtIMst 
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
      Index           =   0
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtMix 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   600
      Width           =   585
   End
   Begin VB.TextBox txtRecName 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtCarNum 
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtDrvName 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txtClntObj 
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtClntBG 
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1335
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtOrd 
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
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtExp 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   600
      Width           =   1305
   End
   Begin VB.TextBox txtDateExp 
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   600
      Width           =   1335
   End
   Begin ComCtl2.UpDown udMix 
      Height          =   375
      Left            =   3120
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtMix"
      BuddyDispid     =   196618
      OrigLeft        =   3240
      OrigTop         =   600
      OrigRight       =   3495
      OrigBottom      =   975
      Max             =   1000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblMade 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMade"
      Height          =   195
      Index           =   3
      Left            =   8040
      TabIndex        =   87
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblMade 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMade"
      Height          =   195
      Index           =   2
      Left            =   8040
      TabIndex        =   86
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblMade 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMade"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   85
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblMade 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMade"
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   84
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblOrdered 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblOrdered"
      Height          =   195
      Index           =   3
      Left            =   6840
      TabIndex        =   83
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblOrdered 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblOrdered"
      Height          =   195
      Index           =   2
      Left            =   6840
      TabIndex        =   82
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblOrdered 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblOrdered"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   81
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblOrdered 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblOrdered"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   80
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblMat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMat"
      Height          =   195
      Index           =   3
      Left            =   4800
      TabIndex        =   79
      Top             =   5640
      Width           =   1875
   End
   Begin VB.Label lblMat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMat"
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   78
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Label lblMat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMat"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   77
      Top             =   6360
      Width           =   1875
   End
   Begin VB.Label lblMat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMat"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   76
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Label lblRecName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblRecName"
      Height          =   195
      Left            =   6360
      TabIndex        =   75
      Top             =   2280
      Width           =   2115
   End
   Begin VB.Label lblCarNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCarNum"
      Height          =   195
      Left            =   3840
      TabIndex        =   74
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblDrvName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblDrvName"
      Height          =   195
      Left            =   480
      TabIndex        =   73
      Top             =   2280
      Width           =   2715
   End
   Begin VB.Label lblClntObj 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblClntObj"
      Height          =   195
      Left            =   5160
      TabIndex        =   72
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label lblBG 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblBG"
      Height          =   195
      Left            =   3240
      TabIndex        =   71
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lblClntName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblClntName"
      Height          =   195
      Left            =   480
      TabIndex        =   70
      Top             =   1320
      Width           =   2115
   End
   Begin VB.Label lblOrd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblOrd"
      Height          =   195
      Left            =   7080
      TabIndex        =   69
      Top             =   360
      Width           =   1395
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblDate"
      Height          =   195
      Left            =   4560
      TabIndex        =   68
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label lblMix 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblMix"
      Height          =   195
      Left            =   2400
      TabIndex        =   67
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblExp"
      Height          =   195
      Left            =   570
      TabIndex        =   66
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label lblChem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblChem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   4800
      TabIndex        =   60
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblChem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblChem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4800
      TabIndex        =   59
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblChem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblChem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4800
      TabIndex        =   57
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblChem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblChem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4800
      TabIndex        =   56
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblChem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblChem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4800
      TabIndex        =   55
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblChem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblChem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4800
      TabIndex        =   54
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblWat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblWat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   51
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblWat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblWat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   50
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblCem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   45
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblCem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblCem 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   43
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblCem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblCem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblIM 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   35
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblIM 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblIM 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   33
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblIM 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblIM 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   31
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblIM 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmCorData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Told            As Single
Dim Tst             As Single
Dim IMold(1 To 6)   As Single
Dim Cemold(1 To 4)  As Single
Dim Watold(1 To 2)  As Single
Dim Chemold(1 To 6) As Single
Dim PointLook11     As Boolean

Private Sub Form_Load()
    
    Dim r       As Integer
    Dim cnCor   As ADODB.Connection
    Dim rsCor   As Recordset
    Dim comm    As String
    Dim i       As Integer

    Me.Caption = frmDataCor
    Me.lblExp.Caption = uniExped
    Me.lblMix.Caption = uniMix
    Me.lblDate.Caption = uniDate
    Me.lblOrd.Caption = uniOrdCode
    Me.lblClntName.Caption = uniClnt
    Me.lblBG.Caption = uniBG
    Me.lblClntObj.Caption = uniObj
    Me.lblDrvName.Caption = uniDrv
    Me.lblCarNum.Caption = uniDrvReg
    Me.lblRecName.Caption = uniRec
    Me.btnSvData.Caption = uniSave

    For r = 0 To 3
        Me.lblMat(r).Caption = uniMat
        Me.lblOrdered(r).Caption = uniOrdered
        Me.lblMade(r).Caption = uniMade
    Next r
    For r = 0 To ns1 - 1
        Me.lblIM(r).Visible = True
        Me.txtIMst(r).Visible = True
        Me.txtIMm(r).Visible = True
        Me.txtIMm(r).MaxLength = 4
    Next r
    For r = 0 To ns3 - 1
        Me.lblCem(r).Visible = True
        Me.txtCemst(r).Visible = True
        Me.txtCemm(r).Visible = True
        Me.txtCemm(r).MaxLength = 4
    Next r
    For r = 0 To ns2 - 1
        Me.lblWat(r).Visible = True
        Me.txtWatst(r).Visible = True
        Me.txtWatm(r).Visible = True
        Me.txtWatm(r).MaxLength = 4
    Next r
    For r = 0 To ns4 - 1
        Me.lblChem(r).Visible = True
        Me.txtChemst(r).Visible = True
        Me.txtChemm(r).Visible = True
        Me.txtChemm(r).MaxLength = 5
    Next r
'------------------------------Start PostgreSQL----------------------------------
    Set cnCor = New ADODB.Connection 'връзка с база данни
        cnCor.ConnectionTimeout = 10
        cnCor.Open ConStr 'отваряме връзката
    
    MousePointer = vbHourglass
        'визуализация на замесите от последната експедиция
    Set rsCor = cnCor.Execute("SELECT mix_num, exp_num, exp_q FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1") 'маркираме последния замес
    If Not rsCor.EOF And Not rsCor.BOF Then
        i = Val(rsCor!exp_num) 'маркираме номера на експедиция от последния замес
    Else
    End If
    comm = "SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & i & " ORDER BY mix_num ASC;"
    Set rsCor = cnCor.Execute(comm) 'маркираме всички замеси от намерената експедиция
    rsCor.MoveFirst 'отиваме в началото на маркираните замеси
    Me.udMix.Min = Val(rsCor!mix_num)
    rsCor.MoveLast
    Me.udMix.Max = Val(rsCor!mix_num)
    Me.txtExp.Text = rsCor!exp_num
    Me.txtMix.Text = rsCor!mix_num
    Me.txtDateExp.Text = Left$(rsCor!time_exp_start, 10)
    Me.txtOrd.Text = rsCor!ord_num
    Me.txtClnt.Text = rsCor!name_clnt
    Me.txtClntBG.Text = rsCor!bg_clnt
    Me.txtClntObj.Text = rsCor!obj_clnt
    Me.txtDrvName.Text = rsCor!name_drv
    Me.txtCarNum.Text = rsCor!reg_drv
    Me.txtRecName.Text = rsCor!name_rec
    Me.lblIM(0).Caption = rsCor!im1_name
    Me.lblIM(1).Caption = rsCor!im2_name
    Me.lblIM(2).Caption = rsCor!im3_name
    Me.lblIM(3).Caption = rsCor!im4_name
    Me.lblIM(4).Caption = rsCor!im5_name
    Me.lblIM(5).Caption = rsCor!im6_name
    Me.txtIMst(0).Text = rsCor!im1z
    Me.txtIMst(1).Text = rsCor!im2z
    Me.txtIMst(2).Text = rsCor!im3z
    Me.txtIMst(3).Text = rsCor!im4z
    Me.txtIMst(4).Text = rsCor!im5z
    Me.txtIMst(5).Text = rsCor!im6z
    Me.txtIMm(0).Text = rsCor!im1i
    Me.txtIMm(1).Text = rsCor!im2i
    Me.txtIMm(2).Text = rsCor!im3i
    Me.txtIMm(3).Text = rsCor!im4i
    Me.txtIMm(4).Text = rsCor!im5i
    Me.txtIMm(5).Text = rsCor!im6i
    IMold(1) = Val(rsCor!im1i)
    IMold(2) = Val(rsCor!im2i)
    IMold(3) = Val(rsCor!im3i)
    IMold(4) = Val(rsCor!im4i)
    IMold(5) = Val(rsCor!im5i)
    IMold(6) = Val(rsCor!im6i)
    Me.lblCem(0).Caption = rsCor!cem1_name
    Me.lblCem(1).Caption = rsCor!cem2_name
    Me.lblCem(2).Caption = rsCor!cem3_name
    Me.lblCem(3).Caption = rsCor!cem4_name
    Me.txtCemst(0).Text = rsCor!cem1z
    Me.txtCemst(1).Text = rsCor!cem2z
    Me.txtCemst(2).Text = rsCor!cem3z
    Me.txtCemst(3).Text = rsCor!cem4z
    Me.txtCemm(0).Text = rsCor!cem1i
    Me.txtCemm(1).Text = rsCor!cem2i
    Me.txtCemm(2).Text = rsCor!cem3i
    Me.txtCemm(3).Text = rsCor!cem4i
    Cemold(1) = Val(rsCor!cem1i)
    Cemold(2) = Val(rsCor!cem2i)
    Cemold(3) = Val(rsCor!cem3i)
    Cemold(4) = Val(rsCor!cem4i)
    Me.lblWat(0).Caption = rsCor!wat1_name
    Me.lblWat(1).Caption = rsCor!wat2_name
    Me.txtWatst(0).Text = rsCor!wat1z
    Me.txtWatst(1).Text = rsCor!wat2z
    Me.txtWatm(0).Text = rsCor!wat1i
    Me.txtWatm(1).Text = rsCor!wat2i
    Watold(1) = Val(rsCor!wat1i)
    Watold(2) = Val(rsCor!wat2i)
    Me.lblChem(0).Caption = rsCor!chem1_name
    Me.lblChem(1).Caption = rsCor!chem2_name
    Me.lblChem(2).Caption = rsCor!chem3_name
    Me.lblChem(3).Caption = rsCor!chem4_name
    Me.lblChem(4).Caption = rsCor!chem5_name
    Me.lblChem(5).Caption = rsCor!chem6_name
    Me.txtChemst(0).Text = rDs(rsCor!chem1z)
    Me.txtChemst(1).Text = rDs(rsCor!chem2z)
    Me.txtChemst(2).Text = rDs(rsCor!chem3z)
    Me.txtChemst(3).Text = rDs(rsCor!chem4z)
    Me.txtChemst(4).Text = rDs(rsCor!chem5z)
    Me.txtChemst(5).Text = rDs(rsCor!chem6z)
    Me.txtChemm(0).Text = rDs(rsCor!chem1i)
    Me.txtChemm(1).Text = rDs(rsCor!chem2i)
    Me.txtChemm(2).Text = rDs(rsCor!chem3i)
    Me.txtChemm(3).Text = rDs(rsCor!chem4i)
    Me.txtChemm(4).Text = rDs(rsCor!chem5i)
    Me.txtChemm(5).Text = rDs(rsCor!chem6i)
    Chemold(1) = CSng(rDs(rsCor!chem1i))
    Chemold(2) = CSng(rDs(rsCor!chem2i))
    Chemold(3) = CSng(rDs(rsCor!chem3i))
    Chemold(4) = CSng(rDs(rsCor!chem4i))
    Chemold(5) = CSng(rDs(rsCor!chem5i))
    Chemold(6) = CSng(rDs(rsCor!chem6i))
    Tst = CSng(rDs(rsCor!total_rec_kg))
    Told = CSng(rDs(rsCor!total_real_kg))
    
    rsCor.Close 'затваряме записите
    Set rsCor = Nothing
    cnCor.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnCor = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------
    For r = 0 To ns1 - 1
        If CSng(rDs(Me.txtIMst(r).Text)) > 0 Then
            Me.txtIMm(r).Locked = False
        Else
            Me.txtIMm(r).Locked = True
        End If
    Next r
    For r = 0 To ns3 - 1
        If CSng(rDs(Me.txtCemst(r).Text)) > 0 Then
            Me.txtCemm(r).Locked = False
        Else
            Me.txtCemm(r).Locked = True
        End If
    Next r
    For r = 0 To ns2 - 1
        If CSng(rDs(Me.txtWatst(r).Text)) > 0 Then
            Me.txtWatm(r).Locked = False
        Else
            Me.txtWatm(r).Locked = True
        End If
    Next r
    For r = 0 To ns4 - 1
        If CSng(rDs(Me.txtChemst(r).Text)) > 0 Then
            Me.txtChemm(r).Locked = False
        Else
            Me.txtChemm(r).Locked = True
        End If
    Next r
End Sub

Private Sub txtMix_Change()

    Dim cnCor   As ADODB.Connection
    Dim rsCor   As Recordset
    Dim comm    As String
    Dim i       As Integer
    Dim r       As Integer
    
'------------------------------Start PostgreSQL----------------------------------
    Set cnCor = New ADODB.Connection 'връзка с база данни
        cnCor.ConnectionTimeout = 10
        cnCor.Open ConStr 'отваряме връзката
    
    MousePointer = vbHourglass
    'визуализация на замесите от последната експедиция
    Set rsCor = cnCor.Execute("SELECT mix_num, exp_num, exp_q FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1") 'маркираме последния замес
    If Not rsCor.EOF And Not rsCor.BOF Then
        i = Val(rsCor!exp_num) 'маркираме номера на експедиция от последния замес
    Else
    End If
    comm = "SELECT * FROM mix_result_bc" & MachineNumber & " WHERE mix_num = " & Val(Me.txtMix.Text) & " LIMIT 1;"
    Set rsCor = cnCor.Execute(comm) 'маркираме всички замеси от намерената експедиция
    Me.txtExp.Text = rsCor!exp_num
    Me.txtMix.Text = rsCor!mix_num
    Me.txtDateExp.Text = Left$(rsCor!time_exp_start, 10)
    Me.txtOrd.Text = rsCor!ord_num
    Me.txtClnt.Text = rsCor!name_clnt
    Me.txtClntBG.Text = rsCor!bg_clnt
    Me.txtClntObj.Text = rsCor!obj_clnt
    Me.txtDrvName.Text = rsCor!name_drv
    Me.txtCarNum.Text = rsCor!reg_drv
    Me.txtRecName.Text = rsCor!name_rec
    Me.lblIM(0).Caption = rsCor!im1_name
    Me.lblIM(1).Caption = rsCor!im2_name
    Me.lblIM(2).Caption = rsCor!im3_name
    Me.lblIM(3).Caption = rsCor!im4_name
    Me.lblIM(4).Caption = rsCor!im5_name
    Me.lblIM(5).Caption = rsCor!im6_name
    Me.txtIMst(0).Text = rsCor!im1z
    Me.txtIMst(1).Text = rsCor!im2z
    Me.txtIMst(2).Text = rsCor!im3z
    Me.txtIMst(3).Text = rsCor!im4z
    Me.txtIMst(4).Text = rsCor!im5z
    Me.txtIMst(5).Text = rsCor!im6z
    Me.txtIMm(0).Text = rsCor!im1i
    Me.txtIMm(1).Text = rsCor!im2i
    Me.txtIMm(2).Text = rsCor!im3i
    Me.txtIMm(3).Text = rsCor!im4i
    Me.txtIMm(4).Text = rsCor!im5i
    Me.txtIMm(5).Text = rsCor!im6i
    IMold(1) = Val(rsCor!im1i)
    IMold(2) = Val(rsCor!im2i)
    IMold(3) = Val(rsCor!im3i)
    IMold(4) = Val(rsCor!im4i)
    IMold(5) = Val(rsCor!im5i)
    IMold(6) = Val(rsCor!im6i)
    Me.lblCem(0).Caption = rsCor!cem1_name
    Me.lblCem(1).Caption = rsCor!cem2_name
    Me.lblCem(2).Caption = rsCor!cem3_name
    Me.lblCem(3).Caption = rsCor!cem4_name
    Me.txtCemst(0).Text = rsCor!cem1z
    Me.txtCemst(1).Text = rsCor!cem2z
    Me.txtCemst(2).Text = rsCor!cem3z
    Me.txtCemst(3).Text = rsCor!cem4z
    Me.txtCemm(0).Text = rsCor!cem1i
    Me.txtCemm(1).Text = rsCor!cem2i
    Me.txtCemm(2).Text = rsCor!cem3i
    Me.txtCemm(3).Text = rsCor!cem4i
    Cemold(1) = Val(rsCor!cem1i)
    Cemold(2) = Val(rsCor!cem2i)
    Cemold(3) = Val(rsCor!cem3i)
    Cemold(4) = Val(rsCor!cem4i)
    Me.lblWat(0).Caption = rsCor!wat1_name
    Me.lblWat(1).Caption = rsCor!wat2_name
    Me.txtWatst(0).Text = rsCor!wat1z
    Me.txtWatst(1).Text = rsCor!wat2z
    Me.txtWatm(0).Text = rsCor!wat1i
    Me.txtWatm(1).Text = rsCor!wat2i
    Watold(1) = Val(rsCor!wat1i)
    Watold(2) = Val(rsCor!wat2i)
    Me.lblChem(0).Caption = rsCor!chem1_name
    Me.lblChem(1).Caption = rsCor!chem2_name
    Me.lblChem(2).Caption = rsCor!chem3_name
    Me.lblChem(3).Caption = rsCor!chem4_name
    Me.lblChem(4).Caption = rsCor!chem5_name
    Me.lblChem(5).Caption = rsCor!chem6_name
    Me.txtChemst(0).Text = rDs(rsCor!chem1z)
    Me.txtChemst(1).Text = rDs(rsCor!chem2z)
    Me.txtChemst(2).Text = rDs(rsCor!chem3z)
    Me.txtChemst(3).Text = rDs(rsCor!chem4z)
    Me.txtChemst(4).Text = rDs(rsCor!chem5z)
    Me.txtChemst(5).Text = rDs(rsCor!chem6z)
    Me.txtChemm(0).Text = rDs(rsCor!chem1i)
    Me.txtChemm(1).Text = rDs(rsCor!chem2i)
    Me.txtChemm(2).Text = rDs(rsCor!chem3i)
    Me.txtChemm(3).Text = rDs(rsCor!chem4i)
    Me.txtChemm(4).Text = rDs(rsCor!chem5i)
    Me.txtChemm(5).Text = rDs(rsCor!chem6i)
    Chemold(1) = CSng(rDs(rsCor!chem1i))
    Chemold(2) = CSng(rDs(rsCor!chem2i))
    Chemold(3) = CSng(rDs(rsCor!chem3i))
    Chemold(4) = CSng(rDs(rsCor!chem4i))
    Chemold(5) = CSng(rDs(rsCor!chem5i))
    Chemold(6) = CSng(rDs(rsCor!chem6i))
    Tst = CSng(rDs(rsCor!total_rec_kg))
    Told = CSng(rDs(rsCor!total_real_kg))

    If rsCor!avstat = "True" Then
        Me.btnSvData.Enabled = False
    Else
        Me.btnSvData.Enabled = True
    End If
    
    rsCor.Close 'затваряме записите
    Set rsCor = Nothing
    cnCor.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnCor = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------
    For r = 0 To ns1 - 1
        If CSng(rDs(Me.txtIMst(r).Text)) > 0 Then
            Me.txtIMm(r).Locked = False
        Else
            Me.txtIMm(r).Locked = True
        End If
    Next r
    For r = 0 To ns3 - 1
        If CSng(rDs(Me.txtCemst(r).Text)) > 0 Then
            Me.txtCemm(r).Locked = False
        Else
            Me.txtCemm(r).Locked = True
        End If
    Next r
    For r = 0 To ns2 - 1
        If CSng(rDs(Me.txtWatst(r).Text)) > 0 Then
            Me.txtWatm(r).Locked = False
        Else
            Me.txtWatm(r).Locked = True
        End If
    Next r
    For r = 0 To ns4 - 1
        If CSng(rDs(Me.txtChemst(r).Text)) > 0 Then
            Me.txtChemm(r).Locked = False
        Else
            Me.txtChemm(r).Locked = True
        End If
    Next r
End Sub

Private Sub txtIMm_KeyPress(Index As Integer, KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtCemm_KeyPress(Index As Integer, KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtWatm_KeyPress(Index As Integer, KeyAscii As Integer)

    If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End Sub

Private Sub txtChemm_GotFocus(Index As Integer)

    If InStr(txtChemm(Index).Text, DecSep) <> 0 Then
        PointLook11 = True
    Else
        PointLook11 = False
    End If
End Sub

Private Sub txtChemm_Change(Index As Integer)

    If InStr(txtChemm(Index).Text, DecSep) <> 0 Then
        PointLook11 = True
    Else
        PointLook11 = False
    End If
End Sub

Private Sub txtChemm_KeyPress(Index As Integer, KeyAscii As Integer)

    If InStr(txtChemm(Index).Text, DecSep) <> 0 Then
        PointLook11 = True
    Else
        PointLook11 = False
    End If
    If txtChemm(Index).SelLength = Len(txtChemm(Index).Text) Then
        PointLook11 = False
    Else
    End If
    If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> "," And Chr$(KeyAscii) <> "." And Chr$(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    Else
    End If
    If (Chr$(KeyAscii) = "," Or Chr$(KeyAscii) = ".") And PointLook11 = True Then
        KeyAscii = 0
    Else
        If Chr$(KeyAscii) = "." Or Chr$(KeyAscii) = "," Then
            KeyAscii = Asc(DecSep)
            PointLook11 = True
        Else
        End If
    End If
End Sub

Private Sub btnSvData_Click()
    
    Dim e               As Integer
    Dim Tcorkg          As Single
    Dim Tcorvol         As Single
    Dim IMcor(1 To 6)   As Single
    Dim Cemcor(1 To 4)  As Single
    Dim Watcor(1 To 2)  As Single
    Dim Chemcor(1 To 6) As Single
    Dim lastRowCor      As Long
    Dim cnCor           As ADODB.Connection
    Dim rsCor           As Recordset
    Dim comm            As String

    Tcorkg = 0
    Tcorvol = 0
    
    For e = 0 To ns1 - 1
        Tcorkg = Tcorkg + Val(Me.txtIMm(e))
        IMcor(e + 1) = Val(txtIMm(e)) - IMold(e + 1)
    Next e
    For e = 0 To ns3 - 1
        Tcorkg = Tcorkg + Val(Me.txtCemm(e))
        Cemcor(e + 1) = Val(txtCemm(e)) - Cemold(e + 1)
    Next e
    For e = 0 To ns2 - 1
        Tcorkg = Tcorkg + Val(Me.txtWatm(e))
        Watcor(e + 1) = Val(txtWatm(e)) - Watold(e + 1)
    Next e
    For e = 0 To ns4 - 1
        Tcorkg = Tcorkg + CSng(rDs(Me.txtChemm(e)))
        Chemcor(e + 1) = CSng(rDs(Me.txtChemm(e))) - Chemold(e + 1)
    Next e
    If Tst > 0 Then
        Tcorvol = ARound(CSng(rDs(nCoefs)) * (Tcorkg / Tst), 3) 'обем на всеки замес
    End If
'------------------------------Start PostgreSQL----------------------------------
    Set cnCor = New ADODB.Connection 'връзка с база данни
        cnCor.ConnectionTimeout = 10
        cnCor.Open ConStr 'отваряме връзката
    
    MousePointer = vbHourglass

    comm = "UPDATE mix_result_bc" & MachineNumber & " SET im1i = '" & Me.txtIMm(0) _
    & "', im2i = '" & Me.txtIMm(1) _
    & "', im3i = '" & Me.txtIMm(2) _
    & "', im4i = '" & Me.txtIMm(3) _
    & "', im5i = '" & Me.txtIMm(4) _
    & "', im6i = '" & Me.txtIMm(5) _
    & "', cem1i = '" & Me.txtCemm(0) _
    & "', cem2i = '" & Me.txtCemm(1) _
    & "', cem3i = '" & Me.txtCemm(2) _
    & "', cem4i = '" & Me.txtCemm(3) _
    & "', wat1i = '" & Me.txtWatm(0) _
    & "', wat2i = '" & Me.txtWatm(1) _
    & "', chem1i = '" & Me.txtChemm(0) _
    & "', chem2i = '" & Me.txtChemm(1) _
    & "', chem3i = '" & Me.txtChemm(2) _
    & "', chem4i = '" & Me.txtChemm(3) _
    & "', chem5i = '" & Me.txtChemm(4) _
    & "', chem6i = '" & Me.txtChemm(5) _
    & "', total_real_kg = '" & Tcorkg _
    & "', total_vol = '" & rDs(Tcorvol) _
    & "', avstat = 'true' WHERE mix_num = " & Val(Me.txtMix.Text) & ";"
    
    Set rsCor = cnCor.Execute(comm)
    
    'корекция на наличност на им
    For e = 1 To ns1
        If Me.lblIM(e - 1).Caption <> "0" And Me.lblIM(e - 1).Caption <> uniEmpty And Me.lblIM(e - 1).Caption <> "" Then
            Set rsCor = cnCor.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & Me.lblIM(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsCor!m_sold)) + IMcor(e) / 1000, 3) & "'WHERE m_name = '" & Me.lblIM(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & Me.lblIM(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            If Not rsCor.EOF And Not rsCor.BOF Then
                Set rsCor = cnCor.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsCor!mat_sold)) + IMcor(e) / 1000, 3) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & Me.lblIM(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            Else
                'откриваме последния запис
                Set rsCor = cnCor.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                If Not rsCor.EOF And Not rsCor.BOF Then
                    lastRowCor = Val(rsCor!row_num) + 1
                Else
                    lastRowCor = 1
                End If
                Set rsCor = cnCor.Execute("INSERT INTO daily_expenses VALUES(" & lastRowCor & ",'" & Me.lblIM(e - 1).Caption & "','" & ARound(IMcor(e) / 1000, 3) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
            End If
        End If
    Next e
    
    'корекция на наличност на цимент
    For e = 1 To ns3
        If Me.lblCem(e - 1).Caption <> "0" And Me.lblCem(e - 1).Caption <> uniEmpty And Me.lblCem(e - 1).Caption <> "" Then
            Set rsCor = cnCor.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & Me.lblCem(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsCor!m_sold)) + Cemcor(e) / 1000, 3) & "'WHERE m_name = '" & Me.lblCem(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & Me.lblCem(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            If Not rsCor.EOF And Not rsCor.BOF Then
                Set rsCor = cnCor.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsCor!mat_sold)) + Cemcor(e) / 1000, 3) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & Me.lblCem(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            Else
                'откриваме последния запис
                Set rsCor = cnCor.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                If Not rsCor.EOF And Not rsCor.BOF Then
                    lastRowCor = Val(rsCor!row_num) + 1
                Else
                    lastRowCor = 1
                End If
                Set rsCor = cnCor.Execute("INSERT INTO daily_expenses VALUES(" & lastRowCor & ",'" & Me.lblCem(e - 1).Caption & "','" & ARound(Cemcor(e) / 1000, 3) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
            End If
        End If
    Next e
    
    'корекция на разход на вода
    For e = 1 To ns2
        If Me.lblWat(e - 1).Caption <> "0" And Me.lblWat(e - 1).Caption <> uniEmpty And Me.lblWat(e - 1).Caption <> "" Then
            Set rsCor = cnCor.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & Me.lblWat(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsCor!m_sold)) + Watcor(e) / 1000, 3) & "'WHERE m_name = '" & Me.lblWat(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & Me.lblWat(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            If Not rsCor.EOF And Not rsCor.BOF Then
                Set rsCor = cnCor.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsCor!mat_sold)) + Watcor(e) / 1000, 3) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & Me.lblWat(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            Else
                'откриваме последния запис
                Set rsCor = cnCor.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                If Not rsCor.EOF And Not rsCor.BOF Then
                    lastRowCor = Val(rsCor!row_num) + 1
                Else
                    lastRowCor = 1
                End If
                Set rsCor = cnCor.Execute("INSERT INTO daily_expenses VALUES(" & lastRowCor & ",'" & Me.lblWat(e - 1).Caption & "','" & ARound(Watcor(e) / 1000, 3) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
            End If
        End If
    Next e
    
    'корекция на наличност на хд
    For e = 1 To ns4
        If Me.lblChem(e - 1).Caption <> "0" And Me.lblChem(e - 1).Caption <> uniEmpty And Me.lblChem(e - 1).Caption <> "" Then
            Set rsCor = cnCor.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & Me.lblChem(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsCor!m_sold)) + Chemcor(e) / 1000, 5) & "'WHERE m_name = '" & Me.lblChem(e - 1).Caption & "';")
            Set rsCor = cnCor.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & Me.lblChem(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            If Not rsCor.EOF And Not rsCor.BOF Then
                Set rsCor = cnCor.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsCor!mat_sold)) + Chemcor(e) / 1000, 5) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & Me.lblChem(e - 1).Caption & "' AND stamp_date = '" & DayToday & "';")
            Else
                'откриваме последния запис
                Set rsCor = cnCor.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                If Not rsCor.EOF And Not rsCor.BOF Then
                    lastRowCor = Val(rsCor!row_num) + 1
                Else
                    lastRowCor = 1
                End If
                Set rsCor = cnCor.Execute("INSERT INTO daily_expenses VALUES(" & lastRowCor & ",'" & Me.lblChem(e - 1).Caption & "','" & ARound(Chemcor(e) / 1000, 5) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
            End If
        End If
    Next e
    
    rsCor.Close 'затваряме записите
    Set rsCor = Nothing
    cnCor.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnCor = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------
    Me.btnSvData.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim TempMix     As Long
    Dim ResCor      As Result
    Dim cnCor       As ADODB.Connection
    Dim rsCor       As Recordset
    Dim i           As Integer
    Dim comm        As String
    Dim PrevSet     As Boolean
    Dim strSubKey   As String
    
    Set ResCor = New Result
    
'------------------------------Start PostgreSQL----------------------------------
    Set cnCor = New ADODB.Connection 'връзка с база данни
        cnCor.ConnectionTimeout = 10
        cnCor.Open ConStr 'отваряме връзката

    'замесите от последната експедиция
    comm = "SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & Me.txtExp.Text & ";"
    Set rsCor = cnCor.Execute(comm) 'маркираме всички замеси от намерената експедиция
    rsCor.MoveFirst 'отиваме в началото на маркираните замеси
    Do While Not rsCor.EOF
        'запис на данните от базата дани в масиви от променливи
        ResCor.TotalMeasuredKG = ResCor.TotalMeasuredKG + CSng(rDs(rsCor!total_real_kg)) 'сумираме теглата по измерено от всеки замес
        ResCor.TotalQuant = ResCor.TotalQuant + CSng(rDs(rsCor!total_vol)) 'сумираме количества от всеки замес
        rsCor.MoveNext
    Loop
        
    'запис на корекцията в таблицата със заявките
    Set rsCor = cnCor.Execute("SELECT order_qmade FROM orders WHERE order_num =  " & Me.txtOrd.Text & ";")
    Set rsCor = cnCor.Execute("UPDATE orders SET order_qmade = '" & ARound(CSng(rDs(rsCor!order_qmade)) + CSng(rDs(ResCor.TotalQuant)), 3) & "' WHERE order_num =" & Me.txtOrd.Text & ";")
    
    'запис на последните временни данни в tempmix_bc1 таблицата
    Set rsCor = cnCor.Execute("SELECT * FROM tempmix_bc" & MachineNumber & ";")
    TempMix = rsCor!mix_id
    Set rsCor = cnCor.Execute("UPDATE tempmix_bc" & MachineNumber & " SET real_q = '" & ResCor.TotalQuant & "',total_kg_temp = '" & ResCor.TotalMeasuredKG & "' WHERE mix_id =" & TempMix & ";")
    
    rsCor.Close 'затваряме записите
    Set rsCor = Nothing
    cnCor.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cnCor = Nothing
    Set ResCor = Nothing
'-------------------------------End PostgreSQL-------------------------------------------------
    
    PrintAnyForm = False
    
    'проверка в регистъра дали има настройки за принтиране на формите 1,2,3 - ако има излиза въпрос за печат
    strSubKey = Trim(PlaceProgSet1)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        rPrint1 = GetSetting(PlaceProgSettings, PlaceForm1, "Print1", ErrRes)
    Else
        rPrint1 = 0
    End If
            
    strSubKey = Trim(PlaceProgSet2)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        rPrint2 = GetSetting(PlaceProgSettings, PlaceForm2, "Print2", ErrRes)
    Else
        rPrint2 = 0
    End If
            
    strSubKey = Trim(PlaceProgSet3)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        rPrint3 = GetSetting(PlaceProgSettings, PlaceForm3, "Print3", ErrRes)
    Else
        rPrint3 = 0
    End If

    If rPrint1 = 1 Or rPrint2 = 1 Or rPrint3 = 1 Then
        DispPanel.FormT.Enabled = True 'извикваме функция за попълване на форма ако е необходимо
    End If

    Call OpenDisp
End Sub

