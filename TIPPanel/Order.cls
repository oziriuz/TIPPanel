VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCode         As Long 'local copy

Private mvarDateBegin    As String 'local copy

Private mvarDateEnd      As String 'local copy

Private mvarStampDate    As Date 'local copy

Private mvarOrderedQuant As Single 'local copy

Private mvarMadeQuant    As Single 'local copy

Private mvarRecCode      As Integer 'local copy

Private mvarRecTitle     As String 'local copy

Private mvarRecClass     As String 'local copy

Private mvarClntCode     As Integer 'local copy

Private mvarClntTitle    As String 'local copy

Private mvarClntWorksite As String 'local copy
'local variable(s) to hold property value(s)

Public Property Let Code(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Code = 5
    mvarCode = vData
End Property

Public Property Get Code() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Code
    Code = mvarCode
End Property

Public Property Let DateBegin(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DateBegin = 5
    mvarDateBegin = vData
End Property

Public Property Get DateBegin() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DateBegin
    DateBegin = mvarDateBegin
End Property

Public Property Let DateEnd(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DateEnd = 5
    mvarDateEnd = vData
End Property

Public Property Get DateEnd() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DateEnd
    DateEnd = mvarDateEnd
End Property

Public Property Let StampDate(ByVal vData As Date)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.StampDate = 5
    mvarStampDate = vData
End Property

Public Property Get StampDate() As Date
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.StampDate
    StampDate = mvarStampDate
End Property

Public Property Let OrderedQuant(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.OrderedQuant = 5
    mvarOrderedQuant = vData
End Property

Public Property Get OrderedQuant() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.OrderedQuant
    OrderedQuant = mvarOrderedQuant
End Property

Public Property Let MadeQuant(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.MadeQuant = 5
    mvarMadeQuant = vData
End Property

Public Property Get MadeQuant() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.MadeQuant
    MadeQuant = mvarMadeQuant
End Property

Public Property Let RecCode(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecCode = 5
    mvarRecCode = vData
End Property

Public Property Get RecCode() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecCode
    RecCode = mvarRecCode
End Property

Public Property Let RecTitle(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecTitle = 5
    mvarRecTitle = vData
End Property

Public Property Get RecTitle() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecTitle
    RecTitle = mvarRecTitle
End Property

Public Property Let RecClass(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecClass = 5
    mvarRecClass = vData
End Property

Public Property Get RecClass() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecClass
    RecClass = mvarRecClass
End Property

Public Property Let ClntCode(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClntCode = 5
    mvarClntCode = vData
End Property

Public Property Get ClntCode() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClntCode
    ClntCode = mvarClntCode
End Property

Public Property Let ClntTitle(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClntTitle = 5
    mvarClntTitle = vData
End Property

Public Property Get ClntTitle() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClntTitle
    ClntTitle = mvarClntTitle
End Property

Public Property Let ClntWorksite(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClntObject = 5
    mvarClntWorksite = vData
End Property

Public Property Get ClntWorksite() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClntObject
    ClntWorksite = mvarClntWorksite
End Property

