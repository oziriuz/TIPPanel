VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCode           As Integer 'local copy

Private mvarTitle          As String 'local copy

Private mvarIdent          As String 'local copy

Private mvarMOL            As String 'local copy

Private mvarAddress        As String 'local copy

Private mvarPhone          As String 'local copy

Private mvarWorksite(1000) As String 'local copy

Private mvarDistance(1000) As Integer 'local copy

'local variable(s) to hold property value(s)
Private mvarVisible        As Boolean 'local copy

Public Property Let Code(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Code = 5
    mvarCode = vData
End Property

Public Property Get Code() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Code
    Code = mvarCode
End Property

Public Property Let Title(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Title = 5
    mvarTitle = vData
End Property

Public Property Get Title() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Title
    Title = mvarTitle
End Property

Public Property Let Ident(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Ident = 5
    mvarIdent = vData
End Property

Public Property Get Ident() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Ident
    Ident = mvarIdent
End Property

Public Property Let MOL(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.MOL = 5
    mvarMOL = vData
End Property

Public Property Get MOL() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.MOL
    MOL = mvarMOL
End Property

Public Property Let Address(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Address = 5
    mvarAddress = vData
End Property

Public Property Get Address() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Address
    Address = mvarAddress
End Property

Public Property Let Phone(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Phone = 5
    mvarPhone = vData
End Property

Public Property Get Phone() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Phone
    Phone = mvarPhone
End Property

Public Property Let Worksite(Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Worksite = 5
    mvarWorksite(Index) = vData
End Property

Public Property Get Worksite(Index As Integer) As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Worksite
    Worksite = mvarWorksite(Index)
End Property

Public Property Let Distance(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Distance = 5
    mvarDistance(Index) = vData
End Property

Public Property Get Distance(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Distance
    Distance = mvarDistance(Index)
End Property

Public Property Let Visible(ByVal vData As Boolean)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Visible = 5
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
End Property

