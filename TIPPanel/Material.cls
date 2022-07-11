VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCode      As Integer 'local copy

Private mvarTitle     As String 'local copy

Private mvarKind      As String 'local copy

Private mvarLoaded    As String 'local copy

Private mvarDelivered As Single 'local copy

Private mvarSold      As Single 'local copy

Private mvarHumidity  As Single 'local copy

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

Public Property Let Kind(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Kind = 5
    mvarKind = vData
End Property

Public Property Get Kind() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Kind
    Kind = mvarKind
End Property

Public Property Let Loaded(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Loaded = 5
    mvarLoaded = vData
End Property

Public Property Get Loaded() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Loaded
    Loaded = mvarLoaded
End Property

Public Property Let Delivered(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Delivered = 5
    mvarDelivered = vData
End Property

Public Property Get Delivered() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Delivered
    Delivered = mvarDelivered
End Property

Public Property Let Sold(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Sold = 5
    mvarSold = vData
End Property

Public Property Get Sold() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Sold
    Sold = mvarSold
End Property

Public Property Let Humidity(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Humidity = 5
    mvarHumidity = vData
End Property

Public Property Get Humidity() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Humidity
    Humidity = mvarHumidity
End Property

Public Property Get Balance() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Balance
    Balance = mvarDelivered - mvarSold
End Property

