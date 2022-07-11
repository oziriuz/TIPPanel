VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Driver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCode     As Integer 'local copy

Private mvarTitle    As String 'local copy

Private mvarCarNum   As String 'local copy

Private mvarCarModel As String 'local copy

Private mvarCapacity As Single 'local copy

Private mvarPhone    As String 'local copy

Private mvarNote     As String 'local copy

Private mvarVisible  As Boolean 'local copy

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

Public Property Let CarModel(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CarModel = 5
    mvarCarModel = vData
End Property

Public Property Get CarModel() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CarModel
    CarModel = mvarCarModel
End Property

Public Property Let CarNum(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CarNum = 5
    mvarCarNum = vData
End Property

Public Property Get CarNum() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CarNum
    CarNum = mvarCarNum
End Property

Public Property Let Capacity(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Capacity = 5
    mvarCapacity = vData
End Property

Public Property Get Capacity() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Capacity
    Capacity = mvarCapacity
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

Public Property Let Note(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Note = 5
    mvarNote = vData
End Property

Public Property Get Note() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Note
    Note = mvarNote
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

