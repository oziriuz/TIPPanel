VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const BIM = 10

Const BWat = 20

Const BScr = 30

Const BChem = 40

'local variable(s) to hold property value(s)
Private mvarCode             As Integer 'local copy

Private mvarTitle            As String 'local copy

Private mvarKind             As String 'local copy

Private mvarClass            As String 'local copy

Private mvarClassK           As String 'local copy

Private mvarClassV           As String 'local copy

Private mvarClassH           As String 'local copy

Private mvarClassP           As String 'local copy

Private mvarEDM              As Integer 'local copy

Private mvarTmix             As Integer 'local copy

Private mvarTpour            As Integer 'local copy

Private mvarinitIM(1 To 6)   As Integer 'local copy

Private mvarinitScr(1 To 4)  As Integer 'local copy

Private mvarinitWat(1 To 2)  As Integer 'local copy

Private mvarinitChem(1 To 6) As Integer 'local copy

Private mvarkgIM(1 To 6)     As Integer 'local copy

Private mvarkgScr(1 To 4)    As Integer 'local copy

Private mvarkgWat(1 To 2)    As Integer 'local copy

Private mvarkgChem(1 To 6)   As Single 'local copy

Private mvarkgTotal          As Single 'local copy

Private mvarVisible          As Boolean 'local copy

Public Property Get AllkgIM(Index As Integer) As Integer

    Dim n As Integer

    AllkgIM = 0

    For n = 1 To 6

        If mvarinitIM(n) = BIM + Index Then
            AllkgIM = AllkgIM + mvarkgIM(n)
        End If

    Next n

End Property

Public Property Get AllkgScr(Index As Integer) As Integer

    Dim n As Integer

    AllkgScr = 0

    For n = 1 To 4

        If mvarinitScr(n) = BScr + Index Then
            AllkgScr = AllkgScr + mvarkgScr(n)
        End If

    Next n

End Property

Public Property Get AllkgWat(Index As Integer) As Single

    Dim n As Integer

    AllkgWat = 0

    For n = 1 To 2

        If mvarinitWat(n) = BWat + Index Then
            AllkgWat = AllkgWat + mvarkgWat(n)
        End If

    Next n

End Property

Public Property Get AllkgChem(Index As Integer) As Single

    Dim n As Integer

    AllkgChem = 0

    For n = 1 To 6

        If mvarinitChem(n) = BChem + Index Then
            AllkgChem = AllkgChem + mvarkgChem(n)
        End If

    Next n

End Property

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

Public Property Let Class(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Class = 5
    mvarClass = vData
End Property

Public Property Get Class() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Class
    Class = mvarClass
End Property

Public Property Let ClassK(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClassK = 5
    mvarClassK = vData
End Property

Public Property Get ClassK() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClassK
    ClassK = mvarClassK
End Property

Public Property Let ClassV(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClassV = 5
    mvarClassV = vData
End Property

Public Property Get ClassV() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClassV
    ClassV = mvarClassV
End Property

Public Property Let ClassH(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClassH = 5
    mvarClassH = vData
End Property

Public Property Get ClassH() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClassH
    ClassH = mvarClassH
End Property

Public Property Let ClassP(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClassP = 5
    mvarClassP = vData
End Property

Public Property Get ClassP() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClassP
    ClassP = mvarClassP
End Property

Public Property Let EDM(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.EDM = 5
    mvarEDM = vData
End Property

Public Property Get EDM() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.EDM
    EDM = mvarEDM
End Property

Public Property Let Tmix(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Tmix = 5
    mvarTmix = vData
End Property

Public Property Get Tmix() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Tmix
    Tmix = mvarTmix
End Property

Public Property Let Tpour(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Tpour = 5
    mvarTpour = vData
End Property

Public Property Get Tpour() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Tpour
    Tpour = mvarTpour
End Property

Public Property Let initIM(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.initIM = 5
    mvarinitIM(Index) = vData
End Property

Public Property Get initIM(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.initIM1
    initIM = mvarinitIM(Index)
End Property

Public Property Let kgIM(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.kgIM = 5
    mvarkgIM(Index) = vData
End Property

Public Property Get kgIM(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.kgIM
    kgIM = mvarkgIM(Index)
End Property

Public Property Let initScr(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.initScr = 5
    mvarinitScr(Index) = vData
End Property

Public Property Get initScr(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.initScr
    initScr = mvarinitScr(Index)
End Property

Public Property Let kgScr(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.kgScr = 5
    mvarkgScr(Index) = vData
End Property

Public Property Get kgScr(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.kgScr
    kgScr = mvarkgScr(Index)
End Property

Public Property Let initWat(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.initWat = 5
    mvarinitWat(Index) = vData
End Property

Public Property Get initWat(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.initWat
    initWat = mvarinitWat(Index)
End Property

Public Property Let kgWat(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.kgWat = 5
    mvarkgWat(Index) = vData
End Property

Public Property Get kgWat(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.kgWat
    kgWat = mvarkgWat(Index)
End Property

Public Property Let initChem(Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.initChem = 5
    mvarinitChem(Index) = vData
End Property

Public Property Get initChem(Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.initChem
    initChem = mvarinitChem(Index)
End Property

Public Property Let kgChem(Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.kgChem = 5
    mvarkgChem(Index) = vData
End Property

Public Property Get kgChem(Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.kgChem
    kgChem = mvarkgChem(Index)
End Property

Public Property Let kgTotal(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.kgTotal = 5
    mvarkgTotal = vData
End Property

Public Property Get kgTotal() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.kgTotal
    kgTotal = mvarkgTotal
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

