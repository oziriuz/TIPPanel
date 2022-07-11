VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SendExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarIMname(1 To 6)   As String 'local copy

Private mvarIMkg(1 To 6)     As Integer 'local copy

Private mvarSCRname(1 To 4)  As String 'local copy

Private mvarSCRkg(1 To 4)    As Integer 'local copy

Private mvarWATname(1 To 2)  As String 'local copy

Private mvarWATkg(1 To 2)    As Integer 'local copy

Private mvarCHEMname(1 To 6) As String 'local copy

Private mvarCHEMkg(1 To 6)   As Single 'local copy

Private mvarTmix             As Integer 'local copy

Private mvarTpour            As Integer 'local copy
'local variable(s) to hold property value(s)

Public Property Get AllIMkg(Index As Integer) As Integer

    Dim n As Integer

    AllIMkg = 0

    For n = 1 To 6

        If mvarIMname(n) = IM(Index) Then
            AllIMkg = AllIMkg + mvarIMkg(n)
        End If

    Next n

End Property

Public Property Get AllSCRkg(Index As Integer) As Integer

    Dim n As Integer

    AllSCRkg = 0

    For n = 1 To 4

        If mvarSCRname(n) = Scr(Index) Then
            AllSCRkg = AllSCRkg + mvarSCRkg(n)
        End If

    Next n

End Property

Public Property Get AllWATkg(Index As Integer) As Single

    Dim n As Integer

    AllWATkg = 0

    For n = 1 To 2

        If mvarWATname(n) = Wat(Index) Then
            AllWATkg = AllWATkg + mvarWATkg(n)
        End If

    Next n

End Property

Public Property Get AllCHEMkg(Index As Integer) As Single

    Dim n As Integer

    AllCHEMkg = 0

    For n = 1 To 6

        If mvarCHEMname(n) = Chem(Index) Then
            AllCHEMkg = AllCHEMkg + mvarCHEMkg(n)
        End If

    Next n

End Property

Public Property Let IMname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.IMname = 5
    mvarIMname(Index) = vData
End Property

Public Property Get IMinit(ByVal Index As Integer) As Integer

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.IMinit
    Select Case mvarIMname(Index)

        Case IM(1)
            IMinit = 11

        Case IM(2)
            IMinit = 12

        Case IM(3)
            IMinit = 13

        Case IM(4)
            IMinit = 14

        Case IM(5)
            IMinit = 15

        Case IM(6)
            IMinit = 16

        Case Else
            IMinit = 0
    End Select

End Property

Public Property Let IMkg(ByVal Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.IMkg = 5
    mvarIMkg(Index) = vData
End Property

Public Property Get IMkg(ByVal Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.IMkg
    IMkg = mvarIMkg(Index)
End Property

Public Property Let SCRname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.SCRname = 5
    mvarSCRname(Index) = vData
End Property

Public Property Get SCRinit(ByVal Index As Integer) As Integer

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.SCRinit
    Select Case mvarSCRname(Index)

        Case Scr(1)
            SCRinit = 31

        Case Scr(2)
            SCRinit = 32

        Case Scr(3)
            SCRinit = 33

        Case Scr(4)
            SCRinit = 34

        Case Else
            SCRinit = 0
    End Select

End Property

Public Property Let SCRkg(ByVal Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.SCRkg = 5
    mvarSCRkg(Index) = vData
End Property

Public Property Get SCRkg(ByVal Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.SCRkg
    SCRkg = mvarSCRkg(Index)
End Property

Public Property Let WATname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.WATname = 5
    mvarWATname(Index) = vData
End Property

Public Property Get WATinit(ByVal Index As Integer) As Integer

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.WATinit
    Select Case mvarWATname(Index)

        Case Wat(1)
            WATinit = 21

        Case Wat(2)
            WATinit = 22

        Case Else
    End Select

End Property

Public Property Let WATkg(ByVal Index As Integer, ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.WATkg = 5
    mvarWATkg(Index) = vData
End Property

Public Property Get WATkg(ByVal Index As Integer) As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.WATkg
    WATkg = mvarWATkg(Index)
End Property

Public Property Let CHEMname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CHEMname = 5
    mvarCHEMname(Index) = vData
End Property

Public Property Get CHEMinit(ByVal Index As Integer) As Integer

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CHEMinit
    Select Case mvarCHEMname(Index)

        Case Chem(1)
            CHEMinit = 41

        Case Chem(2)
            CHEMinit = 42

        Case Chem(3)
            CHEMinit = 43

        Case Chem(4)
            CHEMinit = 44

        Case Chem(5)
            CHEMinit = 45

        Case Chem(6)
            CHEMinit = 46

        Case Else
            CHEMinit = 0
    End Select

End Property

Public Property Let CHEMkg(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CHEMkg = 5
    mvarCHEMkg(Index) = vData
End Property

Public Property Get CHEMkg(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CHEMkg
    CHEMkg = mvarCHEMkg(Index)
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

