VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarMixNum               As Long 'local copy

Private mvarExpNum               As Long 'local copy

Private mvarExpStartTime         As String 'local copy

Private mvarMixReadyTime         As String 'local copy

Private mvarStampDate            As Date 'local copy

Private mvarDispName             As String 'local copy

Private mvarOrderCode            As Long 'local copy

Private mvarOrderDate            As String 'local copy

Private mvarOrderQuant           As Single 'local copy

Private mvarExpQuant             As Single 'local copy

Private mvarExpNumFromOrder      As Long 'local copy

Private mvarMixNumFromOrder      As Long 'local copy

Private mvarClntTitle            As String 'local copy

Private mvarClntIdent            As String 'local copy

Private mvarClntWorksite         As String 'local copy

Private mvarWorksiteDist         As Integer 'local copy

Private mvarDrvTitle             As String 'local copy

Private mvarDrvCarNum            As String 'local copy

Private mvarDrvCapacity          As Single 'local copy

Private mvarRecTitle             As String 'local copy

Private mvarRecKind              As String 'local copy

Private mvarRecClass             As String 'local copy

Private mvarRecClassK            As String 'local copy

Private mvarRecClassV            As String 'local copy

Private mvarRecClassH            As String 'local copy

Private mvarRecClassP            As String 'local copy

Private mvarRecEDM               As Integer 'local copy

Private mvarIMname(1 To 6)       As String 'local copy

Private mvarIMstated(1 To 6)     As Single 'local copy

Private mvarIMmeasured(1 To 6)   As Single 'local copy

Private mvarSCRname(1 To 4)      As String 'local copy

Private mvarSCRstated(1 To 4)    As Single 'local copy

Private mvarSCRmeasured(1 To 4)  As Single 'local copy

Private mvarWATname(1 To 2)      As String 'local copy

Private mvarWATstated(1 To 2)    As Single 'local copy

Private mvarWATmeasured(1 To 2)  As Single 'local copy

Private mvarCHEMname(1 To 6)     As String 'local copy

Private mvarCHEMstated(1 To 6)   As Single 'local copy

Private mvarCHEMmeasured(1 To 6) As Single 'local copy

Private mvarTotalStatedKG        As Single 'local copy

Private mvarTotalMeasuredKG      As Single 'local copy

Private mvarTotalQuant           As Single 'local copy

Public Property Let MixNum(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.MixNum = 5
    mvarMixNum = vData
End Property

Public Property Get MixNum() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.MixNum
    MixNum = mvarMixNum
End Property

Public Property Let ExpNum(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ExpNum = 5
    mvarExpNum = vData
End Property

Public Property Get ExpNum() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ExpNum
    ExpNum = mvarExpNum
End Property

Public Property Let ExpStartTime(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ExpStartTime = 5
    mvarExpStartTime = vData
End Property

Public Property Get ExpStartTime() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ExpStartTime
    ExpStartTime = mvarExpStartTime
End Property

Public Property Let MixReadyTime(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.MixReadyTime = 5
    mvarMixReadyTime = vData
End Property

Public Property Get MixReadyTime() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.MixReadyTime
    MixReadyTime = mvarMixReadyTime
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

Public Property Let OrderCode(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.OrderCode = 5
    mvarOrderCode = vData
End Property

Public Property Get OrderCode() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.OrderCode
    OrderCode = mvarOrderCode
End Property

Public Property Let OrderDate(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.OrderDate = 5
    mvarOrderDate = vData
End Property

Public Property Get OrderDate() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.OrderDate
    OrderDate = mvarOrderDate
End Property

Public Property Let OrderQuant(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.OrderQuant = 5
    mvarOrderQuant = vData
End Property

Public Property Get OrderQuant() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.OrderQuant
    OrderQuant = mvarOrderQuant
End Property

Public Property Let ExpQuant(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ExpQuant = 5
    mvarExpQuant = vData
End Property

Public Property Get ExpQuant() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ExpQuant
    ExpQuant = mvarExpQuant
End Property

Public Property Let ExpNumFromOrder(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ExpNumFromOrder = 5
    mvarExpNumFromOrder = vData
End Property

Public Property Get ExpNumFromOrder() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ExpNumFromOrder
    ExpNumFromOrder = mvarExpNumFromOrder
End Property

Public Property Let MixNumFromOrder(ByVal vData As Long)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.MixNumFromOrder = 5
    mvarMixNumFromOrder = vData
End Property

Public Property Get MixNumFromOrder() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.MixNumFromOrder
    MixNumFromOrder = mvarMixNumFromOrder
End Property

Public Property Let DispName(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DispName = 5
    mvarDispName = vData
End Property

Public Property Get DispName() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DispName
    DispName = mvarDispName
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

Public Property Let ClntIdent(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.ClntIdent = 5
    mvarClntIdent = vData
End Property

Public Property Get ClntIdent() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.ClntIdent
    ClntIdent = mvarClntIdent
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

Public Property Let WorksiteDist(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.WorksiteDist = 5
    mvarWorksiteDist = vData
End Property

Public Property Get WorksiteDist() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.WorksiteDist
    WorksiteDist = mvarWorksiteDist
End Property

Public Property Let DrvTitle(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DrvTitle = 5
    mvarDrvTitle = vData
End Property

Public Property Get DrvTitle() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DrvTitle
    DrvTitle = mvarDrvTitle
End Property

Public Property Let DrvCarNum(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DrvCarNum = 5
    mvarDrvCarNum = vData
End Property

Public Property Get DrvCarNum() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DrvCarNum
    DrvCarNum = mvarDrvCarNum
End Property

Public Property Let DrvCapacity(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DrvCapacity = 5
    mvarDrvCapacity = vData
End Property

Public Property Get DrvCapacity() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DrvCapacity
    DrvCapacity = mvarDrvCapacity
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

Public Property Let RecKind(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecKind = 5
    mvarRecKind = vData
End Property

Public Property Get RecKind() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecKind
    RecKind = mvarRecKind
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

Public Property Let RecClassK(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecClassK = 5
    mvarRecClassK = vData
End Property

Public Property Get RecClassK() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecClassK
    RecClassK = mvarRecClassK
End Property

Public Property Let RecClassV(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecClassV = 5
    mvarRecClassV = vData
End Property

Public Property Get RecClassV() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecClassV
    RecClassV = mvarRecClassV
End Property

Public Property Let RecClassH(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecClassH = 5
    mvarRecClassH = vData
End Property

Public Property Get RecClassH() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecClassH
    RecClassH = mvarRecClassH
End Property

Public Property Let RecClassP(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecClassP = 5
    mvarRecClassP = vData
End Property

Public Property Get RecClassP() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecClassP
    RecClassP = mvarRecClassP
End Property

Public Property Let RecEDM(ByVal vData As Integer)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RecEDM = 5
    mvarRecEDM = vData
End Property

Public Property Get RecEDM() As Integer
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RecEDM
    RecEDM = mvarRecEDM
End Property

Public Property Let IMname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.IMname = 5
    mvarIMname(Index) = vData
End Property

Public Property Get IMname(ByVal Index As Integer) As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.IMname
    IMname = mvarIMname(Index)
End Property

Public Property Let IMstated(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.IMstated = 5
    mvarIMstated(Index) = vData
End Property

Public Property Get IMstated(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.IMstated
    IMstated = mvarIMstated(Index)
End Property

Public Property Let IMmeasured(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.IMmeasured = 5
    mvarIMmeasured(Index) = vData
End Property

Public Property Get IMmeasured(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.IMmeasured
    IMmeasured = mvarIMmeasured(Index)
End Property

Public Property Let SCRname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.SCRname = 5
    mvarSCRname(Index) = vData
End Property

Public Property Get SCRname(ByVal Index As Integer) As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.SCRname
    SCRname = mvarSCRname(Index)
End Property

Public Property Let SCRstated(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.SCRstated = 5
    mvarSCRstated(Index) = vData
End Property

Public Property Get SCRstated(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.SCRstated
    SCRstated = mvarSCRstated(Index)
End Property

Public Property Let SCRmeasured(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.SCRmeasured = 5
    mvarSCRmeasured(Index) = vData
End Property

Public Property Get SCRmeasured(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.SCRmeasured
    SCRmeasured = mvarSCRmeasured(Index)
End Property

Public Property Let WATname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.WATname = 5
    mvarWATname(Index) = vData
End Property

Public Property Get WATname(ByVal Index As Integer) As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.WATname
    WATname = mvarWATname(Index)
End Property

Public Property Let WATstated(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.WATstated = 5
    mvarWATstated(Index) = vData
End Property

Public Property Get WATstated(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.WATstated
    WATstated = mvarWATstated(Index)
End Property

Public Property Let WATmeasured(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.WATmeasured = 5
    mvarWATmeasured(Index) = vData
End Property

Public Property Get WATmeasured(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.WATmeasured
    WATmeasured = mvarWATmeasured(Index)
End Property

Public Property Let CHEMname(ByVal Index As Integer, ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CHEMname = 5
    mvarCHEMname(Index) = vData
End Property

Public Property Get CHEMname(ByVal Index As Integer) As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CHEMname
    CHEMname = mvarCHEMname(Index)
End Property

Public Property Let CHEMstated(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CHEMstated = 5
    mvarCHEMstated(Index) = vData
End Property

Public Property Get CHEMstated(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CHEMstated
    CHEMstated = mvarCHEMstated(Index)
End Property

Public Property Let CHEMmeasured(ByVal Index As Integer, ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CHEMmeasured = 5
    mvarCHEMmeasured(Index) = vData
End Property

Public Property Get CHEMmeasured(ByVal Index As Integer) As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CHEMmeasured
    CHEMmeasured = mvarCHEMmeasured(Index)
End Property

Public Property Let TotalStatedKG(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.TotalStatedKG = 5
    mvarTotalStatedKG = vData
End Property

Public Property Get TotalStatedKG() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.TotalStatedKG
    TotalStatedKG = mvarTotalStatedKG
End Property

Public Property Let TotalMeasuredKG(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.TotalMeasuredKG = 5
    mvarTotalMeasuredKG = vData
End Property

Public Property Get TotalMeasuredKG() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.TotalMeasuredKG
    TotalMeasuredKG = mvarTotalMeasuredKG
End Property

Public Property Let TotalQuant(ByVal vData As Single)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.TotalQuant = 5
    mvarTotalQuant = vData
End Property

Public Property Get TotalQuant() As Single
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.TotalQuant
    TotalQuant = mvarTotalQuant
End Property

Public Function Clear()

    Dim i As Integer

    mvarMixNum = 0
    mvarExpNum = 0
    mvarExpStartTime = ""
    mvarMixReadyTime = ""
    mvarStampDate = "00:00:00"
    mvarDispName = ""
    mvarOrderCode = 0
    mvarOrderDate = ""
    mvarOrderQuant = 0
    mvarExpQuant = 0
    mvarExpNumFromOrder = 0
    mvarMixNumFromOrder = 0
    mvarClntTitle = ""
    mvarClntIdent = ""
    mvarClntWorksite = ""
    mvarWorksiteDist = 0
    mvarDrvTitle = ""
    mvarDrvCarNum = ""
    mvarDrvCapacity = 0
    mvarRecTitle = ""
    mvarRecKind = ""
    mvarRecClass = ""
    mvarRecClassK = ""
    mvarRecClassV = ""
    mvarRecClassH = ""
    mvarRecClassP = ""
    mvarRecEDM = 0
    
    For i = 1 To 6
        mvarIMname(i) = ""
        mvarIMstated(i) = 0
        mvarIMmeasured(i) = 0
        mvarCHEMname(i) = ""
        mvarCHEMstated(i) = 0
        mvarCHEMmeasured(i) = 0
    Next i

    For i = 1 To 4
        mvarSCRname(i) = ""
        mvarSCRstated(i) = 0
        mvarSCRmeasured(i) = 0
    Next i

    For i = 1 To 2
        mvarWATname(i) = ""
        mvarWATstated(i) = 0
        mvarWATmeasured(i) = 0
    Next i

    mvarTotalStatedKG = 0
    mvarTotalMeasuredKG = 0
    mvarTotalQuant = 0
End Function

