Attribute VB_Name = "fnctnsRep"
    
    Public Const CSIDL_COMMON_APPDATA = &H23
    Public Const CB_SHOWDROPDOWN = &H14F
    Public Const CB_FINDSTRING = &H14C
    Public Const LOCALE_SDECIMAL = &HE
    Public Const HKEY_CURRENT_USER = &H80000001
    
    Public Const DbaseName = "postgres"
    Public Const DbaseUser = "reporter"
    Public Const PassConnStr = "reporter"
    Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
    
    Dim picTemp As PictureBox
    
    Private Const WM_USER = &H400
    Private Const EM_FORMATRANGE As Long = WM_USER + 57

    Private Const PHYSICALOFFSETX As Long = 112
    Private Const PHYSICALOFFSETY As Long = 113
    
    Public Const KEY_READ = &H20019

'константи за регистъра
    Public Const ErrRes = "whatthefuckareyoudoinghere"
    Public Const PlaceProgSettings = "TipReport"
    Public Const PlaceProgSet1 = "Software\VB and VBA Program Settings\TipReport\Form1Set"
    Public Const PlaceProgSet2 = "Software\VB and VBA Program Settings\TipReport\Form2Set"
    Public Const PlaceProgSet3 = "Software\VB and VBA Program Settings\TipReport\Form3Set"
    Public Const PlaceForm1 = "Form1Set"
    Public Const PlaceForm2 = "Form2Set"
    Public Const PlaceForm3 = "Form3Set"
    Public Const PlaceProgAllow = "Software\VB and VBA Program Settings\TipReport\Allow"
    Public Const PlaceAllow = "Allow"
    
    Public MachineNumber As Integer
    Public MachineOther As Integer
    
    Public AddMach As Boolean
    Public Choice As Boolean
    
'променливи за настойка на формите за печат на бележки
    Public rDist As Integer
    Public rRecType As Integer
    Public rVol As Integer
    Public rW As Integer
    Public rOrdVol As Integer
    Public rClass As Integer
    Public rClassK As Integer
    Public rClassV As Integer
    Public rClassH As Integer
    Public rClassP As Integer
    Public rCem1 As Integer
    Public rCem2 As Integer
    Public rCem3 As Integer
    Public rChem1 As Integer
    Public rChem2 As Integer
    Public rChem3 As Integer
    Public rEDM As Integer
    Public rMixTime As Integer
    Public rExpTime As Integer
    Public rRealVol As Integer
    Public PrintAnyForm As Boolean
    Public rQinForm  As Integer
    
'------------------------------------------
   
    Public ConStr As String 'connection string за връзка с база данни
    Public IPConnStr As String
    Public MachName As String

    Public frmDBdata As String
    
    Public TxtInfoCap As String
    Public TxtVerCap As String
    
    Public frmReportPanelCap As String
    Public frmAdPanel As String
    
    Public MsgAnotherRun As String
    Public MsgConfigNotFound As String
    Public MsgCallTIP As String
    Public MsgDBConnEst As String
    Public MsgErrBx As String
    Public MsgErrNoRec As String
    Public MsgNoDBConn As String
    Public MsgNoPayment As String
    Public MsgNoRecords As String
    Public MsgTablesNotFound As String
    Public msgChooseFilter As String
    
    Public lblEntIP As String
    Public lblStDate As String
    Public lblEndDate As String
    
    Public UniCancel As String
    Public UniEnter As String
    Public UniExit As String
    Public UniOK As String

    Public uniAdmin As String
    Public uniAdd As String
    Public uniBG As String
    Public uniCap As String
    Public uniCapacity As String
    Public uniChem As String
    Public uniClass As String
    Public uniClassK As String
    Public uniClassV As String
    Public uniClassH As String
    Public uniClassP As String
    Public uniClnt As String
    Public uniClnts As String
    Public uniCode As String
    Public uniComInfo As String
    Public uniConcPlant As String
    Public uniConMat As String
    Public uniDate As String
    Public uniDateOrd As String
    Public uniDateReady As String
    Public uniDay As String
    Public uniDays As String
    Public uniDelivered As String
    Public uniDisp As String
    Public uniDlvrs As String
    Public uniDrv As String
    Public uniDrvReg As String
    Public uniDrvs As String
    Public uniEDM As String
    Public uniEmpty As String
    Public uniEx As String
    Public uniExped As String
    Public uniExpeds As String
    Public uniFax As String
    Public uniFirm As String
    Public uniForm1 As String
    Public uniForm2 As String
    Public uniForm3 As String
    Public uniHave As String
    Public uniHour As String
    Public uniHourOut As String
    Public uniIM As String
    Public uniKmShort As String
    Public uniLoad As String
    Public uniLog As String
    Public uniMade As String
    Public uniMat As String
    Public uniMats As String
    Public uniMeasured As String
    Public uniMix As String
    Public uniMixes As String
    Public uniMod As String
    Public uniMOL As String
    Public uniNew As String
    Public uniNo As String
    Public uniNote As String
    Public uniNotes As String
    Public uniNr As String
    Public uniNm As String
    Public uniObj As String
    Public uniOld As String
    Public uniOrd As String
    Public uniOrdered As String
    Public uniOrds As String
    Public uniOther As String
    Public uniQ As String
    Public uniQuant As String
    Public uniRealQinForm As String
    Public uniRec As String
    Public uniRecType As String
    Public uniRecs As String
    Public uniResults As String
    Public uniRevision As String
    Public uniRevisions As String
    Public uniRevisor As String
    Public uniSave As String
    Public uniSendingPrinter As String
    Public uniSettings As String
    Public uniSheet As String
    Public uniSold As String
    Public uniSup As String
    Public uniSups As String
    Public uniTel As String
    Public uniTempResults As String
    Public uniTimeMixShort As String
    Public uniTimePourShort As String
    Public uniTotal As String
    Public uniTotalKg As String
    Public uniTown As String
    Public uniType As String
    Public uniTypeDoc As String
    Public uniWat As String
    Public uniYes As String

    Public repOperDay As String
    Public repOperAll As String
    Public repDrvExped As String
    Public repDrvDay As String
    Public repDrvAll As String
    Public repMatSold As String
    Public repClntMix As String
    Public repClntExped As String
    Public repClntOrd As String
    Public repDailyProd As String
    Public repDailyExped As String
    Public repDailyRep As String
    
    Public btLoad As String
    Public btPrint As String
    Public btExport As String
    
    
'променливи за работните файлове на програмата
    Public PathCore As String
    Public DBSetFile As String
    Public ConfirmityFile As String
    Public InfoFile As String
    Public LangSetFile As String
    Public LangBgFile As String
    Public LangRusFile As String
    Public LangEnFile As String
'------------------------------------------------

'променливи за параметри на машината
    Public n1s1 As Integer
    Public n1s2 As Integer
    Public n1s3 As Integer
    Public n1s4 As Integer

    Public n2s1 As Integer
    Public n2s2 As Integer
    Public n2s3 As Integer
    Public n2s4 As Integer
'------------------------------------------------
Public Declare Function GetThreadLocale Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
    (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, _
    ByVal cchData As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function SHGetFolderPath _
                        Lib "shfolder.dll" Alias "SHGetFolderPathA" _
                        (ByVal hwndOwner As Long, _
                         ByVal nFolder As Long, _
                         ByVal hToken As Long, _
                         ByVal dwReserved As Long, _
                         ByVal lpszPath As String) As Long

Public Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal nWidth As Long, _
                ByVal nHeight As Long, _
                ByVal hSrcDC As Long, _
                ByVal xSrc As Long, _
                ByVal ySrc As Long, _
                ByVal dwRop As Long) As Long
                
Public Declare Function GetDeviceCaps Lib "gdi32" _
  (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Type RECT
   Left           As Long
   Top            As Long
   Right          As Long
   Bottom         As Long
End Type

Public Type CharRange
   cpMin          As Long
   cpMax          As Long
End Type

Public Type FormatRange
   hdc            As Long
   hdcTarget      As Long
   rc             As RECT
   rcPage         As RECT
   chrg           As CharRange
End Type

Public Function LoadLang()
'функция за зареждане на езика

'    Dim intEmpFileNbr1 As Integer
'    Dim LangSet As String
'
'    intEmpFileNbr1 = FreeFile
'
'    If Dir(LangSetFile) <> "" Then
'        Open LangSetFile For Input As intEmpFileNbr1
'        Do Until EOF(intEmpFileNbr1)
'            Input #intEmpFileNbr1, LangSet
'        Loop
'        Close #intEmpFileNbr1
'    Else
        TxtInfoCap = "Софтуер за визуализация, печат и експорт на справки за бетонови стопанства"
        TxtVerCap = "Справочник ТИП-Репорт v1.2/2014"
        MsgAnotherRun = "Програмата вече е активна!"
        MsgCallTIP = "Свържете се с ТИП-Сервиз ЕООД!"
        MsgConfigNotFound = "Не е открита конфигурация на машината!"
        MsgDBConnEst = "Връзката базата данни е осъществена!"
        MsgErrBx = "Грешка"
        MsgErrNoRec = "Няма записи"
        MsgNoDBConn = "Няма връзка с база данни!"
        MsgNoPayment = "Лицензът на програмата е прекратен поради непостъпило окончателно плащане!"
        MsgNoRecords = "Няма записи за този период / дата!"
        MsgTablesNotFound = "Следните таблици не бяха открити в базата данни: "
        msgChooseFilter = "Изберете основен филтър за търсенето"
        lblEntIP = "IP за достъп до база данни (локално = 127.0.0.1)"
        lblStDate = "от дата"
        lblEndDate = "до дата"
        frmAdPanel = "Администраторски Панел ТИП-Репорт"
        frmReportPanelCap = "Справки - ТИП-Репорт v1.2   --- +359 887507525 ---"
        frmDBdata = "Достъп до база данни"
        UniCancel = "Отказ"
        UniEnter = "Вход"
        UniExit = "Изход"
        UniOK = "OK"
        uniAdd = "Адрес"
        uniAdmin = "Администратор"
        uniBG = "БУЛСТАТ"
        uniCap = "Кап."
        uniCapacity = "Капацитет"
        uniChem = "Химическа добавка"
        uniClass = "Кл. якост"
        uniClassK = "Кл. конс."
        uniClassV = "Кл. въз."
        uniClassH = "Кл. хл."
        uniClassP = "Вод."
        uniClnt = "Клиент"
        uniClnts = "Клиенти"
        uniCode = "Код"
        uniComInfo = "Информация за фирмата"
        uniConcPlant = "Бетонов възел"
        uniConMat = "Свързващо вещество"
        uniDate = "Дата"
        uniDateOrd = "Дата приемане"
        uniDateReady = "Дата готовност"
        uniDay = "ден"
        uniDays = "дни"
        uniDelivered = "Доставено [t]"
        uniDisp = "Диспечер"
        uniDlvrs = "Доставки"
        uniDrv = "Водач"
        uniDrvReg = "Кола"
        uniDrvs = "Водачи"
        uniEDM = "ЕДМ"
        uniEmpty = "празна течка"
        uniExped = "Експедиция"
        uniExpeds = "Експедиции"
        uniEx = "експ."
        uniFax = "Факс"
        uniFirm = "Фирма"
        uniForm1 = "Форма 1"
        uniForm2 = "Форма 2"
        uniForm3 = "Форма 3"
        uniHave = "Наличност [t]"
        uniHour = "час"
        uniHourOut = "час изсипване"
        uniIM = "Инертен материал"
        uniKmShort = "Разст."
        uniLoad = "Заредено"
        uniLog = "Регистър"
        uniMade = "изпълнено"
        uniMat = "Материал"
        uniMats = "Материали"
        uniMeasured = "Измерено"
        uniMix = "замес"
        uniMixes = "замеси"
        uniMod = "Марка/Модел"
        uniMOL = "МОЛ"
        uniNew = "Ново"
        uniNm = "Име"
        uniNo = "не"
        uniNote = "Бележка"
        uniNotes = "Бележки"
        uniNr = "No."
        uniObj = "Обект"
        uniOld = "старо"
        uniOrd = "заявка"
        uniOrdered = "заявено"
        uniOrds = "Заявки"
        uniOther = "Други"
        uniQuant = "количество"
        uniQ = "кол."
        uniRealQinForm = "Визуализация на реалното измерено количество бетон в справките"
        uniRec = "Рецепта"
        uniRecType = "вид р-р"
        uniRecs = "Рецепти"
        uniResults = "Резултати"
        uniRevision = "Ревизия"
        uniRevisions = "Ревизии"
        uniRevisor = "Ревизор"
        uniSave = "Запис"
        uniSendingPrinter = "Изпращане за печат..."
        uniSettings = "Настройки"
        uniSheet = "лист "
        uniSold = "Разход [t]"
        uniSup = "Доставчик"
        uniSups = "Доставчици"
        uniTel = "Телефон"
        uniTempResults = "Временни резултати"
        uniTimeMixShort = "tmix"
        uniTimePourShort = "tout"
        uniTimes = "Времена"
        uniTotal = "Тотал: "
        uniTotalKg = "Общо кг"
        uniTown = "град"
        uniType = "Тип"
        uniTypeDoc = "Документ вид"
        uniWat = "Вода"
        uniYes = "да"
        repOperDay = "Диспечери по дни"
        repOperAll = "Всички диспечери"
        repDrvExped = "Водачи по експедиции"
        repDrvDay = "Водачи по дни"
        repDrvAll = "Всички водачи"
        repMatSold = "Разход на материали"
        repClntMix = "Клиенти по замеси"
        repClntExped = "Клиенти по експедиции"
        repClntOrd = "Клиенти по заявки"
        repDailyProd = "Дневник производство"
        repDailyRep = "Дневна справка бетони"
        repDailyExped = "Дневник експедиции"
        btLoad = "Зареди справка"
        btPrint = "Печат"
        btExport = "Експорт"
        
'    End If
End Function

Public Function CheckRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
'проверка за съществуване на ключ в регистъра
    
    Dim handle As Long
'~~> Try to open the key
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) = 0 Then
'~~> The key exists
        CheckRegistryKey = True
'~~> Close it before exiting
        RegCloseKey handle
    End If
End Function

Public Function strGetCommonAppDataPath() As String
'функция за търсене на път на програма
    Dim strPath As String

    strPath = Space$(512)
    Call SHGetFolderPath(0, CSIDL_COMMON_APPDATA, 0, 0, strPath)
    strPath = Left$(strPath, InStr(strPath, vbNullChar) - 1)

    strGetCommonAppDataPath = strPath
End Function

Public Function strGetDesktopPath() As String

    'функция за търсене на път на програма
    Dim strPath As String

    strPath = Space$(512)
    Call SHGetFolderPath(0, CSIDL_DESKTOP, 0, 0, strPath)
    strPath = Left$(strPath, InStr(strPath, vbNullChar) - 1)

    strGetDesktopPath = strPath
End Function

Public Function isRunning(ByVal Process As String) As Boolean
    'проверка дали е стартиран даден процес

    Dim objWMIService, colProcesses

    Set objWMIService = GetObject("winmgmts:")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name='" & Process & "'")

    If colProcesses.count Then
        isRunning = True
    Else
        isRunning = False
    End If

End Function


Public Function GetDecimalSep() As String
'функция за извикване на сепаратора за дробни числа от настройките на системата
    
    Dim LCID As Integer
    Dim Data As String
    Dim ret As Integer
    Dim DataLen As Long
 
' Get the local decimal seperator
' Find the threads local
    LCID = GetThreadLocale
 
' Find the required size of the output variables
    ret = GetLocaleInfo(LCID, LOCALE_SDECIMAL, Data, DataLen)
 
    If ret <> 0 Then
     ' prepare the output variable
        DataLen = ret
        Data = Space(DataLen)
        ret = GetLocaleInfo(LCID, LOCALE_SDECIMAL, Data, DataLen)
    Else
     ' Error no data found
     ' enter some good error handling here, using GetLastError()
    End If
' Remove the null terminator from the string
    GetDecimalSep = Left(Data, DataLen - 1)
End Function

Public Function rDs(ByVal str As String)
'функция за смяна надесетичния сепаратор при четене спрямо този от настройките на компютъра

    Dim DecSep As String
    
    DecSep = GetDecimalSep
    
    If Len(str) = 0 Then
        rDs = "0"
        Exit Function
    End If
    
    If InStr(str, ",") <> 0 And DecSep <> "," Then
        rDs = Replace(str, ",", DecSep)
    ElseIf InStr(str, ".") <> 0 And DecSep <> "." Then
        rDs = Replace(str, ".", DecSep)
    Else
        rDs = str
    End If
End Function

Public Function ARound(ByVal MyNumber, ByVal Deci)
'функция за закръгление на числа
      
      ARound = Int(MyNumber * 10 ^ Deci + 1 / 2) / 10 ^ Deci
End Function

Public Sub AutoColW(ListViewTemp As ListView)
'автоформатиране на таблица в listview според текста в клетките и заглавките
    
    Const LVM_FIRST = &H1000
    Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
    Const LVSCW_AUTOSIZE_USEHEADER = -2
    Dim I As Long
  
    With ListViewTemp
        SendMessage .hWnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE_USEHEADER
        For I = 1 To .ColumnHeaders.count - 1
            SendMessage .hWnd, LVM_SETCOLUMNWIDTH, I, ByVal LVSCW_AUTOSIZE_USEHEADER
        Next
    End With
End Sub

Public Function cmbAutoComplete(ByRef cboComplete As ComboBox, ByVal KeyAscii As Integer, Optional ByVal bLimitToList As Boolean = False) As Long
    Dim lRetVal As Long
    Dim sSearch As String
    Const CB_ERR = (-1), CB_FINDSTRING = &H14C
    
    On Error GoTo ErrFailed
    If cboComplete.Style <> vbComboDropdown Then
        Debug.Print "Error in ComboAutoComplete. Combo must be of the style vbComboDropdown..."
        Debug.Assert False
        'Return the KeyAscii
        ComboAutoComplete = KeyAscii
        Exit Function
    End If
    If KeyAscii = 8 Then
        'Pressed delete
        If cboComplete.SelStart <= 1 Then
            'Last character, clear combo.
            cboComplete.Text = ""
            ComboAutoComplete = 0
            Exit Function
        End If
        'Delete text
        If cboComplete.SelLength = 0 Then
            'Delete a single character
            sSearch = UCase$(Left$(cboComplete.Text, Len(cboComplete) - 1))
        Else
            'Delete the selected text
            sSearch = Left$(cboComplete.Text, cboComplete.SelStart - 1)
        End If
    ElseIf KeyAscii < 32 Or KeyAscii > 255 Then
        'Invalid keyboard characters
        Exit Function
    Else
        'Append the new text to the combo text
        SendMessage cboComplete.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&
        If cboComplete.SelLength = 0 Then
            'Append a character
            sSearch = UCase$(cboComplete.Text & Chr$(KeyAscii))
        Else
            'Insert a character
            sSearch = Left$(cboComplete.Text, cboComplete.SelStart) & Chr$(KeyAscii)
        End If
    End If
    'Find the closest match
    lRetVal = SendMessage(cboComplete.hWnd, CB_FINDSTRING, -1, ByVal sSearch)

    If lRetVal = CB_ERR Then
        'Did not find a matching item in list
        If bLimitToList = True Then
            'Block the KeyAscii
            ComboAutoComplete = 0
        Else
            'Return the KeyAscii
            ComboAutoComplete = KeyAscii
        End If
    Else
        'Found a matching item in list
        cboComplete.ListIndex = lRetVal
        cboComplete.SelStart = Len(sSearch)
        cboComplete.SelLength = Len(cboComplete.Text) - cboComplete.SelStart
        ComboAutoComplete = 0
    End If
    
    Exit Function
    
ErrFailed:
    'Return the keycode
    ComboAutoComplete = KeyAscii
End Function

Public Function PrintLVPic(lvw As ListView, Orient As Integer, HeadPrnt As Boolean, _
NowPrnt As Boolean, PageNumPrnt As Boolean, Optional NamePage As String = "", _
Optional TopMargPerc As Integer = 500, Optional LeftMargPerc As Integer = 500, _
Optional not1 As Integer = 0, Optional not2 As Integer = 0, Optional not3 As Integer = 0)
'функция за принтиране на ListView към PictureBox
'извиква след всяка страница функция за печат на PictureBox като му прави AutoFit
'към A4 според избрана ориентация 1-портретно, 2-пейзажно
'преработена за печат на много страници
'избира се True / False за печат на хедърите на ListView
'задава се заглавие на страница - по избор
'печат на дата и час - True/False
'пропуска печат до 3 колони по избор на номера им
'за да работи функцията трябва да има форма - frmPrint и PictureBox - pbPrint на нея

Const MARGIN = 70
Const COL_MARGIN = 70

Dim ymin As Single
Dim ymax As Single
Dim xmin As Single
Dim xmax As Single
Dim num_cols As Integer
Dim list_item As ListItem
Dim I As Integer
Dim num_subitems As Integer
Dim col_wid() As Single
Dim X As Single
Dim Xpage As Single
Dim Y As Single
Dim line_hgt As Single
Dim start As Integer
Dim listCount As Integer

frmPrint.Show
frmPrint.barPrint.Value = frmPrint.barPrint.Min
listCount = 0
Line = 1
Again:

    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
ymin = 0
ymax = 0
xmin = 0
xmax = 0
num_cols = 0
I = 0
X = 0
Xpage = 0
Y = 0
line_hgt = 0
listCount = listCount + 1


MousePointer = vbHourglass

    'определяме шрифта сега за да може правилно да изчислим ширините на колоните
    frmPrint.pbPrint.FontSize = 10

    ' ******************
    ' Get column widths.
    num_cols = lvw.ColumnHeaders.count
    ReDim col_wid(1 To num_cols)

    ' Check the column headers.
    For I = 1 To num_cols
        If I <> not1 And I <> not2 And I <> not3 Then
        col_wid(I) = frmPrint.pbPrint.TextWidth(lvw.ColumnHeaders(I).Text)
        End If
    Next I

    ' Check the items.
    num_subitems = num_cols - 1
    For Each list_item In lvw.ListItems
        ' Check the item.
        If col_wid(1) < frmPrint.pbPrint.TextWidth(list_item.Text) Then _
           col_wid(1) = frmPrint.pbPrint.TextWidth(list_item.Text)

        ' Check the subitems.
        For I = 1 To num_subitems
            If I <> not1 - 1 And I <> not2 - 1 And I <> not3 - 1 Then
            If col_wid(I + 1) < frmPrint.pbPrint.TextWidth(list_item.SubItems(I)) Then _
                col_wid(I + 1) = frmPrint.pbPrint.TextWidth(list_item.SubItems(I))
            End If
        Next I
    Next list_item
    
    ' Add a column margin.
    For I = 1 To num_cols
        If I <> not1 And I <> not2 And I <> not3 Then
        col_wid(I) = col_wid(I) + COL_MARGIN
        End If
    Next I
    
    'изчисляваме ширината на PictureBox-a според ширините на колоните
    Xpage = MARGIN
    For I = 1 To num_subitems + 1
        Xpage = Xpage + col_wid(I)
    Next I
    frmPrint.pbPrint.Width = Xpage + MARGIN + 100
    
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = Orient
    Printer.Print

    'ако ширината на принтера е по-голяма от нужното пространство
    'вземаме нея за ширина на PictureBox-a
    If Printer.ScaleWidth >= frmPrint.pbPrint.Width Then frmPrint.pbPrint.Width = Printer.ScaleWidth
    
    'притнираме заглавието на страницата ако има такова
    If NamePage <> "" Then
        frmPrint.pbPrint.CurrentX = X
        frmPrint.pbPrint.CurrentY = Y
        frmPrint.pbPrint.FontSize = 12
        line_hgt = frmPrint.pbPrint.TextHeight("X")
        frmPrint.pbPrint.Print NamePage
        ymin = frmPrint.pbPrint.CurrentY + line_hgt / 2
        X = frmPrint.pbPrint.TextWidth(NamePage) + MARGIN
    End If
    
    'принтираме номер на страница ако е избрано
    If PageNumPrnt = True Then
        frmPrint.pbPrint.CurrentX = X
        frmPrint.pbPrint.CurrentY = Y
        frmPrint.pbPrint.FontSize = 12
        line_hgt = frmPrint.pbPrint.TextHeight("X")
        frmPrint.pbPrint.Print " - " & uniSheet & listCount
        ymin = frmPrint.pbPrint.CurrentY + line_hgt / 2
        X = X + frmPrint.pbPrint.TextWidth(" - " & uniSheet & listCount) + MARGIN
    End If
    
    'принтираме дата и час ако е избрано
    If NowPrnt = True Then
        frmPrint.pbPrint.CurrentX = X
        frmPrint.pbPrint.CurrentY = Y
        frmPrint.pbPrint.FontSize = 12
        line_hgt = frmPrint.pbPrint.TextHeight("X")
        frmPrint.pbPrint.Print " - " & uniDate & ": " & Now
        ymin = frmPrint.pbPrint.CurrentY + line_hgt / 2
        X = frmPrint.pbPrint.TextWidth(NamePage) + MARGIN
    End If
    
    frmPrint.pbPrint.FontSize = 10
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    ' *************************
    ' Print the column headers.
    frmPrint.pbPrint.CurrentY = ymin + MARGIN
    frmPrint.pbPrint.CurrentX = MARGIN
    X = xmin + MARGIN
    For I = 1 To num_cols
        If I <> not1 And I <> not2 And I <> not3 Then
        frmPrint.pbPrint.CurrentX = X
            If HeadPrnt = True Then
                frmPrint.pbPrint.Print FittedText(lvw.ColumnHeaders(I).Text, col_wid(I));
            End If
        X = X + col_wid(I)
        End If
    Next I
    
    xmax = X + MARGIN
    frmPrint.pbPrint.Print
    line_hgt = frmPrint.pbPrint.TextHeight("X")
    Y = ymin
    
    If HeadPrnt = True Then
        Y = frmPrint.pbPrint.CurrentY + line_hgt / 2
        frmPrint.pbPrint.Line (xmin, Y)-(xmax, Y)
        Y = Y + line_hgt / 2
        Y = frmPrint.pbPrint.CurrentY
    End If

    ' Print the rows.
    num_subitems = num_cols - 1
    Printer.Print
    
    If Orient = 2 Then
        frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 0.7
    ElseIf Orient = 1 Then
        frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 1.41
    End If
    
    For start = Line To lvw.ListItems.count
        X = xmin + MARGIN

        ' Print the item.
        frmPrint.pbPrint.CurrentX = X
        frmPrint.pbPrint.CurrentY = Y
        frmPrint.pbPrint.Print FittedText(lvw.ListItems.Item(start), col_wid(1));
        X = X + col_wid(1)

        ' Print the subitems.
        For I = 1 To num_subitems
        If I <> not1 - 1 And I <> not2 - 1 And I <> not3 - 1 Then
            frmPrint.pbPrint.CurrentX = X
            frmPrint.pbPrint.Print FittedText(lvw.ListItems.Item(start).SubItems(I), col_wid(I + 1));
            X = X + col_wid(I + 1)
        End If
        Next I
        
        If lvw.ListItems.Item(start).SubItems(1) = "" Then
            frmPrint.pbPrint.Line (xmin, Y)-(xmax, Y)
        End If
        
        Y = Y + line_hgt * 1.2
        If Y >= frmPrint.pbPrint.Height - (ymin + 100) Then
            PrintLVPic = start
            GoTo EndOfPage
        Else
            frmPrint.barPrint.Value = frmPrint.barPrint.Max
            PrintLVPic = -1
        End If
    Next start
    
EndOfPage:
    ymax = Y
    ' Draw lines around it all.
    frmPrint.pbPrint.Line (xmin, ymin)-(xmax, ymax), , B

    X = xmin + MARGIN / 2
    For I = 1 To num_cols - 1
    If I <> not1 And I <> not2 And I <> not3 Then
        X = X + col_wid(I)
        frmPrint.pbPrint.Line (X, ymin)-(X, ymax)
    End If
    Next I
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    PrintThePicture frmPrint, frmPrint.pbPrint, 95, LeftMargPerc, TopMargPerc
    
    MousePointer = vbDefault
    
    If PrintLVPic > 0 Then
        Printer.NewPage
        Line = start + 1
        GoTo Again
    Else
        Printer.EndDoc
        frmPrint.barPrint.Value = frmPrint.barPrint.Max
        Unload frmPrint
        Exit Function
    End If
    
End Function

Public Function FittedText(ByVal txt As String, ByVal wid As Single) As String
' Return as much text as will fit in this width.
    
    Do While Printer.TextWidth(txt) > wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    FittedText = txt
End Function

Private Sub ScalePic(ByVal Percentage As Single, ByVal sngTop As Single, ByVal sngLeft As Single)
'Dim intHeight As Integer
'Dim intWidth As Integer
Dim sngRatio As Single
Dim intLeft As Integer
Dim intTop As Integer

    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Percentage = Percentage / 100
'    sngTop = sngTop / 100
'    sngLeft = sngLeft / 100
    
    '   Scale the picture to either use the full width
    '   or the full height of the page.
    If frmPrint.pbPrint.Width > (Printer.ScaleWidth * Percentage) Then
        sngRatio = (Printer.ScaleWidth * Percentage) / frmPrint.pbPrint.Width
    Else
        sngRatio = 1
    End If
    
    If frmPrint.pbPrint.Height * sngRatio > (Printer.ScaleHeight * Percentage) Then
        sngRatio = (Printer.ScaleHeight * Percentage) / frmPrint.pbPrint.Height
    Else
    End If
    
    '   Center the picture on the page.
'    intLeft = (Printer.ScaleWidth - (picTemp.Width * sngRatio)) * sngLeft
'    intTop = (Printer.ScaleHeight - (picTemp.Height * sngRatio)) * sngTop
    intLeft = sngLeft
    intTop = sngTop
    
    '   send the picture to the printer
    
    Printer.PaintPicture frmPrint.pbPrint.Image, intLeft, intTop, frmPrint.pbPrint.Width * sngRatio, frmPrint.pbPrint.Height * sngRatio
'    Printer.EndDoc
    
    '   Cleanup
    Set picTemp = Nothing
    frmPrint.pbPrint.Cls
    frmPrint.pbPrint.Refresh
'    frmPrint.Controls.Remove "picTemp"
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    'веднага след приключване на фунцията трябва да зададем Printer.enddoc

End Sub

Public Sub PrintThePicture(frm As Form, Pic As PictureBox, _
                        Optional PercentOfPage As Integer = 100, _
                        Optional LeftMarginPercent As Integer = 0, _
                        Optional TopMarginPercent As Integer = 0)
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

    ScalePic PercentOfPage, TopMarginPercent, LeftMarginPercent
End Sub

Public Function ExportToExcel(lvw As MSComctlLib.ListView) As Boolean
'функция за експортиране на ListView в Excel
 
    Dim objExcel As Object
    Dim objWorkbook As Object
    Dim objWorksheet As Object
    Dim objRange As Object
     
'    Dim lngResults As Long
    Dim I As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    Dim Fname As String
 
    Fname = lvw.Name
    
    MousePointer = vbHourglass
    
' 'If there are no selected items in the listview control
'    If lvw.SelectedItem Is Nothing Then
'        MousePointer = vbDefault
'        MsgBox MsgNoSelection, vbOKOnly Or vbInformation, MsgErrBx
'        GoTo ExitFunction
'    End If
 
' 'Ask the user if they want to export just the selected items
'    MousePointer = vbDefault
'    lngResults = MsgBox(MsgConfExp1, vbYesNoCancel Or vbQuestion, MsgConf1ExpBx)
'    If lngResults = vbCancel Then
'        GoTo ExitFunction
'    End If
 
    Screen.MousePointer = vbHourglass
 
 'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If Err.Number > 0 Then
        MsgBox MsgNoExcel, vbOKOnly Or vbCritical, MsgErrBx
        GoTo ExitFunction
    End If
 
    On Error GoTo HANDLE_ERROR
 ' Don't allow user to affect workbook
    objExcel.Interactive = False
    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
 
    objExcel.WindowState = vbMaximized
 
    Set objWorkbook = objExcel.Workbooks.Add
    Set objWorksheet = objWorkbook.Sheets(1)
 
    intCounter = 0
    Set objRange = objWorksheet.Rows(1)
    objRange.Font.Size = 10
    objRange.Font.Bold = True
    For I = 1 To lvw.ColumnHeaders.count
        If lvw.ColumnHeaders(I).Width <> 0 Then
        ' Create an array of visible column indexes
            intColumns = intColumns + 1
            ReDim Preserve intVisibleColumns(1 To intColumns)
            intVisibleColumns(intColumns) = I
            objRange.cells(1, intColumns) = lvw.ColumnHeaders(I).Text
'            With objWorksheet.Columns(intColumns)
'                Select Case LCase$(lvw.ColumnHeaders(i).Tag)
'                ' If tag is empty, format as text
'                    Case "string" ', ""
'                        .NumberFormat = "@"
'                    Case "number"
'                        .NumberFormat = "###0,000"
'                        .HorizontalAlignment = xlRight
'                    Case "date"
'                        .NumberFormat = "mm/dd/yyyy"
''                        .HorizontalAlignment = xlRight
'                End Select
'            End With
         End If
     Next I
 
 ' Dimension array to number of listitems
    ReDim strArray(1 To lvw.ListItems.count, 1 To intColumns)
 
    intCounter = 0
    intStartRow = 2
    For Each itm In lvw.ListItems
    ' A response of vbNo meant to export all the items
'        If lngResults = vbNo Or itm.Selected Then
        ' increment the number of selected rows
            intCounter = intCounter + 1
            For I = 1 To intColumns
                If intVisibleColumns(I) = 1 Then
                    strArray(intCounter, 1) = itm.Text
                Else
                    strArray(intCounter, I) = itm.SubItems(intVisibleColumns(I) - 1)
                End If
            Next I
'        End If
    Next itm
 
 ' Send entire array to Excel range
    With objWorksheet
        .Range(.cells(2, 1), _
        .cells(2 + intCounter - 1, intColumns)) = strArray
    End With
 
    objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
'    objWorkbook.SaveAs App.Path & "\" & Fname & DayToday '& ".xls"
 
    ExportToExcel = True

ExitFunction:
    Screen.MousePointer = vbDefault
    Exit Function
HANDLE_ERROR:
    MsgBox MsgErrExcel & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description, vbOKOnly Or vbCritical, MsgErrBx
    Set objRange = Nothing
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    objExcel.Quit
    GoTo ExitFunction
End Function

Public Function PrintRTF(rtf As RichTextBox, nnLeftMarginWidth _
 As Long, nnTopMarginHeight As Long, nnRightMarginWidth As _
 Long, nnBottomMarginHeight As Long) As Boolean
'форматирано принтиране на RTFBox

' #VBIDEUtils#************************************************
' * Programmer Name  : Waty Thierry
' * Web Site         : www.geocities.com/ResearchTriangle/6311/
' * E-Mail           : waty.thierry@usa.net
' * Date             : 30/10/98
' * Time             : 14:43
' * Module Name      : Main_Module
' * Module Filename  : Main.bas
' * Procedure Name   : PrintRTF
' * Parameters       :
' *                    rtf As RichTextBox
' *                    nnLeftMarginWidth As Long
' *                    nnTopMarginHeight As Long
' *                    nnRightMarginWidth As Long
' *                    nnBottomMarginHeight As Long
' ***************************************************************
' * Comments         :
' *
' *
' *************************************************************
On Error GoTo ErrorHandler
Dim nLeftOffset      As Long
Dim nTopOffset       As Long
Dim nLeftMargin      As Long
Dim nTopMargin       As Long
Dim nRightMargin     As Long
Dim nBottomMargin    As Long
Dim fr               As FormatRange
Dim rcDrawTo         As RECT
Dim rcPage           As RECT
Dim nTextLength      As Long
Dim nNextCharPos     As Long
Dim nRet             As Long

MousePointer = vbHourglass

frmPrint.pbPrint.Print Space(1)
nLeftOffset = frmPrint.pbPrint.ScaleX(GetDeviceCaps(frmPrint.pbPrint.hdc, _
   PHYSICALOFFSETX), vbPixels, vbTwips)
   
nTopOffset = frmPrint.pbPrint.ScaleY(GetDeviceCaps(frmPrint.pbPrint.hdc, _
   PHYSICALOFFSETY), vbPixels, vbTwips)
   
nLeftMargin = nnLeftMarginWidth - nLeftOffset
nTopMargin = nnTopMarginHeight - nTopOffset
nRightMargin = (frmPrint.pbPrint.Width - nnRightMarginWidth) _
   - nLeftOffset
   
nBottomMargin = (frmPrint.pbPrint.Height - nnBottomMarginHeight) _
   - nTopOffset
   
rcPage.Left = 0
rcPage.Top = 0
rcPage.Right = frmPrint.pbPrint.ScaleWidth
rcPage.Bottom = frmPrint.pbPrint.ScaleHeight
rcDrawTo.Left = nLeftMargin
rcDrawTo.Top = nTopMargin
rcDrawTo.Right = nRightMargin
rcDrawTo.Bottom = nBottomMargin
fr.hdc = frmPrint.pbPrint.hdc
fr.hdcTarget = frmPrint.pbPrint.hdc
fr.rc = rcDrawTo
fr.rcPage = rcPage
fr.chrg.cpMin = 0
fr.chrg.cpMax = -1
nTextLength = Len(rtf.Text)

Do
   fr.hdc = frmPrint.pbPrint.hdc
   fr.hdcTarget = frmPrint.pbPrint.hdc
   nNextCharPos = SendMessage(rtf.hWnd, EM_FORMATRANGE, True, fr)
   If nNextCharPos >= nTextLength Then Exit Do
   fr.chrg.cpMin = nNextCharPos
   frmPrint.pbPrint.Print Space(1)
   
Loop

nRet = SendMessage(rtf.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
MousePointer = vbDefault
PrintRTF = True

Exit Function
ErrorHandler:
    PrintRTF = False
    MousePointer = vbDefault
End Function

Public Function BtnFillForm1()
'попълване на форма 1 от старите експедиции
    
    If MachineNumber = 1 Then
        ns1 = n1s1
        ns2 = n1s2
        ns3 = n1s3
        ns4 = n1s4
    ElseIf MachineNumber = 2 Then
        ns1 = n2s1
        ns2 = n2s2
        ns3 = n2s3
        ns4 = n2s4
    End If
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    Dim ResForm1 As Result
    Set ResForm1 = New Result
    
'почистваме полетата на формата
    prntForm1btn.txtExpNote.Text = ""
    prntForm1btn.txtDate.Text = ""
    prntForm1btn.txtOrd.Text = ""
    prntForm1btn.txtClnt.Text = ""
    prntForm1btn.txtObj.Text = ""
    prntForm1btn.txtDrv.Text = ""
    prntForm1btn.txtDrvNo.Text = ""
    prntForm1btn.txtDist.Text = ""
    prntForm1btn.txtRecType.Text = ""
    prntForm1btn.txtVol.Text = ""
    prntForm1btn.txtW.Text = ""
    prntForm1btn.txtOrdVol.Text = ""
    prntForm1btn.txtClass.Text = ""
    prntForm1btn.txtClassK.Text = ""
    prntForm1btn.txtClassV.Text = ""
    prntForm1btn.txtClassH.Text = ""
    prntForm1btn.txtClassP.Text = ""
    prntForm1btn.txtCem1.Text = ""
    prntForm1btn.txtCem2.Text = ""
    prntForm1btn.txtCem3.Text = ""
    prntForm1btn.txtChem1.Text = ""
    prntForm1btn.txtChem2.Text = ""
    prntForm1btn.txtChem3.Text = ""
    prntForm1btn.txtEDM.Text = ""
    prntForm1btn.txtMixTime.Text = ""
    prntForm1btn.txtExpTime.Text = ""
    prntForm1btn.txtOper.Text = ""
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

'-----------------------Start postgreSQL-----------------------------------
    Dim cnNew1 As ADODB.Connection
    Dim rsNew1 As Recordset
        
    Set cnNew1 = New ADODB.Connection
    cnNew1.ConnectionTimeout = 10
    cnNew1.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме избраната експедиция
    Set rsNew1 = cnNew1.Execute("SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & Val(frmNotes.txtNotes) & " ORDER BY mix_num ASC")
    
    If Not rsNew1.EOF And Not rsNew1.BOF Then
        rsNew1.MoveLast
        
        'попълване на полета на бележката от последния замес
        prntForm1btn.txtExpNote.Text = "M" & MachineNumber & "-" & Format(rsNew1!exp_num, "000000000") 'номер на бележка е номер на последна експедиция
        prntForm1btn.txtDate.Text = Left$(rsNew1!time_mix_ready, 10) 'дата на бележка от последния замес
        prntForm1btn.txtOrd.Text = Format(rsNew1!ord_num, "0000000") & "/" & rsNew1!ord_date 'номер/дата на заявката
        prntForm1btn.txtClnt.Text = rsNew1!name_clnt 'име на клиента
        prntForm1btn.txtObj.Text = rsNew1!obj_clnt 'име на обекта
        prntForm1btn.txtDist.Text = rsNew1!km_clnt 'разстояние до обекта
        prntForm1btn.txtDrv.Text = rsNew1!name_drv 'име на водача
        prntForm1btn.txtDrvNo.Text = rsNew1!reg_drv 'номер на превозно средство
        prntForm1btn.txtRecType.Text = rsNew1!type_rec 'рецепта тип
'        prntForm1btn.txtOrdVol.Text = rDs(rsNew1!ord_q) 'общо количество по заявката
        prntForm1btn.txtClass.Text = rsNew1!class_rec 'клас по якост
        prntForm1btn.txtClassK.Text = rsNew1!classk_rec 'клас по консистенция
        prntForm1btn.txtClassV.Text = rsNew1!classv_rec 'клас по въздействие
        prntForm1btn.txtClassH.Text = rsNew1!classh_rec 'клас по хлориди
        prntForm1btn.txtClassP.Text = rsNew1!classp_rec 'водоплътност
        prntForm1btn.txtEDM.Text = rsNew1!edm_rec 'едм
        prntForm1btn.txtMixTime.Text = Mid$(rsNew1!time_exp_start, 14, 5) 'час на стартиране на експедицията
        prntForm1btn.txtExpTime.Text = Mid$(rsNew1!time_mix_ready, 14, 5) 'час на последния замес по експедицията
        prntForm1btn.txtOper.Text = rsNew1!name_op 'име и фамилия на диспечера
        
        ResForm1.ExpQuant = ARound(CSng(rDs(rsNew1!exp_q)), 2) 'заявен обем за експедицията
        
        'попълваме заявените цименти
        ResForm1.SCRname(1) = rsNew1!cem1_name
        ResForm1.SCRstated(1) = rsNew1!cem1z
        ResForm1.SCRname(2) = rsNew1!cem2_name
        ResForm1.SCRstated(2) = rsNew1!cem2z
        ResForm1.SCRname(3) = rsNew1!cem3_name
        ResForm1.SCRstated(3) = rsNew1!cem3z
        ResForm1.SCRname(4) = rsNew1!cem4_name
        ResForm1.SCRstated(4) = rsNew1!cem4z
        
        'попълваме заявените химически добавки
        ResForm1.CHEMname(1) = rsNew1!chem1_name
        ResForm1.CHEMstated(1) = rDs(rsNew1!chem1z)
        ResForm1.CHEMname(2) = rsNew1!chem2_name
        ResForm1.CHEMstated(2) = rDs(rsNew1!chem2z)
        ResForm1.CHEMname(3) = rsNew1!chem3_name
        ResForm1.CHEMstated(3) = rDs(rsNew1!chem3z)
        ResForm1.CHEMname(4) = rsNew1!chem4_name
        ResForm1.CHEMstated(4) = rDs(rsNew1!chem4z)
        ResForm1.CHEMname(5) = rsNew1!chem5_name
        ResForm1.CHEMstated(5) = rDs(rsNew1!chem5z)
        ResForm1.CHEMname(6) = rsNew1!chem6_name
        ResForm1.CHEMstated(6) = rDs(rsNew1!chem6z)
    End If
    
'зареждане на първите 3 срещнати използвани в рецептата материали от силозите
    Dim ret As Integer
    For r = 1 To ns3
        If ResForm1.SCRstated(r) > 0 Then
            prntForm1btn.txtCem1.Text = ResForm1.SCRname(r)
            ret = r + 1
            Exit For
        Else
            ret = r + 1
        End If
    Next r
    If ret <= ns3 Then
        For r = ret To ns3
            If ResForm1.SCRstated(r) > 0 Then
                prntForm1btn.txtCem2.Text = ResForm1.SCRname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    If ret <= ns3 Then
        For r = ret To ns3
            If ResForm1.SCRstated(r) > 0 Then
                prntForm1btn.txtCem3.Text = ResForm1.SCRname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
        
'зареждане на първите 3 срещнати използвани в рецептата химически добавки
    For r = 1 To ns4
        If ResForm1.CHEMstated(r) > 0 Then
            prntForm1btn.txtChem1.Text = ResForm1.CHEMname(r)
            ret = r + 1
            Exit For
        Else
        End If
    Next r
    If ret <= ns4 Then
        For r = ret To ns4
            If ResForm1.CHEMstated(r) > 0 Then
                prntForm1btn.txtChem2.Text = ResForm1.CHEMname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    If ret <= ns4 Then
        For r = ret To ns4
            If ResForm1.CHEMstated(r) > 0 Then
                prntForm1btn.txtChem3.Text = ResForm1.CHEMname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
        
    If Not rsNew1.BOF And Not rsNew1.EOF Then rsNew1.MoveFirst
    
    ResForm1.TotalStatedKG = 0 'нулираме променливата за кг по рецепта
    ResForm1.TotalMeasuredKG = 0 'нулираме променливата за кг по изпълнение
    ResForm1.TotalQuant = 0 'нулираме променливата за обем по изпълнение
    
    Do While Not rsNew1.EOF
        ResForm1.TotalStatedKG = ResForm1.TotalStatedKG + CSng(rDs(rsNew1!total_rec_kg))
        ResForm1.TotalMeasuredKG = ResForm1.TotalMeasuredKG + CSng(rDs(rsNew1!total_real_kg))
        ResForm1.TotalQuant = ResForm1.TotalQuant + CSng(rDs(rsNew1!total_vol))
        rsNew1.MoveNext
    Loop
    
    Set rsNew1 = cnNew1.Execute("SELECT total_vol FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm1btn.txtOrd.Text) & " ORDER BY mix_num ASC")
    Dim totsmth As Single
    totsmth = 0
    If Not rsNew1.BOF And Not rsNew1.EOF Then rsNew1.MoveFirst
    Do While Not rsNew1.EOF
        totsmth = ARound(totsmth, 2) + ARound(CSng(rDs(rsNew1!total_vol)), 2)
        rsNew1.MoveNext
    Loop
    
    Set rsNew1 = cnNew1.Execute("SELECT DISTINCT ON (exp_num) exp_q FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm1btn.txtOrd.Text) & " ORDER BY exp_num ASC")
    Dim totexpsmth As Single
    totexpsmth = 0
    If Not rsNew1.BOF And Not rsNew1.EOF Then rsNew1.MoveFirst
    Do While Not rsNew1.EOF
        totexpsmth = ARound(totexpsmth, 2) + ARound(CSng(rDs(rsNew1!exp_q)), 2)
        rsNew1.MoveNext
    Loop
    
    rsNew1.Close
    Set rsNew1 = Nothing
    cnNew1.Close
    MousePointer = vbDefault
    Set cnNew1 = Nothing
'--------------------------End PostgreSQL-----------------------------------
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet As Boolean
    Dim strSubKey As String
    
    strSubKey = Trim(PlaceProgSet1)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    If PrevSet = True Then
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm1, "RealVol", ErrRes)
    Else
        rRealVol = 1
    End If
        
    If rRealVol = 1 Then
        prntForm1btn.txtVol.Text = ARound(ResForm1.TotalQuant, 2) 'реален обем на експедицията
        prntForm1btn.txtOrdVol.Text = totsmth
    Else
        prntForm1btn.txtVol.Text = ResForm1.ExpQuant 'заявен обем на експедицията
        prntForm1btn.txtOrdVol.Text = totexpsmth
    End If
        
    prntForm1btn.txtW.Text = ARound(ResForm1.TotalMeasuredKG, 0)
    
    Set ResForm1 = Nothing
    
    MousePointer = vbDefault
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Call PrintBtnForm1(prntForm1btn)
End Function

Public Sub PrintBtnForm1(frm As Form)
'принтиране на форма 1
    
    MousePointer = vbHourglass
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrint.pbPrint.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    
    frmPrint.pbPrint.Width = Printer.Width
    frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 1.41
    
    frmPrint.pbPrint.Line (50, 50)-(frmPrint.pbPrint.Width - 200, 50)
    frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height - 100)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    frmPrint.pbPrint.Line (50, 50)-(50, frmPrint.pbPrint.Height - 100)
    frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height / 2)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height / 2)
    frmPrint.pbPrint.Line (frmPrint.pbPrint.Width - 200, 50)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

    For Each ctr In frm
        If TypeOf ctr Is Label Then
            If ctr.Visible = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + 50
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
            End If
        ElseIf TypeOf ctr Is TextBox Then
            If ctr.Enabled = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + 50
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
                X1 = ctr.Left
                Y1 = ctr.Top + ctr.Height + 30 - 450
                X2 = X1 + ctr.Width
                Y2 = Y1 + ctr.Height - 50
                frmPrint.pbPrint.Line (X1, Y1)-(X2, Y1)
                frmPrint.pbPrint.Line (X1, Y2)-(X2, Y2)
                frmPrint.pbPrint.Line (X1, Y2)-(X1, Y1)
                frmPrint.pbPrint.Line (X2, Y2)-(X2, Y1)
            End If
        End If
    Next ctr

    For Each ctr In frm
        If TypeOf ctr Is Label Then
            If ctr.Visible = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + frmPrint.pbPrint.Height / 2
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
            End If
        ElseIf TypeOf ctr Is TextBox Then
            If ctr.Enabled = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + frmPrint.pbPrint.Height / 2
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
                X1 = ctr.Left
                Y1 = ctr.Top + ctr.Height + 30 + frmPrint.pbPrint.Height / 2 - 500
                X2 = X1 + ctr.Width
                Y2 = Y1 + ctr.Height - 50
                frmPrint.pbPrint.Line (X1, Y1)-(X2, Y1)
                frmPrint.pbPrint.Line (X1, Y2)-(X2, Y2)
                frmPrint.pbPrint.Line (X1, Y2)-(X1, Y1)
                frmPrint.pbPrint.Line (X2, Y2)-(X2, Y1)
            End If
        End If
    Next ctr
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    PrintThePicture frmPrint, frmPrint.pbPrint, 96, 350, 300
    
    MousePointer = vbDefault
    
    Printer.EndDoc
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Max
    Unload frmPrint
    Unload frm
End Sub

Public Function BtnFillForm2()
'попълване на форма 2 от старите експедиции
    If MachineNumber = 1 Then
        ns1 = n1s1
        ns2 = n1s2
        ns3 = n1s3
        ns4 = n1s4
    ElseIf MachineNumber = 2 Then
        ns1 = n2s1
        ns2 = n2s2
        ns3 = n2s3
        ns4 = n2s4
    End If
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    Dim ResForm2 As Result
    Set ResForm2 = New Result
    
    Dim TotalIMKGzfff(0 To 5) As Single
    Dim TotalIMKGifff(0 To 5) As Single
    Dim TotalCemKGzfff(0 To 3) As Single
    Dim TotalCemKGifff(0 To 3) As Single
    Dim TotalWatKGzfff(0 To 1) As Single
    Dim TotalWatKGifff(0 To 1) As Single
    Dim TotalChemKGzfff(0 To 5) As Single
    Dim TotalChemKGifff(0 To 5) As Single

    prntForm2btn.txtExpNote.Text = ""
    prntForm2btn.txtDate.Text = ""
    prntForm2btn.txtOrd.Text = ""
    prntForm2btn.txtClnt.Text = ""
    prntForm2btn.txtObj.Text = ""
    prntForm2btn.txtDrv.Text = ""
    prntForm2btn.txtDrvNo.Text = ""
    prntForm2btn.txtDist.Text = ""
    prntForm2btn.txtRecType.Text = ""
    prntForm2btn.txtVol.Text = ""
    prntForm2btn.txtW.Text = ""
    prntForm2btn.txtOrdVol.Text = ""
    prntForm2btn.txtClass.Text = ""
    prntForm2btn.txtClassK.Text = ""
    prntForm2btn.txtClassV.Text = ""
    prntForm2btn.txtClassH.Text = ""
    prntForm2btn.txtClassP.Text = ""
    prntForm2btn.txtCem1.Text = ""
    prntForm2btn.txtCem2.Text = ""
    prntForm2btn.txtCem3.Text = ""
    prntForm2btn.txtChem1.Text = ""
    prntForm2btn.txtChem2.Text = ""
    prntForm2btn.txtChem3.Text = ""
    prntForm2btn.txtEDM.Text = ""
    prntForm2btn.txtMixTime.Text = ""
    prntForm2btn.txtExpTime.Text = ""
    prntForm2btn.txtOper.Text = ""
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

'-----------------------Start postgreSQL-----------------------------------
    Dim cnNew2 As ADODB.Connection
    Dim rsNew2 As Recordset
        
    Set cnNew2 = New ADODB.Connection
    cnNew2.ConnectionTimeout = 10
    cnNew2.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме последния направен замес като сортираме в обратен ред и вземаме най-големия номер
    Set rsNew2 = cnNew2.Execute("SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & Val(frmNotes.txtNotes) & " ORDER BY mix_num ASC")
    
    If Not rsNew2.EOF And Not rsNew2.BOF Then
        rsNew2.MoveLast
        
        'попълване на полета на бележката от последния замес
        prntForm2btn.txtExpNote.Text = "M" & MachineNumber & "-" & Format(rsNew2!exp_num, "000000000") 'номер на бележка е номер на последна експедиция
        prntForm2btn.txtDate.Text = Left$(rsNew2!time_mix_ready, 10) 'дата на бележка от последния замес
        prntForm2btn.txtOrd.Text = Format(rsNew2!ord_num, "0000000") & "/" & rsNew2!ord_date 'номер/дата на заявката
        prntForm2btn.txtClnt.Text = rsNew2!name_clnt 'име на клиента
        prntForm2btn.txtObj.Text = rsNew2!obj_clnt 'име на обекта
        prntForm2btn.txtDist.Text = rsNew2!km_clnt 'разстояние до обекта
        prntForm2btn.txtDrv.Text = rsNew2!name_drv 'име на водача
        prntForm2btn.txtDrvNo.Text = rsNew2!reg_drv 'номер на превозно средство
        prntForm2btn.txtRecType.Text = rsNew2!type_rec 'рецепта тип
'        prntForm2btn.txtOrdVol.Text = rDs(rsNew2!ord_q) 'общо количество по заявката
        prntForm2btn.txtClass.Text = rsNew2!class_rec 'клас по якост
        prntForm2btn.txtClassK.Text = rsNew2!classk_rec 'клас по консистенция
        prntForm2btn.txtClassV.Text = rsNew2!classv_rec 'клас по въздействие
        prntForm2btn.txtClassH.Text = rsNew2!classh_rec 'клас по хлориди
        prntForm2btn.txtClassP.Text = rsNew2!classp_rec 'водоплътност
        prntForm2btn.txtEDM.Text = rsNew2!edm_rec 'едм
        prntForm2btn.txtMixTime.Text = Mid$(rsNew2!time_exp_start, 14, 5) 'час на стартиране на експедицията
        prntForm2btn.txtExpTime.Text = Mid$(rsNew2!time_mix_ready, 14, 5) 'час на последния замес по експедицията
        prntForm2btn.txtOper.Text = rsNew2!name_op 'име и фамилия на диспечера
        
        ResForm2.ExpQuant = ARound(CSng(rDs(rsNew2!exp_q)), 2) 'заявен обем за експедицията
        
        'попълване на всички имена на материал за таблицата във форма 2
        ResForm2.IMname(1) = rsNew2!im1_name
        ResForm2.IMname(2) = rsNew2!im2_name
        ResForm2.IMname(3) = rsNew2!im3_name
        ResForm2.IMname(4) = rsNew2!im4_name
        ResForm2.IMname(5) = rsNew2!im5_name
        ResForm2.IMname(6) = rsNew2!im6_name
        
        'попълваме заявените цименти
        ResForm2.SCRname(1) = rsNew2!cem1_name
        ResForm2.SCRstated(1) = rsNew2!cem1z
        ResForm2.SCRname(2) = rsNew2!cem2_name
        ResForm2.SCRstated(2) = rsNew2!cem2z
        ResForm2.SCRname(3) = rsNew2!cem3_name
        ResForm2.SCRstated(3) = rsNew2!cem3z
        ResForm2.SCRname(4) = rsNew2!cem4_name
        ResForm2.SCRstated(4) = rsNew2!cem4z
        
        'попълваме заявените води
        ResForm2.WATname(1) = rsNew2!wat1_name
        ResForm2.WATstated(1) = rsNew2!wat1z
        ResForm2.WATname(2) = rsNew2!wat2_name
        ResForm2.WATstated(2) = rsNew2!wat2z
        
        'попълваме заявените химически добавки
        ResForm2.CHEMname(1) = rsNew2!chem1_name
        ResForm2.CHEMstated(1) = rDs(rsNew2!chem1z)
        ResForm2.CHEMname(2) = rsNew2!chem2_name
        ResForm2.CHEMstated(2) = rDs(rsNew2!chem2z)
        ResForm2.CHEMname(3) = rsNew2!chem3_name
        ResForm2.CHEMstated(3) = rDs(rsNew2!chem3z)
        ResForm2.CHEMname(4) = rsNew2!chem4_name
        ResForm2.CHEMstated(4) = rDs(rsNew2!chem4z)
        ResForm2.CHEMname(5) = rsNew2!chem5_name
        ResForm2.CHEMstated(5) = rDs(rsNew2!chem5z)
        ResForm2.CHEMname(6) = rsNew2!chem6_name
        ResForm2.CHEMstated(6) = rDs(rsNew2!chem6z)
    End If
    
'зареждане на първите 3 срещнати използвани в рецептата материали от силозите
    Dim ret As Integer
    For r = 1 To ns3
        If ResForm2.SCRstated(r) > 0 Then
            prntForm2btn.txtCem1.Text = ResForm2.SCRname(r)
            ret = r + 1
            Exit For
        Else
            ret = r + 1
        End If
    Next r
    If ret <= ns3 Then
        For r = ret To ns3
            If ResForm2.SCRstated(r) > 0 Then
                prntForm2btn.txtCem2.Text = ResForm2.SCRname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    If ret <= ns3 Then
        For r = ret To ns3
            If ResForm2.SCRstated(r) > 0 Then
                prntForm2btn.txtCem3.Text = ResForm2.SCRname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
        
'зареждане на първите 3 срещнати използвани в рецептата химически добавки
    For r = 1 To ns4
        If ResForm2.CHEMstated(r) > 0 Then
            prntForm2btn.txtChem1.Text = ResForm2.CHEMname(r)
            ret = r + 1
            Exit For
        Else
            ret = r + 1
        End If
    Next r
    If ret <= ns4 Then
        For r = ret To ns4
            If ResForm2.CHEMstated(r) > 0 Then
                prntForm2btn.txtChem2.Text = ResForm2.CHEMname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    If ret <= ns4 Then
        For r = ret To ns4
            If ResForm2.CHEMstated(r) > 0 Then
                prntForm2btn.txtChem3.Text = ResForm2.CHEMname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    
    If Not rsNew2.BOF And Not rsNew2.EOF Then rsNew2.MoveFirst
    
    ResForm2.TotalStatedKG = 0 'нулираме променливата за кг по рецепта
    ResForm2.TotalMeasuredKG = 0 'нулираме променливата за кг по изпълнение
    ResForm2.TotalQuant = 0 'нулираме променливата за обем по изпълнение
    For I = 0 To 5
        TotalIMKGzfff(I) = 0
        TotalIMKGifff(I) = 0
    Next I
    For I = 0 To 3
        TotalCemKGzfff(I) = 0
        TotalCemKGifff(I) = 0
    Next I
    For I = 0 To 1
        TotalWatKGzfff(I) = 0
        TotalWatKGifff(I) = 0
    Next I
    For I = 0 To 5
        TotalChemKGzfff(I) = 0
        TotalChemKGifff(I) = 0
    Next I
    Do While Not rsNew2.EOF
        ResForm2.TotalStatedKG = ResForm2.TotalStatedKG + CSng(rDs(rsNew2!total_rec_kg)) 'сума кг по зададено
        ResForm2.TotalMeasuredKG = ResForm2.TotalMeasuredKG + CSng(rDs(rsNew2!total_real_kg)) 'сума кг по изпълнено
        ResForm2.TotalQuant = ResForm2.TotalQuant + CSng(rDs(rsNew2!total_vol)) 'сума реален обем
        
        'сума на отделните ИМ по зададено
        TotalIMKGzfff(0) = TotalIMKGzfff(0) + Val(rsNew2!im1z)
        TotalIMKGzfff(1) = TotalIMKGzfff(1) + Val(rsNew2!im2z)
        TotalIMKGzfff(2) = TotalIMKGzfff(2) + Val(rsNew2!im3z)
        TotalIMKGzfff(3) = TotalIMKGzfff(3) + Val(rsNew2!im4z)
        TotalIMKGzfff(4) = TotalIMKGzfff(4) + Val(rsNew2!im5z)
        TotalIMKGzfff(5) = TotalIMKGzfff(5) + Val(rsNew2!im6z)
        
        'сума на отделните ИМ по изпълнено
        TotalIMKGifff(0) = TotalIMKGifff(0) + Val(rsNew2!im1i)
        TotalIMKGifff(1) = TotalIMKGifff(1) + Val(rsNew2!im2i)
        TotalIMKGifff(2) = TotalIMKGifff(2) + Val(rsNew2!im3i)
        TotalIMKGifff(3) = TotalIMKGifff(3) + Val(rsNew2!im4i)
        TotalIMKGifff(4) = TotalIMKGifff(4) + Val(rsNew2!im5i)
        TotalIMKGifff(5) = TotalIMKGifff(5) + Val(rsNew2!im6i)
        
        'сума на отделните цименти по зададено
        TotalCemKGzfff(0) = TotalCemKGzfff(0) + Val(rsNew2!cem1z)
        TotalCemKGzfff(1) = TotalCemKGzfff(1) + Val(rsNew2!cem2z)
        TotalCemKGzfff(2) = TotalCemKGzfff(2) + Val(rsNew2!cem3z)
        TotalCemKGzfff(3) = TotalCemKGzfff(3) + Val(rsNew2!cem4z)
        
        'сума на отделните цименти по изпълнено
        TotalCemKGifff(0) = TotalCemKGifff(0) + Val(rsNew2!cem1i)
        TotalCemKGifff(1) = TotalCemKGifff(1) + Val(rsNew2!cem2i)
        TotalCemKGifff(2) = TotalCemKGifff(2) + Val(rsNew2!cem3i)
        TotalCemKGifff(3) = TotalCemKGifff(3) + Val(rsNew2!cem4i)
        
        'сума на вода по зададено
        TotalWatKGzfff(0) = TotalWatKGzfff(0) + Val(rsNew2!wat1z)
        TotalWatKGzfff(1) = TotalWatKGzfff(1) + Val(rsNew2!wat2z)
        
        'сума на вода по изпълнено
        TotalWatKGifff(0) = TotalWatKGifff(0) + Val(rsNew2!wat1i)
        TotalWatKGifff(1) = TotalWatKGifff(1) + Val(rsNew2!wat2i)
        
        'сума на отделните хд по зададено
        TotalChemKGzfff(0) = TotalChemKGzfff(0) + CSng(rDs(rsNew2!chem1z))
        TotalChemKGzfff(1) = TotalChemKGzfff(1) + CSng(rDs(rsNew2!chem2z))
        TotalChemKGzfff(2) = TotalChemKGzfff(2) + CSng(rDs(rsNew2!chem3z))
        TotalChemKGzfff(3) = TotalChemKGzfff(3) + CSng(rDs(rsNew2!chem4z))
        TotalChemKGzfff(4) = TotalChemKGzfff(4) + CSng(rDs(rsNew2!chem5z))
        TotalChemKGzfff(5) = TotalChemKGzfff(5) + CSng(rDs(rsNew2!chem6z))
        
        'сума на отделните хд по изпълнено
        TotalChemKGifff(0) = TotalChemKGifff(0) + CSng(rDs(rsNew2!chem1i))
        TotalChemKGifff(1) = TotalChemKGifff(1) + CSng(rDs(rsNew2!chem2i))
        TotalChemKGifff(2) = TotalChemKGifff(2) + CSng(rDs(rsNew2!chem3i))
        TotalChemKGifff(3) = TotalChemKGifff(3) + CSng(rDs(rsNew2!chem4i))
        TotalChemKGifff(4) = TotalChemKGifff(4) + CSng(rDs(rsNew2!chem5i))
        TotalChemKGifff(5) = TotalChemKGifff(5) + CSng(rDs(rsNew2!chem6i))
        
        rsNew2.MoveNext
    Loop
    
    Set rsNew2 = cnNew2.Execute("SELECT total_vol FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm2btn.txtOrd.Text) & " ORDER BY mix_num ASC")
    Dim totsmth As Single
    totsmth = 0
    If Not rsNew2.BOF And Not rsNew2.EOF Then rsNew2.MoveFirst
    Do While Not rsNew2.EOF
        totsmth = ARound(totsmth, 2) + ARound(CSng(rDs(rsNew2!total_vol)), 2)
        rsNew2.MoveNext
    Loop
    
    Set rsNew2 = cnNew2.Execute("SELECT DISTINCT ON (exp_num) exp_q FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm2btn.txtOrd.Text) & " ORDER BY exp_num ASC")
    Dim totexpsmth As Single
    totexpsmth = 0
    If Not rsNew2.BOF And Not rsNew2.EOF Then rsNew2.MoveFirst
    Do While Not rsNew2.EOF
        totexpsmth = ARound(totexpsmth, 2) + ARound(CSng(rDs(rsNew2!exp_q)), 2)
        rsNew2.MoveNext
    Loop
    
    rsNew2.Close
    Set rsNew2 = Nothing
    cnNew2.Close
    MousePointer = vbDefault
    Set cnNew2 = Nothing
'--------------------------End PostgreSQL-----------------------------------
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet As Boolean
    Dim strSubKey As String
    strSubKey = Trim(PlaceProgSet2)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm2, "RealVol", ErrRes)
    Else
        rRealVol = 1
    End If
        
    If rRealVol = 1 Then
        prntForm2btn.txtVol.Text = ARound(ResForm2.TotalQuant, 2) 'реален обем на експедицията
        prntForm2btn.txtOrdVol.Text = totsmth
    Else
        prntForm2btn.txtVol.Text = ResForm2.ExpQuant 'заявен обем на експедицията
        prntForm2btn.txtOrdVol.Text = totexpsmth
    End If
        
    prntForm2btn.txtW.Text = ARound(ResForm2.TotalMeasuredKG, 0)

    For e = 1 To ns1
        prntForm2btn.txtIMname(e - 1).Text = ResForm2.IMname(e)
        prntForm2btn.txtIMkgR(e - 1).Text = TotalIMKGifff(e - 1) 'кг по измерено ИМ за всеки материал от всички замеси
        prntForm2btn.txtIMkg(e - 1).Text = TotalIMKGzfff(e - 1) 'кг по зададено ИМ за всеки материал от всички замеси
        If TotalIMKGzfff(e - 1) > 0 Then
            prntForm2btn.txtIMDiff(e - 1).Text = ARound(100 * (TotalIMKGifff(e - 1) - TotalIMKGzfff(e - 1)) / TotalIMKGzfff(e - 1), 2)
        Else
            prntForm2btn.txtIMDiff(e - 1).Text = 0
        End If
        If CSng(rDs(prntForm2btn.txtIMDiff(e - 1).Text)) < 3 And CSng(rDs(prntForm2btn.txtIMDiff(e - 1).Text)) > -3 Then
            prntForm2btn.txtIMOK(e - 1).Text = uniYes
        Else
            prntForm2btn.txtIMOK(e - 1).Text = uniNo
        End If
    Next e
        
    For e = 1 To ns3
        prntForm2btn.txtCemname(e - 1).Text = ResForm2.SCRname(e)
        prntForm2btn.txtCemkgR(e - 1).Text = TotalCemKGifff(e - 1) 'кг по измерено цимент за всеки материал от всички замеси
        prntForm2btn.txtCemkg(e - 1).Text = TotalCemKGzfff(e - 1) 'кг по зададено цимент за всеки материал от всички замеси
        If TotalCemKGzfff(e - 1) > 0 Then
            prntForm2btn.txtCemDiff(e - 1).Text = ARound(100 * (TotalCemKGifff(e - 1) - TotalCemKGzfff(e - 1)) / TotalCemKGzfff(e - 1), 2)
        Else
            prntForm2btn.txtCemDiff(e - 1).Text = 0
        End If
        If CSng(rDs(prntForm2btn.txtCemDiff(e - 1).Text)) < 3 And CSng(rDs(prntForm2btn.txtCemDiff(e - 1).Text)) > -3 Then
            prntForm2btn.txtCemOK(e - 1).Text = uniYes
        Else
            prntForm2btn.txtCemOK(e - 1).Text = uniNo
        End If
    Next e
        
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    For e = 1 To ns2
        prntForm2btn.txtWatname(e - 1).Text = ResForm2.WATname(e)
        prntForm2btn.txtWatkgR(e - 1).Text = TotalWatKGifff(e - 1) 'кг по измерено вода от всички замеси
        prntForm2btn.txtWatkg(e - 1).Text = TotalWatKGzfff(e - 1) 'кг по зададено вода от всички замеси
        If TotalWatKGzfff(e - 1) > 0 Then
            prntForm2btn.txtWatDiff(e - 1).Text = ARound(100 * (TotalWatKGifff(e - 1) - TotalWatKGzfff(e - 1)) / TotalWatKGzfff(e - 1), 2)
        Else
            prntForm2btn.txtWatDiff(e - 1).Text = 0
        End If
        If CSng(rDs(prntForm2btn.txtWatDiff(e - 1).Text)) < 3 And CSng(rDs(prntForm2btn.txtWatDiff(e - 1).Text)) > -3 Then
            prntForm2btn.txtWatOK(e - 1).Text = uniYes
        Else
            prntForm2btn.txtWatOK(e - 1).Text = uniNo
        End If
    Next e
    
    For e = 1 To ns4
        prntForm2btn.txtChemname(e - 1).Text = ResForm2.CHEMname(e)
        prntForm2btn.txtChemkgR(e - 1).Text = CSng(rDs(TotalChemKGifff(e - 1))) 'кг по измерено химия за всеки материал от всички замеси
        prntForm2btn.txtChemkg(e - 1).Text = CSng(rDs(TotalChemKGzfff(e - 1))) 'кг по зададено химия за всеки материал от всички замеси
        If CSng(rDs(TotalChemKGzfff(e - 1))) > 0 Then
            prntForm2btn.txtChemDiff(e - 1).Text = ARound(100 * (TotalChemKGifff(e - 1) - CSng(rDs(TotalChemKGzfff(e - 1)))) / CSng(rDs(TotalChemKGzfff(e - 1))), 2)
        Else
            prntForm2btn.txtChemDiff(e - 1).Text = 0
        End If
        If CSng(rDs(prntForm2btn.txtChemDiff(e - 1).Text)) < 5 And CSng(rDs(prntForm2btn.txtChemDiff(e - 1).Text)) > -5 Then
            prntForm2btn.txtChemOK(e - 1).Text = uniYes
        Else
            prntForm2btn.txtChemOK(e - 1).Text = uniNo
        End If
    Next e
    
    MousePointer = vbDefault
    Set ResForm2 = Nothing
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Call PrintBtnForm2(prntForm2btn)
End Function

Public Sub PrintBtnForm2(frm As Form)
'принтиране на форма 2
    
    MousePointer = vbHourglass
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrint.pbPrint.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    
    frmPrint.pbPrint.Width = Printer.Width
    frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 1.41
    
    frmPrint.pbPrint.Line (50, 50)-(frmPrint.pbPrint.Width - 200, 50)
    frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height - 100)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    frmPrint.pbPrint.Line (50, 50)-(50, frmPrint.pbPrint.Height - 100)
    frmPrint.pbPrint.Line (frmPrint.pbPrint.Width - 200, 50)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

    For Each ctr In frm
        If TypeOf ctr Is Label Then
            If ctr.Visible = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + 50
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
            End If
        ElseIf TypeOf ctr Is TextBox Then
            If ctr.Enabled = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + 50
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
                X1 = ctr.Left
                Y1 = ctr.Top + ctr.Height + 30 - 450
                X2 = X1 + ctr.Width
                Y2 = Y1 + ctr.Height - 50
                frmPrint.pbPrint.Line (X1, Y1)-(X2, Y1)
                frmPrint.pbPrint.Line (X1, Y2)-(X2, Y2)
                frmPrint.pbPrint.Line (X1, Y2)-(X1, Y1)
                frmPrint.pbPrint.Line (X2, Y2)-(X2, Y1)
            End If
        End If
    Next ctr
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    PrintThePicture frmPrint, frmPrint.pbPrint, 96, 350, 300
    
    MousePointer = vbDefault
    
    Printer.EndDoc
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Max
    Unload frmPrint
    Unload frm
End Sub

Public Function BtnFillForm3()
'попълване на форма 3 от старите експедиции
    
    If MachineNumber = 1 Then
        ns1 = n1s1
        ns2 = n1s2
        ns3 = n1s3
        ns4 = n1s4
    ElseIf MachineNumber = 2 Then
        ns1 = n2s1
        ns2 = n2s2
        ns3 = n2s3
        ns4 = n2s4
    End If
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    Dim ResForm3 As Result
    Set ResForm3 = New Result

    prntForm3btn.txtExpNote.Text = ""
    prntForm3btn.txtDate.Text = ""
    prntForm3btn.txtOrd.Text = ""
    prntForm3btn.txtClnt.Text = ""
    prntForm3btn.txtObj.Text = ""
    prntForm3btn.txtDrv.Text = ""
    prntForm3btn.txtDrvNo.Text = ""
    prntForm3btn.txtDist.Text = ""
    prntForm3btn.txtRecType.Text = ""
    prntForm3btn.txtVol.Text = ""
    prntForm3btn.txtW.Text = ""
    prntForm3btn.txtOrdVol.Text = ""
    prntForm3btn.txtClass.Text = ""
    prntForm3btn.txtClassK.Text = ""
    prntForm3btn.txtClassV.Text = ""
    prntForm3btn.txtClassH.Text = ""
    prntForm3btn.txtClassP.Text = ""
    prntForm3btn.txtCem1.Text = ""
    prntForm3btn.txtCem2.Text = ""
    prntForm3btn.txtCem3.Text = ""
    prntForm3btn.txtChem1.Text = ""
    prntForm3btn.txtChem2.Text = ""
    prntForm3btn.txtChem3.Text = ""
    prntForm3btn.txtEDM.Text = ""
    prntForm3btn.txtMixTime.Text = ""
    prntForm3btn.txtExpTime.Text = ""
    prntForm3btn.txtOper.Text = ""
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

'-----------------------Start postgreSQL-----------------------------------
    Dim cnNew3 As ADODB.Connection
    Dim rsNew3 As Recordset
        
    Set cnNew3 = New ADODB.Connection
    cnNew3.ConnectionTimeout = 10
    cnNew3.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме избраната експедиция
    Set rsNew3 = cnNew3.Execute("SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & Val(frmNotes.txtNotes) & " ORDER BY mix_num ASC")
    
    If Not rsNew3.EOF And Not rsNew3.BOF Then
        rsNew3.MoveLast
        
        'попълване на полета на бележката от последния замес
        prntForm3btn.txtExpNote.Text = "M" & MachineNumber & "-" & Format(rsNew3!exp_num, "000000000") 'номер на бележка е номер на последна експедиция
        prntForm3btn.txtDate.Text = Left$(rsNew3!time_mix_ready, 10) 'дата на бележка от последния замес
        prntForm3btn.txtOrd.Text = Format(rsNew3!ord_num, "0000000") & "/" & rsNew3!ord_date 'номер/дата на заявката
        prntForm3btn.txtClnt.Text = rsNew3!name_clnt 'име на клиента
        prntForm3btn.txtObj.Text = rsNew3!obj_clnt 'име на обекта
        prntForm3btn.txtDist.Text = rsNew3!km_clnt 'разстояние до обекта
        prntForm3btn.txtDrv.Text = rsNew3!name_drv 'име на водача
        prntForm3btn.txtDrvNo.Text = rsNew3!reg_drv 'номер на превозно средство
        prntForm3btn.txtRecType.Text = rsNew3!type_rec 'рецепта тип
'        prntForm3btn.txtOrdVol.Text = rDs(rsNew3!ord_q) 'общо количество по заявката
        prntForm3btn.txtClass.Text = rsNew3!class_rec 'клас по якост
        prntForm3btn.txtClassK.Text = rsNew3!classk_rec 'клас по консистенция
        prntForm3btn.txtClassV.Text = rsNew3!classv_rec 'клас по въздействие
        prntForm3btn.txtClassH.Text = rsNew3!classh_rec 'клас по хлориди
        prntForm3btn.txtClassP.Text = rsNew3!classp_rec 'водоплътност
        prntForm3btn.txtEDM.Text = rsNew3!edm_rec 'едм
        prntForm3btn.txtMixTime.Text = Mid$(rsNew3!time_exp_start, 14, 5) 'час на стартиране на експедицията
        prntForm3btn.txtExpTime.Text = Mid$(rsNew3!time_mix_ready, 14, 5) 'час на последния замес по експедицията
        prntForm3btn.txtOper.Text = rsNew3!name_op 'име и фамилия на диспечера
        
        ResForm3.ExpQuant = ARound(CSng(rDs(rsNew3!exp_q)), 2) 'заявен обем за експедицията
        
        'попълваме заявените цименти
        ResForm3.SCRname(1) = rsNew3!cem1_name
        ResForm3.SCRstated(1) = rsNew3!cem1z
        ResForm3.SCRname(2) = rsNew3!cem2_name
        ResForm3.SCRstated(2) = rsNew3!cem2z
        ResForm3.SCRname(3) = rsNew3!cem3_name
        ResForm3.SCRstated(3) = rsNew3!cem3z
        ResForm3.SCRname(4) = rsNew3!cem4_name
        ResForm3.SCRstated(4) = rsNew3!cem4z
        
        'попълваме заявените химически добавки
        ResForm3.CHEMname(1) = rsNew3!chem1_name
        ResForm3.CHEMstated(1) = rDs(rsNew3!chem1z)
        ResForm3.CHEMname(2) = rsNew3!chem2_name
        ResForm3.CHEMstated(2) = rDs(rsNew3!chem2z)
        ResForm3.CHEMname(3) = rsNew3!chem3_name
        ResForm3.CHEMstated(3) = rDs(rsNew3!chem3z)
        ResForm3.CHEMname(4) = rsNew3!chem4_name
        ResForm3.CHEMstated(4) = rDs(rsNew3!chem4z)
        ResForm3.CHEMname(5) = rsNew3!chem5_name
        ResForm3.CHEMstated(5) = rDs(rsNew3!chem5z)
        ResForm3.CHEMname(6) = rsNew3!chem6_name
        ResForm3.CHEMstated(6) = rDs(rsNew3!chem6z)
    End If
    
'зареждане на първите 3 срещнати използвани в рецептата материали от силозите
    Dim ret As Integer
    For r = 1 To ns3
        If ResForm3.SCRstated(r) > 0 Then
            prntForm3btn.txtCem1.Text = ResForm3.SCRname(r)
            ret = r + 1
            Exit For
        Else
            ret = r + 1
        End If
    Next r
    If ret <= ns3 Then
        For r = ret To ns3
            If ResForm3.SCRstated(r) > 0 Then
                prntForm3btn.txtCem2.Text = ResForm3.SCRname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    If ret <= ns3 Then
        For r = ret To ns3
            If ResForm3.SCRstated(r) > 0 Then
                prntForm3btn.txtCem3.Text = ResForm3.SCRname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
        
'зареждане на първите 3 срещнати използвани в рецептата химически добавки
    For r = 1 To ns4
        If ResForm3.CHEMstated(r) > 0 Then
            prntForm3btn.txtChem1.Text = ResForm3.CHEMname(r)
            ret = r + 1
            Exit For
        Else
            ret = r + 1
        End If
    Next r
    If ret <= ns4 Then
        For r = ret To ns4
            If ResForm3.CHEMstated(r) > 0 Then
                prntForm3btn.txtChem2.Text = ResForm3.CHEMname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    If ret <= ns4 Then
        For r = ret To ns4
            If ResForm3.CHEMstated(r) > 0 Then
                prntForm3btn.txtChem3.Text = ResForm3.CHEMname(r)
                ret = r + 1
                Exit For
            Else
                ret = r + 1
            End If
        Next r
    End If
    
    If Not rsNew3.BOF And Not rsNew3.EOF Then rsNew3.MoveFirst
    
    ResForm3.TotalStatedKG = 0 'нулираме променливата за кг по рецепта
    ResForm3.TotalMeasuredKG = 0 'нулираме променливата за кг по изпълнение
    ResForm3.TotalQuant = 0 'нулираме променливата за обем по изпълнение
    
    Do While Not rsNew3.EOF
        ResForm3.TotalStatedKG = ResForm3.TotalStatedKG + CSng(rDs(rsNew3!total_rec_kg))
        ResForm3.TotalMeasuredKG = ResForm3.TotalMeasuredKG + CSng(rDs(rsNew3!total_real_kg))
        ResForm3.TotalQuant = ResForm3.TotalQuant + CSng(rDs(rsNew3!total_vol))
        rsNew3.MoveNext
    Loop
    
    Set rsNew3 = cnNew3.Execute("SELECT total_vol FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm3btn.txtOrd.Text) & " ORDER BY mix_num ASC")
    Dim totsmth As Single
    totsmth = 0
    If Not rsNew3.BOF And Not rsNew3.EOF Then rsNew3.MoveFirst
    Do While Not rsNew3.EOF
        totsmth = ARound(totsmth, 2) + ARound(CSng(rDs(rsNew3!total_vol)), 2)
        rsNew3.MoveNext
    Loop
    
    Set rsNew3 = cnNew3.Execute("SELECT DISTINCT ON (exp_num) exp_q FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm3btn.txtOrd.Text) & " ORDER BY exp_num ASC")
    Dim totexpsmth As Single
    totexpsmth = 0
    If Not rsNew3.BOF And Not rsNew3.EOF Then rsNew3.MoveFirst
    Do While Not rsNew3.EOF
        totexpsmth = ARound(totexpsmth, 2) + ARound(CSng(rDs(rsNew3!exp_q)), 2)
        rsNew3.MoveNext
    Loop
    
    rsNew3.Close
    Set rsNew3 = Nothing
    cnNew3.Close
    MousePointer = vbDefault
    Set cnNew3 = Nothing
'--------------------------End PostgreSQL-----------------------------------

    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet As Boolean
    Dim strSubKey As String
    strSubKey = Trim(PlaceProgSet3)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    If PrevSet = True Then
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm3, "RealVol", ErrRes)
    Else
        rRealVol = 1
    End If
        
    If rRealVol = 1 Then
        prntForm3btn.txtVol.Text = ARound(ResForm3.TotalQuant, 2) 'реален обем на експедицията
        prntForm3btn.txtOrdVol.Text = totsmth
    Else
        prntForm3btn.txtVol.Text = ResForm3.ExpQuant 'заявен обем на експедицията
        prntForm3btn.txtOrdVol.Text = totexpsmth
    End If
        
    prntForm3btn.txtW.Text = ARound(ResForm3.TotalMeasuredKG, 0)
    
    MousePointer = vbDefault
    
    Set ResForm3 = Nothing
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Call PrintBtnForm3(prntForm3btn)
End Function

Public Sub PrintBtnForm3(frm As Form)
'принтиране на форма 3
    
    MousePointer = vbHourglass
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrint.pbPrint.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    
    frmPrint.pbPrint.Width = Printer.Width
    frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 1.41
    
    frmPrint.pbPrint.Line (50, 50)-(frmPrint.pbPrint.Width - 200, 50)
    frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height - 100)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    frmPrint.pbPrint.Line (50, 50)-(50, frmPrint.pbPrint.Height - 100)
    frmPrint.pbPrint.Line (frmPrint.pbPrint.Width - 200, 50)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

    For Each ctr In frm
        If TypeOf ctr Is Label Then
            If ctr.Visible = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + 50
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
            End If
        ElseIf TypeOf ctr Is TextBox Then
            If ctr.Enabled = True Then
                frmPrint.pbPrint.CurrentX = ctr.Left + 50
                frmPrint.pbPrint.CurrentY = ctr.Top + 50
                frmPrint.pbPrint.Font = ctr.Font
                frmPrint.pbPrint.FontSize = ctr.FontSize
                frmPrint.pbPrint.FontBold = ctr.FontBold
                frmPrint.pbPrint.FontItalic = ctr.FontItalic
                frmPrint.pbPrint.Print ctr
                X1 = ctr.Left
                Y1 = ctr.Top + ctr.Height + 30 - 450
                X2 = X1 + ctr.Width
                Y2 = Y1 + ctr.Height - 50
                frmPrint.pbPrint.Line (X1, Y1)-(X2, Y1)
                frmPrint.pbPrint.Line (X1, Y2)-(X2, Y2)
                frmPrint.pbPrint.Line (X1, Y2)-(X1, Y1)
                frmPrint.pbPrint.Line (X2, Y2)-(X2, Y1)
            End If
        End If
    Next ctr
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Call PrintRTF(prntForm3btn.Confirmity, 850, 7300, 800, 300)
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then _
    frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

    PrintThePicture frmPrint, frmPrint.pbPrint, 96, 350, 300
    
    MousePointer = vbDefault
    
    Printer.EndDoc
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Max
    Unload frmPrint
    Unload frm
End Sub

