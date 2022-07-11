Attribute VB_Name = "procedures"

'константa машина номер
Public Const MachineNumber = 1

'константи системни
Public Const KEY_READ = &H20019

Public Const HKEY_CURRENT_USER = &H80000001

Public Const LOCALE_SDECIMAL = &HE

Public Const CB_FINDSTRING = &H14C

Public Const CB_SHOWDROPDOWN = &H14F

Public Const CB_ERR = (-1)

Public Const FrmPlace = 2800

Public Const CSIDL_COMMON_APPDATA = &H23

Public Const CSIDL_DESKTOP = &H0

'--------------------------------------
'константи за регистъра
Public Const PlaceKey = "PoliciesAdmin"

Public Const PlaceKeyAdd = "AddKey"

Public Const PlcRegLicNum = "LicenseNumber"

Public Const ErrRes = "whatthefuckareyoudoinghere"

Public Const PlacePass = "PoliciesAdmin"

Public Const PlacePassAdd = "Addons"

Public Const PlaceProgSettings = "TipPanel"

Public Const PlaceProgSet1 = "Software\VB and VBA Program Settings\TipPanel\Form1Set"

Public Const PlaceProgSet2 = "Software\VB and VBA Program Settings\TipPanel\Form2Set"

Public Const PlaceProgSet3 = "Software\VB and VBA Program Settings\TipPanel\Form3Set"

Public Const PlaceForm1 = "Form1Set"

Public Const PlaceForm2 = "Form2Set"

Public Const PlaceForm3 = "Form3Set"

Public Const PlaceProgAllow = "Software\VB and VBA Program Settings\TipPanel\Allow"

Public Const PlaceAllow = "Allow"

Public Const PlaceProgPrint = "Software\VB and VBA Program Settings\TipPanel\AutoPrint"

Public Const PlacePrint = "AutoPrint"

Public Const Place1SilosSet = "Software\VB and VBA Program Settings\TipPanel\Silos1Set"

Public Const Place1Silos = "Silos1Set"

Public Const Place2SilosSet = "Software\VB and VBA Program Settings\TipPanel\Silos2Set"

Public Const Place2Silos = "Silos2Set"

Public Const Place1SilosQ = "Software\VB and VBA Program Settings\TipPanel\Silos1Question"

Public Const Place1Q = "Silos1Question"

Public Const Place2SilosQ = "Software\VB and VBA Program Settings\TipPanel\Silos2Question"

Public Const Place2Q = "Silos2Question"

Public Const PlaceEditor = "Software\VB and VBA Program Settings\TipPanel\NotesEditor"

Public Const PlaceEd = "NotesEditor"

Public Const PlaceShit = "Software\VB and VBA Program Settings\TipPanel\Shit"

Public Const Shit = "Shit"

'---------------------------------

'константи за течките на машината
Public Const BaseIM = 11

Public Const BaseWat = 21

Public Const BaseScr = 31

Public Const BaseChem = 41

'-------------------------------------

'константи за базата данни
Public Const DbaseName = "postgres"

Public Const DbaseUser = "postgres"

'променлива за базата данни

Public IPConnStr      As String

Public PassConnStr    As String

'------------------------------------

Public dispatcher As Boolean

'променливи за настойка на формите за печат на бележки
Public rDist                As Integer

Public rRecType             As Integer

Public rVol                 As Integer

Public rW                   As Integer

Public rOrdVol              As Integer

Public rClass               As Integer

Public rClassK              As Integer

Public rClassV              As Integer

Public rClassH              As Integer

Public rClassP              As Integer

Public rCem1                As Integer

Public rCem2                As Integer

Public rCem3                As Integer

Public rChem1               As Integer

Public rChem2               As Integer

Public rChem3               As Integer

Public rEDM                 As Integer

Public rMixTime             As Integer

Public rExpTime             As Integer

Public rRealVol             As Integer

Public rPrint1              As Integer

Public rPrint2              As Integer

Public rPrint3              As Integer

Public PrintAnyForm         As Boolean

Public PrintRightBut        As Boolean

Public numSheetsForm1       As Integer

Public numSheetsForm2       As Integer

Public numSheetsForm3       As Integer
'------------------------------------------

'променливи за настойка на формите за печат на бележки
Public rSilos1              As Integer

Public rSilos2              As Integer

Public rSilos3              As Integer

Public rSilos4              As Integer

'------------------------------------------

'променливи за разрешения от админа
Public rActDel              As Integer

Public rActForm1            As Integer

Public rActForm2            As Integer

Public rActForm3            As Integer

Public rDeactNRPass         As Integer

Public rDeactDRPass         As Integer
'-------------------------------------------

'променлива за въпроса за количество в силозите
Public QuestSilos           As Boolean

'променлива за редактиране на бележки
Public ShowEditor           As Boolean

'променлива за сговнясване на програмата

Public ShitEnabled          As Boolean

'------------------------------------------------
'променливи за работните файлове на програмата
Public PathCore             As String

Public BackPath             As String

Public OPCSetFile           As String

Public DBSetFile            As String

Public InfoFile             As String

Public SilosFile            As String

Public ConfirmityFile       As String

'Public LangSetFile          As String
'
'Public LangBgFile           As String
'
'Public LangRusFile          As String
'
'Public LangEnFile           As String

'------------------------------------------------

'променливи за OPC server
Public Servers              As Variant

Public MyServer             As String
    
Public handyRec(39)         As Long

Public handyReady(4)        As Long

Public handyConfig(1)       As Long
'-------------------------------------

'Променливи за езикови преводи
Public TxtInfoCap           As String

Public TxtVerCap            As String
    
Public MsgAdminSuccess      As String

Public MsgAnotherRun        As String

Public MsgArhBx             As String

Public MsgArhivError        As String

Public MsgArhivReady        As String

Public MsgArhivQuest        As String

Public MsgAvaria            As String

Public MsgBusyFlow          As String

Public MsgCallTIP           As String

Public MsgCantDelRec        As String

Public MsgCantDelClnt       As String

Public MsgCantSvFutDelivery As String

Public MsgClose             As String

Public MsgClWatRev          As String

Public MsgCodeZero          As String

Public MsgConfCrTables      As String

Public MsgConfDel           As String

Public MsgConfEdit          As String

Public MsgConnEst           As String

Public MsgContinue          As String

Public MsgCorData           As String

Public MsgDBConnEst         As String

Public MsgDaysLeft          As String

Public MsgDelBx             As String

Public MsgDelSuccess        As String

Public MsgEditBx            As String

Public MsgEndOnCancel       As String

Public MsgErrBx             As String

Public MsgErrBxFatal        As String

Public MsgErrNoData         As String

Public MsgErrExcel          As String

Public MsgExpWait           As String

Public MsgFillAll           As String

Public MsgInvalLic          As String

Public MsgKey               As String

Public MsgMatDelivery       As String

Public MsgMatLoaded         As String

Public MsgMatNotQuant       As String

Public MsgMatSold           As String

Public MsgMaxL15            As String

Public MsgMkAdmin           As String

Public MsgNewName           As String

Public MsgNoClnt            As String

Public MsgNoDBConn          As String

Public MsgNoExcel           As String

Public MsgNoLic             As String

Public MsgNoMat             As String

Public MsgNoOps             As String

Public MsgNoPayment         As String

Public MsgNoPaymentDays     As String

Public MsgNoResOnExit       As String

Public MsgNoSelection       As String

Public MsgNotEnQuant        As String

Public MsgNotRespOPC        As String

Public MsgNotWorkDB         As String

Public MsgOffline           As String

Public MsgOverCapDrv        As String

Public MsgPassNotConf       As String

Public MsgQuantNotMat       As String

Public MsgRecErrCem         As String

Public MsgRecErrChem        As String

Public MsgRecErrIM          As String

Public MsgRecErrWat         As String

Public MsgResBx             As String

Public MsgRestoreError      As String

Public MsgRestoreReady      As String

Public MsgSameNmFamOp       As String

Public MsgSaveSuccess       As String

Public MsgTablesNotFound    As String

Public MsgTablesReady       As String

Public MsgWrong             As String

Public MsgWrongPass         As String
    
Public lblChooseArhFile     As String

Public lblChooseFolder      As String

Public lblEntIP             As String

Public lblEntPass           As String

Public lblLabPassCap        As String

Public lblOperCap           As String

Public lblPassCap           As String

Public lblPassConfCap       As String

Public lblRevInfo           As String

Public lblRevision          As String

Public lblSetClients        As String

Public lblSetObjects        As String

Public lblSetDrivers        As String

Public lblSetRecepies       As String

Public lblSetSuppliers      As String
    
Public frmAdPanel           As String

Public frmConfSend          As String

Public frmDataCor           As String

Public frmDispPanelCap      As String

Public frmLabPassCap        As String

Public frmNewAd             As String

Public frmNmSilos           As String

Public frmDBdata            As String
    
Public btnCreateOp          As String

Public btnEditAd            As String

Public btnEditOp            As String

Public btnParamSys          As String

Public btnSendControllerCap As String
    
Public statAuto             As String

Public statMan              As String

Public statAv               As String

Public statAvaria           As String

Public statAvStop           As String

Public statReqStarted       As String

Public statReadyReqNew      As String

Public statMixOn            As String

Public statMixOff           As String

Public statMixOpened        As String

Public statMixClosed        As String

Public statBeltOn           As String

Public statBeltOff          As String

Public statCemOpened        As String

Public statCemClosed        As String

Public statWatOpened        As String

Public statWatClosed        As String

Public statChemOpened       As String

Public statChemClosed       As String
    
Public UniCancel            As String

Public UniEnter             As String

Public UniExit              As String

Public UniOK                As String
    
Public uniAbout             As String

Public uniActDel            As String

Public uniActForm1          As String

Public uniActForm2          As String

Public uniActForm3          As String

Public uniAdd               As String

Public uniAdmin             As String

Public uniAllow             As String

Public uniAutoPrint         As String
    
Public uniBG                As String

Public uniCapacity          As String

Public uniCem               As String

Public uniCement            As String

Public uniCemShort          As String

Public uniChem              As String

Public uniChemShort         As String

Public uniClass             As String

Public uniClassK            As String

Public uniClassV            As String

Public uniClassH            As String

Public uniClassP            As String

Public uniClnt              As String

Public uniClntCode          As String

Public uniClnts             As String

Public uniCode              As String

Public uniComInfo           As String

Public uniConMat            As String

Public uniConcPlant         As String
    
Public uniDate              As String

Public uniDateDlvr          As String

Public uniDateOrd           As String

Public uniDateReadyShort    As String

Public uniDateReady         As String

Public uniDeactNRPass       As String

Public uniDeactDRPass       As String

Public uniDel               As String

Public uniDelivered         As String

Public uniDisp              As String

Public uniDispQuant         As String

Public uniDlvr              As String

Public uniDlvrs             As String

Public uniDrv               As String

Public uniDrvCode           As String

Public uniDrvReg            As String

Public uniDrvs              As String
    
Public uniEDM               As String

Public uniEmpty             As String

Public uniEnterExp          As String

Public uniExped             As String
    
Public uniFam               As String

Public uniFax               As String

Public uniFirm              As String

Public uniFlow              As String

Public uniForm1             As String

Public uniForm2             As String

Public uniForm3             As String
    
Public uniHave              As String
    
Public uniIM                As String

Public uniIMShort           As String

Public uniInfoClnt          As String

Public uniInfoDrv           As String

Public uniInfoMix           As String
    
Public uniKGstated          As String

Public uniKGmeasured        As String

Public uniKm                As String

Public uniKmShort           As String
    
Public uniLoad              As String

Public uniLoaded            As String

Public uniLoading           As String

Public uniLog               As String
    
Public uniMade              As String

Public uniMat               As String

Public uniMats              As String

Public uniMeasured          As String

Public uniMix               As String

Public uniMixCap            As String

Public uniMod               As String

Public uniMOL               As String
    
Public uniNew               As String

Public uniNewa              As String

Public uniNewNm             As String

Public uniNm                As String

Public uniNo                As String

Public uniNoDoc             As String

Public uniNote              As String

Public uniNotes             As String

Public uniNr                As String

Public uniNumCem            As String

Public uniNumChem           As String

Public uniNumIM             As String

Public uniNumMix            As String

Public uniNumWat            As String
    
Public uniObj               As String

Public uniOpenMix           As String

Public uniOrdCode           As String

Public uniOrdered           As String

Public uniOrds              As String

Public uniOrdsVert          As String

Public uniOrdQuant          As String

Public uniOther             As String
    
Public uniPrint             As String

Public uniPump              As String
    
Public uniQuant             As String

Public uniQuantMix          As String

Public uniQuestPrint        As String
    
Public uniReadyVert         As String

Public uniReady2Vert        As String

Public uniRec               As String

Public uniRecCode           As String

Public uniRecNote           As String

Public uniRecType           As String

Public uniRecs              As String

Public uniResults           As String

Public uniRevision          As String

Public uniRevisor           As String
    
Public uniSave              As String

Public uniSendingPrinter    As String

Public uniSettings          As String

Public uniSheet             As String

Public uniSold              As String

Public uniSrchBG            As String

Public uniSTART             As String

Public uniStatus            As String

Public uniSup               As String

Public uniSups              As String
    
Public uniTank              As String

Public uniTel               As String

Public uniTempResults       As String

Public uniTimeMix           As String

Public uniTimeMixShort      As String

Public uniTimePour          As String

Public uniTimePourShort     As String

Public uniTotalKg           As String

Public uniTown              As String

Public uniType              As String

Public uniTypeDoc           As String
    
Public uniVolmeasured       As String
    
Public uniWat               As String

Public uniWat1cb            As String
    
Public uniYes               As String
'---------------------------------------------------------------------------

'променливи за рецептата
Public IM(6)                As String

Public Scr(4)               As String

Public Wat(2)               As String

Public Chem(6)              As String
'-------------------------------------------

'променливи за изпращане на рецепта към контролера
Public cyc                  As Integer

Public kgIMs(6)             As Integer

Public kgSCRs(4)            As Integer

Public kgWATs(2)            As Integer

Public kgCHEMs(6)           As Single

Public VIPAkgIMs(6)         As Integer

Public VIPAkgSCRs(4)        As Integer

Public VIPAkgWATs(2)        As Integer

Public VIPAkgCHEMs(6)       As Single

Public OrdData(3)           As String

Public ClientData(3)        As String

Public DriverData(2)        As String

Public Recs                 As Integer

Public RecNames             As String

Public RecTypes             As String

Public RecClasss            As String

Public RecClassKs           As String

Public RecClassVs           As String

Public RecClassHs           As String

Public RecClassPs           As String

Public RecEDMs              As Integer

Public nCoefs               As Single
'-------------------------------------------

'променливи за параметри на машината
Public MixCap               As Single

Public TMd                  As Single

Public TPd                  As Single

Public ns1                  As Integer

Public ns2                  As Integer

Public ns3                  As Integer

Public ns4                  As Integer
'--------------------------------------
    
'променливи за записи на резултат от замесите
Public ExpeditionStarted    As Boolean

Public CountMix             As Integer

Public HelpRes              As Integer

Public HelpResAggr          As Integer

Public HelpResCem           As Integer

Public HelpResWat           As Integer

Public HelpResHD            As Integer

Public resMatrix(30, 16)    As Single

Public ReqTime              As String

Public tRealQuant           As Single

Public tTotalKGs            As Single

Public EmptyData            As Boolean

Public tempQQQ              As Single
'---------------------------------------
    
'други променливи
Public SecondStart          As Boolean 'второ стартиране

Public DayToday             As String 'днес

Public OffMode              As Boolean 'офлайн OPC

Public ConStr               As String 'connection string за връзка с база данни

Public DecSep               As String 'десетичен знак

Public OperName             As String 'име на логнатия оператор
    
Public FlagButRec           As Integer 'флаг за натиснат бутон от екран рецепти (запис или изтрий)

Public DontAskExit          As Boolean 'флаг за изход без въпрос

Public VipaActive           As Boolean 'флаг за работа с VIPA313

Public WasAuto              As Boolean 'флаг дали машината е била в автомат

Public okAggr               As Boolean 'флаг за изчитане на ИМ

Public okCem                As Boolean 'флаг за изчитане на цимент

Public okWat                As Boolean 'флаг за изчитане на вода

Public okHD                 As Boolean 'флаг за изчитане на ХД

Public RecMin               As Integer 'минимален номер на рецепта

Public ErrDaily             As Boolean 'флаг за грешка в дневния отчет
'---------------------------------
    
'----------- декларации на функции от системни библиотеки ---------------------------------------------------------
Public Declare Function GetThreadLocale Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function GetLocaleInfo _
               Lib "kernel32" _
               Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                       ByVal LCType As Long, _
                                       ByVal lpLCData As String, _
                                       ByVal cchData As Long) As Long

Public Declare Function RegOpenKeyEx _
               Lib "advapi32.dll" _
               Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String, _
                                      ByVal ulOptions As Long, _
                                      ByVal samDesired As Long, _
                                      phkResult As Long) As Long

Public Declare Function GetVolumeInformation _
               Lib "kernel32" _
               Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                              ByVal pVolumeNameBuffer As String, _
                                              ByVal nVolumeNameSize As Long, _
                                              lpVolumeSerialNumber As Long, _
                                              lpMaximumComponentLength As Long, _
                                              lpFileSystemFlags As Long, _
                                              ByVal lpFileSystemNameBuffer As String, _
                                              ByVal nFileSystemNameSize As Long) As Long
   
Public Declare Function BitBlt Lib "GDI32.DLL" _
(ByVal hDestDC As Long, ByVal X As Integer, ByVal Y As Integer, _
 ByVal nWid As Integer, ByVal nHt As Integer, ByVal hSrcDC As Long, _
 ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) _
As Integer

Public Const SRCCOPY = &HCC0020

Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long
    
Public Declare Function SHGetFolderPath _
               Lib "shfolder.dll" _
               Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, _
                                         ByVal nFolder As Long, _
                                         ByVal hToken As Long, _
                                         ByVal dwReserved As Long, _
                                         ByVal lpszPath As String) As Long

Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal nIndex As Long) As Long

Public Declare Function FindWindow _
                Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Integer

Public Declare Function ShowWindow _
                Lib "user32.dll" (ByVal hwnd As Long, _
                ByVal nCmdShow As Long) As Long

Private Const WM_USER = &H400

Private Const EM_FORMATRANGE  As Long = WM_USER + 57

Private Const PHYSICALOFFSETX As Long = 112

Private Const PHYSICALOFFSETY As Long = 113

Private Type RECT

    Left           As Long
    Top            As Long
    Right          As Long
    Bottom         As Long

End Type

Private Type CharRange

    cpMin          As Long
    cpMax          As Long

End Type

Private Type FormatRange

    hdc            As Long
    hdcTarget      As Long
    rc             As RECT
    rcPage         As RECT
    chrg           As CharRange

End Type

'-----------------------------------------------------------------------------------------------------------

'нов тип променлива - използва се при автонастройката на flexgrid-a
Public Type SIZE

    cx As Long
    cy As Long

End Type

Public Function LoadLang()
    'функция за зареждане на езика

    '    Dim intEmpFileNbr1 As Integer
    '    Dim LangSet As String
    '
    '    intEmpFileNbr1 = FreeFile
    
    '    If Dir(LangSetFile) <> "" Then
    '        Open LangSetFile For Input As intEmpFileNbr1
    '        Do Until EOF(intEmpFileNbr1)
    '            Input #intEmpFileNbr1, LangSet
    '        Loop
    '        Close #intEmpFileNbr1
    '    Else
    TxtInfoCap = "Софтуер за диспечеризация и управление на заявки към бетонови стопанства"
    TxtVerCap = "Диспечер ТИП-Панел v1.2/2014 мод VIPA"
            
    MsgAdminSuccess = "Създаденият администратор е:"
    MsgArhBx = "Архивиране БД"
    MsgArhivError = "Резервното копие на базата данни не е създадено"
    MsgArhivReady = "Резервното копие на базата данни е създадено"
    MsgArhivQuest = "Искате ли да направите архив на базата данни до момента!"
    MsgAnotherRun = "Програмата вече е активна!"
    MsgAvaria = "Има постъпил сигнал за авария от машина " & MachineNumber & "!"
    MsgBusyFlow = "Има материал в тази течка!"
    MsgCallTIP = "Свържете се с ТИП-Сервиз ЕООД!"
    MsgCantDelRec = "По тази рецепта има текуща заявка! Изтриването е неуспешно!"
    MsgCantDelClnt = "За този клиент има текуща заявка! Изтриването е неуспешно!"
    MsgCantSvFutDelivery = "Не може да запишете доставка с бъдеща дата!"
    MsgClose = "Изход от Диспечерския Панел - Машина " & MachineNumber & "?"
    MsgClWatRev = "Да нулирам ли разхода на вода до момента?"
    MsgCodeZero = "Въведете код различен от 0!"
    MsgConfCrTables = "Да създам ли таблиците?"
    MsgConfDel = "Изтриване на запис?"
    MsgConfEdit = "Съществуващ запис в базата данни! Потвърдете редактирането!"
    MsgContinue = "Желаете ли да продължите изчакването за резултати от машината?"
    MsgCorData = "Има нулеви резултати от контролера! Желаете ли да коригирате данните?"
    MsgDBConnEst = "Връзката базата данни е осъществена!"
    MsgConnEst = "Връзката с OPC-Server е възстановена! Рестартирайте програмата!"
    MsgDaysLeft = "Оставащи дни активност на програмата"
    MsgDelBx = "Изтриване"
    MsgDelSuccess = "Записът е изтрит!"
    MsgEditBx = "Редактиране"
    MsgEndOnCancel = "При отказ програма спира!"
    MsgErrBx = "Грешка"
    MsgErrBxFatal = "Грешка - Датата на компютъра не помага на блокировката!"
    MsgErrNoData = "Няма данни от контролера"
    MsgErrExcel = "Експортирането е неуспешно. Случиха се следните грешки :"
    MsgExpWait = "Имате чакаща заявка в контролера на машина " & MachineNumber & "!"
    MsgFillAll = "Моля, попълнете всички полета!"
    MsgInvalLic = "Грешен лицензен ключ."
    MsgKey = "Лицензен ключ: "
    MsgMatLoaded = "Материалът е зареден в течка и не може да бъде изтрит!"
    MsgMatDelivery = "Има доставено количество и материалът не може да бъде изтрит!"
    MsgMatNotQuant = "Има избран материал без да е посочено количество!"
    MsgMatSold = "Има направен разход и материалът не може да бъде изтрит!"
    MsgMaxL15 = "Максималната дължина на име/парола е 15 символа!"
    MsgMkAdmin = "Няма администратор! Моля, създайте администратор."
    MsgNewName = "Изберете друго име!"
    MsgNoClnt = "Няма такъв номер на клиент!"
    MsgNoDBConn = "Няма връзка с база данни!"
    MsgNoExcel = "Нямате инсталиран MS Excel!"
    MsgNoLic = "Нямате лицензен ключ."
    MsgNoMat = "Няма въведените материали!"
    MsgNoOps = "Няма въведени оператори!"
    MsgNoPayment = "Лицензът на програмата е прекратен поради непостъпило окончателно плащане!"
    MsgNoPaymentDays = "Лицензът на програмата ще бъде прекратен поради непостъпило окончателно плащане!"
    MsgNoResOnExit = "При изход програмата няма да следи за резултати!"
    MsgNoSelection = "Няма маркиран запис!"
    MsgNotEnQuant = "Няма достатъчна наличност за изписване"
    MsgNotRespOPC = "Няма връзка!"
    MsgNotWorkDB = "Няма работеща база данни на Вашия компютър!"
    MsgOffline = "Няма връзка с OPC-Server! Да продължа ли в офлайн режим?"
    MsgOverCapDrv = "Експедицията ще превиши капацитета на превозното средство! Да продължа ли?"
    MsgPassNotConf = "Паролата не е потвърдена!"
    MsgQuantNotMat = "Има посочено количество без да е избран материал!"
    MsgRecErrCem = "Грешка в рецептата за цимента!"
    MsgRecErrChem = "Грешка в рецептата за химическите добавки!"
    MsgRecErrIM = "Грешка в рецептата за инертния материал!"
    MsgRecErrWat = "Грешка в рецептата за водата!"
    MsgResBx = "Въстановяване БД"
    MsgRestoreError = "Базата данни не е възстановена"
    MsgRestoreReady = "Базата данни е възстановена"
    MsgSameNmFamOp = "Има такъв оператор! Въведете друго име / фамилия!"
    MsgSaveSuccess = "Успешен запис!"
    MsgTablesNotFound = "Следните таблици не бяха открити в базата данни: "
    MsgTablesReady = "Следните таблици бяха успешно създадени в базата данни: "
    MsgWrong = "Грешно име / парола!"
    MsgWrongPass = "Грешна парола!"
        
    lblChooseArhFile = "Посочете архивен файл. Датата и часа на архива се съдържат в името на файла във формат - YYYYMMDD-HHMMSS"
    lblChooseFolder = "Посочете папка за архивното копие"
    lblEntIP = "IP за достъп до база данни (локално = 127.0.0.1)"
    lblEntPass = "Парола за достъп до база данни"
    lblLabPassCap = "Въведете парола за контрол на достъп за корекция на рецепти:"
    lblOperCap = "Оператор код:"
    lblPassCap = "Парола:"
    lblPassConfCap = "Потвърдете:"
    lblRevInfo = "След запис автоматично се праща разпечатка към принтер!"
    lblRevision = "Въведете всички данни от ревизията в тонове!"
    lblSetClients = "Настройка на видими и скрити клиенти в списъка"
    lblSetDrivers = "Настройка на видими и скрити водачи в списъка"
    lblSetObjects = "Настройка на видими и скрити обекти в списъка"
    lblSetRecepies = "Настройка на видими и скрити рецепти в списъка"
    lblSetSuppliers = "Настройка на видими и скрити доставчици в списъка"
        
    frmAdPanel = "Администраторски Панел ТИП-Панел v1.2"
    frmConfSend = "Потвърждение за стартиране на производствен цикъл"
    frmDispPanelCap = "Машина " & MachineNumber & " - ТИП-Панел v1.2"
    frmDataCor = "Корекция на пропуснати данни за експедиция"
    frmLabPassCap = "Парола за Лаборант"
    frmNewAd = "Нов Администратор"
    frmNmSilos = "Наименования на течките"
    frmDBdata = "Достъп до база данни"
        
    btnCreateOp = "Създай оператор"
    btnEditAd = "Промяна на администратор"
    btnEditOp = "Редактирай оператор"
    btnSendControllerCap = "ИЗПРАТИ КЪМ МАШИНА " & MachineNumber & " !"
    btnParamSys = "Параметри на машина " & MachineNumber
        
    statAuto = "автоматичен режим"
    statMan = "ръчен режим"
    statAv = "авариен режим"
    statAvaria = "АВАРИЯ"
    statAvStop = "авариен стоп"
    statReqStarted = "стартирана заявка"
    statReadyReqNew = "готов за заявка"
    statMixOn = "миксер включен"
    statMixOff = "миксер изключен"
    statMixOpened = "шибър отворен"
    statMixClosed = "шибър затворен"
    statBeltOn = "лентата изсипва"
    statBeltOff = "лента спряна"
    statCemOpened = "цимент отворен"
    statCemClosed = "цимент затворен"
    statWatOpened = "вода отворена"
    statWatClosed = "вода затворена"
    statChemOpened = "химия отворена"
    statChemClosed = "химия затворена"
        
    UniCancel = "Отказ"
    UniEnter = "Вход"
    UniExit = "Изход"
    UniOK = "OK"
        
    uniAbout = "За Програмата..."
    uniActDel = "Активиране на бутоните за изтриване на записи в режим оператор"
    uniActForm1 = "Редактиране на съдържанието на Форма 1 в режим оператор"
    uniActForm2 = "Редактиране на съдържанието на Форма 2 в режим оператор"
    uniActForm3 = "Редактиране на съдържанието на Форма 3 в режим оператор"
    uniAdd = "Адрес"
    uniAdmin = "Администратор"
    uniAllow = "Разрешения"
    uniAutoPrint = "Авто-печат на избраните форми след експедицията"
        
    uniBG = "БУЛСТАТ"
    
    uniCapacity = "Капацитет"
    uniCem = "Силоз"
    uniCement = "Цимент"
    uniCemShort = "Сил"
    uniChem = "Химическа добавка"
    uniChemShort = "ХД"
    uniClass = "Клас по якост"
    uniClassK = "Клас по консист."
    uniClassV = "Клас по възд."
    uniClassH = "Клас по с-е хлориди"
    uniClassP = "Водоплътност"
    uniClnt = "Клиент"
    uniClntCode = "Клиент код"
    uniClnts = "Клиенти"
    uniCode = "Код"
    uniComInfo = "Информация за фирмата"
    uniConcPlant = "Бетонов възел"
    uniConMat = "Свързващо вещество"
        
    uniDate = "Дата"
    uniDateDlvr = "Дата на доставка"
    uniDateOrd = "Дата приемане"
    uniDateReady = "Дата и час за изпълнение"
    uniDateReadyShort = "Дата готовност"
    uniDeactNRPass = "Деактивиране на паролата за запис и редакция на рецепти в режим оператор"
    uniDeactDRPass = "Деактивиране на паролата за изтриване на рецепти в режим оператор"
    uniDel = "Изтрий"
    uniDelivered = "Доставено [t]"
    uniDisp = "Диспечер"
    uniDispQuant = "Количество експедиция"
    uniDlvr = "Доставка ..."
    uniDlvrs = "Доставки"
    uniDrv = "Водач"
    uniDrvCode = "Водач код"
    uniDrvReg = "Кола рег. No."
    uniDrvs = "Водачи"
            
    uniEDM = "ЕДМ"
    uniEmpty = "празна течка"
    uniEnterExp = "Въведи разход"
    uniExped = "Експедиция"
        
    uniFam = "Фамилия"
    uniFax = "Факс"
    uniFirm = "Фирма"
    uniFlow = "Течка"
    uniForm1 = "Форма 1"
    uniForm2 = "Форма 2"
    uniForm3 = "Форма 3"
        
    uniHave = "Наличност [t]"
        
    uniIM = "Инертен материал"
    uniIMShort = "ИМ"
    uniInfoClnt = "Данни за клиент"
    uniInfoDrv = "Данни за водач"
    uniInfoMix = "Данни за 1 замес"
        
    uniKGmeasured = "тегло измерено"
    uniKGstated = "тегло заявено"
    uniKm = "Разстояние до обекта"
    uniKmShort = "Разстояние"
        
    uniLoad = "Заредено в"
    uniLoaded = "- Свързан -"
    uniLoading = "Свързване"
    uniLog = "Регистър"
        
    uniMade = "Изпълнено"
    uniMat = "Материал"
    uniMats = "Материали"
    uniMeasured = "Измерено"
    uniMix = "Замес"
    uniMixCap = "Миксер капацитет"
    uniMod = "Марка/Модел"
    uniMOL = "МОЛ"
        
    uniNew = "Нов"
    uniNewa = "Нова"
    uniNewNm = "Ново име"
    uniNm = "Име"
    uniNo = "не"
    uniNoDoc = "Документ номер"
    uniNote = "Бележка"
    uniNotes = "Бележки"
    uniNr = "No."
    uniNumCem = "Брой течки цимент"
    uniNumChem = "Брой течки ХД"
    uniNumIM = "Брой течки ИМ"
    uniNumMix = "Брой замеси"
    uniNumWat = "Брой течки вода"
        
    uniObj = "Обект"
    uniOpenMix = "Време шибър отворен"
    uniOrdCode = "Заявка код"
    uniOrdered = "Заявено"
    uniOrds = "Заявки"
    uniOrdsVert = "з а я в к и"
    uniOrdQuant = "Количество заявка"
    uniOther = "Други"
            
    uniPrint = "Печат..."
    uniPump = "Помпа"
        
    uniQuant = "Количество"
    uniQuantMix = "Количество за 1 замес"
    uniQuestPrint = "Експедицията е готова! Да принтирам ли експедиционна бележка?"
        
    uniReadyVert = "п о с л е д н а"
    uniReady2Vert = "е к с п е д и ц и я"
    uniRec = "Рецепта"
    uniRecCode = "Рецепта код"
    uniRecNote = "Рецептите се въвеждат за 1 кубически метър бетонов разтвор и се записват с номерата на течките, а не имената на материалите в тях!"
    uniRecType = "Вид р-р"
    uniRecs = "Рецепти"
    uniResults = "Резултати"
    uniRevision = "Ревизия"
    uniRevisor = "Ревизор"
        
    uniSave = "Запис"
    uniSendingPrinter = "Изпращане за печат..."
    uniSettings = "Настройки"
    uniSheet = "лист "
    uniSold = "Разход [t]"
    uniSrchBG = "Търсене"
    uniSTART = "СТАРТ !"
    uniStatus = "Общо състояние на системата"
    uniSup = "Доставчик"
    uniSups = "Доставчици"
        
    uniTank = "Цистерна"
    uniTel = "Телефон"
    uniTempResults = "Временни резултати"
    uniTimeMix = "Време замес"
    uniTimeMixShort = "tmix"
    uniTimePour = "Време изсипване"
    uniTimePourShort = "tout"
    uniTotalKg = "Общо кг"
    uniTown = "град"
    uniType = "Тип"
    uniTypeDoc = "Документ вид"
        
    uniVolmeasured = "обем измерено"
        
    uniWat = "Вода"
    uniWat1cb = "Вода за 1 м3"
        
    uniYes = "да"
    '    End If
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

Public Function GetSerialNumber(ByVal sDrive As String) As Long
    'функция за прочитане на серен номер на харддиск

    If Len(sDrive) Then
        If InStr(sDrive, "\\") = 1 Then

            ' Make sure we end in backslash for UNC
            If Right$(sDrive, 1) <> "\" Then
                sDrive = sDrive & "\"
            End If

        Else
            ' If not UNC, take first letter as drive
            sDrive = Left$(sDrive, 1) & ":\"
        End If

    Else
        ' Else just use current drive
        sDrive = vbNullString
    End If

    ' Grab S/N -- Most params can be NULL
    Call GetVolumeInformation(sDrive, vbNullString, 0, GetSerialNumber, ByVal 0&, ByVal 0&, vbNullString, 0)
End Function

Public Function ProdKeyGen() As String
    'функция за генериране на лицензен номер от хард-диск с:
    
    Dim Drive     As String

    Dim VolSerNum As String

    Dim VolSerDec As Long

    Dim CodeCalc1 As Long

    Dim CodeCalc2 As Long

    Dim CodeCalc3 As Long

    Const CardNum = 139967

    Const Pi = 3.14

    Const Dat1 = 1.9

    Const Dat2 = 29.04

    'read volume C: s/n
    Drive = "C"
    VolSerNum = GetSerialNumber(Drive)
    VolSerDec = CDec(Mid(VolSerNum, 1, 10))

    If VolSerDec < 0 Then
        VolSerDec = VolSerDec * -1
    Else
    End If

    CodeCalc1 = (VolSerDec + CardNum) / Pi
    CodeCalc2 = Dat1 * (CodeCalc1 + CardNum)
    CodeCalc3 = 2 * (CodeCalc1 + CardNum) / Dat2
    ProdKeyGen = Hex$(CodeCalc1) & "-" & Hex$(CodeCalc2) & "-" & Hex$(CodeCalc3)
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

Public Sub AutoColW(ListViewTemp As ListView)
    'автоформатиране на таблица в listview според текста в клетките и заглавките
    
    Const LVM_FIRST = &H1000

    Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)

    Const LVSCW_AUTOSIZE_USEHEADER = -2

    Dim i As Long
  
    With ListViewTemp
        SendMessage .hwnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE_USEHEADER

        For i = 1 To .ColumnHeaders.count - 1
            SendMessage .hwnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE_USEHEADER
        Next

    End With

End Sub

Public Sub FlexGrid_AutoSizeColumns(ByRef pGrid As MSFlexGrid, _
                                    ByRef pForm As Form, _
                                    Optional ByVal pIncludeHeaderRows As Boolean = True, _
                                    Optional ByVal pAllowShrink As Boolean = True, _
                                    Optional ByVal pMinCol As Long = 0, _
                                    Optional ByVal pMaxCol As Long = -1, _
                                    Optional ByVal pBorderSize As Long = 8)
    'Set flexgrid column widths to the minimum for viewing all text
 
    'Note that this will not be accurate if Cells have different fonts,
    'or if .FontWidth (or .CellFontWidth) has been set
 
    'Parameters:
    '  pGrid              - the grid to work with
    '  pForm              - the form the grid is on
    '  pIncludeHeaderRows - whether to take the width of text in FixedRows into account
    '  pAllowShrink       - allow column widths to get smaller than current?
    '  pMinCol            - the first column to work with
    '  pMaxCol            - the last column to work with (-1 means the right-most column)
    '  pBorderSize        - the number of pixels used as a border around text (seems like 8 to me!)
 
    Dim lngMinCol   As Long, lngMaxCol As Long, lngCurrRow As Long

    Dim lngMinRow   As Long, lngMaxRow As Long, lngCurrCol As Long

    Dim lngMaxWidth As Long, lngCurrWidth As Long

    Dim fntFormFont As StdFont
 
    'Store current form font (so can restore later)
    Set fntFormFont = New StdFont
    Call CopyFont(pForm.Font, fntFormFont)
    'Set font of form to same as grid, to get accurate values
    Call CopyFont(pGrid.Font, pForm.Font)
 
    With pGrid                'Set rows/columns to check
        lngMinCol = pMinCol
        lngMaxCol = IIf(pMaxCol = -1, .Cols - 1, pMaxCol)
        lngMinRow = IIf(pIncludeHeaderRows, 0, .FixedRows)
        lngMaxRow = .Rows - 1

        'For each column in specified range..
        For lngCurrCol = lngMinCol To lngMaxCol
            '..set min allowed size based on options
            lngMaxWidth = IIf(pAllowShrink, 0, pForm.ScaleX(.ColWidth(lngCurrCol), vbTwips, pForm.ScaleMode))
 
            For lngCurrRow = lngMinRow To lngMaxRow   '..find widest text (in scalemode of the form)
                lngCurrWidth = pForm.TextWidth(.TextMatrix(lngCurrRow, lngCurrCol))

                If lngMaxWidth < lngCurrWidth Then lngMaxWidth = lngCurrWidth
            Next lngCurrRow

            '..as the scalemode of the form may differ, convert to twips
            lngMaxWidth = pForm.ScaleX(lngMaxWidth, pForm.ScaleMode, vbTwips)
            '..resize the column as apt (with specified border size)
            .ColWidth(lngCurrCol) = lngMaxWidth + (pBorderSize * Screen.TwipsPerPixelX)
        Next lngCurrCol

    End With

    'Restore form font
    Call CopyFont(fntFormFont, pForm.Font)
 
End Sub

Public Sub CopyFont(ByVal pFontFrom As StdFont, ByRef pFontTo As StdFont)
    'Copy the properties of a font object to another
 
    With pFontFrom
        pFontTo.Bold = .Bold
        pFontTo.Charset = .Charset
        pFontTo.Italic = .Italic
        pFontTo.Name = .Name
        pFontTo.SIZE = .SIZE
        pFontTo.Strikethrough = .Strikethrough
        pFontTo.Underline = .Underline
        pFontTo.Weight = .Weight
    End With
 
End Sub

Public Function GetDecimalSep() As String
    'функция за извикване на сепаратора за дробни числа от настройките на системата
    
    Dim LCID    As Integer

    Dim Data    As String

    Dim ret     As Integer

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
    'функция за смяна на десетичния сепаратор при четене спрямо този от настройките на компютъра

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

Public Function rDsNew(ByVal str As String)
    'функция за смяна на десетичния сепаратор със точка
    
    If Len(str) = 0 Then
        rDsNew = "0"

        Exit Function

    End If
    
    If InStr(str, ",") <> 0 Then
        rDsNew = Replace(str, ",", ".")
    Else
        rDsNew = str
    End If

End Function


Public Function ARound(ByVal MyNumber, ByVal Deci)
    'функция за закръгление на числа
      
    ARound = Int(MyNumber * 10 ^ Deci + 1 / 2) / 10 ^ Deci
End Function

Public Function DecToBin(ByVal DeciValue As Single, _
                         Optional NoOfBits As Integer = 8) As String
    'функция за конвертиране на десетични числа в двоични
    
    Dim i As Integer
    Dim bay As String
    Dim newDec As Long
    newDec = ARound(DeciValue, 0)
    
    If DeciValue >= 3.5 And DeciValue < 4 Then newDec = 3
    If DeciValue >= 7.5 And DeciValue < 8 Then newDec = 7
    If DeciValue >= 15.5 And DeciValue < 15 Then newDec = 15
    If DeciValue >= 31.5 And DeciValue < 32 Then newDec = 31
    If DeciValue >= 63.5 And DeciValue < 64 Then newDec = 63
    
    Do While DeciValue > (2 ^ NoOfBits) - 1
        NoOfBits = NoOfBits + 8
    Loop
    
    DecToBin = vbNullString
    
    For i = 0 To (NoOfBits - 1)

        If (DeciValue < 2 ^ i) Then
            DecToBin = "0" & DecToBin
        Else
            bay = newDec And 2 ^ i
            DecToBin = CStr(bay / 2 ^ i) & DecToBin
        End If

    Next i

End Function

Public Function BinToDec(Binary As String) As Long
    'функция за конвертиране на двоични числа в десетични
    
    Dim n As Long

    Dim s As Integer

    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ (s - 1)))
    Next s

    BinToDec = n
End Function

Public Function Mantisse(Binary As String) As Single
    'функция за изчисляване на мантисата на двоични числа
    
    Dim M1, M2, M3, M4, M5 As Single

    M1 = ((((Mid$(Binary, 23, 1) / 2 + Mid$(Binary, 22, 1)) / 2 + Mid$(Binary, 21, 1)) / 2 + Mid$(Binary, 20, 1)) / 2 + Mid$(Binary, 19, 1)) / 2
    M2 = (((((M1 + Mid(Binary, 18, 1)) / 2 + Mid$(Binary, 17, 1)) / 2 + Mid$(Binary, 16, 1)) / 2 + Mid$(Binary, 15, 1)) / 2 + Mid$(Binary, 14, 1)) / 2
    M3 = (((((M2 + Mid$(Binary, 13, 1)) / 2 + Mid$(Binary, 12, 1)) / 2 + Mid$(Binary, 11, 1)) / 2 + Mid$(Binary, 10, 1)) / 2 + Mid$(Binary, 9, 1)) / 2
    M4 = (((((M3 + Mid$(Binary, 8, 1)) / 2 + Mid$(Binary, 7, 1)) / 2 + Mid$(Binary, 6, 1)) / 2 + Mid$(Binary, 5, 1)) / 2 + Mid$(Binary, 4, 1)) / 2
    M5 = (((M4 + Mid$(Binary, 3, 1)) / 2 + Mid$(Binary, 2, 1)) / 2 + Mid$(Binary, 1, 1)) / 2
    Mantisse = M5
End Function

Public Function IEEE754(BCD As Single) As Variant
    'функция за конвертиране на числа от формат IEEE754 към десетични реални стойности
    
    Dim Bin     As String

    Dim SignBin As String

    Dim ExpoBin As String

    Dim MantBin As String

    Dim Sign    As Integer

    Dim Expo    As Variant

    Dim Mant    As Single
    
    If BCD < 1000000000 Then
        IEEE754 = 0
    Else
        Bin = DecToBin(BCD, 32)
        SignBin = Mid$(Bin, 1, 1)
        ExpoBin = Mid$(Bin, 2, 8)
        MantBin = Mid$(Bin, 10, 23)
        Sign = (-1) ^ SignBin
        Expo = 2 ^ (BinToDec(ExpoBin) - 127)
        Mant = 1 + Mantisse(MantBin)
        IEEE754 = Sign * Expo * Mant
    End If

End Function

Public Function ToIEEE754(realNum As Single) As Long
    'функция за конвертиране на десетични реални числа към формата IEEE754
    
    Dim M(1 To 23)    As Integer

    Dim ExpoDec       As Single

    Dim ToIEEE754Bin  As String

    Dim MantCalc      As Single

    Dim SignBit       As Long

    Dim BinPrep       As String

    Dim ExpoBit       As Long

    Dim ExpoReady     As Long

    Dim ExpoDiv       As Single

    Dim MantPrep      As Single

    Dim MantisseReady As String

    Dim MantReduct    As Single
    
    MantReduct = 1
    
    If realNum = 0 Then
        ToIEEE754 = 0

        Exit Function

    Else
    End If
    
    If realNum > 0 Then
        SignBit = 0
    Else
        ToIEEE754 = 0

        Exit Function

    End If
    
    If realNum >= 1 Then
        BinPrep = DecToBin(realNum, 32)
        i = 0

        Do While Mid$(BinPrep, 1 + i, 1) = 0
            counter = counter + 1
            i = i + 1

            If i = 32 Then
                counter = counter - 1
                GoTo Jump
            Else
            End If

        Loop

    Else
        i = 0

        Do While 2 ^ counter > realNum
            counter = counter - 1
            i = i + 1

            If i = 32 Then
                counter = counter - 1
                GoTo Jump
            Else
            End If

        Loop

    End If
                
Jump:

    If realNum >= 1 Then
        ExpoBit = Len(BinPrep) - counter - 1
    Else
        ExpoBit = counter
    End If
    
    ExpoDec = ExpoBit + 127
    ExpoReady = DecToBin(ExpoDec, 8)
    ExpoDiv = 2 ^ ExpoBit
    MantPrep = realNum / ExpoDiv
    MantCalc = (MantPrep - MantReduct)
    i = 1

    For i = 1 To 23

        If (MantCalc * 2) > 1 Then
            M(i) = "1"
            MantCalc = (2 * MantCalc) - MantReduct
            GoTo FlagOut
        Else
        End If
        
        If (MantCalc * 2) = 1 Then
            M(i) = "1"
            MantCalc = 0
            GoTo FlagOut
        Else
        End If
        
        If (MantCalc * 2) < 1 And MantCalc > 0 Then
            M(i) = "0"
            MantCalc = 2 * MantCalc
            GoTo FlagOut
        Else
        End If
        
        If MantCalc < 0 Then
            M(i) = "0"
            GoTo FlagOut
        Else
        End If

FlagOut:
    Next i

    MantisseReady = M(1) & M(2) & M(3) & M(4) & M(5) & M(6) & M(7) & M(8) & M(9) & M(10) & M(11) & M(12) & M(13) & M(14) & M(15) & M(16) & M(17) & M(18) & M(19) & M(20) & M(21) & M(22) & M(23)
    ToIEEE754Bin = SignBit & ExpoReady & MantisseReady
    ToIEEE754 = BinToDec(ToIEEE754Bin)
End Function

Public Function PrintLVPic(lvw As ListView, _
                           Orient As Integer, _
                           HeadPrnt As Boolean, _
                           NowPrnt As Boolean, _
                           PageNumPrnt As Boolean, _
                           Optional NamePage As String = "", _
                           Optional TopMargPerc As Integer = 500, _
                           Optional LeftMargPerc As Integer = 500, _
                           Optional not1 As Integer = 0, _
                           Optional not2 As Integer = 0, _
                           Optional not3 As Integer = 0)
    'функция за принтиране на ListView към PictureBox
    'извиква след всяка страница функция за печат на PictureBox като му прави AutoFit
    'към A4 според избрана ориентация 1-портретно, 2-пейзажно
    'преработена за печат на много страници
    'избира се True / False за печат на хедърите на ListView
    'задава се заглавие на страница - по избор
    'печат на дата и час - True/False
    'пропуска печат до 3 колони по избор на номера им
    'за да работи функцията трябва да има форма - frmPrint и PictureBox - pbPrint на нея

    Const MARGIN = 60

    Const COL_MARGIN = 150

    Dim ymin         As Single

    Dim ymax         As Single

    Dim xmin         As Single

    Dim xmax         As Single

    Dim num_cols     As Integer

    Dim list_item    As ListItem

    Dim i            As Integer

    Dim num_subitems As Integer

    Dim col_wid()    As Single

    Dim X            As Single

    Dim Xpage        As Single

    Dim Y            As Single

    Dim line_hgt     As Single

    Dim start        As Integer

    Dim listCount    As Integer

    frmPrint.Show
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    listCount = 0
    Line = 1
Again:

    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    ymin = 0
    ymax = 0
    xmin = 0
    xmax = 0
    num_cols = 0
    i = 0
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
    For i = 1 To num_cols

        If i <> not1 And i <> not2 And i <> not3 Then
            col_wid(i) = frmPrint.pbPrint.TextWidth(lvw.ColumnHeaders(i).Text)
        End If

    Next i

    ' Check the items.
    num_subitems = num_cols - 1

    For Each list_item In lvw.ListItems

        ' Check the item.
        If col_wid(1) < frmPrint.pbPrint.TextWidth(list_item.Text) Then col_wid(1) = frmPrint.pbPrint.TextWidth(list_item.Text)

        ' Check the subitems.
        For i = 1 To num_subitems

            If i <> not1 - 1 And i <> not2 - 1 And i <> not3 - 1 Then
                If col_wid(i + 1) < frmPrint.pbPrint.TextWidth(list_item.SubItems(i)) Then col_wid(i + 1) = frmPrint.pbPrint.TextWidth(list_item.SubItems(i))
            End If

        Next i
    Next list_item
    
    ' Add a column margin.
    For i = 1 To num_cols

        If i <> not1 And i <> not2 And i <> not3 Then
            col_wid(i) = col_wid(i) + COL_MARGIN
        End If

    Next i
    
    'изчисляваме ширината на PictureBox-a според ширините на колоните
    Xpage = MARGIN

    For i = 1 To num_subitems + 1
        Xpage = Xpage + col_wid(i)
    Next i

    frmPrint.pbPrint.Width = Xpage + MARGIN + 100
    
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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    ' *************************
    ' Print the column headers.
    frmPrint.pbPrint.CurrentY = ymin + MARGIN
    frmPrint.pbPrint.CurrentX = MARGIN
    X = xmin + MARGIN

    For i = 1 To num_cols

        If i <> not1 And i <> not2 And i <> not3 Then
            frmPrint.pbPrint.CurrentX = X

            If HeadPrnt = True Then
                frmPrint.pbPrint.Print FittedText(lvw.ColumnHeaders(i).Text, col_wid(i));
            End If

            X = X + col_wid(i)
        End If

    Next i
    
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
        For i = 1 To num_subitems

            If i <> not1 - 1 And i <> not2 - 1 And i <> not3 - 1 Then
                frmPrint.pbPrint.CurrentX = X
                frmPrint.pbPrint.Print FittedText(lvw.ListItems.Item(start).SubItems(i), col_wid(i + 1));
                X = X + col_wid(i + 1)
            End If

        Next i
        
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

    For i = 1 To num_cols - 1

        If i <> not1 And i <> not2 And i <> not3 Then
            X = X + col_wid(i)
            frmPrint.pbPrint.Line (X, ymin)-(X, ymax)
        End If

    Next i
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
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

Private Sub ScalePic(ByVal Percentage As Single, _
                     ByVal sngTop As Single, _
                     ByVal sngLeft As Single)

    'Dim intHeight As Integer
    'Dim intWidth As Integer
    Dim sngRatio As Single

    Dim intLeft  As Integer

    Dim intTop   As Integer

    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    'веднага след приключване на фунцията трябва да зададем Printer.enddoc

End Sub

Public Sub PrintThePicture(frm As Form, _
                           Pic As PictureBox, _
                           Optional PercentOfPage As Integer = 100, _
                           Optional LeftMarginPercent As Integer = 0, _
                           Optional TopMarginPercent As Integer = 0)
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

    ScalePic PercentOfPage, TopMarginPercent, LeftMarginPercent
End Sub

Private Sub ScalePicMix(ByVal Percentage As Single, _
                        ByVal sngTop As Single, _
                        ByVal sngLeft As Single)

    'Dim intHeight As Integer
    'Dim intWidth As Integer
    Dim sngRatio As Single

    Dim intLeft  As Integer

    Dim intTop   As Integer

    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    Percentage = Percentage / 100
    '    sngTop = sngTop / 100
    '    sngLeft = sngLeft / 100
    
    '   Scale the picture to either use the full width
    '   or the full height of the page.
    If frmPrintMix.pbPrint.Width > (Printer.ScaleWidth * Percentage) Then
        sngRatio = (Printer.ScaleWidth * Percentage) / frmPrintMix.pbPrint.Width
    Else
        sngRatio = 1
    End If
    
    If frmPrintMix.pbPrint.Height * sngRatio > (Printer.ScaleHeight * Percentage) Then
        sngRatio = (Printer.ScaleHeight * Percentage) / frmPrintMix.pbPrint.Height
    Else
    End If
    
    '   Center the picture on the page.
    '    intLeft = (Printer.ScaleWidth - (picTemp.Width * sngRatio)) * sngLeft
    '    intTop = (Printer.ScaleHeight - (picTemp.Height * sngRatio)) * sngTop
    intLeft = sngLeft
    intTop = sngTop
      
    '   send the picture to the printer
    Printer.PaintPicture frmPrintMix.pbPrint.Image, intLeft, intTop, frmPrintMix.pbPrint.Width * sngRatio, frmPrintMix.pbPrint.Height * sngRatio
    '    Printer.EndDoc
    
    '   Cleanup
    Set picTemp = Nothing
    frmPrintMix.pbPrint.Cls
    frmPrintMix.pbPrint.Refresh
'    frmPrintMix.Controls.Remove "picTemp"
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    'веднага след приключване на фунцията трябва да зададем Printer.enddoc

End Sub

Public Sub PrintThePictureMix(frm As Form, _
                              Pic As PictureBox, _
                              Optional PercentOfPage As Integer = 100, _
                              Optional LeftMarginPercent As Integer = 0, _
                              Optional TopMarginPercent As Integer = 0)
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

    ScalePicMix PercentOfPage, TopMarginPercent, LeftMarginPercent
End Sub

Public Function ExportToExcel(lvw As MSComctlLib.ListView) As Boolean
    'функция за експортиране на ListView в Excel
 
    Dim objExcel            As Object

    Dim objWorkbook         As Object

    Dim objWorksheet        As Object

    Dim objRange            As Object
     
    '    Dim lngResults As Long
    Dim i                   As Integer

    Dim intCounter          As Integer

'    Dim intStartRow         As Integer

    Dim strArray()          As String

    Dim intVisibleColumns() As Integer

    Dim intColumns          As Integer

    Dim itm                 As ListItem

'    Dim Fname               As String
 
'    Fname = lvw.Name
    
    MousePointer = vbHourglass
    
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
    objRange.Font.SIZE = 10
    objRange.Font.Bold = True

    For i = 1 To lvw.ColumnHeaders.count

        If lvw.ColumnHeaders(i).Width <> 0 Then
            ' Create an array of visible column indexes
            intColumns = intColumns + 1
            ReDim Preserve intVisibleColumns(1 To intColumns)
            intVisibleColumns(intColumns) = i
            objRange.cells(1, intColumns) = lvw.ColumnHeaders(i).Text
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

    Next i
 
    ' Dimension array to number of listitems
    ReDim strArray(1 To lvw.ListItems.count, 1 To intColumns)
 
    intCounter = 0
'    intStartRow = 2

    For Each itm In lvw.ListItems

        ' A response of vbNo meant to export all the items
        '        If lngResults = vbNo Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1

        For i = 1 To intColumns

            If intVisibleColumns(i) = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If

        Next i

        '        End If
    Next itm
 
    ' Send entire array to Excel range
    With objWorksheet
        .Range(.cells(2, 1), .cells(2 + intCounter - 1, intColumns)) = strArray
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

Public Function PrintRTF(rtf As RichTextBox, _
                         nnLeftMarginWidth As Long, _
                         nnTopMarginHeight As Long, _
                         nnRightMarginWidth As Long, _
                         nnBottomMarginHeight As Long) As Boolean
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

    Dim nLeftOffset   As Long

    Dim nTopOffset    As Long

    Dim nLeftMargin   As Long

    Dim nTopMargin    As Long

    Dim nRightMargin  As Long

    Dim nBottomMargin As Long

    Dim fr            As FormatRange

    Dim rcDrawTo      As RECT

    Dim rcPage        As RECT

    Dim nTextLength   As Long

    Dim nNextCharPos  As Long

    Dim nRet          As Long

    MousePointer = vbHourglass

    frmPrint.pbPrint.Print Space(1)
    nLeftOffset = frmPrint.pbPrint.ScaleX(GetDeviceCaps(frmPrint.pbPrint.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
   
    nTopOffset = frmPrint.pbPrint.ScaleY(GetDeviceCaps(frmPrint.pbPrint.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)
   
    nLeftMargin = nnLeftMarginWidth - nLeftOffset
    nTopMargin = nnTopMarginHeight - nTopOffset
    nRightMargin = (frmPrint.pbPrint.Width - nnRightMarginWidth) - nLeftOffset
   
    nBottomMargin = (frmPrint.pbPrint.Height - nnBottomMarginHeight) - nTopOffset
   
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
        nNextCharPos = SendMessage(rtf.hwnd, EM_FORMATRANGE, True, fr)

        If nNextCharPos >= nTextLength Then Exit Do
        fr.chrg.cpMin = nNextCharPos
        frmPrint.pbPrint.Print Space(1)
   
    Loop

    nRet = SendMessage(rtf.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    MousePointer = vbDefault
    PrintRTF = True

    Exit Function

ErrorHandler:
    PrintRTF = False
    MousePointer = vbDefault
End Function

Public Function PrintRTFMix(rtf As RichTextBox, _
                            nnLeftMarginWidth As Long, _
                            nnTopMarginHeight As Long, _
                            nnRightMarginWidth As Long, _
                            nnBottomMarginHeight As Long) As Boolean
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

    Dim nLeftOffset   As Long

    Dim nTopOffset    As Long

    Dim nLeftMargin   As Long

    Dim nTopMargin    As Long

    Dim nRightMargin  As Long

    Dim nBottomMargin As Long

    Dim fr            As FormatRange

    Dim rcDrawTo      As RECT

    Dim rcPage        As RECT

    Dim nTextLength   As Long

    Dim nNextCharPos  As Long

    Dim nRet          As Long

    MousePointer = vbHourglass

    frmPrintMix.pbPrint.Print Space(1)
    nLeftOffset = frmPrintMix.pbPrint.ScaleX(GetDeviceCaps(frmPrintMix.pbPrint.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
   
    nTopOffset = frmPrintMix.pbPrint.ScaleY(GetDeviceCaps(frmPrintMix.pbPrint.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)
   
    nLeftMargin = nnLeftMarginWidth - nLeftOffset
    nTopMargin = nnTopMarginHeight - nTopOffset
    nRightMargin = (frmPrintMix.pbPrint.Width - nnRightMarginWidth) - nLeftOffset
   
    nBottomMargin = (frmPrintMix.pbPrint.Height - nnBottomMarginHeight) - nTopOffset
   
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = frmPrintMix.pbPrint.ScaleWidth
    rcPage.Bottom = frmPrintMix.pbPrint.ScaleHeight
    rcDrawTo.Left = nLeftMargin
    rcDrawTo.Top = nTopMargin
    rcDrawTo.Right = nRightMargin
    rcDrawTo.Bottom = nBottomMargin
    fr.hdc = frmPrintMix.pbPrint.hdc
    fr.hdcTarget = frmPrintMix.pbPrint.hdc
    fr.rc = rcDrawTo
    fr.rcPage = rcPage
    fr.chrg.cpMin = 0
    fr.chrg.cpMax = -1
    nTextLength = Len(rtf.Text)

    Do
        fr.hdc = frmPrintMix.pbPrint.hdc
        fr.hdcTarget = frmPrintMix.pbPrint.hdc
        nNextCharPos = SendMessage(rtf.hwnd, EM_FORMATRANGE, True, fr)

        If nNextCharPos >= nTextLength Then Exit Do
        fr.chrg.cpMin = nNextCharPos
        frmPrintMix.pbPrint.Print Space(1)
   
    Loop

    nRet = SendMessage(rtf.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    MousePointer = vbDefault
    PrintRTFMix = True

    Exit Function

ErrorHandler:
    PrintRTFMix = False
    MousePointer = vbDefault
End Function

Public Function GetStatVIPA()

    'функция за следене на статуса на машината през OPC сървъра
    
    If OffMode = True Then
        DispPanel.indMode.Caption = ""
        DispPanel.indAvaria.Caption = ""
        DispPanel.indReq.Caption = ""
        DispPanel.indValveMix.Caption = ""
    Else

        DispPanel.indMode.Caption = ""
        DispPanel.indAvaria.Caption = ""
        DispPanel.indReq.Caption = ""
        DispPanel.indValveMix.Caption = ""

        For i = 0 To 5

            If frmOPC.Stat(0).Text = "True" Then
                DispPanel.indMode.Caption = statAuto
                DispPanel.indMode.ForeColor = &HC000&
                DispPanel.indReq.Caption = statReqStarted
            Else
                DispPanel.indMode.Caption = statMan
                DispPanel.indMode.ForeColor = &HC00000
            End If

            If frmOPC.Stat(1).Text = "False" Then
                DispPanel.indAvaria.Caption = statAvaria
                DispPanel.indAvaria.ForeColor = &HFF&
            Else
                DispPanel.indAvaria.Caption = ""
            End If
            
            If frmOPC.Stat(2).Text = "True" Then
                DispPanel.indValveMix.Caption = statMixClosed
                DispPanel.indValveMix.ForeColor = &HFF&
            ElseIf frmOPC.Stat(3).Text = "True" Then
                DispPanel.indValveMix.Caption = statMixOpened
                DispPanel.indValveMix.ForeColor = &HC000&
            End If
            
        Next i

    End If

End Function

Public Function MaintConn() As Boolean
    'функция за следене на връзката с OPC сървъра

    If frmOPC.Stat(4).Text = "False" Then
        DispPanel.lblLoading.ForeColor = &HC000&
        DispPanel.lblLoading.Caption = uniLoaded
        DispPanel.lblLoading.Refresh
        MaintConn = True
    Else
        DispPanel.lblLoading.ForeColor = &HFF&
        DispPanel.lblLoading.Caption = MsgNotRespOPC
        DispPanel.lblLoading.Refresh
        MaintConn = False
    End If

End Function

Public Function OpenDisp()
    'зареждане на меню диспечер

    Dim ResDisp As Result

    Set ResDisp = New Result
    
    Dim CountMixd As Integer
    
    Dim colx      As ColumnHeader

    Dim itmX      As ListItem
    
    Const colw = 1000

    Const colwn = 1300
    
    MixCap = CSng(rDs(frmOPC.Config(0)))
    
    'преместване на рамката Диспечер във видимата част на формата
    DispPanel.frDisp.Left = 120
    DispPanel.frDisp.Top = DispPanel.Height \ 2 - FrmPlace

    'настройка на табулаторите
'    DispPanel.cmbDispOrd.SetFocus
    DispPanel.cmbDispOrd.TabIndex = 0
    DispPanel.cmbDispDrv.TabIndex = 1
    DispPanel.cmbDispDrvName.TabIndex = 2
    DispPanel.txtDispWat.TabIndex = 3
    DispPanel.txtDispQuant.TabIndex = 4
    DispPanel.chPrintConf.TabIndex = 5
    DispPanel.btnDispStart.TabIndex = 6
    DispPanel.chPrintConf.TabIndex = 7
    DispPanel.btnDisp.TabIndex = 8
    DispPanel.btnOrders.TabIndex = 9
    DispPanel.btnRecepies.TabIndex = 10
    DispPanel.btnClients.TabIndex = 11
    DispPanel.btnDrivers.TabIndex = 12
    DispPanel.btnSuppliers.TabIndex = 13
    DispPanel.btnMaterials.TabIndex = 14
    DispPanel.btnNotes.TabIndex = 15
    DispPanel.btnAdminPanel.TabIndex = 16
    DispPanel.btnExit.TabIndex = 17
    
    'почистване на заглавките и клетките на заявките
    DispPanel.lstOrdWait.ColumnHeaders.Clear
    DispPanel.lstOrdWait.ListItems.Clear
    
    'зареждане на заглавките на заявките
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 1000
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniDateOrd
    colx.Width = 1750
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniDateReadyShort
    colx.Width = 1750
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniOrdered
    colx.Width = 850
    colx.Tag = Number
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniMade
    colx.Width = 850
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniRecCode
    colx.Width = 1150
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniNm & " " & uniRec
    colx.Width = 1200
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniClass
    colx.Width = 1300
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniClntCode
    colx.Width = 1100
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniNm & " " & uniClnt
    colx.Width = 1300
    Set colx = DispPanel.lstOrdWait.ColumnHeaders.Add()
    colx.Text = uniObj
    colx.Width = 1300

    If DispPanel.indReq.Caption <> statReqStarted Then
        'почистване на заглавките и клетките на последната експедиция
        DispPanel.lstMixReady.ColumnHeaders.Clear
        DispPanel.lstMixReady.ListItems.Clear
    
        'зареждане на заглавките на експедициите от файл
        intEmpFileNbr1 = FreeFile
        Open SilosFile For Input As #intEmpFileNbr1

        Do Until EOF(intEmpFileNbr1)
            Input #intEmpFileNbr1, IM(1), IM(2), IM(3), IM(4), IM(5), IM(6), Scr(1), Scr(2), Scr(3), Scr(4), Wat(1), Wat(2), Chem(1), Chem(2), Chem(3), Chem(4), Chem(5), Chem(6)
        Loop

        Close #intEmpFileNbr1

        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniNr
        colx.Width = 500
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniOrdCode
        colx.Width = 1150
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniRec & " " & uniNm
        colx.Width = 1200
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniClass
        colx.Width = 1300

        For count = 1 To ns1
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = IM(count) 'имена на течките на им
            colx.Width = colwn
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        For count = 1 To ns3
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = Scr(count) 'имена на течки цимент
            colx.Width = colwn
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        For count = 1 To ns2
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = Wat(count) 'име на течка вода
            colx.Width = colw
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        For count = 1 To ns4
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = Chem(count) 'имена на течки хд
            colx.Width = colwn
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = "тегло заявено" 'тегло по заявено
        colx.Width = 1200
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = "тегло измерено" 'тегло по измерено
        colx.Width = 1200
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = "обем измерено" 'обем по измерено
        colx.Width = 1200
    End If

    'почистване на полетата за въвеждане на данни
    DispPanel.cmbDispOrd.Clear
    DispPanel.cmbDispDrv.Clear
    DispPanel.cmbDispDrvName.Clear
    DispPanel.txtDispDrvReg.Text = ""
    DispPanel.txtDispDrvCap.Text = 0
    DispPanel.txtDispRec.Text = ""
    DispPanel.txtDispRecName.Text = ""
    DispPanel.txtDispRecClass.Text = ""
    DispPanel.txtDispClnt.Text = ""
    DispPanel.txtDispClntName.Text = ""
    DispPanel.txtDispClntObj.Text = ""
    DispPanel.txtDispOrdDate.Text = ""
    DispPanel.txtDispOrdQuant.Text = 0
    DispPanel.txtDispQuant.Text = 0
    DispPanel.txtDispQuant.MaxLength = 5
    DispPanel.txtDispWat.Text = 0
    DispPanel.txtDispWat.MaxLength = 3
    
    'нулиране на брояча на замесите и сумарните тегла
    CountMixd = 0
    ResDisp.TotalStatedKG = 0
    ResDisp.TotalMeasuredKG = 0
    ResDisp.TotalQuant = 0
    
    '------------------------------Start PostgreSQL----------------------------------
    Dim cn   As ADODB.Connection

    Dim rs   As Recordset

    Dim comm As String

    Dim i    As Integer
    
    Set cn = New ADODB.Connection 'връзка с база данни
    cn.ConnectionTimeout = 10
    cn.Open ConStr 'отваряме връзката
    MousePointer = vbHourglass
    'визуализация на заявките от днешния ден
    comm = "SELECT * FROM orders WHERE stamp_date >= '" & DayToday & "';"
    Set rs = cn.Execute(comm) 'маркиране на заявките отговарящи на дата днес
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Set itmX = DispPanel.lstOrdWait.ListItems.Add(1, , Format(rs!order_num, "0000000")) 'зареждане в ListView
        itmX.SubItems(1) = rs!order_date
        itmX.SubItems(2) = rs!order_date_que
        itmX.SubItems(3) = rDs(rs!order_q)
        itmX.SubItems(4) = rDs(rs!order_qmade)
        itmX.SubItems(5) = Format(rs!order_rec, "0000")
        itmX.SubItems(6) = rs!order_rec_name
        itmX.SubItems(7) = rs!order_rec_class
        itmX.SubItems(8) = Format(rs!order_clnt, "0000")
        itmX.SubItems(9) = rs!order_clnt_name
        itmX.SubItems(10) = rs!order_clnt_obj
        DispPanel.cmbDispOrd.AddItem Format(rs!order_num, "0000000") 'зареждане на комбото със номерата на активните заявки
        rs.MoveNext
    Loop

    If DispPanel.indReq.Caption <> statReqStarted Then
        'визуализация на замесите от последната експедиция
        Set rs = cn.Execute("SELECT mix_num, exp_num, exp_q FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1") 'маркираме последния замес
    
        If Not rs.EOF And Not rs.BOF Then
            i = Val(rs!exp_num) 'маркираме номера на експедиция от последния замес
        Else 'ако няма замеси
            GoTo LoadDrv
        End If
    
        'обновяване на статус бара
        DispPanel.StatusBar.Panels(3) = "Замеси: " & rs!mix_num
        DispPanel.StatusBar.Panels(4) = "Експедиции: " & rs!exp_num
        DispPanel.StatusBar.Panels(5) = "Заявена експедиция: " & rDs(rs!exp_q) & " m3"
        rs.Close
    
        comm = "SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & i & " ORDER BY mix_num ASC;"
    
        Set rs = cn.Execute(comm) 'маркираме всички замеси от намерената експедиция
        rs.MoveFirst 'отиваме в началото на маркираните замеси

        Do While Not rs.EOF

            If Val(rs!exp_num) <> i Then Exit Do 'излизаме ако номера на експедицията не отговаря на търсенето
        
            ResDisp.ExpQuant = rDs(rs!exp_q) 'запис на данните от базата дани в масиви от променливи
            ResDisp.IMname(1) = rs!im1_name
            ResDisp.IMname(2) = rs!im2_name
            ResDisp.IMname(3) = rs!im3_name
            ResDisp.IMname(4) = rs!im4_name
            ResDisp.IMname(5) = rs!im5_name
            ResDisp.IMname(6) = rs!im6_name
            ResDisp.IMstated(1) = rs!im1z
            ResDisp.IMstated(2) = rs!im2z
            ResDisp.IMstated(3) = rs!im3z
            ResDisp.IMstated(4) = rs!im4z
            ResDisp.IMstated(5) = rs!im5z
            ResDisp.IMstated(6) = rs!im6z
            ResDisp.IMmeasured(1) = rs!im1i
            ResDisp.IMmeasured(2) = rs!im2i
            ResDisp.IMmeasured(3) = rs!im3i
            ResDisp.IMmeasured(4) = rs!im4i
            ResDisp.IMmeasured(5) = rs!im5i
            ResDisp.IMmeasured(6) = rs!im6i
            ResDisp.SCRname(1) = rs!cem1_name
            ResDisp.SCRname(2) = rs!cem2_name
            ResDisp.SCRname(3) = rs!cem3_name
            ResDisp.SCRname(4) = rs!cem4_name
            ResDisp.SCRstated(1) = rs!cem1z
            ResDisp.SCRstated(2) = rs!cem2z
            ResDisp.SCRstated(3) = rs!cem3z
            ResDisp.SCRstated(4) = rs!cem4z
            ResDisp.SCRmeasured(1) = rs!cem1i
            ResDisp.SCRmeasured(2) = rs!cem2i
            ResDisp.SCRmeasured(3) = rs!cem3i
            ResDisp.SCRmeasured(4) = rs!cem4i
            ResDisp.WATname(1) = rs!wat1_name
            ResDisp.WATstated(1) = rs!wat1z
            ResDisp.WATmeasured(1) = rs!wat1i
            ResDisp.WATname(2) = rs!wat2_name
            ResDisp.WATstated(2) = rs!wat2z
            ResDisp.WATmeasured(2) = rs!wat2i
            ResDisp.CHEMname(1) = rs!chem1_name
            ResDisp.CHEMname(2) = rs!chem2_name
            ResDisp.CHEMname(3) = rs!chem3_name
            ResDisp.CHEMname(4) = rs!chem4_name
            ResDisp.CHEMname(5) = rs!chem5_name
            ResDisp.CHEMname(6) = rs!chem6_name
            ResDisp.CHEMstated(1) = rDs(rs!chem1z)
            ResDisp.CHEMstated(2) = rDs(rs!chem2z)
            ResDisp.CHEMstated(3) = rDs(rs!chem3z)
            ResDisp.CHEMstated(4) = rDs(rs!chem4z)
            ResDisp.CHEMstated(5) = rDs(rs!chem5z)
            ResDisp.CHEMstated(6) = rDs(rs!chem6z)
            ResDisp.CHEMmeasured(1) = rDs(rs!chem1i)
            ResDisp.CHEMmeasured(2) = rDs(rs!chem2i)
            ResDisp.CHEMmeasured(3) = rDs(rs!chem3i)
            ResDisp.CHEMmeasured(4) = rDs(rs!chem4i)
            ResDisp.CHEMmeasured(5) = rDs(rs!chem5i)
            ResDisp.CHEMmeasured(6) = rDs(rs!chem6i)
            ResDisp.TotalStatedKG = ResDisp.TotalStatedKG + CSng(rDs(rs!total_rec_kg)) 'сумираме теглата по заявено от всеки замес
            ResDisp.TotalMeasuredKG = ResDisp.TotalMeasuredKG + CSng(rDs(rs!total_real_kg)) 'сумираме теглата по измерено от всеки замес
            ResDisp.TotalQuant = ResDisp.TotalQuant + CSng(rDs(rs!total_vol)) 'сумираме количества от всеки замес
        
            CountMixd = CountMixd + 1 'брояч на замесите
        
            Set itmX = DispPanel.lstMixReady.ListItems.Add(1, , Format(CountMixd, "00")) 'запис в ListView
            itmX.SubItems(1) = Format(rs!ord_num, "0000000")
            itmX.SubItems(2) = rs!name_rec
            itmX.SubItems(3) = rs!class_rec

            For e = 1 To ns1
                itmX.SubItems(2 * e + 2) = ResDisp.IMstated(e)
                itmX.SubItems(2 * e + 3) = ResDisp.IMmeasured(e)
            Next e

            For e = 1 To ns3
                itmX.SubItems(2 * (e + ns1) + 2) = ResDisp.SCRstated(e)
                itmX.SubItems(2 * (e + ns1) + 3) = ResDisp.SCRmeasured(e)
            Next e

            For e = 1 To ns2
                itmX.SubItems(2 * (e + ns1 + ns3) + 2) = ResDisp.WATstated(1)
                itmX.SubItems(2 * (e + ns1 + ns3) + 3) = ResDisp.WATmeasured(1)
            Next e

            For e = 1 To ns4
                itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 2) = ResDisp.CHEMstated(e)
                itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 3) = ResDisp.CHEMmeasured(e)
            Next e

            itmX.SubItems(2 * (ns1 + ns3 + ns2 + ns4 + 1) + 2) = rDs(rs!total_rec_kg)
            itmX.SubItems(2 * (ns1 + ns3 + ns2 + ns4 + 1) + 3) = rDs(rs!total_real_kg)
            itmX.SubItems(2 * (ns1 + ns3 + ns2 + ns4 + 1) + 4) = rDs(rs!total_vol)

            rs.MoveNext
        Loop
        
        'почистване на заглавките
        DispPanel.lstMixReady.ColumnHeaders.Clear
    
        'зареждане на заглавките на експедициите от базата данни

        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniNr
        colx.Width = 500
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniOrdCode
        colx.Width = 1150
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniRec & " " & uniNm
        colx.Width = 1200
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniClass
        colx.Width = 1300

        For count = 1 To ns1
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = ResDisp.IMname(count) 'имена на течките на им
            colx.Width = colwn
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        For count = 1 To ns3
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = ResDisp.SCRname(count) 'имена на течки цимент
            colx.Width = colwn
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        For count = 1 To ns2
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = ResDisp.WATname(count) 'име на течка вода
            colx.Width = colw
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        For count = 1 To ns4
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = ResDisp.CHEMname(count) 'имена на течки хд
            colx.Width = colwn
            Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
            colx.Text = uniMeasured
            colx.Width = colw
        Next count

        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniKGstated 'тегло по заявено
        colx.Width = 1200
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniKGmeasured 'тегло по измерено
        colx.Width = 1200
        Set colx = DispPanel.lstMixReady.ColumnHeaders.Add()
        colx.Text = uniVolmeasured 'обем по измерено
        colx.Width = 1200
    End If

LoadDrv:
    'зареждане на комбобоксовете на водача
    Set rs = cn.Execute("SELECT d_num, d_name FROM drivers WHERE d_show = '1' ORDER BY d_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        DispPanel.cmbDispDrv.AddItem Format(Val(rs!d_num), "0000")
        DispPanel.cmbDispDrvName.AddItem rs!d_name
        rs.MoveNext
    Loop
    
    rs.Close 'затваряме записите
    Set rs = Nothing
    cn.Close 'прекъсваме връзката с базата данни
    MousePointer = vbDefault
    Set cn = Nothing
    '-------------------------------End PostgreSQL-------------------------------------------------

    'обновяване на статус бара
    If ExpeditionStarted = False Then
        DispPanel.StatusBar.Panels(6) = "Обем експедиция: " & ResDisp.TotalQuant & " m3"
        DispPanel.StatusBar.Panels(7) = "Тегло експедиция: " & ResDisp.TotalMeasuredKG & " kg"
        DispPanel.StatusBar.Refresh
    End If
    
    Set ResDisp = Nothing
    
    'автонастройка на ListView
    If DispPanel.lstMixReady.ListItems.count > 0 Then AutoColW DispPanel.lstMixReady
    If DispPanel.lstOrdWait.ListItems.count > 0 Then AutoColW DispPanel.lstOrdWait
End Function

Public Function ChangeDispOrd()
    'функция за смяна на заявката от диспечера и зареждане на данните
    
    If DispPanel.cmbDispOrd.Text <> "" Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
    
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        Set rs = cn.Execute("SELECT * FROM orders WHERE order_num = " & Val(DispPanel.cmbDispOrd.Text) & ";")
    
        DispPanel.txtDispRec.Text = Format(rs!order_rec, "0000")
        DispPanel.txtDispRecName.Text = rs!order_rec_name
        DispPanel.txtDispRecClass.Text = rs!order_rec_class
        DispPanel.txtDispClnt.Text = Format(rs!order_clnt, "0000")
        DispPanel.txtDispClntName.Text = rs!order_clnt_name
        DispPanel.txtDispClntObj.Text = rs!order_clnt_obj
        DispPanel.txtDispOrdDate.Text = Left(rs!order_date, 10)
        DispPanel.txtDispOrdQuant.Text = rDs(rs!order_q)
        
        'зареждане данни за водата по рецептата на съответната заявка
        Set rs = cn.Execute("SELECT init_wat1, kg_wat1, init_wat2, kg_wat2 FROM recepies WHERE r_num = " & Val(DispPanel.txtDispRec.Text) & ";")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
        
        Do While Not rs.EOF

            If Val(rs!init_wat1) = 21 Then
                DispPanel.txtDispWat.Text = Val(rs!kg_wat1)
            ElseIf Val(rs!init_wat2) = 21 Then
                DispPanel.txtDispWat.Text = Val(DispPanel.txtDispWat.Text) + Val(rs!kg_wat2)
            Else
                DispPanel.txtDispWat.Text = Val(DispPanel.txtDispWat.Text) + 0
            End If

            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------

    Else 'празни полета ако няма маркирана заявка
        DispPanel.txtDispRec.Text = ""
        DispPanel.txtDispRecName.Text = ""
        DispPanel.txtDispRecClass.Text = ""
        DispPanel.txtDispClnt.Text = ""
        DispPanel.txtDispClntName.Text = ""
        DispPanel.txtDispClntObj.Text = ""
        DispPanel.txtDispQuant.Text = 0
        DispPanel.txtDispOrdDate.Text = ""
        DispPanel.txtDispOrdQuant.Text = 0
        DispPanel.txtDispWat.Text = 0
    End If

End Function

Public Function ChangeDispDrv()
    'функция за смяна на водача от диспечера и зареждане на данните
    
    If DispPanel.cmbDispDrv.Text <> "" Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
    
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        Set rs = cn.Execute("SELECT d_name, d_reg, d_cap FROM drivers WHERE d_num = " & Val(DispPanel.cmbDispDrv.Text) & ";")
        
        DispPanel.cmbDispDrvName.Text = rs!d_name
        DispPanel.txtDispDrvReg.Text = rs!d_reg
        DispPanel.txtDispDrvCap.Text = rDs(rs!d_cap)
        
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------
    
    Else
        DispPanel.cmbDispDrvName.Text = ""
        DispPanel.txtDispDrvReg.Text = ""
        DispPanel.txtDispDrvCap.Text = 0
    End If

End Function

Public Function ChangeDispDrvName()
    'функция за смяна на името на водача от диспечера и зареждане на данните
    
    If DispPanel.cmbDispDrvName.Text <> "" Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
    
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        Set rs = cn.Execute("SELECT d_num, d_reg, d_cap FROM drivers WHERE d_name = '" & DispPanel.cmbDispDrvName.Text & "';")
        
        If Not rs.EOF Or Not rs.BOF Then
            DispPanel.cmbDispDrv.Text = Format(Val(rs!d_num), "0000")
            DispPanel.txtDispDrvReg.Text = rs!d_reg
            DispPanel.txtDispDrvCap.Text = rDs(rs!d_cap)
        End If
        
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------

    Else
        DispPanel.cmbDispDrv.ListIndex = -1
        DispPanel.txtDispDrvReg.Text = ""
        DispPanel.txtDispDrvCap.Text = 0
    End If

End Function

Public Function ListOrdWaitClick()
    'функция за зареждане на количество вода по рецепта при маркиране на запис от таблицата
    
    If DispPanel.frDisp.Enabled = True And DispPanel.frDisp.Visible = True Then
        If DispPanel.lstOrdWait.ListItems.count > 0 Then
            DispPanel.cmbDispOrd.Text = Format(Val(DispPanel.lstOrdWait.ListItems(DispPanel.lstOrdWait.SelectedItem.Index).Text), "0000000")
        End If

    Else
    End If

End Function

Public Function DispConfSend()

    'функция за визуализация на данните преди изпращане към контролера
        
    Dim RecVis  As Recipe

    Dim ClntVis As Client

    Dim DrvVis  As Driver

    Set RecVis = New Recipe
    Set ClntVis = New Client
    Set DrvVis = New Driver
        
    Dim nZ             As Single

    Dim nCoef          As Single

    Dim intEmpFileNbr1 As Integer
    
    Dim ChSilosKg(1 To 4) As Single
    
    Dim ChSilosOrd(0 To 3) As Single
    
    Dim TotalSilosExpKg(0 To 3) As Single
    
    Dim i As Integer
    
    DispPanel.btnDispStart.Visible = False
    
    MousePointer = vbDefault
    
    DispConfirm.txtConfDisp.Text = DispPanel.cmbDispOrd.Text
    DispConfirm.txtConfDispDate.Text = DispPanel.txtDispOrdDate.Text
    DispConfirm.txtConfOrdQuant.Text = DispPanel.txtDispOrdQuant.Text
    DispConfirm.txtConfDispQuant.Text = DispPanel.txtDispQuant.Text
    DispConfirm.txtConfRec.Text = DispPanel.txtDispRec.Text
    DispConfirm.txtConfClnt.Text = DispPanel.txtDispClnt.Text
    DispConfirm.txtConfDrv.Text = DispPanel.cmbDispDrv.Text
    
    nQuant = ARound(CSng(rDs(DispConfirm.txtConfDispQuant.Text)), 2)
    MixCap = CSng(rDs(frmOPC.Config(0)))
    nZ = nQuant / MixCap
    nZ = IIf(Int(nZ + 1) - nZ = 1, nZ, Int(nZ + 1))
    nCycle = nZ
    DispConfirm.txtConfDispCount.Text = nCycle
    
    If nQuant = 0 Then
        nCoef = 0
    Else
        nCoef = ARound(nQuant / nCycle, 4)
    End If
    
    DispConfirm.txtConfCoef.Text = nCoef

    For i = 0 To 5
        DispConfirm.txtConfRecKg1(i).Visible = False
        DispConfirm.txtConfRec1(i).Visible = False
    Next i

    For i = 0 To ns1 - 1
        DispConfirm.txtConfRecKg1(i).Visible = True
        DispConfirm.txtConfRec1(i).Visible = True
    Next i
    
    For i = 0 To 3
        DispConfirm.txtConfRecKg3(i).Visible = False
        DispConfirm.txtConfRec3(i).Visible = False
    Next i

    For i = 0 To ns3 - 1
        DispConfirm.txtConfRecKg3(i).Visible = True
        DispConfirm.txtConfRec3(i).Visible = True
    Next i

    For i = 0 To 1
        DispConfirm.txtConfRecKg2(i).Visible = False
        DispConfirm.txtConfRec2(i).Visible = False
    Next i

    For i = 0 To ns2 - 1
        DispConfirm.txtConfRecKg2(i).Visible = True
        DispConfirm.txtConfRec2(i).Visible = True
    Next i

    For i = 0 To 5
        DispConfirm.txtConfRecKg4(i).Visible = False
        DispConfirm.txtConfRec4(i).Visible = False
    Next i

    For i = 0 To ns4 - 1
        DispConfirm.txtConfRecKg4(i).Visible = True
        DispConfirm.txtConfRec4(i).Visible = True
    Next i
    
    'прочитане на файл с имена на течките
    intEmpFileNbr1 = FreeFile
    
    Open SilosFile For Input As #intEmpFileNbr1

    Do Until EOF(intEmpFileNbr1)
        Input #intEmpFileNbr1, IM(1), IM(2), IM(3), IM(4), IM(5), IM(6), Scr(1), Scr(2), Scr(3), Scr(4), Wat(1), Wat(2), Chem(1), Chem(2), Chem(3), Chem(4), Chem(5), Chem(6)
    Loop

    Close #intEmpFileNbr1

    If DispConfirm.txtConfRec.Text <> "" And DispConfirm.txtConfClnt.Text <> "" And DispConfirm.txtConfDrv.Text <> "" Then

        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
        
        Dim rsChSilos As Recordset
        
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        'маркираме необходимата рецепта от базата данни и я прочитаме с променливите
        Set rs = cn.Execute("SELECT * FROM recepies WHERE r_num = " & Val(DispConfirm.txtConfRec.Text) & ";")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
        
        Do While Not rs.EOF
            RecVis.Code = rs!r_num
            RecVis.Title = rs!r_name
            RecVis.Kind = rs!r_type
            RecVis.Class = rs!r_class
            RecVis.ClassK = rs!r_classk
            RecVis.ClassV = rs!r_classv
            RecVis.ClassH = rs!r_classh
            RecVis.ClassP = rs!r_classp
            RecVis.EDM = rs!r_edm
            RecVis.Tpour = rs!r_tpour
            RecVis.Tmix = rs!r_tmix
            RecVis.initIM(1) = rs!init_im1
            RecVis.kgIM(1) = rs!kg_im1
            RecVis.initIM(2) = rs!init_im2
            RecVis.kgIM(2) = rs!kg_im2
            RecVis.initIM(3) = rs!init_im3
            RecVis.kgIM(3) = rs!kg_im3
            RecVis.initIM(4) = rs!init_im4
            RecVis.kgIM(4) = rs!kg_im4
            RecVis.initIM(5) = rs!init_im5
            RecVis.kgIM(5) = rs!kg_im5
            RecVis.initIM(6) = rs!init_im6
            RecVis.kgIM(6) = rs!kg_im6
            If CInt(rs!init_scr1) > 0 Then
                RecVis.initScr(1) = 30 + CInt(DispPanel.numSilos(CInt(Right(rs!init_scr1, 1)) - 1))
            Else
                RecVis.initScr(1) = rs!init_scr1
            End If
            RecVis.kgScr(1) = rs!kg_scr1
            If CInt(rs!init_scr2) > 0 Then
                RecVis.initScr(2) = 30 + CInt(DispPanel.numSilos(CInt(Right(rs!init_scr2, 1)) - 1))
            Else
                RecVis.initScr(2) = rs!init_scr2
            End If
            RecVis.kgScr(2) = rs!kg_scr2
            If CInt(rs!init_scr3) > 0 Then
                RecVis.initScr(3) = 30 + CInt(DispPanel.numSilos(CInt(Right(rs!init_scr3, 1)) - 1))
            Else
                RecVis.initScr(3) = rs!init_scr3
            End If
            RecVis.kgScr(3) = rs!kg_scr3
            If CInt(rs!init_scr4) > 0 Then
                RecVis.initScr(4) = 30 + CInt(DispPanel.numSilos(CInt(Right(rs!init_scr4, 1)) - 1))
            Else
                RecVis.initScr(4) = rs!init_scr4
            End If
            RecVis.kgScr(4) = rs!kg_scr4
            RecVis.initWat(1) = rs!init_wat1
            RecVis.kgWat(1) = rs!kg_wat1
            RecVis.initWat(2) = rs!init_wat2
            RecVis.kgWat(2) = rs!kg_wat2
            RecVis.initChem(1) = rs!init_chem1
            RecVis.kgChem(1) = rDs(rs!kg_chem1)
            RecVis.initChem(2) = rs!init_chem2
            RecVis.kgChem(2) = rDs(rs!kg_chem2)
            RecVis.initChem(3) = rs!init_chem3
            RecVis.kgChem(3) = rDs(rs!kg_chem3)
            RecVis.initChem(4) = rs!init_chem4
            RecVis.kgChem(4) = rDs(rs!kg_chem4)
            RecVis.initChem(5) = rs!init_chem5
            RecVis.kgChem(5) = rDs(rs!kg_chem5)
            RecVis.initChem(6) = rs!init_chem6
            RecVis.kgChem(6) = rDs(rs!kg_chem6)
            
            rs.MoveNext
        Loop
        
        'маркираме необходимия клиент от базата данни и го прочитаме с променливите
        Set rs = cn.Execute("SELECT * FROM clients where c_num =" & Val(DispConfirm.txtConfClnt.Text) & ";")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
        
        Do While Not rs.EOF
            ClntVis.Code = Val(rs!c_num)
            ClntVis.Title = rs!c_name
            ClntVis.Ident = rs!c_bg
            ClntVis.MOL = rs!c_mol
            ClntVis.Address = rs!c_add
            ClntVis.Phone = rs!c_tel
        
            rs.MoveNext
        Loop

        'маркираме необходимия обект от базата данни и го прочитаме
        Set rs = cn.Execute("SELECT w_name, w_km FROM worksites where w_cnum ='" & Val(ClntVis.Code) & "' AND w_name = '" & DispPanel.txtDispClntObj.Text & "';")
    
        If Not rs.EOF And Not rs.BOF Then
            rs.MoveFirst
        Else
            ClntVis.Worksite(1) = ""
            ClntVis.Distance(1) = 0
        End If
        
        Do While Not rs.EOF
            ClntVis.Worksite(1) = rs!w_name
            ClntVis.Distance(1) = Val(rs!w_km)
        
            rs.MoveNext
        Loop

        'маркираме необходимия водач от базата данни и го прочитаме с променливите
        Set rs = cn.Execute("SELECT * FROM drivers where d_num =" & Val(DispConfirm.txtConfDrv.Text) & ";")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
        
        Do While Not rs.EOF
            DrvVis.Code = Val(rs!d_num)
            DrvVis.Title = rs!d_name
            DrvVis.CarNum = rs!d_reg
            DrvVis.Capacity = rDs(rs!d_cap)
            DrvVis.CarModel = rs!d_mod
            DrvVis.Phone = rs!d_tel
        
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing
        
        'проверка на остатъчното количество материал в силозите
        Set rsChSilos = cn.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_type = '1' AND m_load <> ' 0 0 0 0' ORDER BY m_load DESC;")
        
        If Not rsChSilos.BOF And Not rsChSilos.EOF Then rsChSilos.MoveFirst
        
        i = 0
        
        Do While Not rsChSilos.EOF
            
            If rsChSilos!m_load = " 1 0 0 0" Then i = 1
            If rsChSilos!m_load = " 0 1 0 0" Then i = 2
            If rsChSilos!m_load = " 0 0 1 0" Then i = 3
            If rsChSilos!m_load = " 0 0 0 1" Then i = 4
            
            ChSilosKg(i) = CSng(rDs(rsChSilos!m_del)) - CSng(rDs(rsChSilos!m_sold))
            
            rsChSilos.MoveNext
        Loop
        
        rsChSilos.Close
        Set rsChSilos = Nothing
        
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------

        MousePointer = vbHourglass
        
        'повторна проверка за правилно маркирана рецепта от базата данни  и визуализация във формата
        If Val(DispConfirm.txtConfRec.Text) = RecVis.Code Then
            DispConfirm.txtConfRecName.Text = RecVis.Title
            DispConfirm.txtConfRecType.Text = RecVis.Kind
            DispConfirm.txtConfRecClass.Text = RecVis.Class
            DispConfirm.txtConfRecClassK.Text = RecVis.ClassK
            DispConfirm.txtConfRecClassV.Text = RecVis.ClassV
            DispConfirm.txtConfRecClassH.Text = RecVis.ClassH
            DispConfirm.txtConfRecClassP.Text = RecVis.ClassP
            DispConfirm.txtConfRecEDM.Text = RecVis.EDM
            DispConfirm.txtConfRecTimePour.Text = RecVis.Tpour
            DispConfirm.txtConfRecTimeMix.Text = RecVis.Tmix

            'предаваме данните за рецептата след като се уверим, че
            'имената на използваните течки са различни от "", "0" и "празна течка"
            For i = 0 To ns1 - 1
                For j = 1 To ns1

                    If RecVis.initIM(i + 1) = 10 + j And IM(j) <> "0" And IM(j) <> uniEmpty And IM(j) <> "" Then
                        DispConfirm.txtConfRec1(i).Text = IM(j)
                        DispConfirm.initIM(i).Text = RecVis.initIM(i + 1)
                    End If

                Next j
                VIPAkgIMs(i) = 0
            Next i

            For i = 0 To ns1 - 1

                If DispConfirm.txtConfRec1(i).Text <> "0" And DispConfirm.txtConfRec1(i).Text <> uniEmpty And DispConfirm.txtConfRec1(i).Text <> "" Then
                    DispConfirm.txtConfRecKg1(i).Text = ARound(RecVis.kgIM(i + 1) * nCoef, 0)
                    VIPAkgIMs(i) = RecVis.kgIM(i + 1)
                End If

            Next i
                
            For i = 0 To ns3 - 1
                For j = 1 To ns3

                    If RecVis.initScr(i + 1) = 30 + j And Scr(j) <> "0" And Scr(j) <> uniEmpty And Scr(j) <> "" Then
                        DispConfirm.txtConfRec3(i).Text = Scr(j)
                        DispConfirm.initCem(i).Text = RecVis.initScr(i + 1)
                        
                        ChSilosOrd(i) = ChSilosKg(j) * 1000
                    End If

                Next j
                VIPAkgSCRs(i) = 0
            Next i

            For i = 0 To ns3 - 1

                If DispConfirm.txtConfRec3(i).Text <> "0" And DispConfirm.txtConfRec3(i).Text <> uniEmpty And DispConfirm.txtConfRec3(i).Text <> "" Then
                    DispConfirm.txtConfRecKg3(i).Text = ARound(RecVis.kgScr(i + 1) * nCoef, 0)
                    VIPAkgSCRs(i) = RecVis.kgScr(i + 1)
                    
                    TotalSilosExpKg(i) = ARound(RecVis.kgScr(i + 1) * nCoef, 0) * nCycle
                    
                    'проверка за достатъчно количество цимент
                    If QuestSilos = True And TotalSilosExpKg(i) > 0 Then
                        If TotalSilosExpKg(i) > ChSilosOrd(i) Then
                            Result = MsgBox("Няма достатъчно количество материал в избрания силоз!" & vbNewLine & "Желаете ли да направите смяна на силозите?", vbQuestion Or vbYesNo, MsgErrBx)
                            If Result = vbYes Then
                                frmChSilos.Show
                                Unload DispConfirm
                                DispPanel.btnDispStart.Visible = True
                                DispPanel.btnDispStart.Enabled = True
                                
                                Set RecVis = Nothing
                                Set ClntVis = Nothing
                                Set DrvVis = Nothing

                                Exit Function
                            End If
                        End If
                    End If
                    
                End If

            Next i
            
            For i = 0 To ns2 - 1
                VIPAkgWATs(0) = 0
                VIPAkgWATs(1) = 0
                
                If RecVis.initWat(i + 1) = 21 And Wat(1) <> "0" And Wat(1) <> uniEmpty And Wat(1) <> "" Then
                    DispConfirm.txtConfRec2(i).Text = Wat(1)
                    DispConfirm.initWat(i).Text = 21
                    DispConfirm.txtConfRecKg2(i).Text = ARound(Val(DispPanel.txtDispWat.Text) * nCoef, 0)
                    VIPAkgWATs(0) = Val(DispPanel.txtDispWat.Text)
                End If

                If RecVis.initWat(i + 1) = 22 And Wat(2) <> "0" And Wat(2) <> uniEmpty And Wat(2) <> "" Then
                    DispConfirm.txtConfRec2(i).Text = Wat(2)
                    DispConfirm.initWat(i).Text = 22
                    DispConfirm.txtConfRecKg2(i).Text = ARound(RecVis.kgWat(2) * nCoef, 0)
                    VIPAkgWATs(1) = RecVis.kgWat(2)
                End If

            Next i
            
            For i = 0 To ns4 - 1
                For j = 1 To ns4

                    If RecVis.initChem(i + 1) = 40 + j And Chem(j) <> "0" And Chem(j) <> uniEmpty And Chem(j) <> "" Then
                        DispConfirm.txtConfRec4(i).Text = Chem(j)
                        DispConfirm.initChem(i).Text = RecVis.initChem(i + 1)
                    End If

                Next j
                VIPAkgCHEMs(i) = 0
            Next i

            For i = 0 To ns4 - 1

                If DispConfirm.txtConfRec4(i).Text <> "0" And DispConfirm.txtConfRec4(i).Text <> uniEmpty And DispConfirm.txtConfRec4(i).Text <> "" Then
                    DispConfirm.txtConfRecKg4(i).Text = ARound(RecVis.kgChem(i + 1) * nCoef, 2)
                    VIPAkgCHEMs(i) = RecVis.kgChem(i + 1)
                End If

            Next i

        Else
        End If

        'повторна проверка за правилно маркиран клиент от базата данни и визуализация във формата
        If Val(DispConfirm.txtConfClnt.Text) = ClntVis.Code Then
            DispConfirm.txtConfClntName.Text = ClntVis.Title
            DispConfirm.txtConfClntBG.Text = ClntVis.Ident
            DispConfirm.txtConfClntMOL.Text = ClntVis.MOL
            DispConfirm.txtConfClntAdd.Text = ClntVis.Address
            DispConfirm.txtConfClntTel.Text = ClntVis.Phone
            DispConfirm.txtConfClntObj.Text = ClntVis.Worksite(1)
            DispConfirm.txtConfClntKm.Text = ClntVis.Distance(1)
        Else
        End If
        
        'повторна проверка за правилно маркиран водач от базата данни и визуализация във формата
        If Val(DispConfirm.txtConfDrv.Text) = DrvVis.Code Then
            DispConfirm.txtConfDrvName.Text = DrvVis.Title
            DispConfirm.txtConfDrvReg.Text = DrvVis.CarNum
            DispConfirm.txtConfDrvCap.Text = DrvVis.Capacity
            DispConfirm.txtConfDrvMod.Text = DrvVis.CarModel
            DispConfirm.txtConfDrvTel.Text = DrvVis.Phone
        Else
        End If

    Else
        DispPanel.btnDispStart.Visible = True
        MousePointer = vbDefault

        Exit Function

    End If
    
    Set RecVis = Nothing
    Set ClntVis = Nothing
    Set DrvVis = Nothing
    MousePointer = vbDefault
End Function

Public Function SendToController()
'функция за изпращане на данните към контролера VIPA S-7 CPU-313
    
    MousePointer = vbHourglass
    
    Dim SendRec As SendExp
    Set SendRec = New SendExp
    
    Dim IMready(1 To 6) As Boolean
    Dim SCRready(1 To 4) As Boolean
    Dim WATready(1 To 2) As Boolean
    Dim CHEMready(1 To 4) As Boolean
    
    Dim SyncItemValuesRec(1 To 39) As Variant
    Dim SyncItemSrvHandlesRec(39) As Long
    Dim SyncItemSrvErrRec() As Long
    
    Dim SyncItemValuesReady(1 To 4) As Variant
    Dim SyncItemSrvHandlesReady(4) As Long
    Dim SyncItemSrvErrReady() As Long

    Dim i As Integer
    Dim nQuant As Single
    
    Dim rName(0 To 7) As String
    Dim rNameReady As String
    Dim rNameReady1 As String
    Dim rNameReady2 As String
    Dim lenRec As Integer
    
    'попълване на всички носещи масиви от формата за потвърждение до получаване на резултата
    cyc = Val(DispConfirm.txtConfDispCount.Text) 'брой замеси от експедицията
    
    If cyc <= 0 Then
        Unload DispConfirm
        Exit Function
    End If
    
    OrdData(0) = DispConfirm.txtConfDisp.Text
    DispPanel.stOrd.Caption = "Заявка: " & OrdData(0)
    OrdData(1) = DispConfirm.txtConfDispDate.Text
    OrdData(2) = DispConfirm.txtConfOrdQuant.Text
    OrdData(3) = DispConfirm.txtConfDispQuant.Text
    DispPanel.stExp.Caption = "Кол: " & OrdData(3)
    
    ClientData(0) = DispConfirm.txtConfClntName.Text
    DispPanel.stClnt.Caption = "Клиент: " & ClientData(0)
    ClientData(1) = DispConfirm.txtConfClntBG.Text
    ClientData(2) = DispConfirm.txtConfClntObj.Text
    ClientData(3) = DispConfirm.txtConfClntKm.Text
    
    DriverData(0) = DispConfirm.txtConfDrvName.Text
    DriverData(1) = DispConfirm.txtConfDrvReg.Text
    DriverData(2) = DispConfirm.txtConfDrvCap.Text
    
    For i = 1 To ns1
        SendRec.IMname(i) = DispConfirm.txtConfRec1(i - 1).Text
        SendRec.IMkg(i) = VIPAkgIMs(i - 1)
    Next i
    For i = 1 To ns3
        SendRec.SCRname(i) = DispConfirm.txtConfRec3(i - 1).Text
        SendRec.SCRkg(i) = VIPAkgSCRs(i - 1)
    Next i
    For i = 1 To ns2
        SendRec.WATname(i) = Wat(i)
        SendRec.WATkg(i) = VIPAkgWATs(i - 1)
    Next i
    For i = 1 To ns4
        SendRec.CHEMname(i) = DispConfirm.txtConfRec4(i - 1).Text
        SendRec.CHEMkg(i) = VIPAkgCHEMs(i - 1)
    Next i
    
    For g = 1 To 6
        kgIMs(g) = 0
        kgCHEMs(g) = 0
    Next g
    For g = 1 To 4
        kgSCRs(g) = 0
    Next g
    For g = 1 To 2
        kgWATs(g) = 0
    Next g
    
    For i = 1 To ns1
        Select Case SendRec.IMinit(i)
            Case 11
                kgIMs(1) = kgIMs(1) + Val(DispConfirm.txtConfRecKg1(i - 1).Text)
            Case 12
                kgIMs(2) = kgIMs(2) + Val(DispConfirm.txtConfRecKg1(i - 1).Text)
            Case 13
                kgIMs(3) = kgIMs(3) + Val(DispConfirm.txtConfRecKg1(i - 1).Text)
            Case 14
                kgIMs(4) = kgIMs(4) + Val(DispConfirm.txtConfRecKg1(i - 1).Text)
            Case 15
                kgIMs(5) = kgIMs(5) + Val(DispConfirm.txtConfRecKg1(i - 1).Text)
            Case 16
                kgIMs(6) = kgIMs(6) + Val(DispConfirm.txtConfRecKg1(i - 1).Text)
            Case Else
        End Select
    Next i
    For i = 1 To ns3
        Select Case SendRec.SCRinit(i)
            Case 31
                kgSCRs(1) = kgSCRs(1) + Val(DispConfirm.txtConfRecKg3(i - 1).Text)
            Case 32
                kgSCRs(2) = kgSCRs(2) + Val(DispConfirm.txtConfRecKg3(i - 1).Text)
            Case 33
                kgSCRs(3) = kgSCRs(3) + Val(DispConfirm.txtConfRecKg3(i - 1).Text)
            Case 34
                kgSCRs(4) = kgSCRs(4) + Val(DispConfirm.txtConfRecKg3(i - 1).Text)
            Case Else
        End Select
    Next i
    For i = 1 To ns2
        Select Case SendRec.WATinit(i)
            Case 21
                kgWATs(1) = kgWATs(1) + Val(DispConfirm.txtConfRecKg2(i - 1).Text)
            Case 22
                kgWATs(2) = kgWATs(2) + Val(DispConfirm.txtConfRecKg2(i - 1).Text)
            Case Else
        End Select
    Next i
    For i = 1 To ns4
        Select Case SendRec.CHEMinit(i)
            Case 41
                kgCHEMs(1) = ARound(kgCHEMs(1) + CSng(rDs(DispConfirm.txtConfRecKg4(i - 1).Text)), 2)
            Case 42
                kgCHEMs(2) = ARound(kgCHEMs(2) + CSng(rDs(DispConfirm.txtConfRecKg4(i - 1).Text)), 2)
            Case 43
                kgCHEMs(3) = ARound(kgCHEMs(3) + CSng(rDs(DispConfirm.txtConfRecKg4(i - 1).Text)), 2)
            Case 44
                kgCHEMs(4) = ARound(kgCHEMs(4) + CSng(rDs(DispConfirm.txtConfRecKg4(i - 1).Text)), 2)
            Case 45
                kgCHEMs(5) = ARound(kgCHEMs(5) + CSng(rDs(DispConfirm.txtConfRecKg4(i - 1).Text)), 2)
            Case 46
                kgCHEMs(6) = ARound(kgCHEMs(6) + CSng(rDs(DispConfirm.txtConfRecKg4(i - 1).Text)), 2)
            Case Else
        End Select
    Next i
    
    nCoefs = CSng(rDs(DispConfirm.txtConfCoef.Text))
    Recs = Val(DispConfirm.txtConfRec.Text)
    RecNames = DispConfirm.txtConfRecName.Text
    RecTypes = DispConfirm.txtConfRecType.Text
    RecClasss = DispConfirm.txtConfRecClass.Text
    RecClassKs = DispConfirm.txtConfRecClassK.Text
    RecClassVs = DispConfirm.txtConfRecClassV.Text
    RecClassHs = DispConfirm.txtConfRecClassH.Text
    RecClassPs = DispConfirm.txtConfRecClassP.Text
    RecEDMs = DispConfirm.txtConfRecEDM.Text
  
    If Recs <= 0 Then
        Unload DispConfirm
        DispPanel.TimerStartReq.Enabled = False
        DispPanel.TimerRes.Enabled = False
        Exit Function
    End If
  
'запис на променливите, които ще носят информацията към контролера-------------------------
'ако имена на течки участващи в рецептата са "", "0" или "празна течка" - записваме стойности 0 и не предаваме за дозиране към контролера
    
    SendRec.Tpour = Val(DispConfirm.txtConfRecTimePour.Text)
    SendRec.Tmix = Val(DispConfirm.txtConfRecTimeMix.Text)
    
    'показване на бутона старт след изпращане на данните
    DispPanel.btnDispStart.Visible = True
    
    'изчистване на клетките на екран диспечер
    DispPanel.cmbDispOrd.ListIndex = -1
    DispPanel.cmbDispDrv.ListIndex = -1
    DispPanel.cmbDispDrvName.Text = ""
    DispPanel.txtDispQuant.Text = "0"
    DispPanel.txtDispWat.Text = "0"
    
    nQuant = CSng(rDs(DispConfirm.txtConfDispQuant.Text))

'извеждаме съобщение за несъвместимост на капацитет на превозното средство и количеството за експедиция
    If CSng(rDs(DispConfirm.txtConfDrvCap.Text)) < nQuant Then
        Dim response As Integer
        MousePointer = vbDefault
        response = MsgBox(MsgOverCapDrv, vbQuestion Or vbYesNo, frmConfSend)
        If response = vbYes Then 'при потвърждение отиваме на изпращането на данните
            MousePointer = vbHourglass
            GoTo SendingData
        Else
            Unload DispConfirm 'при отказ прекъсваме
            DispPanel.TimerStartReq.Enabled = False
            DispPanel.TimerRes.Enabled = False
            Exit Function
        End If
    End If

'процес на зареждане на данните
SendingData:

'почистваме клетките на OPC-Server
    For i = 0 To DispPanel.ItemCountRec - 1
        frmOPC.RecInput(i).Text = 0
        SyncItemSrvHandlesRec(i + 1) = handyRec(i + 1)
        SyncItemValuesRec(i + 1) = Val(frmOPC.RecInput(i).Text)
    Next i

    DispPanel.ConGroupRec.SyncWrite DispPanel.ItemCountRec, SyncItemSrvHandlesRec, _
    SyncItemValuesRec, SyncItemSrvErrRec
    
'почистване на битовете за готовност от дозиранията
    frmOPC.Ready(0).Text = "False"
    frmOPC.Ready(1).Text = "False"
    frmOPC.Ready(2).Text = "False"
    frmOPC.Ready(3).Text = "False"
    
    For i = 0 To DispPanel.ItemCountReady - 1
        SyncItemSrvHandlesReady(i + 1) = handyReady(i + 1)
        SyncItemValuesReady(i + 1) = Val(frmOPC.Ready(i).Text)
    Next i
    
    DispPanel.ConGroupReady.SyncWrite DispPanel.ItemCountReady, SyncItemSrvHandlesReady, _
    SyncItemValuesReady, SyncItemSrvErrReady

'попълваме данни
    frmOPC.RecInput(0).Text = Recs
    
    rNameReady = ""
    
    If Len(RecNames) >= 8 Then
        lenRec = 8
        RecNames = Left(RecNames, 8)
    Else
        lenRec = Len(RecNames)
    End If
    
    RecNames1 = Left(RecNames, 4)
    For i = 0 To 3
        rName(i) = Hex(Asc(Right$(RecNames1, 4 - i)))
        rNameReady1 = rName(i) & rNameReady1
    Next i
    
    If lenRec > 4 Then
        RecNames2 = Mid(RecNames, 5, lenRec)
        For i = 4 To lenRec - 1
            rName(i) = Hex(Asc(Right$(RecNames2, lenRec - i)))
            rNameReady2 = rName(i) & rNameReady2
        Next i
        rNameReady2 = "&H" & rNameReady2
        rNameReady2 = CDbl(rNameReady2)
    Else
        RecNames2 = ""
    End If
        
    rNameReady1 = "&H" & rNameReady1
    rNameReady1 = CDbl(rNameReady1)
    
    frmOPC.RecInput(1).Text = rNameReady1
    frmOPC.RecInput(38).Text = rNameReady2
    frmOPC.RecInput(2).Text = rDsNew(DispConfirm.txtConfDispQuant.Text) 'зареждаме количество за изпълнение
    
    For i = 1 To 6
        IMready(i) = False
    Next i
    For i = 1 To ns1 'зареждаме инертните материали
        If SendRec.IMinit(i) > 0 And SendRec.IMkg(i) > 0 Then
            Select Case SendRec.IMinit(i) 'зареждаме инициализатори на течките от контролера
                Case 11
                    If Not IMready(1) Then
                        frmOPC.RecInput(3).Text = SendRec.AllIMkg(1)
                        frmOPC.RecInput(19).Text = i
                        IMready(1) = True
                    End If
                Case 12
                    If Not IMready(2) Then
                        frmOPC.RecInput(4).Text = SendRec.AllIMkg(2)
                        frmOPC.RecInput(20).Text = i
                        IMready(2) = True
                    End If
                Case 13
                    If Not IMready(3) Then
                        frmOPC.RecInput(5).Text = SendRec.AllIMkg(3)
                        frmOPC.RecInput(21).Text = i
                        IMready(3) = True
                    End If
                Case 14
                    If Not IMready(4) Then
                        frmOPC.RecInput(6).Text = SendRec.AllIMkg(4)
                        frmOPC.RecInput(22).Text = i
                        IMready(4) = True
                    End If
                Case 15
                    If Not IMready(5) Then
                        frmOPC.RecInput(7).Text = SendRec.AllIMkg(5)
                        frmOPC.RecInput(23).Text = i
                        IMready(5) = True
                    End If
                Case 15
                    If Not IMready(6) Then
                        frmOPC.RecInput(8).Text = SendRec.AllIMkg(6)
                        frmOPC.RecInput(24).Text = i
                        IMready(6) = True
                    End If
                Case Else
                    MousePointer = vbDefault
                    MsgBox MsgRecErrIM, vbOKOnly Or vbCritical, MsgErrBx
                    Unload DispConfirm
                    DispPanel.TimerStartReq.Enabled = False
                    DispPanel.TimerRes.Enabled = False
                    Exit Function
            End Select
        Else
        End If
    Next i
    
    For i = 1 To 4
        SCRready(i) = False
    Next i
    For i = 1 To ns3 'зареждаме цимент - аналогично на инерния материал
        If SendRec.SCRinit(i) > 0 And SendRec.SCRkg(i) > 0 Then
            Select Case SendRec.SCRinit(i)
                Case 31
                    If Not SCRready(1) Then
                        frmOPC.RecInput(9).Text = SendRec.AllSCRkg(1)
                        frmOPC.RecInput(25).Text = i
                        SCRready(1) = True
                    End If
                Case 32
                    If Not SCRready(2) Then
                        frmOPC.RecInput(10).Text = SendRec.AllSCRkg(2)
                        frmOPC.RecInput(26).Text = i
                        SCRready(2) = True
                    End If
                Case 33
                    If Not SCRready(3) Then
                        frmOPC.RecInput(11).Text = SendRec.AllSCRkg(3)
                        frmOPC.RecInput(27).Text = i
                        SCRready(3) = True
                    End If
                Case 34
                    If Not SCRready(4) Then
                        frmOPC.RecInput(12).Text = SendRec.AllSCRkg(4)
                        frmOPC.RecInput(28).Text = i
                        SCRready(4) = True
                    End If
                Case Else
                    MousePointer = vbDefault
                    MsgBox MsgRecErrCem, vbOKOnly Or vbCritical, MsgErrBx
                    Unload DispConfirm
                    DispPanel.TimerStartReq.Enabled = False
                    DispPanel.TimerRes.Enabled = False
                    Exit Function
            End Select
        Else
        End If
    Next i
    
    For i = 1 To 2
        WATready(i) = False
    Next i
    For i = 1 To ns2 'зареждаме вода - аналогично на инерния материал
        If SendRec.WATinit(i) > 0 And SendRec.WATkg(i) > 0 Then
            Select Case SendRec.WATinit(1)
                Case 21
                    If Not WATready(1) Then
                        frmOPC.RecInput(13).Text = SendRec.AllWATkg(1)
                        frmOPC.RecInput(29).Text = i
                        WATready(1) = True
                    End If
                Case 22
                    If Not WATready(2) Then
                        frmOPC.RecInput(14).Text = SendRec.AllWATkg(2)
                        frmOPC.RecInput(30).Text = i
                        WATready(2) = True
                    End If
                Case Else
                    MousePointer = vbDefault
                    MsgBox MsgRecErrWat, vbOKOnly Or vbCritical, MsgErrBx
                    Unload DispConfirm
                    DispPanel.TimerStartReq.Enabled = False
                    DispPanel.TimerRes.Enabled = False
                    Exit Function
            End Select
        Else
        End If
    Next i
    
    For i = 1 To 4
        CHEMready(i) = False
    Next i
    For i = 1 To ns4 'зареждаме химия - аналогично на инерния материал
        If SendRec.CHEMinit(i) > 0 And SendRec.CHEMkg(i) > 0 Then
            Select Case SendRec.CHEMinit(i)
                Case 41
                    If Not CHEMready(1) Then
                        frmOPC.RecInput(15).Text = rDsNew(str(SendRec.AllCHEMkg(1)))
                        frmOPC.RecInput(31).Text = i
                        CHEMready(1) = True
                    End If
                Case 42
                    If Not CHEMready(2) Then
                        frmOPC.RecInput(16).Text = rDsNew(str(SendRec.AllCHEMkg(2)))
                        frmOPC.RecInput(32).Text = i
                        CHEMready(2) = True
                    End If
                Case 43
                    If Not CHEMready(3) Then
                        frmOPC.RecInput(17).Text = rDsNew(str(SendRec.AllCHEMkg(3)))
                        frmOPC.RecInput(33).Text = i
                        CHEMready(3) = True
                    End If
                Case 44
                    If Not CHEMready(4) Then
                        frmOPC.RecInput(18).Text = rDsNew(str(SendRec.AllCHEMkg(4)))
                        frmOPC.RecInput(34).Text = i
                        CHEMready(4) = True
                    End If
                Case Else
                    MousePointer = vbDefault
                    MsgBox MsgRecErrChem, vbOKOnly Or vbCritical, MsgErrBx
                    Unload DispConfirm
                    DispPanel.TimerStartReq.Enabled = False
                    DispPanel.TimerRes.Enabled = False
                    Exit Function
            End Select
        Else
        End If
    Next i

'зареждаме времената в контролера
    frmOPC.RecInput(35).Text = SendRec.Tmix
    frmOPC.RecInput(36).Text = SendRec.Tpour

'вдигаме бит на контролера, че сме готови със зареждането на рецептата
    frmOPC.RecInput(37).Text = "34"

'попълваме клетките на OPC-Server
    If Recs > 0 Then
        For i = 0 To DispPanel.ItemCountRec - 1
            SyncItemSrvHandlesRec(i + 1) = handyRec(i + 1)
            SyncItemValuesRec(i + 1) = Val(frmOPC.RecInput(i).Text)
        Next i
    
        DispPanel.ConGroupRec.SyncWrite DispPanel.ItemCountRec, SyncItemSrvHandlesRec, _
        SyncItemValuesRec, SyncItemSrvErrRec
    'подготвяме променлива за готовност на функцията за следене на резулата
        HelpRes = 1
        HelpResAggr = 1
        HelpResCem = 1
        HelpResWat = 1
        HelpResHD = 1
        
        For i = 0 To 30
            For j = 0 To 16
                resMatrix(i, j) = 0
            Next j
        Next i
    End If

    
'почистваме списъка с готовите замеси
    DispPanel.lstMixReady.ListItems.Clear
    
    'нулираме тотал tempQQQ
    tempQQQ = 0

    MousePointer = vbDefault
    
'скриваме формата
    Unload DispConfirm
End Function

Public Function GetAggrVIPA()

    Dim SyncItemValuesReady(1 To 4) As Variant
    Dim SyncItemSrvHandlesReady(4) As Long
    Dim SyncItemSrvErrReady() As Long

    If frmOPC.Ready(0).Text = "True" Then
'        For e = 1 To 6
'            frmOPC.Result(e).Refresh
'        Next e
    
        For e = 1 To 6
            resMatrix(HelpResAggr, e) = ARound(CSng(rDs(frmOPC.Result(e).Text)), 2)
        Next e
        HelpResAggr = HelpResAggr + 1
        
        frmOPC.Ready(0).Text = "False"
        For i = 0 To DispPanel.ItemCountReady - 1
            SyncItemSrvHandlesReady(i + 1) = handyReady(i + 1)
            SyncItemValuesReady(i + 1) = Val(frmOPC.Ready(i).Text)
        Next i
    
        DispPanel.ConGroupReady.SyncWrite DispPanel.ItemCountReady, SyncItemSrvHandlesReady, _
        SyncItemValuesReady, SyncItemSrvErrReady
    
    Else
    End If
    
    okAggr = True
End Function

Public Function GetCemVIPA()
    
    Dim SyncItemValuesReady(1 To 4) As Variant
    Dim SyncItemSrvHandlesReady(4) As Long
    Dim SyncItemSrvErrReady() As Long
    
    If frmOPC.Ready(1).Text = "True" Then
'        For e = 7 To 10
'            frmOPC.Result(e).Refresh
'        Next e
        
        For e = 7 To 10
            resMatrix(HelpResCem, e) = ARound(CSng(rDs(frmOPC.Result(e).Text)), 2)
        Next e
        HelpResCem = HelpResCem + 1
        
        frmOPC.Ready(1).Text = "False"
        For i = 0 To DispPanel.ItemCountReady - 1
            SyncItemSrvHandlesReady(i + 1) = handyReady(i + 1)
            SyncItemValuesReady(i + 1) = Val(frmOPC.Ready(i).Text)
        Next i
    
        DispPanel.ConGroupReady.SyncWrite DispPanel.ItemCountReady, SyncItemSrvHandlesReady, _
        SyncItemValuesReady, SyncItemSrvErrReady

    Else
    End If
    
    okCem = True
End Function

Public Function GetWatVIPA()

    Dim SyncItemValuesReady(1 To 4) As Variant
    Dim SyncItemSrvHandlesReady(4) As Long
    Dim SyncItemSrvErrReady() As Long

    If frmOPC.Ready(2).Text = "True" Then
'        For e = 11 To 12
'            frmOPC.Result(e).Refresh
'        Next e
        
        For e = 11 To 12
            resMatrix(HelpResWat, e) = ARound(CSng(rDs(frmOPC.Result(e).Text)), 2)
        Next e
        HelpResWat = HelpResWat + 1
        
        frmOPC.Ready(2).Text = "False"
        For i = 0 To DispPanel.ItemCountReady - 1
            SyncItemSrvHandlesReady(i + 1) = handyReady(i + 1)
            SyncItemValuesReady(i + 1) = Val(frmOPC.Ready(i).Text)
        Next i
    
        DispPanel.ConGroupReady.SyncWrite DispPanel.ItemCountReady, SyncItemSrvHandlesReady, _
        SyncItemValuesReady, SyncItemSrvErrReady

    Else
    End If
    
    okWat = True
End Function

Public Function GetHDVIPA()

    Dim SyncItemValuesReady(1 To 4) As Variant
    Dim SyncItemSrvHandlesReady(4) As Long
    Dim SyncItemSrvErrReady() As Long

    If frmOPC.Ready(3).Text = "True" Then
'        For e = 13 To 16
'            frmOPC.Result(e).Refresh
'        Next e
        
        For e = 13 To 16
            resMatrix(HelpResHD, e) = ARound(CSng(rDs(frmOPC.Result(e).Text)), 2)
        Next e
        HelpResHD = HelpResHD + 1

        frmOPC.Ready(3).Text = "False"
        For i = 0 To DispPanel.ItemCountReady - 1
            SyncItemSrvHandlesReady(i + 1) = handyReady(i + 1)
            SyncItemValuesReady(i + 1) = Val(frmOPC.Ready(i).Text)
        Next i
    
        DispPanel.ConGroupReady.SyncWrite DispPanel.ItemCountReady, SyncItemSrvHandlesReady, _
        SyncItemValuesReady, SyncItemSrvErrReady

    Else
    End If
    
    okHD = True
End Function

Public Function GetResultVIPA()

    'функция за прихващане на резултат от последния направен замес при изпратена заявка
    Dim TempResult(16) As Single
    Dim ResNow As Result
    Dim ResYes As Boolean
    Dim lastRow As Long
'    ResYes = frmOPC.Result(17)
'
'    If Not ResYes Then Exit Function
    
    Set ResNow = New Result
    
    If (Val(frmOPC.Result(0)) = HelpRes And Recs > 0) Or (Val(frmOPC.Result(0)) = 0 And Recs > 0 And WasAuto = True) Then
                
        ResNow.Clear
                
        MousePointer = vbHourglass
        
        Dim hexdm(0 To 49) As String
        
        Dim ttt As String
        
'разчитане на замеса-----------------------------------------------------------------
        For B = 1 To 16
            TempResult(B) = resMatrix(HelpRes, B)
        Next B

'------------------------------------------------------------------------------------

        'подреждане на данни за заявка в класа
        ResNow.OrderCode = OrdData(0) 'заявка код
        ResNow.OrderDate = OrdData(1) 'заявка дата
        ResNow.OrderQuant = OrdData(2) 'заявено количество по заявката
        ResNow.ExpQuant = OrdData(3)  'заявено количество от диспечера за текущата експедиция
        
        ResNow.MixReadyTime = Format(Now, "DD.MM.YYYY - HH:MM:SS") 'час и дата на разтоварване на смесителя за всеки цикъл
        
        ResNow.DispName = OperName 'име и фамилия на действащия оператор

        'попълване на класа за запис
        ResNow.ClntTitle = ClientData(0) 'клиент фирма
        ResNow.ClntIdent = ClientData(1) 'клиент булстат
        ResNow.ClntWorksite = ClientData(2) 'клиент обект
        ResNow.WorksiteDist = ClientData(3) 'км до обект
        
        ResNow.DrvTitle = DriverData(0) 'водач име
        ResNow.DrvCarNum = DriverData(1) 'кола рег. номер
        ResNow.DrvCapacity = DriverData(2) 'кола капацитет
        
        ResNow.RecTitle = RecNames 'рецепта име
        ResNow.RecKind = RecTypes 'рецепта вид разтвор
        ResNow.RecClass = RecClasss 'рецепта клас якост
        ResNow.RecClassK = RecClassKs 'рецепта клас консистенция
        ResNow.RecClassV = RecClassVs 'рецепта клас въздействие
        ResNow.RecClassH = RecClassHs 'рецепта клас хлориди
        ResNow.RecClassP = RecClassPs 'рецепта водоплътност
        ResNow.RecEDM = RecEDMs 'рецепта едм

        For e = 1 To ns1
            ResNow.IMname(e) = IM(e) 'имена на материали в бункерите
            ResNow.IMstated(e) = kgIMs(e) 'кг по заявка
            ResNow.IMmeasured(e) = ARound(TempResult(e), 0) 'кг по измерено им
            ResNow.TotalStatedKG = ResNow.TotalStatedKG + ResNow.IMstated(e)   'кг по рецепта им
            ResNow.TotalMeasuredKG = ResNow.TotalMeasuredKG + ResNow.IMmeasured(e)   'кг по измерено им
            If ResNow.IMstated(e) > 0 And ResNow.IMmeasured(e) = 0 Then
            'проверка за нулеви данни ако има заявено количество им
                EmptyData = True 'вдигаме флаг за нулеви данни за им
            Else
            End If
            
        Next e

        For e = 1 To ns3
            ResNow.SCRname(e) = Scr(e) 'имена на материали цимент
            ResNow.SCRstated(e) = kgSCRs(e) 'кг по заявка
            ResNow.SCRmeasured(e) = ARound(TempResult(e + 6), 0) 'кг по измерено цимент
            ResNow.TotalStatedKG = ResNow.TotalStatedKG + ResNow.SCRstated(e)  'кг по рецепта цимент
            ResNow.TotalMeasuredKG = ResNow.TotalMeasuredKG + ResNow.SCRmeasured(e)  'кг по измерено цимент
            If ResNow.SCRstated(e) > 0 And ResNow.SCRmeasured(e) = 0 Then
            'проверка за нулеви данни ако има заявено количество цименти
                EmptyData = True 'вдигаме флаг за нулеви данни за цименти
            Else
            End If
        Next e

        For e = 1 To ns2
            ResNow.WATname(e) = Wat(e) 'име на материал вода
            ResNow.WATstated(e) = kgWATs(e) 'кг по заявка
            ResNow.WATmeasured(e) = ARound(TempResult(e + 10), 0) 'кг по измерено вода
            ResNow.TotalStatedKG = ResNow.TotalStatedKG + ResNow.WATstated(e)  'кг по рецепта вода
            ResNow.TotalMeasuredKG = ResNow.TotalMeasuredKG + ResNow.WATmeasured(e)  'кг по измерено вода
            If ResNow.WATstated(e) > 0 And ResNow.WATmeasured(e) = 0 Then
            'проверка за нулеви данни ако има заявено количество вода
                EmptyData = True 'вдигаме флаг за нулеви данни за вода
            Else
            End If
        Next e

        For e = 1 To ns4
            ResNow.CHEMname(e) = Chem(e) 'име на материали химия
            ResNow.CHEMstated(e) = kgCHEMs(e) 'кг по заявка
            ResNow.CHEMmeasured(e) = ARound(TempResult(e + 12), 2)  'кг по измерено химия
            ResNow.TotalStatedKG = ResNow.TotalStatedKG + CSng(ResNow.CHEMstated(e)) 'кг по рецепта химия
            ResNow.TotalMeasuredKG = ResNow.TotalMeasuredKG + CSng(ResNow.CHEMmeasured(e))  'кг по измерено химия
            If ResNow.CHEMstated(e) > 0 And ResNow.CHEMmeasured(e) = 0 Then
            'проверка за нулеви данни ако има заявено количество хд
                EmptyData = True 'вдигаме флаг за нулеви данни за хд
            Else
            End If
        Next e
        
        ResNow.TotalStatedKG = ARound(ResNow.TotalStatedKG, 2)
        ResNow.TotalMeasuredKG = ARound(ResNow.TotalMeasuredKG, 2)

        If ResNow.TotalStatedKG > 0 Then
            ResNow.TotalQuant = ARound(CSng(rDs(nCoefs)) * (ResNow.TotalMeasuredKG / ResNow.TotalStatedKG), 3) 'обем на всеки замес
        End If
'EmptyData = True
        '-----------------------Start postgreSQL-----------------------------------
        Dim cnUnique     As ADODB.Connection 'уникално име на връзката и записа за да не я затворим случайно

        Dim rsUnique     As Recordset

        Dim comInsUnique As String

        Dim commUnique   As String
        
        Set cnUnique = New ADODB.Connection
        cnUnique.ConnectionTimeout = 10
        cnUnique.Open ConStr
        MousePointer = vbHourglass
        
        'откриване на последния запис
        Set rsUnique = cnUnique.Execute("SELECT mix_num, exp_num FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1")
        
        If Not rsUnique.EOF And Not rsUnique.BOF Then
            ResNow.MixNum = Val(rsUnique!mix_num) + 1 'номер на последен замес + 1

            If HelpRes = 1 Then 'само при първи замес от експедиция
                ResNow.ExpNum = Val(rsUnique!exp_num) + 1 'номер на последна експедция +1
                tRealQuant = 0
                tTotalKGs = 0
            Else
                ResNow.ExpNum = Val(rsUnique!exp_num) 'номер на експедиция
            End If

        Else 'ако таблицата със замесите е празна
            ResNow.MixNum = 1
            ResNow.ExpNum = 1
            tRealQuant = 0
            tTotalKGs = 0
        End If

        'откриване на последния замес по текуща заявка ако има такъв
        Set rsUnique = cnUnique.Execute("SELECT mix_ord_num, exp_ord_num FROM mix_result_bc" & MachineNumber & " WHERE ord_num =" & ResNow.OrderCode & "ORDER BY mix_num DESC LIMIT 1")
        
        If Not rsUnique.EOF And Not rsUnique.BOF Then
            ResNow.MixNumFromOrder = Val(rsUnique!mix_ord_num) + 1 'номер на последен замес по заявката + 1

            If HelpRes = 1 Then 'само при първи замес от експедиция по заявката
                ResNow.ExpNumFromOrder = Val(rsUnique!exp_ord_num) + 1 'номер на последна експедиция по заявката +1
                DispPanel.lstMixReady.ListItems.Clear
            Else
                ResNow.ExpNumFromOrder = Val(rsUnique!exp_ord_num) 'номер на експедиция по заявката
            End If

        Else 'ако няма замеси по заявката
            ResNow.MixNumFromOrder = 1
            ResNow.ExpNumFromOrder = 1
        End If
    
        comInsUnique = "INSERT INTO mix_result_bc" & MachineNumber & " VALUES(" & ResNow.MixNum & "," & ResNow.ExpNum & ",'" & ReqTime & "','" & ResNow.MixReadyTime & "','" & Format(Now, "DD-MM-YYYY") & "','" & ResNow.DispName _
           & "'," & ResNow.OrderCode & ",'" & ResNow.OrderDate & "','" & ResNow.OrderQuant & "','" & ResNow.ExpQuant & "','" & ResNow.ExpNumFromOrder & "','" & ResNow.MixNumFromOrder _
           & "','" & ResNow.ClntTitle & "','" & ResNow.ClntIdent & "','" & ResNow.ClntWorksite & "','" & ResNow.WorksiteDist & "','" & ResNow.DrvTitle & "','" & ResNow.DrvCarNum & "','" & ResNow.DrvCapacity _
           & "','" & ResNow.RecTitle & "','" & ResNow.RecKind & "','" & ResNow.RecClass & "','" & ResNow.RecClassK & "','" & ResNow.RecClassV _
           & "','" & ResNow.RecClassH & "','" & ResNow.RecClassP & "','" & ResNow.RecEDM _
           & "','" & ResNow.IMname(1) & "','" & ResNow.IMstated(1) & "','" & ResNow.IMmeasured(1) _
           & "','" & ResNow.IMname(2) & "','" & ResNow.IMstated(2) & "','" & ResNow.IMmeasured(2) _
           & "','" & ResNow.IMname(3) & "','" & ResNow.IMstated(3) & "','" & ResNow.IMmeasured(3) _
           & "','" & ResNow.IMname(4) & "','" & ResNow.IMstated(4) & "','" & ResNow.IMmeasured(4) _
           & "','" & ResNow.IMname(5) & "','" & ResNow.IMstated(5) & "','" & ResNow.IMmeasured(5) _
           & "','" & ResNow.IMname(6) & "','" & ResNow.IMstated(6) & "','" & ResNow.IMmeasured(6) _
           & "','" & ResNow.SCRname(1) & "','" & ResNow.SCRstated(1) & "','" & ResNow.SCRmeasured(1) _
           & "','" & ResNow.SCRname(2) & "','" & ResNow.SCRstated(2) & "','" & ResNow.SCRmeasured(2) _
           & "','" & ResNow.SCRname(3) & "','" & ResNow.SCRstated(3) & "','" & ResNow.SCRmeasured(3) _
           & "','" & ResNow.SCRname(4) & "','" & ResNow.SCRstated(4) & "','" & ResNow.SCRmeasured(4) _
           & "','" & ResNow.WATname(1) & "','" & ResNow.WATstated(1) & "','" & ResNow.WATmeasured(1) _
           & "','" & ResNow.WATname(2) & "','" & ResNow.WATstated(2) & "','" & ResNow.WATmeasured(2) _
           & "','" & ResNow.CHEMname(1) & "','" & ResNow.CHEMstated(1) & "','" & ResNow.CHEMmeasured(1) _
           & "','" & ResNow.CHEMname(2) & "','" & ResNow.CHEMstated(2) & "','" & ResNow.CHEMmeasured(2) _
           & "','" & ResNow.CHEMname(3) & "','" & ResNow.CHEMstated(3) & "','" & ResNow.CHEMmeasured(3) _
           & "','" & ResNow.CHEMname(4) & "','" & ResNow.CHEMstated(4) & "','" & ResNow.CHEMmeasured(4) _
           & "','" & ResNow.CHEMname(5) & "','" & ResNow.CHEMstated(5) & "','" & ResNow.CHEMmeasured(5) _
           & "','" & ResNow.CHEMname(6) & "','" & ResNow.CHEMstated(6) & "','" & ResNow.CHEMmeasured(6) _
           & "','" & ResNow.TotalStatedKG & "','" & ResNow.TotalMeasuredKG & "','" & ResNow.TotalQuant _
           & "','false')"
        Set rsUnique = cnUnique.Execute(comInsUnique)

        If EmptyData = False Then
            'запис на изпълненото по експедицията количество в таблицата със заявките
            Set rsUnique = cnUnique.Execute("SELECT order_qmade FROM orders WHERE order_num =  " & ResNow.OrderCode & ";")
            Set rsUnique = cnUnique.Execute("UPDATE orders SET order_qmade = '" & CSng(rDs(rsUnique!order_qmade)) + CSng(rDs(ResNow.TotalQuant)) & "' WHERE order_num =" & ResNow.OrderCode & ";")
        Else
            tempQQQ = tempQQQ + CSng(rDs(ResNow.TotalQuant))
        End If

        'запис на последните временни данни в tempmix_bc1 таблицата
        tRealQuant = tRealQuant + ResNow.TotalQuant
        tTotalKGs = tTotalKGs + ResNow.TotalMeasuredKG

        Set rsUnique = cnUnique.Execute("SELECT * FROM tempmix_bc" & MachineNumber & ";")
        
        If Not rsUnique.BOF And Not rsUnique.EOF Then TempVal = Val(rsUnique!mix_id)
        
        If Val(TempVal) = 0 Then
            Set rsUnique = cnUnique.Execute("INSERT INTO tempmix_bc" & MachineNumber & " VALUES(" & ResNow.MixNum & "," & ResNow.ExpNum & ",'" & ResNow.ExpQuant & "','" & tRealQuant & "','" & tTotalKGs & "')")
        Else
            Set rsUnique = cnUnique.Execute("UPDATE tempmix_bc" & MachineNumber & " SET mix_id = " & ResNow.MixNum & ",exp_id = " & ResNow.ExpNum & ",ordered_q = '" & ResNow.ExpQuant & "',real_q = '" & tRealQuant & "',total_kg_temp = '" & tTotalKGs & "' WHERE mix_id =" & TempVal & ";")
        End If
        
        'корекция на наличните материали след всеки прихванат замес------------------------------------------------
        'корекция на наличност на им
        For e = 1 To ns1

            If ResNow.IMname(e) <> "0" And ResNow.IMname(e) <> uniEmpty And ResNow.IMname(e) <> "" Then
                Set rsUnique = cnUnique.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & ResNow.IMname(e) & "';")
                Set rsUnique = cnUnique.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsUnique!m_sold)) + ResNow.IMmeasured(e) / 1000, 3) & "'WHERE m_name = '" & ResNow.IMname(e) & "';")
                
                Set rsUnique = cnUnique.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & ResNow.IMname(e) & "' AND stamp_date = '" & DayToday & "';")
                If Not rsUnique.EOF And Not rsUnique.BOF Then
                    Set rsUnique = cnUnique.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsUnique!mat_sold)) + ResNow.IMmeasured(e) / 1000, 3) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & ResNow.IMname(e) & "' AND stamp_date = '" & DayToday & "';")
                Else
                    'откриваме последния запис
                    Set rsUnique = cnUnique.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                    If Not rsUnique.EOF And Not rsUnique.BOF Then
                        lastRow = Val(rsUnique!row_num) + 1
                    Else
                        lastRow = 1
                    End If
                    Set rsUnique = cnUnique.Execute("INSERT INTO daily_expenses VALUES(" & lastRow & ",'" & ResNow.IMname(e) & "','" & ARound(ResNow.IMmeasured(e) / 1000, 3) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
                End If
            End If

        Next e
    
        'корекция на наличност на цимент
        For e = 1 To ns3

            If ResNow.SCRname(e) <> "0" And ResNow.SCRname(e) <> uniEmpty And ResNow.SCRname(e) <> "" Then
                Set rsUnique = cnUnique.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & ResNow.SCRname(e) & "';")
                Set rsUnique = cnUnique.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsUnique!m_sold)) + ResNow.SCRmeasured(e) / 1000, 3) & "'WHERE m_name = '" & ResNow.SCRname(e) & "';")
            
                Set rsUnique = cnUnique.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & ResNow.SCRname(e) & "' AND stamp_date = '" & DayToday & "';")
                If Not rsUnique.EOF And Not rsUnique.BOF Then
                    Set rsUnique = cnUnique.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsUnique!mat_sold)) + ResNow.SCRmeasured(e) / 1000, 3) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & ResNow.SCRname(e) & "' AND stamp_date = '" & DayToday & "';")
                Else
                    'откриваме последния запис
                    Set rsUnique = cnUnique.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                    If Not rsUnique.EOF And Not rsUnique.BOF Then
                        lastRow = Val(rsUnique!row_num) + 1
                    Else
                        lastRow = 1
                    End If
                    Set rsUnique = cnUnique.Execute("INSERT INTO daily_expenses VALUES(" & lastRow & ",'" & ResNow.SCRname(e) & "','" & ARound(ResNow.SCRmeasured(e) / 1000, 3) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
                End If
            End If

        Next e
    
        'корекция на разход на вода
        For e = 1 To ns2

            If ResNow.WATname(e) <> "0" And ResNow.WATname(e) <> uniEmpty And ResNow.WATname(e) <> "" Then
                Set rsUnique = cnUnique.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & ResNow.WATname(e) & "';")
                Set rsUnique = cnUnique.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsUnique!m_sold)) + ResNow.WATmeasured(e) / 1000, 3) & "'WHERE m_name = '" & ResNow.WATname(e) & "';")
            
                Set rsUnique = cnUnique.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & ResNow.WATname(e) & "' AND stamp_date = '" & DayToday & "';")
                If Not rsUnique.EOF And Not rsUnique.BOF Then
                    Set rsUnique = cnUnique.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsUnique!mat_sold)) + ResNow.WATmeasured(e) / 1000, 3) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & ResNow.WATname(e) & "' AND stamp_date = '" & DayToday & "';")
                Else
                    'откриваме последния запис
                    Set rsUnique = cnUnique.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                    If Not rsUnique.EOF And Not rsUnique.BOF Then
                        lastRow = Val(rsUnique!row_num) + 1
                    Else
                        lastRow = 1
                    End If
                    Set rsUnique = cnUnique.Execute("INSERT INTO daily_expenses VALUES(" & lastRow & ",'" & ResNow.WATname(e) & "','" & ARound(ResNow.WATmeasured(e) / 1000, 3) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
                End If
            End If

        Next e
    
        'корекция на наличност на хд
        For e = 1 To ns4

            If ResNow.CHEMname(e) <> "0" And ResNow.CHEMname(e) <> uniEmpty And ResNow.CHEMname(e) <> "" Then
                Set rsUnique = cnUnique.Execute("SELECT * FROM materials_bc" & MachineNumber & " WHERE m_name = '" & ResNow.CHEMname(e) & "';")
                Set rsUnique = cnUnique.Execute("UPDATE materials_bc" & MachineNumber & " SET m_sold = '" & ARound(CSng(rDs(rsUnique!m_sold)) + ResNow.CHEMmeasured(e) / 1000, 5) & "'WHERE m_name = '" & ResNow.CHEMname(e) & "';")
            
                Set rsUnique = cnUnique.Execute("SELECT * FROM daily_expenses WHERE mat_name = '" & ResNow.CHEMname(e) & "' AND stamp_date = '" & DayToday & "';")
                If Not rsUnique.EOF And Not rsUnique.BOF Then
                    Set rsUnique = cnUnique.Execute("UPDATE daily_expenses SET mat_sold = '" & ARound(CSng(rDs(rsUnique!mat_sold)) + ResNow.CHEMmeasured(e) / 1000, 5) & "', date_sold = '" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "' WHERE mat_name = '" & ResNow.CHEMname(e) & "' AND stamp_date = '" & DayToday & "';")
                Else
                    'откриваме последния запис
                    Set rsUnique = cnUnique.Execute("SELECT row_num FROM daily_expenses ORDER BY row_num DESC LIMIT 1")
                    If Not rsUnique.EOF And Not rsUnique.BOF Then
                        lastRow = Val(rsUnique!row_num) + 1
                    Else
                        lastRow = 1
                    End If
                    Set rsUnique = cnUnique.Execute("INSERT INTO daily_expenses VALUES(" & lastRow & ",'" & ResNow.CHEMname(e) & "','" & ARound(ResNow.CHEMmeasured(e) / 1000, 5) & "','" & Format(Now, "DD.MM.YYYY - HH:MM:SS") & "','" & DayToday & "')")
                End If
            End If

        Next e

        '-----------------------------------

        'презареждане на таблиците с текущите заявки
        DispPanel.lstOrdWait.ListItems.Clear
        commUnique = "SELECT * FROM orders WHERE stamp_date >= '" & DayToday & "';"
    
        Set rsUnique = cnUnique.Execute(commUnique) 'маркиране на заявките отговарящи на дата днес
    
        If Not rsUnique.EOF And Not rsUnique.BOF Then rsUnique.MoveFirst
    
        Do While Not rsUnique.EOF
            Set itmX = DispPanel.lstOrdWait.ListItems.Add(1, , Format(rsUnique!order_num, "0000000")) 'зареждане в ListView
            itmX.SubItems(1) = rsUnique!order_date
            itmX.SubItems(2) = rsUnique!order_date_que
            itmX.SubItems(3) = rDs(rsUnique!order_q)
            itmX.SubItems(4) = rDs(rsUnique!order_qmade)
            itmX.SubItems(5) = Format(rsUnique!order_rec, "00000")
            itmX.SubItems(6) = rsUnique!order_rec_name
            itmX.SubItems(7) = rsUnique!order_rec_class
            itmX.SubItems(8) = Format(rsUnique!order_clnt, "000000")
            itmX.SubItems(9) = rsUnique!order_clnt_name
            itmX.SubItems(10) = rsUnique!order_clnt_obj
            rsUnique.MoveNext
        Loop
        
        rsUnique.Close
        Set rsUnique = Nothing
        cnUnique.Close
        MousePointer = vbDefault
        Set cnUnique = Nothing
        '--------------------------End PostgreSQL-----------------------------------

        MousePointer = vbHourglass
    
        'визуализация на последния замес в таблицата
        CountMix = CountMix + 1
        Set itmX = DispPanel.lstMixReady.ListItems.Add(1, , Format(CountMix, "00"))
        itmX.SubItems(1) = Format(ResNow.OrderCode, "0000000")
        itmX.SubItems(2) = ResNow.RecTitle
        itmX.SubItems(3) = ResNow.RecClass

        For e = 1 To ns1
            itmX.SubItems(2 * e + 2) = ResNow.IMstated(e)
            itmX.SubItems(2 * e + 3) = ResNow.IMmeasured(e)
        Next e

        For e = 1 To ns3
            itmX.SubItems(2 * (e + ns1) + 2) = ResNow.SCRstated(e)
            itmX.SubItems(2 * (e + ns1) + 3) = ResNow.SCRmeasured(e)
        Next e

        For e = 1 To ns2
            itmX.SubItems(2 * (ns1 + ns3 + e) + 2) = ResNow.WATstated(e)
            itmX.SubItems(2 * (ns1 + ns3 + e) + 3) = ResNow.WATmeasured(e)
        Next e

        For e = 1 To ns4
            itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 2) = ResNow.CHEMstated(e)
            itmX.SubItems(2 * (e + ns1 + ns3 + ns2) + 3) = ResNow.CHEMmeasured(e)
        Next e

        itmX.SubItems(2 * (ns1 + ns3 + ns2 + ns4 + 1) + 2) = ResNow.TotalStatedKG
        itmX.SubItems(2 * (ns1 + ns3 + ns2 + ns4 + 1) + 3) = ResNow.TotalMeasuredKG
        itmX.SubItems(2 * (ns1 + ns3 + ns2 + ns4 + 1) + 4) = ResNow.TotalQuant

        If DispPanel.lstMixReady.ListItems.count > 0 Then
            AutoColW DispPanel.lstMixReady
        Else
        End If
        
        'обновяване на статус бара
        DispPanel.StatusBar.Panels(3) = "Замеси: " & ResNow.MixNum
        DispPanel.StatusBar.Panels(4) = "Експедиции: " & ResNow.ExpNum
        DispPanel.StatusBar.Panels(5) = "Заявена експедиция: " & ResNow.ExpQuant & " m3"
        DispPanel.StatusBar.Panels(6) = "Обем експедиция: " & tRealQuant & " m3"
        DispPanel.StatusBar.Panels(7) = "Тегло експедиция: " & tTotalKGs & " kg"
        DispPanel.StatusBar.Refresh

        'изключваме таймера за контрол на тази функция, когато са изпълнени заявените цикли
        If cyc = HelpRes Then
            DispPanel.TimerStartReq.Enabled = False
            DispPanel.TimerRes.Enabled = False
            ExpeditionStarted = False
            WasAuto = False
            okAggr = False
            okCem = False
            okWat = False
            okHD = False
            
            'проверка дали има вдигнат флаг за пропуснати данни
            If EmptyData = True Then
                Dim response As Integer
                response = MsgBox(MsgCorData, vbYesNo Or vbQuestion, ErrNoData)
                If response = vbYes Then
                    frmCorData.Show
                    EmptyData = False
                    Exit Function
                Else
                    EmptyData = False
                    'запис на изпълненото по експедицията количество в таблицата със заявките
'-----------------------Start postgreSQL-----------------------------------
                    Set cnUnique = New ADODB.Connection
                    cnUnique.ConnectionTimeout = 10
                    cnUnique.Open ConStr

                    Set rsUnique = cnUnique.Execute("SELECT order_qmade FROM orders WHERE order_num =  " & ResNow.OrderCode & ";")
                    Set rsUnique = cnUnique.Execute("UPDATE orders SET order_qmade = '" & ARound(CSng(rDs(rsUnique!order_qmade)) + CSng(rDs(tempQQQ)), 3) & "' WHERE order_num =" & ResNow.OrderCode & ";")
                        
                    'презареждане на таблиците с текущите заявки
                    DispPanel.lstOrdWait.ListItems.Clear
                    commUnique = "SELECT * FROM orders WHERE stamp_date >= '" & DayToday & "';"
    
                    Set rsUnique = cnUnique.Execute(commUnique) 'маркиране на заявките отговарящи на дата днес
    
                    If Not rsUnique.EOF And Not rsUnique.BOF Then rsUnique.MoveFirst
    
                    Do While Not rsUnique.EOF
                        Set itmX = DispPanel.lstOrdWait.ListItems.Add(1, , Format(rsUnique!order_num, "0000000")) 'зареждане в ListView
                        itmX.SubItems(1) = rsUnique!order_date
                        itmX.SubItems(2) = rsUnique!order_date_que
                        itmX.SubItems(3) = rDs(rsUnique!order_q)
                        itmX.SubItems(4) = rDs(rsUnique!order_qmade)
                        itmX.SubItems(5) = Format(rsUnique!order_rec, "00000")
                        itmX.SubItems(6) = rsUnique!order_rec_name
                        itmX.SubItems(7) = rsUnique!order_rec_class
                        itmX.SubItems(8) = Format(rsUnique!order_clnt, "000000")
                        itmX.SubItems(9) = rsUnique!order_clnt_name
                        itmX.SubItems(10) = rsUnique!order_clnt_obj
                        rsUnique.MoveNext
                    Loop
    
                    rsUnique.Close
                    Set rsUnique = Nothing
                    cnUnique.Close
                    MousePointer = vbDefault
                    Set cnUnique = Nothing
'--------------------------End PostgreSQL-----------------------------------

                    Set ResNow = Nothing
                End If
            End If
            
            'проверка в регистъра дали има настройки за принтиране на формите 1,2,3 - ако има излиза въпрос за печат
            Dim PrevSet1   As Boolean

            Dim strSubKey1 As String

            strSubKey1 = Trim(PlaceProgSet1)
            PrevSet1 = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey1)

            If PrevSet1 = True Then
                rPrint1 = GetSetting(PlaceProgSettings, PlaceForm1, "Print1", ErrRes)
            Else
                rPrint1 = 0
            End If
            
            Dim PrevSet2   As Boolean

            Dim strSubKey2 As String

            strSubKey2 = Trim(PlaceProgSet2)
            PrevSet2 = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey2)

            If PrevSet2 = True Then
                rPrint2 = GetSetting(PlaceProgSettings, PlaceForm2, "Print2", ErrRes)
            Else
                rPrint2 = 0
            End If
            
            Dim PrevSet3   As Boolean

            Dim strSubKey3 As String

            strSubKey3 = Trim(PlaceProgSet3)
            PrevSet3 = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey3)

            If PrevSet3 = True Then
                rPrint3 = GetSetting(PlaceProgSettings, PlaceForm3, "Print3", ErrRes)
            Else
                rPrint3 = 0
            End If
            
            PrintAnyForm = False

            If rPrint1 = 1 Or rPrint2 = 1 Or rPrint3 = 1 Then

                DispPanel.FormT.Enabled = True 'извикваме функция за попълване на форма ако е необходимо

                Exit Function

            End If
        End If

        MousePointer = vbDefault
        HelpRes = HelpRes + 1
    End If

    MousePointer = vbDefault
End Function

Public Function FillForm1()

    DispPanel.FormT.Enabled = False
    
    'функция за попълване на форма 1
    If PrintRightBut = True Then GoTo SkipAuto
    
    'проверка от регистъра за принтиране на форма 1 след експедицията
    If rPrint1 = 0 And PrintAnyForm = True Then
        Call FillForm2

        Exit Function

    End If
    
SkipAuto:

    MousePointer = vbHourglass
    
    frmPrintMix.Show
    
    frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Min
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    Dim ResForm1 As Result

    Set ResForm1 = New Result
        
    Dim ExpBufferff As Long
    
    'почистваме полетата на формата
    prntForm1.txtExpNote.Text = ""
    prntForm1.txtDate.Text = ""
    prntForm1.txtOrd.Text = ""
    prntForm1.txtClnt.Text = ""
    prntForm1.txtObj.Text = ""
    prntForm1.txtDrv.Text = ""
    prntForm1.txtDrvNo.Text = ""
    prntForm1.txtDist.Text = ""
    prntForm1.txtRecType.Text = ""
    prntForm1.txtVol.Text = ""
    prntForm1.txtW.Text = ""
    prntForm1.txtOrdVol.Text = ""
    prntForm1.txtClass.Text = ""
    prntForm1.txtClassK.Text = ""
    prntForm1.txtClassV.Text = ""
    prntForm1.txtClassH.Text = ""
    prntForm1.txtClassP.Text = ""
    prntForm1.txtCem1.Text = ""
    prntForm1.txtCem2.Text = ""
    prntForm1.txtCem3.Text = ""
    prntForm1.txtChem1.Text = ""
    prntForm1.txtChem2.Text = ""
    prntForm1.txtChem3.Text = ""
    prntForm1.txtEDM.Text = ""
    prntForm1.txtMixTime.Text = ""
    prntForm1.txtExpTime.Text = ""
    prntForm1.txtOper.Text = ""
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10

    '-----------------------Start postgreSQL-----------------------------------
    Dim cnUnique As ADODB.Connection

    Dim rsUnique As Recordset
        
    Set cnUnique = New ADODB.Connection
    cnUnique.ConnectionTimeout = 10
    cnUnique.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме последния направен замес като сортираме в обратен ред и вземаме най-големия номер
    Set rsUnique = cnUnique.Execute("SELECT * FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1")
    
    If Not rsUnique.EOF And Not rsUnique.BOF Then
        ExpBufferff = rsUnique!exp_num 'буфер за номера на последната експедиция
        'попълване на полета на бележката от последния замес
        prntForm1.txtExpNote.Text = "M" & MachineNumber & "-" & Format(rsUnique!exp_num, "000000000") 'номер на бележка е номер на последна експедиция
        prntForm1.txtDate.Text = Left$(rsUnique!time_mix_ready, 10) 'дата на бележка от последния замес
        prntForm1.txtOrd.Text = Format(rsUnique!ord_num, "0000000") & "/" & rsUnique!ord_date 'номер/дата на заявката
        prntForm1.txtClnt.Text = rsUnique!name_clnt 'име на клиента
        prntForm1.txtObj.Text = rsUnique!obj_clnt 'име на обекта
        prntForm1.txtDist.Text = rsUnique!km_clnt 'разстояние до обекта
        prntForm1.txtDrv.Text = rsUnique!name_drv 'име на водача
        prntForm1.txtDrvNo.Text = rsUnique!reg_drv 'номер на превозно средство
        prntForm1.txtRecType.Text = rsUnique!type_rec 'рецепта тип
'        prntForm1.txtOrdVol.Text = rDs(rsUnique!ord_q) 'общо количество по заявката
        prntForm1.txtClass.Text = rsUnique!class_rec 'клас по якост
        prntForm1.txtClassK.Text = rsUnique!classk_rec 'клас по консистенция
        prntForm1.txtClassV.Text = rsUnique!classv_rec 'клас по въздействие
        prntForm1.txtClassH.Text = rsUnique!classh_rec 'клас по хлориди
        prntForm1.txtClassP.Text = rsUnique!classp_rec 'водоплътност
        prntForm1.txtEDM.Text = rsUnique!edm_rec 'едм
        prntForm1.txtMixTime.Text = Mid$(rsUnique!time_exp_start, 14, 5) 'час на стартиране на експедицията
        prntForm1.txtExpTime.Text = Mid$(rsUnique!time_mix_ready, 14, 5) 'час на последния замес по експедицията
        prntForm1.txtOper.Text = rsUnique!name_op 'име и фамилия на диспечера
        
        ResForm1.ExpQuant = ARound(CSng(rDs(rsUnique!exp_q)), 2) 'заявен обем за експедицията
        
        'попълваме заявените цименти
        ResForm1.SCRname(1) = rsUnique!cem1_name
        ResForm1.SCRstated(1) = Val(rsUnique!cem1z)
        ResForm1.SCRname(2) = rsUnique!cem2_name
        ResForm1.SCRstated(2) = Val(rsUnique!cem2z)
        ResForm1.SCRname(3) = rsUnique!cem3_name
        ResForm1.SCRstated(3) = Val(rsUnique!cem3z)
        ResForm1.SCRname(4) = rsUnique!cem4_name
        ResForm1.SCRstated(4) = Val(rsUnique!cem4z)
        
        'попълваме заявените химически добавки
        ResForm1.CHEMname(1) = rsUnique!chem1_name
        ResForm1.CHEMstated(1) = rDs(rsUnique!chem1z)
        ResForm1.CHEMname(2) = rsUnique!chem2_name
        ResForm1.CHEMstated(2) = rDs(rsUnique!chem2z)
        ResForm1.CHEMname(3) = rsUnique!chem3_name
        ResForm1.CHEMstated(3) = rDs(rsUnique!chem3z)
        ResForm1.CHEMname(4) = rsUnique!chem4_name
        ResForm1.CHEMstated(4) = rDs(rsUnique!chem4z)
        ResForm1.CHEMname(5) = rsUnique!chem5_name
        ResForm1.CHEMstated(5) = rDs(rsUnique!chem5z)
        ResForm1.CHEMname(6) = rsUnique!chem6_name
        ResForm1.CHEMstated(6) = rDs(rsUnique!chem6z)
    End If
    
    'зареждане на първите 3 срещнати използвани в рецептата материали от силозите
    Dim ret As Integer

    For r = 1 To ns3

        If ResForm1.SCRstated(r) > 0 Then
            prntForm1.txtCem1.Text = ResForm1.SCRname(r)
            ret = r + 1

            Exit For

        Else
            ret = r + 1
        End If

    Next r

    If ret <= ns3 Then

        For r = ret To ns3

            If ResForm1.SCRstated(r) > 0 Then
                prntForm1.txtCem2.Text = ResForm1.SCRname(r)
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
                prntForm1.txtCem3.Text = ResForm1.SCRname(r)
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
            prntForm1.txtChem1.Text = ResForm1.CHEMname(r)
            ret = r + 1

            Exit For

        Else
            ret = r + 1
        End If

    Next r

    If ret <= ns4 Then

        For r = ret To ns4

            If ResForm1.CHEMstated(r) > 0 Then
                prntForm1.txtChem2.Text = ResForm1.CHEMname(r)
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
                prntForm1.txtChem3.Text = ResForm1.CHEMname(r)
                ret = r + 1

                Exit For

            Else
                ret = r + 1
            End If

        Next r

    End If
        
    'прочитане на последната експедиция за да пресметнем теглата и обемите за форма 1
    Set rsUnique = cnUnique.Execute("SELECT total_rec_kg, total_real_kg, total_vol FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & ExpBufferff & " ORDER BY mix_num ASC")
    
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    
    ResForm1.TotalStatedKG = 0 'нулираме променливата за кг по рецепта
    ResForm1.TotalMeasuredKG = 0 'нулираме променливата за кг по изпълнение
    ResForm1.TotalQuant = 0 'нулираме променливата за обем по изпълнение
    
    Do While Not rsUnique.EOF
        ResForm1.TotalStatedKG = ResForm1.TotalStatedKG + CSng(rDs(rsUnique!total_rec_kg))
        ResForm1.TotalMeasuredKG = ResForm1.TotalMeasuredKG + CSng(rDs(rsUnique!total_real_kg))
        ResForm1.TotalQuant = ResForm1.TotalQuant + CSng(rDs(rsUnique!total_vol))
        rsUnique.MoveNext
    Loop
    
    Set rsUnique = cnUnique.Execute("SELECT total_vol FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm1.txtOrd.Text) & " ORDER BY mix_num ASC")
    Dim totsmth As Single
    totsmth = 0
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    Do While Not rsUnique.EOF
        totsmth = ARound(totsmth, 2) + ARound(CSng(rDs(rsUnique!total_vol)), 2)
        rsUnique.MoveNext
    Loop
    
    Set rsUnique = cnUnique.Execute("SELECT DISTINCT ON (exp_num) exp_q FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm1.txtOrd.Text) & " ORDER BY exp_num ASC")
    Dim totexpsmth As Single
    totexpsmth = 0
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    Do While Not rsUnique.EOF
        totexpsmth = ARound(totexpsmth, 2) + ARound(CSng(rDs(rsUnique!exp_q)), 2)
        rsUnique.MoveNext
    Loop
    
    rsUnique.Close
    Set rsUnique = Nothing
    cnUnique.Close
    MousePointer = vbDefault
    Set cnUnique = Nothing
    '--------------------------End PostgreSQL-----------------------------------
        
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet   As Boolean

    Dim strSubKey As String
    
    strSubKey = Trim(PlaceProgSet1)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)
    
    If PrevSet = True Then
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm1, "RealVol", ErrRes)
    Else
        rRealVol = 1
    End If
        
    If rRealVol = 1 Then
        prntForm1.txtVol.Text = ARound(ResForm1.TotalQuant, 2) 'реален обем на експедицията
        prntForm1.txtOrdVol.Text = totsmth
    Else
        prntForm1.txtVol.Text = ResForm1.ExpQuant 'заявен обем на експедицията
        prntForm1.txtOrdVol.Text = totexpsmth
    End If
        
    prntForm1.txtW.Text = ARound(ResForm1.TotalMeasuredKG, 0)
    
    Set ResForm1 = Nothing
    
    MousePointer = vbDefault
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    'показване на редактора за бележки ако е включена опцията
    If ShowEditor = True Then
        prntForm1.Show
    Else
        Call PrintThisForm1(prntForm1)
    End If
End Function

Public Sub PrintThisForm1(frm As Form)
    'принтиране на форма 1

    MousePointer = vbHourglass
    
    If frmPrint.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrintMix.pbPrint.ScaleMode = 1
    Printer.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    Printer.PrintQuality = -4
    
    frmPrintMix.pbPrint.Width = Printer.Width
    frmPrintMix.pbPrint.Height = frmPrintMix.pbPrint.Width * 1.41
    
    For i = 1 To numSheetsForm1
        frmPrintMix.pbPrint.Line (50, 50)-(frmPrintMix.pbPrint.Width - 200, 50)
        frmPrintMix.pbPrint.Line (50, frmPrintMix.pbPrint.Height - 100)-(frmPrintMix.pbPrint.Width - 200, frmPrintMix.pbPrint.Height - 100)
        frmPrintMix.pbPrint.Line (50, 50)-(50, frmPrintMix.pbPrint.Height - 100)
        frmPrintMix.pbPrint.Line (50, frmPrintMix.pbPrint.Height / 2)-(frmPrintMix.pbPrint.Width - 200, frmPrintMix.pbPrint.Height / 2)
        frmPrintMix.pbPrint.Line (frmPrintMix.pbPrint.Width - 200, 50)-(frmPrintMix.pbPrint.Width - 200, frmPrintMix.pbPrint.Height - 100)
    
        If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10

        For Each ctr In frm

            If TypeOf ctr Is Label Then
                If ctr.Visible = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + 50
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                End If

            ElseIf TypeOf ctr Is TextBox Then

                If ctr.Enabled = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + 50
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                    X1 = ctr.Left
                    Y1 = ctr.Top + ctr.Height + 30 - 450
                    X2 = X1 + ctr.Width
                    Y2 = Y1 + ctr.Height - 50
                    frmPrintMix.pbPrint.Line (X1, Y1)-(X2, Y1)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X2, Y2)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X1, Y1)
                    frmPrintMix.pbPrint.Line (X2, Y2)-(X2, Y1)
                End If
            End If

        Next ctr

        For Each ctr In frm

            If TypeOf ctr Is Label Then
                If ctr.Visible = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + frmPrintMix.pbPrint.Height / 2
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                End If

            ElseIf TypeOf ctr Is TextBox Then

                If ctr.Enabled = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + frmPrintMix.pbPrint.Height / 2
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                    X1 = ctr.Left
                    Y1 = ctr.Top + ctr.Height + 30 + frmPrintMix.pbPrint.Height / 2 - 500
                    X2 = X1 + ctr.Width
                    Y2 = Y1 + ctr.Height - 50
                    frmPrintMix.pbPrint.Line (X1, Y1)-(X2, Y1)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X2, Y2)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X1, Y1)
                    frmPrintMix.pbPrint.Line (X2, Y2)-(X2, Y1)
                End If
            End If

        Next ctr
    
        If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
        
        PrintThePictureMix frmPrintMix, frmPrintMix.pbPrint, 96, 350, 300
        
        If i < numSheetsForm1 Then
            Printer.NewPage
        Else
        End If
    Next i
    
    MousePointer = vbDefault
    
    Unload frm

    If PrintRightBut = True Then
        frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Max
        Printer.EndDoc
        Unload frmPrintMix
        PrintRightBut = False

        Exit Sub

    End If
    
    If PrintAnyForm = True Then
        Printer.NewPage
        Call FillForm2
    Else
        frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Max
        Printer.EndDoc
        Unload frmPrintMix
    End If

End Sub

Public Function FillForm2()

    'попълване на форма 2
    If PrintRightBut = True Then GoTo SkipAuto2
    
    'проверка от регистъра за принтиране на форма 2 след експедицията
    If rPrint2 = 0 And PrintAnyForm = True Then
        Call FillForm3

        Exit Function

    End If

SkipAuto2:

    frmPrintMix.Show
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    Dim ResForm2 As Result

    Set ResForm2 = New Result
    
    Dim TotalIMKGz(0 To 5)   As Single

    Dim TotalIMKGi(0 To 5)   As Single

    Dim TotalCemKGz(0 To 3)  As Single

    Dim TotalCemKGi(0 To 3)  As Single

    Dim TotalWatKGz(0 To 1)  As Single

    Dim TotalWatKGi(0 To 1)  As Single

    Dim TotalChemKGz(0 To 5) As Single

    Dim TotalChemKGi(0 To 5) As Single

    prntForm2.txtExpNote.Text = ""
    prntForm2.txtDate.Text = ""
    prntForm2.txtOrd.Text = ""
    prntForm2.txtClnt.Text = ""
    prntForm2.txtObj.Text = ""
    prntForm2.txtDrv.Text = ""
    prntForm2.txtDrvNo.Text = ""
    prntForm2.txtDist.Text = ""
    prntForm2.txtRecType.Text = ""
    prntForm2.txtVol.Text = ""
    prntForm2.txtW.Text = ""
    prntForm2.txtOrdVol.Text = ""
    prntForm2.txtClass.Text = ""
    prntForm2.txtClassK.Text = ""
    prntForm2.txtClassV.Text = ""
    prntForm2.txtClassH.Text = ""
    prntForm2.txtClassP.Text = ""
    prntForm2.txtCem1.Text = ""
    prntForm2.txtCem2.Text = ""
    prntForm2.txtCem3.Text = ""
    prntForm2.txtChem1.Text = ""
    prntForm2.txtChem2.Text = ""
    prntForm2.txtChem3.Text = ""
    prntForm2.txtEDM.Text = ""
    prntForm2.txtMixTime.Text = ""
    prntForm2.txtExpTime.Text = ""
    prntForm2.txtOper.Text = ""
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10

    '-----------------------Start postgreSQL-----------------------------------
    Dim cnUnique As ADODB.Connection

    Dim rsUnique As Recordset
        
    Set cnUnique = New ADODB.Connection
    cnUnique.ConnectionTimeout = 10
    cnUnique.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме последния направен замес като сортираме в обратен ред и вземаме най-големия номер
    Set rsUnique = cnUnique.Execute("SELECT * FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1")
    
    If Not rsUnique.EOF And Not rsUnique.BOF Then
        ExpBufferff = rsUnique!exp_num 'буфер за номера на последната експедиция
        'попълване на полета на бележката от последния замес
        prntForm2.txtExpNote.Text = "M" & MachineNumber & "-" & Format(rsUnique!exp_num, "000000000") 'номер на бележка е номер на последна експедиция
        prntForm2.txtDate.Text = Left$(rsUnique!time_mix_ready, 10) 'дата на бележка от последния замес
        prntForm2.txtOrd.Text = Format(rsUnique!ord_num, "0000000") & "/" & rsUnique!ord_date 'номер/дата на заявката
        prntForm2.txtClnt.Text = rsUnique!name_clnt 'име на клиента
        prntForm2.txtObj.Text = rsUnique!obj_clnt 'име на обекта
        prntForm2.txtDist.Text = rsUnique!km_clnt 'разстояние до обекта
        prntForm2.txtDrv.Text = rsUnique!name_drv 'име на водача
        prntForm2.txtDrvNo.Text = rsUnique!reg_drv 'номер на превозно средство
        prntForm2.txtRecType.Text = rsUnique!type_rec 'рецепта тип
'        prntForm2.txtOrdVol.Text = rDs(rsUnique!ord_q) 'общо количество по заявката
        prntForm2.txtClass.Text = rsUnique!class_rec 'клас по якост
        prntForm2.txtClassK.Text = rsUnique!classk_rec 'клас по консистенция
        prntForm2.txtClassV.Text = rsUnique!classv_rec 'клас по въздействие
        prntForm2.txtClassH.Text = rsUnique!classh_rec 'клас по хлориди
        prntForm2.txtClassP.Text = rsUnique!classp_rec 'водоплътност
        prntForm2.txtEDM.Text = rsUnique!edm_rec 'едм
        prntForm2.txtMixTime.Text = Mid$(rsUnique!time_exp_start, 14, 5) 'час на стартиране на експедицията
        prntForm2.txtExpTime.Text = Mid$(rsUnique!time_mix_ready, 14, 5) 'час на последния замес по експедицията
        prntForm2.txtOper.Text = rsUnique!name_op 'име и фамилия на диспечера
        
        ResForm2.ExpQuant = ARound(CSng(rDs(rsUnique!exp_q)), 2) 'заявен обем за експедицията
        
        'попълване на всички имена на материал за таблицата във форма 2
        ResForm2.IMname(1) = rsUnique!im1_name
        ResForm2.IMname(2) = rsUnique!im2_name
        ResForm2.IMname(3) = rsUnique!im3_name
        ResForm2.IMname(4) = rsUnique!im4_name
        ResForm2.IMname(5) = rsUnique!im5_name
        ResForm2.IMname(6) = rsUnique!im6_name
        
        'попълваме заявените цименти
        ResForm2.SCRname(1) = rsUnique!cem1_name
        ResForm2.SCRstated(1) = Val(rsUnique!cem1z)
        ResForm2.SCRname(2) = rsUnique!cem2_name
        ResForm2.SCRstated(2) = Val(rsUnique!cem2z)
        ResForm2.SCRname(3) = rsUnique!cem3_name
        ResForm2.SCRstated(3) = Val(rsUnique!cem3z)
        ResForm2.SCRname(4) = rsUnique!cem4_name
        ResForm2.SCRstated(4) = Val(rsUnique!cem4z)
        
        'попълваме заявените води
        ResForm2.WATname(1) = rsUnique!wat1_name
        ResForm2.WATstated(1) = Val(rsUnique!wat1z)
        ResForm2.WATname(2) = rsUnique!wat2_name
        ResForm2.WATstated(2) = Val(rsUnique!wat2z)
        
        'попълваме заявените химически добавки
        ResForm2.CHEMname(1) = rsUnique!chem1_name
        ResForm2.CHEMstated(1) = rDs(rsUnique!chem1z)
        ResForm2.CHEMname(2) = rsUnique!chem2_name
        ResForm2.CHEMstated(2) = rDs(rsUnique!chem2z)
        ResForm2.CHEMname(3) = rsUnique!chem3_name
        ResForm2.CHEMstated(3) = rDs(rsUnique!chem3z)
        ResForm2.CHEMname(4) = rsUnique!chem4_name
        ResForm2.CHEMstated(4) = rDs(rsUnique!chem4z)
        ResForm2.CHEMname(5) = rsUnique!chem5_name
        ResForm2.CHEMstated(5) = rDs(rsUnique!chem5z)
        ResForm2.CHEMname(6) = rsUnique!chem6_name
        ResForm2.CHEMstated(6) = rDs(rsUnique!chem6z)
    End If
    
    'зареждане на първите 3 срещнати използвани в рецептата материали от силозите
    Dim ret As Integer

    For r = 1 To ns3

        If ResForm2.SCRstated(r) > 0 Then
            prntForm2.txtCem1.Text = ResForm2.SCRname(r)
            ret = r + 1

            Exit For

        Else
            ret = r + 1
        End If

    Next r

    If ret <= ns3 Then

        For r = ret To ns3

            If ResForm2.SCRstated(r) > 0 Then
                prntForm2.txtCem2.Text = ResForm2.SCRname(r)
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
                prntForm2.txtCem3.Text = ResForm2.SCRname(r)
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
            prntForm2.txtChem1.Text = ResForm2.CHEMname(r)
            ret = r + 1

            Exit For

        Else
            ret = r + 1
        End If

    Next r

    If ret <= ns4 Then

        For r = ret To ns4

            If ResForm2.CHEMstated(r) > 0 Then
                prntForm2.txtChem2.Text = ResForm2.CHEMname(r)
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
                prntForm2.txtChem3.Text = ResForm2.CHEMname(r)
                ret = r + 1

                Exit For

            Else
                ret = r + 1
            End If

        Next r

    End If
        
    'прочитане на последната експедиция от файла за да пресметнем теглата и обемите за форма 2
    Set rsUnique = cnUnique.Execute("SELECT * FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & ExpBufferff & " ORDER BY mix_num ASC")
    
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    
    ResForm2.TotalStatedKG = 0 'нулираме променливата за кг по рецепта
    ResForm2.TotalMeasuredKG = 0 'нулираме променливата за кг по изпълнение
    ResForm2.TotalQuant = 0 'нулираме променливата за обем по изпълнение

    For i = 0 To 5
        TotalIMKGz(i) = 0
        TotalIMKGi(i) = 0
    Next i

    For i = 0 To 3
        TotalCemKGz(i) = 0
        TotalCemKGi(i) = 0
    Next i

    For i = 0 To 1
        TotalWatKGz(i) = 0
        TotalWatKGi(i) = 0
    Next i

    For i = 0 To 5
        TotalChemKGz(i) = 0
        TotalChemKGi(i) = 0
    Next i

    Do While Not rsUnique.EOF
        ResForm2.TotalStatedKG = ResForm2.TotalStatedKG + CSng(rDs(rsUnique!total_rec_kg)) 'сума кг по зададено
        ResForm2.TotalMeasuredKG = ResForm2.TotalMeasuredKG + CSng(rDs(rsUnique!total_real_kg)) 'сума кг по изпълнено
        ResForm2.TotalQuant = ResForm2.TotalQuant + CSng(rDs(rsUnique!total_vol)) 'сума реален обем
        
        'сума на отделните ИМ по зададено
        TotalIMKGz(0) = TotalIMKGz(0) + Val(rsUnique!im1z)
        TotalIMKGz(1) = TotalIMKGz(1) + Val(rsUnique!im2z)
        TotalIMKGz(2) = TotalIMKGz(2) + Val(rsUnique!im3z)
        TotalIMKGz(3) = TotalIMKGz(3) + Val(rsUnique!im4z)
        TotalIMKGz(4) = TotalIMKGz(4) + Val(rsUnique!im5z)
        TotalIMKGz(5) = TotalIMKGz(5) + Val(rsUnique!im6z)
        
        'сума на отделните ИМ по изпълнено
        TotalIMKGi(0) = TotalIMKGi(0) + Val(rsUnique!im1i)
        TotalIMKGi(1) = TotalIMKGi(1) + Val(rsUnique!im2i)
        TotalIMKGi(2) = TotalIMKGi(2) + Val(rsUnique!im3i)
        TotalIMKGi(3) = TotalIMKGi(3) + Val(rsUnique!im4i)
        TotalIMKGi(4) = TotalIMKGi(4) + Val(rsUnique!im5i)
        TotalIMKGi(5) = TotalIMKGi(5) + Val(rsUnique!im6i)
        
        'сума на отделните цименти по зададено
        TotalCemKGz(0) = TotalCemKGz(0) + Val(rsUnique!cem1z)
        TotalCemKGz(1) = TotalCemKGz(1) + Val(rsUnique!cem2z)
        TotalCemKGz(2) = TotalCemKGz(2) + Val(rsUnique!cem3z)
        TotalCemKGz(3) = TotalCemKGz(3) + Val(rsUnique!cem4z)
        
        'сума на отделните цименти по изпълнено
        TotalCemKGi(0) = TotalCemKGi(0) + Val(rsUnique!cem1i)
        TotalCemKGi(1) = TotalCemKGi(1) + Val(rsUnique!cem2i)
        TotalCemKGi(2) = TotalCemKGi(2) + Val(rsUnique!cem3i)
        TotalCemKGi(3) = TotalCemKGi(3) + Val(rsUnique!cem4i)
        
        'сума на вода по зададено
        TotalWatKGz(0) = TotalWatKGz(0) + Val(rsUnique!wat1z)
        TotalWatKGz(1) = TotalWatKGz(1) + Val(rsUnique!wat2z)
        
        'сума на вода по изпълнено
        TotalWatKGi(0) = TotalWatKGi(0) + Val(rsUnique!wat1i)
        TotalWatKGi(1) = TotalWatKGi(1) + Val(rsUnique!wat2i)
        
        'сума на отделните хд по зададено
        TotalChemKGz(0) = TotalChemKGz(0) + CSng(rDs(rsUnique!chem1z))
        TotalChemKGz(1) = TotalChemKGz(1) + CSng(rDs(rsUnique!chem2z))
        TotalChemKGz(2) = TotalChemKGz(2) + CSng(rDs(rsUnique!chem3z))
        TotalChemKGz(3) = TotalChemKGz(3) + CSng(rDs(rsUnique!chem4z))
        TotalChemKGz(4) = TotalChemKGz(4) + CSng(rDs(rsUnique!chem5z))
        TotalChemKGz(5) = TotalChemKGz(5) + CSng(rDs(rsUnique!chem6z))
        
        'сума на отделните хд по изпълнено
        TotalChemKGi(0) = TotalChemKGi(0) + CSng(rDs(rsUnique!chem1i))
        TotalChemKGi(1) = TotalChemKGi(1) + CSng(rDs(rsUnique!chem2i))
        TotalChemKGi(2) = TotalChemKGi(2) + CSng(rDs(rsUnique!chem3i))
        TotalChemKGi(3) = TotalChemKGi(3) + CSng(rDs(rsUnique!chem4i))
        TotalChemKGi(4) = TotalChemKGi(4) + CSng(rDs(rsUnique!chem5i))
        TotalChemKGi(5) = TotalChemKGi(5) + CSng(rDs(rsUnique!chem6i))
        
        rsUnique.MoveNext
    Loop
    
    Set rsUnique = cnUnique.Execute("SELECT total_vol FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm2.txtOrd.Text) & " ORDER BY mix_num ASC")
    Dim totsmth As Single
    totsmth = 0
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    Do While Not rsUnique.EOF
        totsmth = ARound(totsmth, 2) + ARound(CSng(rDs(rsUnique!total_vol)), 2)
        rsUnique.MoveNext
    Loop
    
    Set rsUnique = cnUnique.Execute("SELECT DISTINCT ON (exp_num) exp_q FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm2.txtOrd.Text) & " ORDER BY exp_num ASC")
    Dim totexpsmth As Single
    totexpsmth = 0
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    Do While Not rsUnique.EOF
        totexpsmth = ARound(totexpsmth, 2) + ARound(CSng(rDs(rsUnique!exp_q)), 2)
        rsUnique.MoveNext
    Loop
    
    rsUnique.Close
    Set rsUnique = Nothing
    cnUnique.Close
    MousePointer = vbDefault
    Set cnUnique = Nothing
    '--------------------------End PostgreSQL-----------------------------------
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet   As Boolean

    Dim strSubKey As String

    strSubKey = Trim(PlaceProgSet2)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)

    If PrevSet = True Then
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm2, "RealVol", ErrRes)
    Else
        rRealVol = 1
    End If
        
    If rRealVol = 1 Then
        prntForm2.txtVol.Text = ARound(ResForm2.TotalQuant, 2) 'реален обем на експедицията
        prntForm2.txtOrdVol.Text = totsmth
    Else
        prntForm2.txtVol.Text = ResForm2.ExpQuant 'заявен обем на експедицията
        prntForm2.txtOrdVol.Text = totexpsmth
    End If
        
    prntForm2.txtW.Text = ARound(ResForm2.TotalMeasuredKG, 0)

    For e = 1 To ns1
        prntForm2.txtIMname(e - 1).Text = ResForm2.IMname(e)
        prntForm2.txtIMkgR(e - 1).Text = TotalIMKGi(e - 1) 'кг по измерено ИМ за всеки материал от всички замеси
        prntForm2.txtIMkg(e - 1).Text = TotalIMKGz(e - 1) 'кг по зададено ИМ за всеки материал от всички замеси

        If TotalIMKGz(e - 1) > 0 Then
            prntForm2.txtIMDiff(e - 1).Text = ARound(100 * (TotalIMKGi(e - 1) - TotalIMKGz(e - 1)) / TotalIMKGz(e - 1), 2)
        Else
            prntForm2.txtIMDiff(e - 1).Text = 0
        End If

        If CSng(rDs(prntForm2.txtIMDiff(e - 1).Text)) < 3 And CSng(rDs(prntForm2.txtIMDiff(e - 1).Text)) > -3 Then
            prntForm2.txtIMOK(e - 1).Text = uniYes
        Else
            prntForm2.txtIMOK(e - 1).Text = uniNo
        End If

    Next e
        
    For e = 1 To ns3
        prntForm2.txtCemname(e - 1).Text = ResForm2.SCRname(e)
        prntForm2.txtCemkgR(e - 1).Text = TotalCemKGi(e - 1) 'кг по измерено цимент за всеки материал от всички замеси
        prntForm2.txtCemkg(e - 1).Text = TotalCemKGz(e - 1) 'кг по зададено цимент за всеки материал от всички замеси

        If TotalCemKGz(e - 1) > 0 Then
            prntForm2.txtCemDiff(e - 1).Text = ARound(100 * (TotalCemKGi(e - 1) - TotalCemKGz(e - 1)) / TotalCemKGz(e - 1), 2)
        Else
            prntForm2.txtCemDiff(e - 1).Text = 0
        End If

        If CSng(rDs(prntForm2.txtCemDiff(e - 1).Text)) < 3 And CSng(rDs(prntForm2.txtCemDiff(e - 1).Text)) > -3 Then
            prntForm2.txtCemOK(e - 1).Text = uniYes
        Else
            prntForm2.txtCemOK(e - 1).Text = uniNo
        End If

    Next e
        
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    For e = 1 To ns2
        prntForm2.txtWatname(e - 1).Text = ResForm2.WATname(e)
        prntForm2.txtWatkgR(e - 1).Text = TotalWatKGi(e - 1) 'кг по измерено вода от всички замеси
        prntForm2.txtWatkg(e - 1).Text = TotalWatKGz(e - 1) 'кг по зададено вода от всички замеси

        If TotalWatKGz(e - 1) > 0 Then
            prntForm2.txtWatDiff(e - 1).Text = ARound(100 * (TotalWatKGi(e - 1) - TotalWatKGz(e - 1)) / TotalWatKGz(e - 1), 2)
        Else
            prntForm2.txtWatDiff(e - 1).Text = 0
        End If

        If CSng(rDs(prntForm2.txtWatDiff(e - 1).Text)) < 3 And CSng(rDs(prntForm2.txtWatDiff(e - 1).Text)) > -3 Then
            prntForm2.txtWatOK(e - 1).Text = uniYes
        Else
            prntForm2.txtWatOK(e - 1).Text = uniNo
        End If

    Next e
    
    For e = 1 To ns4
        prntForm2.txtChemname(e - 1).Text = ResForm2.CHEMname(e)
        prntForm2.txtChemkgR(e - 1).Text = CSng(rDs(TotalChemKGi(e - 1))) 'кг по измерено химия за всеки материал от всички замеси
        prntForm2.txtChemkg(e - 1).Text = CSng(rDs(TotalChemKGz(e - 1))) 'кг по зададено химия за всеки материал от всички замеси

        If CSng(rDs(TotalChemKGz(e - 1))) > 0 Then
            prntForm2.txtChemDiff(e - 1).Text = ARound(100 * (TotalChemKGi(e - 1) - CSng(rDs(TotalChemKGz(e - 1)))) / CSng(rDs(TotalChemKGz(e - 1))), 2)
        Else
            prntForm2.txtChemDiff(e - 1).Text = 0
        End If

        If CSng(rDs(prntForm2.txtChemDiff(e - 1).Text)) < 5 And CSng(rDs(prntForm2.txtChemDiff(e - 1).Text)) > -5 Then
            prntForm2.txtChemOK(e - 1).Text = uniYes
        Else
            prntForm2.txtChemOK(e - 1).Text = uniNo
        End If

    Next e
    
    Set ResForm2 = Nothing
    
    MousePointer = vbDefault
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    'показване на редактора за бележки ако е включена опцията
    If ShowEditor = True Then
        prntForm2.Show
    Else
        Call PrintThisForm2(prntForm2)
    End If
End Function

Public Sub PrintThisForm2(frm As Form)
    'принтиране на форма 2
    
    MousePointer = vbHourglass
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrintMix.pbPrint.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    
    frmPrintMix.pbPrint.Width = Printer.Width
    frmPrintMix.pbPrint.Height = frmPrintMix.pbPrint.Width * 1.41
    
    For i = 1 To numSheetsForm2
        frmPrintMix.pbPrint.Line (50, 50)-(frmPrintMix.pbPrint.Width - 200, 50)
        frmPrintMix.pbPrint.Line (50, frmPrintMix.pbPrint.Height - 100)-(frmPrintMix.pbPrint.Width - 200, frmPrintMix.pbPrint.Height - 100)
        frmPrintMix.pbPrint.Line (50, 50)-(50, frmPrintMix.pbPrint.Height - 100)
        frmPrintMix.pbPrint.Line (frmPrintMix.pbPrint.Width - 200, 50)-(frmPrintMix.pbPrint.Width - 200, frmPrintMix.pbPrint.Height - 100)
    
        If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10

        For Each ctr In frm

            If TypeOf ctr Is Label Then
                If ctr.Visible = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + 50
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                End If

            ElseIf TypeOf ctr Is TextBox Then

                If ctr.Enabled = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + 50
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                    X1 = ctr.Left
                    Y1 = ctr.Top + ctr.Height + 30 - 450
                    X2 = X1 + ctr.Width
                    Y2 = Y1 + ctr.Height - 50
                    frmPrintMix.pbPrint.Line (X1, Y1)-(X2, Y1)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X2, Y2)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X1, Y1)
                    frmPrintMix.pbPrint.Line (X2, Y2)-(X2, Y1)
                End If
            End If

        Next ctr
    
        If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
        PrintThePictureMix frmPrintMix, frmPrintMix.pbPrint, 96, 350, 300
        
        If i < numSheetsForm2 Then
            Printer.NewPage
        Else
        End If
    Next i
    
    MousePointer = vbDefault
    
    Unload frm
    
    If PrintRightBut = True Then
        frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Max
        Printer.EndDoc
        Unload frmPrintMix
        PrintRightBut = False

        Exit Sub

    End If
    
    If PrintAnyForm = True Then
        Printer.NewPage
        Call FillForm3
    Else
        frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Max
        Printer.EndDoc
        Unload frmPrintMix
    End If

End Sub

Public Function FillForm3()

    'попълване на форма 3
    If PrintRightBut = True Then GoTo SkipAuto3
    
    'проверка от регистъра за принтиране на форма 3 след експедицията
    If rPrint3 = 0 And PrintAnyForm = True Then
        frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Max
        Printer.EndDoc
        Unload frmPrintMix

        Exit Function

    End If
    
SkipAuto3:

    MousePointer = vbHourglass
    
    frmPrintMix.Show
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    Dim ResForm3 As Result

    Set ResForm3 = New Result

    prntForm3.txtExpNote.Text = ""
    prntForm3.txtDate.Text = ""
    prntForm3.txtOrd.Text = ""
    prntForm3.txtClnt.Text = ""
    prntForm3.txtObj.Text = ""
    prntForm3.txtDrv.Text = ""
    prntForm3.txtDrvNo.Text = ""
    prntForm3.txtDist.Text = ""
    prntForm3.txtRecType.Text = ""
    prntForm3.txtVol.Text = ""
    prntForm3.txtW.Text = ""
    prntForm3.txtOrdVol.Text = ""
    prntForm3.txtClass.Text = ""
    prntForm3.txtClassK.Text = ""
    prntForm3.txtClassV.Text = ""
    prntForm3.txtClassH.Text = ""
    prntForm3.txtClassP.Text = ""
    prntForm3.txtCem1.Text = ""
    prntForm3.txtCem2.Text = ""
    prntForm3.txtCem3.Text = ""
    prntForm3.txtChem1.Text = ""
    prntForm3.txtChem2.Text = ""
    prntForm3.txtChem3.Text = ""
    prntForm3.txtEDM.Text = ""
    prntForm3.txtMixTime.Text = ""
    prntForm3.txtExpTime.Text = ""
    prntForm3.txtOper.Text = ""
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10

    '-----------------------Start postgreSQL-----------------------------------
    Dim cnUnique As ADODB.Connection

    Dim rsUnique As Recordset
        
    Set cnUnique = New ADODB.Connection
    cnUnique.ConnectionTimeout = 10
    cnUnique.Open ConStr
    MousePointer = vbHourglass
    
    'маркираме последния направен замес като сортираме в обратен ред и вземаме най-големия номер
    Set rsUnique = cnUnique.Execute("SELECT * FROM mix_result_bc" & MachineNumber & " ORDER BY mix_num DESC LIMIT 1")
    
    If Not rsUnique.EOF And Not rsUnique.BOF Then
        ExpBufferff = rsUnique!exp_num 'буфер за номера на последната експедиция
        'попълване на полета на бележката от последния замес
        prntForm3.txtExpNote.Text = "M" & MachineNumber & "-" & Format(rsUnique!exp_num, "000000000") 'номер на бележка е номер на последна експедиция
        prntForm3.txtDate.Text = Left$(rsUnique!time_mix_ready, 10) 'дата на бележка от последния замес
        prntForm3.txtOrd.Text = Format(rsUnique!ord_num, "0000000") & "/" & rsUnique!ord_date 'номер/дата на заявката
        prntForm3.txtClnt.Text = rsUnique!name_clnt 'име на клиента
        prntForm3.txtObj.Text = rsUnique!obj_clnt 'име на обекта
        prntForm3.txtDist.Text = rsUnique!km_clnt 'разстояние до обекта
        prntForm3.txtDrv.Text = rsUnique!name_drv 'име на водача
        prntForm3.txtDrvNo.Text = rsUnique!reg_drv 'номер на превозно средство
        prntForm3.txtRecType.Text = rsUnique!type_rec 'рецепта тип
'        prntForm3.txtOrdVol.Text = rDs(rsUnique!ord_q) 'общо количество по заявката
        prntForm3.txtClass.Text = rsUnique!class_rec 'клас по якост
        prntForm3.txtClassK.Text = rsUnique!classk_rec 'клас по консистенция
        prntForm3.txtClassV.Text = rsUnique!classv_rec 'клас по въздействие
        prntForm3.txtClassH.Text = rsUnique!classh_rec 'клас по хлориди
        prntForm3.txtClassP.Text = rsUnique!classp_rec 'водоплътност
        prntForm3.txtEDM.Text = rsUnique!edm_rec 'едм
        prntForm3.txtMixTime.Text = Mid$(rsUnique!time_exp_start, 14, 5) 'час на стартиране на експедицията
        prntForm3.txtExpTime.Text = Mid$(rsUnique!time_mix_ready, 14, 5) 'час на последния замес по експедицията
        prntForm3.txtOper.Text = rsUnique!name_op 'име и фамилия на диспечера
        
        ResForm3.ExpQuant = ARound(CSng(rDs(rsUnique!exp_q)), 2) 'заявен обем за експедицията
        
        'попълваме заявените цименти
        ResForm3.SCRname(1) = rsUnique!cem1_name
        ResForm3.SCRstated(1) = rsUnique!cem1z
        ResForm3.SCRname(2) = rsUnique!cem2_name
        ResForm3.SCRstated(2) = rsUnique!cem2z
        ResForm3.SCRname(3) = rsUnique!cem3_name
        ResForm3.SCRstated(3) = rsUnique!cem3z
        ResForm3.SCRname(4) = rsUnique!cem4_name
        ResForm3.SCRstated(4) = rsUnique!cem4z
        
        'попълваме заявените химически добавки
        ResForm3.CHEMname(1) = rsUnique!chem1_name
        ResForm3.CHEMstated(1) = rDs(rsUnique!chem1z)
        ResForm3.CHEMname(2) = rsUnique!chem2_name
        ResForm3.CHEMstated(2) = rDs(rsUnique!chem2z)
        ResForm3.CHEMname(3) = rsUnique!chem3_name
        ResForm3.CHEMstated(3) = rDs(rsUnique!chem3z)
        ResForm3.CHEMname(4) = rsUnique!chem4_name
        ResForm3.CHEMstated(4) = rDs(rsUnique!chem4z)
        ResForm3.CHEMname(5) = rsUnique!chem5_name
        ResForm3.CHEMstated(5) = rDs(rsUnique!chem5z)
        ResForm3.CHEMname(6) = rsUnique!chem6_name
        ResForm3.CHEMstated(6) = rDs(rsUnique!chem6z)
    End If
    
    'зареждане на първите 3 срещнати използвани в рецептата материали от силозите
    Dim ret As Integer

    For r = 1 To ns3

        If ResForm3.SCRstated(r) > 0 Then
            prntForm3.txtCem1.Text = ResForm3.SCRname(r)
            ret = r + 1

            Exit For

        Else
            ret = r + 1
        End If

    Next r

    If ret <= ns3 Then

        For r = ret To ns3

            If ResForm3.SCRstated(r) > 0 Then
                prntForm3.txtCem2.Text = ResForm3.SCRname(r)
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
                prntForm3.txtCem3.Text = ResForm3.SCRname(r)
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
            prntForm3.txtChem1.Text = ResForm3.CHEMname(r)
            ret = r + 1

            Exit For

        Else
            ret = r + 1
        End If

    Next r

    If ret <= ns4 Then

        For r = ret To ns4

            If ResForm3.CHEMstated(r) > 0 Then
                prntForm3.txtChem2.Text = ResForm3.CHEMname(r)
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
                prntForm3.txtChem3.Text = ResForm3.CHEMname(r)
                ret = r + 1

                Exit For

            Else
                ret = r + 1
            End If

        Next r

    End If
        
    'прочитане на последната експедиция от файла за да пресметнем теглата и обемите за форма 3
    Set rsUnique = cnUnique.Execute("SELECT total_rec_kg, total_real_kg, total_vol FROM mix_result_bc" & MachineNumber & " WHERE exp_num = " & ExpBufferff & " ORDER BY mix_num ASC")
    
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    
    ResForm3.TotalStatedKG = 0 'нулираме променливата за кг по рецепта
    ResForm3.TotalMeasuredKG = 0 'нулираме променливата за кг по изпълнение
    ResForm3.TotalQuant = 0 'нулираме променливата за обем по изпълнение
    
    Do While Not rsUnique.EOF
        ResForm3.TotalStatedKG = ResForm3.TotalStatedKG + CSng(rDs(rsUnique!total_rec_kg))
        ResForm3.TotalMeasuredKG = ResForm3.TotalMeasuredKG + CSng(rDs(rsUnique!total_real_kg))
        ResForm3.TotalQuant = ResForm3.TotalQuant + CSng(rDs(rsUnique!total_vol))
        rsUnique.MoveNext
    Loop
    
    Set rsUnique = cnUnique.Execute("SELECT total_vol FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm3.txtOrd.Text) & " ORDER BY mix_num ASC")
    Dim totsmth As Single
    totsmth = 0
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    Do While Not rsUnique.EOF
        totsmth = ARound(totsmth, 2) + ARound(CSng(rDs(rsUnique!total_vol)), 2)
        rsUnique.MoveNext
    Loop
    
    Set rsUnique = cnUnique.Execute("SELECT DISTINCT ON (exp_num) exp_q FROM mix_result_bc" & MachineNumber & " WHERE ord_num = " & Val(prntForm3.txtOrd.Text) & " ORDER BY exp_num ASC")
    Dim totexpsmth As Single
    totexpsmth = 0
    If Not rsUnique.BOF And Not rsUnique.EOF Then rsUnique.MoveFirst
    Do While Not rsUnique.EOF
        totexpsmth = ARound(totexpsmth, 2) + ARound(CSng(rDs(rsUnique!exp_q)), 2)
        rsUnique.MoveNext
    Loop
    
    rsUnique.Close
    Set rsUnique = Nothing
    cnUnique.Close
    MousePointer = vbDefault
    Set cnUnique = Nothing
    '--------------------------End PostgreSQL-----------------------------------
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet   As Boolean

    Dim strSubKey As String

    strSubKey = Trim(PlaceProgSet3)
    PrevSet = CheckRegistryKey(HKEY_CURRENT_USER, strSubKey)

    If PrevSet = True Then
        rRealVol = GetSetting(PlaceProgSettings, PlaceForm3, "RealVol", ErrRes)
    Else
        rRealVol = 1
    End If
        
    If rRealVol = 1 Then
        prntForm3.txtVol.Text = ARound(ResForm3.TotalQuant, 2) 'реален обем на експедицията
        prntForm3.txtOrdVol.Text = totsmth
    Else
        prntForm3.txtVol.Text = ResForm3.ExpQuant 'заявен обем на експедицията
        prntForm3.txtOrdVol.Text = totexpsmth
    End If
        
    prntForm3.txtW.Text = ARound(ResForm3.TotalMeasuredKG, 0)
    
    Set ResForm3 = Nothing
    
    MousePointer = vbDefault
    
    'показване на редактора за бележки ако е включена опцията
    If ShowEditor = True Then
        prntForm3.Show
    Else
        Call PrintThisForm3(prntForm3)
    End If
End Function

Public Sub PrintThisForm3(frm As Form)
    'принтиране на форма 3
    
    MousePointer = vbHourglass
    
    If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrintMix.pbPrint.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    
    frmPrintMix.pbPrint.Width = Printer.Width
    frmPrintMix.pbPrint.Height = frmPrintMix.pbPrint.Width * 1.41
    
    For i = 1 To numSheetsForm3
        frmPrintMix.pbPrint.Line (50, 50)-(frmPrintMix.pbPrint.Width - 200, 50)
        frmPrintMix.pbPrint.Line (50, frmPrintMix.pbPrint.Height - 100)-(frmPrintMix.pbPrint.Width - 200, frmPrintMix.pbPrint.Height - 100)
        frmPrintMix.pbPrint.Line (50, 50)-(50, frmPrintMix.pbPrint.Height - 100)
        frmPrintMix.pbPrint.Line (frmPrintMix.pbPrint.Width - 200, 50)-(frmPrintMix.pbPrint.Width - 200, frmPrintMix.pbPrint.Height - 100)
    
        If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10

        For Each ctr In frm

            If TypeOf ctr Is Label Then
                If ctr.Visible = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + 50
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                End If

            ElseIf TypeOf ctr Is TextBox Then

                If ctr.Enabled = True Then
                    frmPrintMix.pbPrint.CurrentX = ctr.Left + 50
                    frmPrintMix.pbPrint.CurrentY = ctr.Top + 50
                    frmPrintMix.pbPrint.Font = ctr.Font
                    frmPrintMix.pbPrint.FontSize = ctr.FontSize
                    frmPrintMix.pbPrint.FontBold = ctr.FontBold
                    frmPrintMix.pbPrint.FontItalic = ctr.FontItalic
                    frmPrintMix.pbPrint.Print ctr
                    X1 = ctr.Left
                    Y1 = ctr.Top + ctr.Height + 30 - 450
                    X2 = X1 + ctr.Width
                    Y2 = Y1 + ctr.Height - 50
                    frmPrintMix.pbPrint.Line (X1, Y1)-(X2, Y1)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X2, Y2)
                    frmPrintMix.pbPrint.Line (X1, Y2)-(X1, Y1)
                    frmPrintMix.pbPrint.Line (X2, Y2)-(X2, Y1)
                End If
            End If

        Next ctr
    
        If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10
    
        Call PrintRTFMix(prntForm3btn.Confirmity, 850, 7300, 800, 300)
    
        If frmPrintMix.barPrint.Value < frmPrintMix.barPrint.Max Then frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Value + 10

        PrintThePictureMix frmPrintMix, frmPrintMix.pbPrint, 96, 350, 300
        
        If i < numSheetsForm3 Then
            Printer.NewPage
        Else
        End If
    Next i
    
    MousePointer = vbDefault
    Printer.EndDoc
    
    If PrintRightBut = True Then PrintRightBut = False
    
    frmPrintMix.barPrint.Value = frmPrintMix.barPrint.Max
    Unload frmPrintMix
    Unload frm
End Sub

Public Function OpenOrders()
    'зареждане на меню заявки

    Dim colx As ColumnHeader

    Dim itmX As ListItem
    
    DispPanel.cmbOrdRec.SetFocus
    DispPanel.cmbOrdRec.TabIndex = 0
    DispPanel.cmbOrdClnt.TabIndex = 1
    DispPanel.cmbOrdClntObj.TabIndex = 2
    DispPanel.txtOrdQuant.TabIndex = 3
    DispPanel.queOrdDate.TabIndex = 4
    DispPanel.queOrdTime.TabIndex = 5
    DispPanel.btnClearOrd.TabIndex = 6
    DispPanel.btnSvNwOrd.TabIndex = 7
    DispPanel.btnDelOrd.TabIndex = 8
    DispPanel.btnDisp.TabIndex = 9
    DispPanel.btnOrders.TabIndex = 10
    DispPanel.btnRecepies.TabIndex = 11
    DispPanel.btnClients.TabIndex = 12
    DispPanel.btnDrivers.TabIndex = 13
    DispPanel.btnSuppliers.TabIndex = 14
    DispPanel.btnMaterials.TabIndex = 15
    DispPanel.btnNotes.TabIndex = 16
    DispPanel.btnAdminPanel.TabIndex = 17
    DispPanel.btnExit.TabIndex = 18
    
    DispPanel.lstOrd.ColumnHeaders.Clear
    DispPanel.lstOrd.ListItems.Clear
    
    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 1000
    
    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniDateOrd
    colx.Width = 1750
    
    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniDateReadyShort
    colx.Width = 1750
    
    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniOrdered
    colx.Width = 1100
    colx.Tag = "number"
    
    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniMade
    colx.Width = 1100

    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniRecCode
    colx.Width = 1200

    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniNm & " " & uniRec
    colx.Width = 1200
    
    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniClass
    colx.Width = 1300
    
    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniClntCode
    colx.Width = 1100

    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniNm & " " & uniClnt
    colx.Width = 1300

    Set colx = DispPanel.lstOrd.ColumnHeaders.Add()
    colx.Text = uniObj
    colx.Width = 1300
        
    DispPanel.nowOrdDate.Value = Now
    DispPanel.queOrdDate = Now
    DispPanel.queOrdTime = Now
    
    DispPanel.txtOrd.MaxLength = 10
    DispPanel.txtOrdQuant.MaxLength = 4

    DispPanel.frOrders.Left = 120
    DispPanel.frOrders.Top = DispPanel.Height \ 2 - FrmPlace
    
    DispPanel.cmbOrdRec.Clear
    DispPanel.cmbOrdClnt.Clear
    DispPanel.cmbOrdClntName.Clear
    
    '------------------------------Start PostgreSQL----------------------------------
    Dim cn As ADODB.Connection

    Dim rs As Recordset
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    'зареждане на всички заявки в ListView
    Set rs = cn.Execute("SELECT * FROM orders WHERE stamp_date >= '" & Format(Now - Day(31), "DD-MM-YYYY") & "' ORDER BY order_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Set itmX = DispPanel.lstOrd.ListItems.Add(1, , Format(rs!order_num, "0000000"))
        itmX.SubItems(1) = rs!order_date
        itmX.SubItems(2) = rs!order_date_que
        itmX.SubItems(3) = rDs(rs!order_q)
        itmX.SubItems(4) = rDs(rs!order_qmade)
        itmX.SubItems(5) = Format(rs!order_rec, "0000")
        itmX.SubItems(6) = rs!order_rec_name
        itmX.SubItems(7) = rs!order_rec_class
        itmX.SubItems(8) = Format(rs!order_clnt, "0000")
        itmX.SubItems(9) = rs!order_clnt_name
        itmX.SubItems(10) = rs!order_clnt_obj
        DispPanel.txtOrd.Text = Format(Val(rs!order_num) + 1, "0000000")
        rs.MoveNext
    Loop
    
    'зареждане на комбото с рецептите
    Set rs = cn.Execute("SELECT r_num FROM recepies ORDER BY r_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        DispPanel.cmbOrdRec.AddItem Format(rs!r_num, "0000")
        rs.MoveNext
    Loop
    
    'зареждане на комбото с клиентите
    Set rs = cn.Execute("SELECT c_num, c_name FROM clients WHERE c_show = '1' ORDER BY c_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        DispPanel.cmbOrdClnt.AddItem Format(rs!c_num, "0000")
        DispPanel.cmbOrdClntName.AddItem rs!c_name
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
    '--------------------------End PostgreSQL------------------------------------------

    If DispPanel.lstOrd.ListItems.count > 0 Then
        AutoColW DispPanel.lstOrd
    Else
        DispPanel.txtOrd.Text = Format(DispPanel.lstOrd.ListItems.count + 1, "0000000")
    End If
    
    If DispPanel.lstRec.ListItems.count > 0 And DispPanel.txtRec <> "" Then
        DispPanel.cmbOrdRec.Text = DispPanel.lstRec.ListItems(DispPanel.lstRec.SelectedItem.Index).Text
    Else
    End If
    
    If DispPanel.lstClnt.ListItems.count > 0 And DispPanel.txtClnt <> "" Then
        DispPanel.cmbOrdClnt.Text = DispPanel.lstClnt.ListItems(DispPanel.lstClnt.SelectedItem.Index).Text
    Else
    End If
    
End Function

Public Function ClearOrdBut()
    'функция за почистване на клетките на заявките за въвеждане на нова

    If DispPanel.frOrders.Enabled = True And DispPanel.frOrders.Visible = True Then
        DispPanel.cmbOrdRec.ListIndex = -1
        DispPanel.cmbOrdClnt.ListIndex = -1
        DispPanel.cmbOrdClntObj.Clear
        DispPanel.txtOrdQuant.Text = 0
        DispPanel.nowOrdDate.Value = Now
    
        If DispPanel.lstOrd.ListItems.count > 0 Then
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn As ADODB.Connection

            Dim rs As Recordset
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT order_num FROM orders ORDER BY order_num DESC LIMIT 1") 'маркираме последния запис
    
            DispPanel.txtOrd.Text = Format(Val(rs!order_num) + 1, "0000000")
    
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------
        
        Else
            DispPanel.txtOrd.Text = Format(Val(DispPanel.lstOrd.ListItems.count) + 1, "0000000")
        End If

    Else
    End If

End Function

Public Function SvNwOrdBut()
    'функция за запис на заявка

    If DispPanel.frOrders.Enabled = True And DispPanel.frOrders.Visible = True Then

        Dim OrdNew As Order

        Set OrdNew = New Order
        
        DispPanel.queOrdDate.Hour = DispPanel.queOrdTime.Hour
        DispPanel.queOrdDate.Minute = DispPanel.queOrdTime.Minute
        DispPanel.queOrdDate.Second = DispPanel.queOrdTime.Second
    
        DispPanel.nowOrdDate.Value = Now

        If DispPanel.queOrdDate.Value < DispPanel.nowOrdDate.Value Then
            DispPanel.queOrdDate.Value = DispPanel.nowOrdDate.Value
        Else
        End If

        If Len(DispPanel.txtOrd.Text) > 0 And Len(DispPanel.txtOrdQuant.Text) > 0 And Len(DispPanel.cmbOrdRec.Text) > 0 And Len(DispPanel.cmbOrdClnt.Text) > 0 Then
            DispPanel.nowOrdDate.Value = Now
            OrdNew.DateBegin = Format(Now, "DD.MM.YYYY - HH:MM:SS")
            OrdNew.DateEnd = Format(DispPanel.queOrdDate.Value, "DD.MM.YYYY - HH:MM:SS")
            OrdNew.OrderedQuant = ARound(CSng(rDs(DispPanel.txtOrdQuant.Text)), 2)
            OrdNew.RecCode = Val(DispPanel.cmbOrdRec.Text)
            OrdNew.RecTitle = DispPanel.txtOrdRecName.Text
            OrdNew.RecClass = DispPanel.txtOrdRecClass.Text
            OrdNew.ClntCode = Val(DispPanel.cmbOrdClnt.Text)
            OrdNew.ClntTitle = DispPanel.cmbOrdClntName.Text
            OrdNew.ClntWorksite = DispPanel.cmbOrdClntObj.Text
            OrdNew.MadeQuant = 0
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn     As ADODB.Connection

            Dim rs     As Recordset

            Dim comIns As String
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
    
            Set rs = cn.Execute("SELECT order_num FROM orders ORDER BY order_num DESC LIMIT 1") 'маркираме последния запис
            
            If Not rs.EOF And Not rs.BOF Then
                OrdNew.Code = Val(rs!order_num) + 1
            Else
                OrdNew.Code = 1
            End If
    
            comIns = "INSERT INTO orders VALUES(" & OrdNew.Code & ",'" & OrdNew.DateBegin & "','" & OrdNew.DateEnd & "','" & Format(DispPanel.queOrdDate.Value, "DD-MM-YYYY") & "','" & OrdNew.OrderedQuant & "','" & OrdNew.MadeQuant & "'," & OrdNew.RecCode & ",'" & OrdNew.RecTitle & "','" & OrdNew.RecClass & "'," & OrdNew.ClntCode & ",'" & OrdNew.ClntTitle & "','" & OrdNew.ClntWorksite & "')"
            Set rs = cn.Execute(comIns)
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------
            
            Set itmX = DispPanel.lstOrd.ListItems.Add(1, , Format(OrdNew.Code, "0000000"))
            itmX.SubItems(1) = OrdNew.DateBegin
            itmX.SubItems(2) = OrdNew.DateEnd
            itmX.SubItems(3) = OrdNew.OrderedQuant
            itmX.SubItems(4) = OrdNew.MadeQuant
            itmX.SubItems(5) = Format(OrdNew.RecCode, "0000")
            itmX.SubItems(6) = OrdNew.RecTitle
            itmX.SubItems(7) = OrdNew.RecClass
            itmX.SubItems(8) = Format(OrdNew.ClntCode, "0000")
            itmX.SubItems(9) = OrdNew.ClntTitle
            itmX.SubItems(10) = OrdNew.ClntWorksite
                
            If DispPanel.lstOrd.ListItems.count > 0 Then AutoColW DispPanel.lstOrd
            
            DispPanel.txtOrd.Text = Format(OrdNew.Code + 1, "0000000")
            DispPanel.nowOrdDate.Value = Now
            DispPanel.queOrdDate.Value = Now
            DispPanel.txtOrdQuant.Text = 0
            DispPanel.cmbOrdRec.ListIndex = -1
            DispPanel.txtOrdRecName.Text = ""
            DispPanel.txtOrdRecClass.Text = ""
            DispPanel.cmbOrdClnt.ListIndex = -1
            DispPanel.cmbOrdClntName.Text = ""
            DispPanel.cmbOrdClntObj.ListIndex = -1
            
            MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave

            Exit Function

        Else
            MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx

            Exit Function

        End If

    Else
    End If

End Function

Public Function DelOrdBut()
    'функция за изтриване на заявка

    If DispPanel.frOrders.Enabled = True And DispPanel.frOrders.Visible = True Then
        If Len(DispPanel.txtOrd.Text) > 0 And Len(DispPanel.txtOrdQuant.Text) > 0 And Len(DispPanel.cmbOrdRec.Text) > 0 And Len(DispPanel.cmbOrdClnt.Text) > 0 Then
            response = MsgBox(MsgConfDel, vbYesNo Or vbQuestion, MsgEditBx)

            If response = vbYes Then
            
                '------------------------------Start PostgreSQL----------------------------------
                Dim cn As ADODB.Connection

                Dim rs As Recordset
    
                Set cn = New ADODB.Connection
                cn.ConnectionTimeout = 10
                cn.Open ConStr
                MousePointer = vbHourglass
                
                Set rs = cn.Execute("DELETE FROM orders WHERE order_num = " & Val(DispPanel.txtOrd.Text) & ";") 'изтриване на запис
    
                DispPanel.lstOrd.ListItems.Clear
    
                Set rs = cn.Execute("SELECT * FROM orders ORDER BY order_num ASC;")
                
                If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
                Do While Not rs.EOF
                    Set itmX = DispPanel.lstOrd.ListItems.Add(1, , Format(rs!order_num, "0000000"))
                    itmX.SubItems(1) = rs!order_date
                    itmX.SubItems(2) = rs!order_date_que
                    itmX.SubItems(3) = rDs(rs!order_q)
                    itmX.SubItems(4) = rDs(rs!order_qmade)
                    itmX.SubItems(5) = Format(rs!order_rec, "0000")
                    itmX.SubItems(6) = rs!order_rec_name
                    itmX.SubItems(7) = rs!order_rec_class
                    itmX.SubItems(8) = Format(rs!order_clnt, "0000")
                    itmX.SubItems(9) = rs!order_clnt_name
                    itmX.SubItems(10) = rs!order_clnt_obj
                    rs.MoveNext
                Loop
                
                If Not rs.EOF And Not rs.BOF Then
                    DispPanel.txtOrd.Text = Format(Val(rs!order_num) + 1, "0000000")
                Else
                    DispPanel.txtOrd.Text = "0000001"
                End If

                rs.Close
                Set rs = Nothing
                cn.Close
                MousePointer = vbDefault
                Set cn = Nothing
                '--------------------------End PostgreSQL----------------------------------------

                DispPanel.nowOrdDate.Value = Now
                DispPanel.queOrdDate.Value = Now
                DispPanel.txtOrdQuant.Text = 0
                DispPanel.cmbOrdRec.ListIndex = -1
                DispPanel.txtOrdRecName.Text = ""
                DispPanel.txtOrdRecClass.Text = ""
                DispPanel.cmbOrdClnt.ListIndex = -1
                DispPanel.cmbOrdClntName.Text = ""
                DispPanel.cmbOrdClntObj.ListIndex = -1

                If DispPanel.lstOrd.ListItems.count > 0 Then AutoColW DispPanel.lstOrd
                MsgBox MsgDelSuccess, vbOKOnly Or vbInformation, MsgDelBx
            Else

                Exit Function

            End If

        Else
            MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function ListOrdClick()

    'функция за зареждане на данни при маркиране на запис от таблицата
    On Error Resume Next

    If DispPanel.frOrders.Enabled = True And DispPanel.frOrders.Visible = True Then
        If DispPanel.lstOrd.ListItems.count > 0 Then
            DispPanel.txtOrd.Text = Format(Val(DispPanel.lstOrd.ListItems(DispPanel.lstOrd.SelectedItem.Index).Text), "0000000")
            DispPanel.txtOrdQuant.Text = DispPanel.lstOrd.ListItems(DispPanel.lstOrd.SelectedItem.Index).ListSubItems(3).Text
            DispPanel.cmbOrdRec.Text = Format(Val(DispPanel.lstOrd.ListItems(DispPanel.lstOrd.SelectedItem.Index).ListSubItems(5).Text), "0000")
            DispPanel.cmbOrdClnt.Text = Format(Val(DispPanel.lstOrd.ListItems(DispPanel.lstOrd.SelectedItem.Index).ListSubItems(8).Text), "0000")
            DispPanel.cmbOrdClntObj.Text = DispPanel.lstOrd.ListItems(DispPanel.lstOrd.SelectedItem.Index).ListSubItems(10).Text
        End If

    Else
    End If

End Function

Public Function ChangeOrdRec()
    'функция за избор на рецепта по заявка и зареждане на данни

    If DispPanel.cmbOrdRec.Text <> "" Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
        
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        'зареждане на рецептите
        Set rs = cn.Execute("SELECT r_name, r_class FROM recepies WHERE r_num = " & Val(DispPanel.cmbOrdRec.Text) & ";")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
        Do While Not rs.EOF
            DispPanel.txtOrdRecName.Text = rs!r_name
            DispPanel.txtOrdRecClass.Text = rs!r_class
            rs.MoveNext
        Loop
    
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------

    Else
        DispPanel.txtOrdRecName.Text = ""
        DispPanel.txtOrdRecClass.Text = ""
    End If

End Function

Public Function ChangeOrdClnt()
    'функция за смяна на клиент по заявка и зареждане на данни

    If DispPanel.cmbOrdClnt.Text <> "" Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
    
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        'зареждане на клиентите
        Set rs = cn.Execute("SELECT c_name FROM clients WHERE c_num = " & Val(DispPanel.cmbOrdClnt.Text) & ";")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
        Do While Not rs.EOF
            DispPanel.cmbOrdClntName.Text = rs!c_name
            rs.MoveNext
        Loop
    
        'зареждане на обектите
        DispPanel.cmbOrdClntObj.Clear
        Set rs = cn.Execute("SELECT w_name FROM worksites WHERE w_cnum = '" & Val(DispPanel.cmbOrdClnt.Text) & "' AND w_show = 'true';")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
        Do While Not rs.EOF
            DispPanel.cmbOrdClntObj.AddItem rs!w_name
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------
    
    Else
        DispPanel.cmbOrdClntName.Text = ""
        '        DispPanel.txtOrdClntObj.Text = ""
    End If

End Function
Public Function ChangeOrdClntName()
    'функция за смяна на клиент по заявка и зареждане на данни
    On Error Resume Next
    
    If DispPanel.cmbOrdClntName.Text <> "" Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
    
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        'зареждане на клиентите
        Set rs = cn.Execute("SELECT c_num FROM clients WHERE c_name = '" & DispPanel.cmbOrdClntName.Text & "';")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
        Do While Not rs.EOF
            DispPanel.cmbOrdClnt.Text = Format(rs!c_num, "0000")
            rs.MoveNext
        Loop
        DispPanel.cmbOrdClnt.Refresh
        
        'зареждане на обектите
        DispPanel.cmbOrdClntObj.Clear
        Set rs = cn.Execute("SELECT w_name FROM worksites WHERE w_cnum = '" & Val(DispPanel.cmbOrdClnt.Text) & "' AND w_show = 'true';")
    
        If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
        Do While Not rs.EOF
            DispPanel.cmbOrdClntObj.AddItem rs!w_name
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------
    
    Else
        DispPanel.cmbOrdClnt.ListIndex = -1
        '        DispPanel.txtOrdClntObj.Text = ""
    End If

End Function

Public Function OpenRecepies()
    'зареждане на меню рецепти
    
    Dim Rec As Recipe
        
    Set Rec = New Recipe
    
    Const colw = 1300
    
    Dim colx           As ColumnHeader

    Dim itmX           As ListItem

    Dim count          As Integer

    Dim intEmpFileNbr1 As Integer
    
    'настройка на табулаторите
    DispPanel.txtRec.SetFocus
    DispPanel.txtRec.TabIndex = 0
    DispPanel.txtNameRec.TabIndex = 1
    DispPanel.txtTypeRec.TabIndex = 2
    DispPanel.txtClassRec.TabIndex = 3
    DispPanel.txtClassRecK.TabIndex = 4
    DispPanel.txtClassRecV.TabIndex = 5
    DispPanel.txtClassRecH.TabIndex = 6
    DispPanel.txtClassRecP.TabIndex = 7
    DispPanel.txtEDMRec.TabIndex = 8
    DispPanel.txtTimeMixRec.TabIndex = 9
    DispPanel.txtTimePourRec.TabIndex = 10
    DispPanel.cmbRec1(0).TabIndex = 11
    DispPanel.txtRec1(0).TabIndex = 12
    DispPanel.cmbRec1(1).TabIndex = 13
    DispPanel.txtRec1(1).TabIndex = 14
    DispPanel.cmbRec1(2).TabIndex = 15
    DispPanel.txtRec1(2).TabIndex = 16
    DispPanel.cmbRec1(3).TabIndex = 17
    DispPanel.txtRec1(3).TabIndex = 18
    DispPanel.cmbRec1(4).TabIndex = 19
    DispPanel.txtRec1(4).TabIndex = 20
    DispPanel.cmbRec1(5).TabIndex = 21
    DispPanel.txtRec1(5).TabIndex = 22
    DispPanel.cmbRec3(0).TabIndex = 23
    DispPanel.txtRec3(0).TabIndex = 24
    DispPanel.cmbRec3(1).TabIndex = 25
    DispPanel.txtRec3(1).TabIndex = 26
    DispPanel.cmbRec3(2).TabIndex = 27
    DispPanel.txtRec3(2).TabIndex = 28
    DispPanel.cmbRec3(3).TabIndex = 29
    DispPanel.txtRec3(3).TabIndex = 30
    DispPanel.cmbRec2(0).TabIndex = 31
    DispPanel.txtRec2(0).TabIndex = 32
    DispPanel.cmbRec2(1).TabIndex = 33
    DispPanel.txtRec2(1).TabIndex = 34
    DispPanel.cmbRec4(0).TabIndex = 35
    DispPanel.txtRec4(0).TabIndex = 36
    DispPanel.cmbRec4(1).TabIndex = 37
    DispPanel.txtRec4(1).TabIndex = 38
    DispPanel.cmbRec4(2).TabIndex = 39
    DispPanel.txtRec4(2).TabIndex = 40
    DispPanel.cmbRec4(3).TabIndex = 41
    DispPanel.txtRec4(3).TabIndex = 42
    DispPanel.cmbRec4(4).TabIndex = 43
    DispPanel.txtRec4(4).TabIndex = 44
    DispPanel.cmbRec4(5).TabIndex = 45
    DispPanel.txtRec4(5).TabIndex = 46
    DispPanel.btnClearRec.TabIndex = 47
    DispPanel.btnSvNwRec.TabIndex = 48
    DispPanel.btnDelRec.TabIndex = 49
    DispPanel.btnShowRec.TabIndex = 50
    DispPanel.btnDisp.TabIndex = 51
    DispPanel.btnOrders.TabIndex = 52
    DispPanel.btnRecepies.TabIndex = 53
    DispPanel.btnClients.TabIndex = 54
    DispPanel.btnDrivers.TabIndex = 55
    DispPanel.btnSuppliers.TabIndex = 56
    DispPanel.btnMaterials.TabIndex = 57
    DispPanel.btnNotes.TabIndex = 58
    DispPanel.btnAdminPanel.TabIndex = 59
    DispPanel.btnExit.TabIndex = 60

    intEmpFileNbr1 = FreeFile

    'позициониране на рамката
    DispPanel.frRecepies.Left = 120
    DispPanel.frRecepies.Top = DispPanel.Height \ 2 - FrmPlace
    
    'почистване на таблицата
    DispPanel.lstRec.ColumnHeaders.Clear
    DispPanel.lstRec.ListItems.Clear

    'настройка на заглавките на таблицата
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 700
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniNm
    colx.Width = 1600
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniRecType
    colx.Width = 1200
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniClass
    colx.Width = 1500
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniClassK
    colx.Width = 1700
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniClassV
    colx.Width = 1700
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniClassH
    colx.Width = 1700
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniClassP
    colx.Width = 1700
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniEDM
    colx.Width = 1200
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniTimePourShort
    colx.Width = 500
        
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniTimeMixShort
    colx.Width = 500
    
    'параметри на текстовите клетки
    DispPanel.txtRec.MaxLength = 4
    DispPanel.txtRec.Text = ""
    DispPanel.txtNameRec.MaxLength = 25
    DispPanel.txtNameRec.Text = ""
    DispPanel.txtTypeRec.MaxLength = 20
    DispPanel.txtTypeRec.Text = ""
    DispPanel.txtClassRec.MaxLength = 10
    DispPanel.txtClassRec.Text = ""
    DispPanel.txtClassRecK.MaxLength = 10
    DispPanel.txtClassRecK.Text = ""
    DispPanel.txtClassRecV.MaxLength = 10
    DispPanel.txtClassRecV.Text = ""
    DispPanel.txtClassRecH.MaxLength = 10
    DispPanel.txtClassRecH.Text = ""
    DispPanel.txtClassRecP.MaxLength = 10
    DispPanel.txtClassRecP.Text = ""
    DispPanel.txtEDMRec.MaxLength = 10
    DispPanel.txtEDMRec.Text = ""
    DispPanel.txtTimePourRec.MaxLength = 2
    DispPanel.txtTimePourRec.Text = TPd
    DispPanel.txtTimeMixRec.MaxLength = 2
    DispPanel.txtTimeMixRec.Text = TMd
    DispPanel.txtTotalKg.Text = 0

    'настройка за визуализациата на необходимия брой клетки за въвеждане на рецепта спрямо течките на машината
    For i = 0 To 5
        DispPanel.cmbRec1(i).Visible = False
        DispPanel.cmbRec1(i).Clear
        DispPanel.txtRec1(i).Visible = False
        DispPanel.txtRec1(i).MaxLength = 4
    Next i
    
    For i = 0 To 3
        DispPanel.cmbRec3(i).Visible = False
        DispPanel.cmbRec3(i).Clear
        DispPanel.txtRec3(i).Visible = False
        DispPanel.txtRec3(i).MaxLength = 4
    Next i
    
    For i = 0 To 1
        DispPanel.cmbRec2(i).Visible = False
        DispPanel.cmbRec2(i).Clear
        DispPanel.txtRec2(i).Visible = False
        DispPanel.txtRec2(i).MaxLength = 4
    Next i
    
    For i = 0 To 5
        DispPanel.cmbRec4(i).Visible = False
        DispPanel.cmbRec4(i).Clear
        DispPanel.txtRec4(i).Visible = False
        DispPanel.txtRec4(i).MaxLength = 5
    Next i

    For i = 0 To ns1 - 1
        DispPanel.cmbRec1(i).Visible = True
        DispPanel.txtRec1(i).Visible = True
    Next i
    
    For i = 0 To ns3 - 1
        DispPanel.cmbRec3(i).Visible = True
        DispPanel.txtRec3(i).Visible = True
    Next i
    
    For i = 0 To ns2 - 1
        DispPanel.cmbRec2(i).Visible = True
        DispPanel.txtRec2(i).Visible = True
    Next i
    
    For i = 0 To ns4 - 1
        DispPanel.cmbRec4(i).Visible = True
        DispPanel.txtRec4(i).Visible = True
    Next i
    
    IM(0) = ""
    Scr(0) = ""
    Wat(0) = ""
    Chem(0) = ""

    'прочитане на файла с имената на материалите
    If Dir(SilosFile) = "" Then
        Open SilosFile For Output As #intEmpFileNbr1
        Close #intEmpFileNbr1
    Else
        Open SilosFile For Input As #intEmpFileNbr1

        Do Until EOF(intEmpFileNbr1)
            Input #intEmpFileNbr1, IM(1), IM(2), IM(3), IM(4), IM(5), IM(6), Scr(1), Scr(2), Scr(3), Scr(4), Wat(1), Wat(2), Chem(1), Chem(2), Chem(3), Chem(4), Chem(5), Chem(6)
        Loop

        Close #intEmpFileNbr1
    End If

    'добавяне на имената на материалите в комбобоксовете на рецептата
    For i = 0 To ns1
        For n = 0 To ns1 - 1
            DispPanel.cmbRec1(n).AddItem IM(i)
        Next n
    Next i
    
    For i = 0 To ns3
        For n = 0 To ns3 - 1
            DispPanel.cmbRec3(n).AddItem Scr(i)
        Next n
    Next i
    
    For i = 0 To ns2
        For n = 0 To ns2 - 1
            DispPanel.cmbRec2(n).AddItem Wat(i)
        Next n
    Next i
    
    For i = 0 To ns4
        For n = 0 To ns4 - 1
            DispPanel.cmbRec4(n).AddItem Chem(i)
        Next n
    Next i

    'запис на имената на течките в таблицата
    For count = 1 To ns1
        Set colx = DispPanel.lstRec.ColumnHeaders.Add()
        colx.Text = IM(count)
        colx.Width = colw
    Next count
    
    For count = 1 To ns3
        Set colx = DispPanel.lstRec.ColumnHeaders.Add()
        colx.Text = Scr(count)
        colx.Width = colw
    Next count
    
    For count = 1 To ns2
        Set colx = DispPanel.lstRec.ColumnHeaders.Add()
        colx.Text = Wat(count)
        colx.Width = colw
    Next count
    
    For count = 1 To ns4
        Set colx = DispPanel.lstRec.ColumnHeaders.Add()
        colx.Text = Chem(count)
        colx.Width = colw
    Next count
    
    Set colx = DispPanel.lstRec.ColumnHeaders.Add()
    colx.Text = uniTotalKg
    colx.Width = 1200

    '------------------------------Start PostgreSQL----------------------------------
    Dim cn       As ADODB.Connection

    Dim rs       As Recordset

    Dim frCheck  As Boolean

    Dim RecCount As Long
    
    frCheck = False
    RecCount = 0
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT r_num FROM recepies ORDER BY r_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        RecCount = RecCount + 1

        If Val(rs!r_num) <> RecCount And frCheck = False Then
            DispPanel.txtRec.Text = Format(RecCount, "0000")
            frCheck = True
        Else
            DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
        End If

        rs.MoveNext
    Loop
    
    Set rs = cn.Execute("SELECT * FROM recepies WHERE r_show = '1' ORDER BY r_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Rec.Code = Val(rs!r_num)
        Rec.Title = rs!r_name
        Rec.Kind = rs!r_type
        Rec.Class = rs!r_class
        Rec.ClassK = rs!r_classk
        Rec.ClassV = rs!r_classv
        Rec.ClassH = rs!r_classh
        Rec.ClassP = rs!r_classp
        Rec.EDM = Val(rs!r_edm)
        Rec.Tpour = Val(rs!r_tpour)
        Rec.Tmix = Val(rs!r_tmix)
        Rec.initIM(1) = Val(rs!init_im1)
        Rec.kgIM(1) = Val(rs!kg_im1)
        Rec.initIM(2) = Val(rs!init_im2)
        Rec.kgIM(2) = Val(rs!kg_im2)
        Rec.initIM(3) = Val(rs!init_im3)
        Rec.kgIM(3) = Val(rs!kg_im3)
        Rec.initIM(4) = Val(rs!init_im4)
        Rec.kgIM(4) = Val(rs!kg_im4)
        Rec.initIM(5) = Val(rs!init_im5)
        Rec.kgIM(5) = Val(rs!kg_im5)
        Rec.initIM(6) = Val(rs!init_im6)
        Rec.kgIM(6) = Val(rs!kg_im6)
        Rec.initScr(1) = Val(rs!init_scr1)
        Rec.kgScr(1) = Val(rs!kg_scr1)
        Rec.initScr(2) = Val(rs!init_scr2)
        Rec.kgScr(2) = Val(rs!kg_scr2)
        Rec.initScr(3) = Val(rs!init_scr3)
        Rec.kgScr(3) = Val(rs!kg_scr3)
        Rec.initScr(4) = Val(rs!init_scr4)
        Rec.kgScr(4) = Val(rs!kg_scr4)
        Rec.initWat(1) = Val(rs!init_wat1)
        Rec.kgWat(1) = Val(rs!kg_wat1)
        Rec.initWat(2) = Val(rs!init_wat2)
        Rec.kgWat(2) = Val(rs!kg_wat2)
        Rec.initChem(1) = Val(rs!init_chem1)
        Rec.kgChem(1) = CSng(rDs(rs!kg_chem1))
        Rec.initChem(2) = Val(rs!init_chem2)
        Rec.kgChem(2) = CSng(rDs(rs!kg_chem2))
        Rec.initChem(3) = Val(rs!init_chem3)
        Rec.kgChem(3) = CSng(rDs(rs!kg_chem3))
        Rec.initChem(4) = Val(rs!init_chem4)
        Rec.kgChem(4) = CSng(rDs(rs!kg_chem4))
        Rec.initChem(5) = Val(rs!init_chem5)
        Rec.kgChem(5) = CSng(rDs(rs!kg_chem5))
        Rec.initChem(6) = Val(rs!init_chem6)
        Rec.kgChem(6) = CSng(rDs(rs!kg_chem6))
        
        Set itmX = DispPanel.lstRec.ListItems.Add(1, , Format(Rec.Code, "0000"))
        itmX.SubItems(1) = Rec.Title
        itmX.SubItems(2) = Rec.Kind
        itmX.SubItems(3) = Rec.Class
        itmX.SubItems(4) = Rec.ClassK
        itmX.SubItems(5) = Rec.ClassV
        itmX.SubItems(6) = Rec.ClassH
        itmX.SubItems(7) = Rec.ClassP
        itmX.SubItems(8) = Rec.EDM
        itmX.SubItems(9) = Rec.Tpour
        itmX.SubItems(10) = Rec.Tmix
        
        For n = 0 To ns1 - 1
            itmX.SubItems(11 + n) = Rec.AllkgIM(n + 1)
        Next n
        
        For n = 0 To ns3 - 1
            itmX.SubItems(11 + ns1 + n) = Rec.AllkgScr(n + 1)
        Next n
        
        For n = 0 To ns2 - 1
            itmX.SubItems(11 + ns1 + ns3 + n) = Rec.AllkgWat(n + 1)
        Next n
        
        For n = 0 To ns4 - 1
            itmX.SubItems(11 + ns1 + ns3 + ns2 + n) = Rec.AllkgChem(n + 1)
        Next n
        
        itmX.SubItems(11 + ns1 + ns3 + ns2 + ns4) = rDs(rs!kg_total)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
    '--------------------------End PostgreSQL------------------------------------------
    
    For i = 0 To ns1 - 1
        DispPanel.txtRec1(i) = "0"
        DispPanel.cmbRec1(i).ListIndex = -1
        DispPanel.cmbRec1(i).Refresh
    Next i
    
    For i = 0 To ns3 - 1
        DispPanel.txtRec3(i) = "0"
        DispPanel.cmbRec3(i).ListIndex = -1
        DispPanel.cmbRec3(i).Refresh
    Next i
        
    For i = 0 To ns2 - 1
        DispPanel.txtRec2(i) = "0"
        DispPanel.cmbRec2(i).ListIndex = -1
        DispPanel.cmbRec2(i).Refresh
    Next i
    
    For i = 0 To ns4 - 1
        DispPanel.txtRec4(i) = "0"
        DispPanel.cmbRec4(i).ListIndex = -1
        DispPanel.cmbRec4(i).Refresh
    Next i
    
    If DispPanel.lstRec.ListItems.count > 0 Then
        AutoColW DispPanel.lstRec
    Else
        DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
    End If

End Function

Public Function ClearRecBut()
    'функция за почистване на клетките на рецептата за въвеждане на нова

    If DispPanel.frRecepies.Enabled = True And DispPanel.frRecepies.Visible = True Then
        If DispPanel.lstRec.ListItems.count > 0 Then
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn       As ADODB.Connection

            Dim rs       As Recordset

            Dim frCheck  As Boolean

            Dim RecCount As Integer
    
            frCheck = False
            RecCount = 0
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT r_num FROM recepies ORDER BY r_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                RecCount = RecCount + 1

                If Val(rs!r_num) <> RecCount And frCheck = False Then
                    DispPanel.txtRec.Text = Format(RecCount, "0000")
                    frCheck = True
                Else
                    DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
                End If
        
                Rec = rs!r_num
                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------
            
        Else
            DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
        End If
        
        DispPanel.txtNameRec.Text = ""
        DispPanel.txtTypeRec.Text = ""
        DispPanel.txtClassRec.Text = ""
        DispPanel.txtClassRecK.Text = ""
        DispPanel.txtClassRecV.Text = ""
        DispPanel.txtClassRecH.Text = ""
        DispPanel.txtClassRecP.Text = ""
        DispPanel.txtEDMRec.Text = ""
        DispPanel.txtTimeMixRec.Text = TMd
        DispPanel.txtTimePourRec.Text = TPd
        DispPanel.txtTotalKg.Text = 0

        For i = 0 To ns1 - 1
            DispPanel.txtRec1(i) = "0"
            DispPanel.cmbRec1(i).ListIndex = -1
        Next i
        
        For i = 0 To ns3 - 1
            DispPanel.txtRec3(i) = "0"
            DispPanel.cmbRec3(i).ListIndex = -1
        Next i
        
        For i = 0 To ns2 - 1
            DispPanel.txtRec2(i) = "0"
            DispPanel.cmbRec2(i).ListIndex = -1
        Next i
        
        For i = 0 To ns4 - 1
            DispPanel.txtRec4(i) = "0"
            DispPanel.cmbRec4(i).ListIndex = -1
        Next i
        
        DispPanel.txtRec.SetFocus
    Else
    End If

End Function

Public Function SvNwRecBut()
    'функция за запис на рецепта
    
    Dim RecCh As Recipe

    Set RecCh = New Recipe

    Dim RecNew As Recipe

    Set RecNew = New Recipe
        
    If DispPanel.frRecepies.Enabled = True And DispPanel.frRecepies.Visible = True Then
        If Val(DispPanel.txtRec.Text) < RecMin Then
            MsgBox MsgCodeZero, vbOKOnly Or vbCritical, MsgErrBx

            Exit Function

        Else
        End If
        
        Dim intEmpFileNbr1 As Integer

        Dim response       As Integer

        intEmpFileNbr1 = FreeFile
        
        Open SilosFile For Input As #intEmpFileNbr1

        Do Until EOF(intEmpFileNbr1)
            Input #intEmpFileNbr1, IM(1), IM(2), IM(3), IM(4), IM(5), IM(6), Scr(1), Scr(2), Scr(3), Scr(4), Wat(1), Wat(2), Chem(1), Chem(2), Chem(3), Chem(4), Chem(5), Chem(6)
        Loop

        Close #intEmpFileNbr1
        
        For n = 0 To ns1 - 1
            For i = 0 To ns1 - 1

                If DispPanel.cmbRec1(n).List(DispPanel.cmbRec1(n).ListIndex) = IM(i + 1) And IM(i + 1) <> "" Then
                    RecNew.initIM(n + 1) = BaseIM + i
                End If

            Next i
        Next n
    
        For n = 0 To ns3 - 1
            For i = 0 To ns3 - 1

                If DispPanel.cmbRec3(n).List(DispPanel.cmbRec3(n).ListIndex) = Scr(i + 1) And Scr(i + 1) <> "" Then
                    RecNew.initScr(n + 1) = BaseScr + i
                End If

            Next i
        Next n
    
        For n = 0 To ns2 - 1
            For i = 0 To ns2 - 1

                If DispPanel.cmbRec2(n).List(DispPanel.cmbRec2(n).ListIndex) = Wat(i + 1) And Wat(i + 1) <> "" Then
                    RecNew.initWat(n + 1) = BaseWat + i
                End If

            Next i
        Next n
    
        For n = 0 To ns4 - 1
            For i = 0 To ns4 - 1

                If DispPanel.cmbRec4(n).List(DispPanel.cmbRec4(n).ListIndex) = Chem(i + 1) And Chem(i + 1) <> "" Then
                    RecNew.initChem(n + 1) = BaseChem + i
                End If

            Next i
        Next n
    
        If Len(DispPanel.txtRec.Text) > 0 And Len(DispPanel.txtNameRec.Text) > 0 And Len(DispPanel.txtTimePourRec.Text) > 0 And Len(DispPanel.txtTimeMixRec.Text) > 0 Then
            RecNew.Code = Val(DispPanel.txtRec.Text)
            RecNew.Title = DispPanel.txtNameRec.Text
            RecNew.Kind = DispPanel.txtTypeRec.Text
            RecNew.Class = DispPanel.txtClassRec.Text
            RecNew.ClassK = DispPanel.txtClassRecK.Text
            RecNew.ClassV = DispPanel.txtClassRecV.Text
            RecNew.ClassH = DispPanel.txtClassRecH.Text
            RecNew.ClassP = DispPanel.txtClassRecP.Text
            RecNew.EDM = Val(DispPanel.txtEDMRec.Text)
            RecNew.Tmix = Val(DispPanel.txtTimeMixRec.Text)
            RecNew.Tpour = Val(DispPanel.txtTimePourRec.Text)
        
            For i = 0 To ns1 - 1

                If DispPanel.txtRec1(i).Text = "" Then 'ако клетката е празна записваме 0
                    DispPanel.txtRec1(i).Text = 0
                Else
                End If

                If RecNew.initIM(i + 1) = 0 Then 'в първия случай записваме данните в кг за да изведем съобщение за грешка
                    RecNew.kgIM(i + 1) = Val(DispPanel.txtRec1(i).Text)
                Else
                    RecNew.kgIM(i + 1) = Val(DispPanel.txtRec1(i).Text)

                    If RecNew.kgIM(i + 1) = 0 Then
                        MsgBox MsgMatNotQuant, vbOKOnly Or vbCritical, MsgErrBx

                        Exit Function

                    End If
                End If

                If RecNew.kgIM(i + 1) > 0 And RecNew.initIM(i + 1) = 0 Then
                    MsgBox MsgQuantNotMat, vbOKOnly Or vbCritical, MsgErrBx

                    Exit Function

                Else
                End If

            Next i
        
            For i = 0 To ns3 - 1

                If DispPanel.txtRec3(i).Text = "" Then
                    DispPanel.txtRec3(i).Text = 0
                Else
                End If

                If RecNew.initScr(i + 1) = 0 Then
                    RecNew.kgScr(i + 1) = Val(DispPanel.txtRec3(i).Text)
                Else
                    RecNew.kgScr(i + 1) = Val(DispPanel.txtRec3(i).Text)

                    If RecNew.kgScr(i + 1) = 0 Then
                        MsgBox MsgMatNotQuant, vbOKOnly Or vbCritical, MsgErrBx

                        Exit Function

                    End If
                End If

                If RecNew.kgScr(i + 1) > 0 And RecNew.initScr(i + 1) = 0 Then
                    MsgBox MsgQuantNotMat, vbOKOnly Or vbCritical, MsgErrBx

                    Exit Function

                Else
                End If

            Next i
        
            For i = 0 To ns2 - 1

                If DispPanel.txtRec2(i).Text = "" Then 'ако клетката е празна записваме 0
                    DispPanel.txtRec2(i).Text = 0
                Else
                End If

                If RecNew.initWat(i + 1) = 0 Then 'в първия случай записваме данните в кг за да изведем съобщение за грешка
                    RecNew.kgWat(i + 1) = Val(DispPanel.txtRec2(i).Text)
                Else
                    RecNew.kgWat(i + 1) = Val(DispPanel.txtRec2(i).Text)

                    If RecNew.kgWat(i + 1) = 0 Then
                        MsgBox MsgMatNotQuant, vbOKOnly Or vbCritical, MsgErrBx

                        Exit Function

                    End If
                End If

                If RecNew.kgWat(i + 1) > 0 And RecNew.initWat(i + 1) = 0 Then
                    MsgBox MsgQuantNotMat, vbOKOnly Or vbCritical, MsgErrBx

                    Exit Function

                Else
                End If

            Next i
        
            For i = 0 To ns4 - 1

                If DispPanel.txtRec4(i).Text = "" Then
                    DispPanel.txtRec4(i).Text = 0
                Else
                End If

                If RecNew.initChem(i + 1) = 0 Then
                    RecNew.kgChem(i + 1) = ARound(CSng(rDs(DispPanel.txtRec4(i).Text)), 2)
                Else
                    RecNew.kgChem(i + 1) = ARound(CSng(rDs(DispPanel.txtRec4(i).Text)), 2)

                    If RecNew.kgChem(i + 1) = 0 Then
                        MsgBox MsgMatNotQuant, vbOKOnly Or vbCritical, MsgErrBx

                        Exit Function

                    End If
                End If

                If RecNew.kgChem(i + 1) > 0 And RecNew.initChem(i + 1) = 0 Then
                    MsgBox MsgQuantNotMat, vbOKOnly Or vbCritical, MsgErrBx

                    Exit Function

                Else
                End If

            Next i
            
            RecNew.kgTotal = RecNew.kgIM(1) + RecNew.kgIM(2) + RecNew.kgIM(3) + RecNew.kgIM(4) + RecNew.kgIM(5) + RecNew.kgIM(6) + RecNew.kgScr(1) + RecNew.kgScr(2) + RecNew.kgScr(3) + RecNew.kgScr(4) + RecNew.kgWat(1) + RecNew.kgWat(2) + RecNew.kgChem(1) + RecNew.kgChem(2) + RecNew.kgChem(3) + RecNew.kgChem(4) + RecNew.kgChem(5) + RecNew.kgChem(6)
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn         As ADODB.Connection

            Dim rs         As Recordset

            Dim rs1        As Recordset

            Dim comIns     As String

            Dim comEdit    As String

            Dim numCheck   As Integer

            Dim nameCheck  As String

            Dim nameCheck1 As String

            Dim frCheck    As Boolean

            Dim RecCount   As Long
    
            frCheck = False
            RecCount = 0
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT r_num, r_name FROM recepies WHERE r_num = " & Val(DispPanel.txtRec) & ";") 'рецепта код ако я има
            
            If Not rs.BOF And Not rs.EOF Then
                numCheck = rs!r_num
                nameCheck = rs!r_name
            End If
            
            Set rs1 = cn.Execute("SELECT r_name FROM recepies WHERE r_name = '" & DispPanel.txtNameRec & "';") 'рецепта име ако я има
            
            If Not rs1.BOF And Not rs1.EOF Then nameCheck1 = rs1!r_name
    
            If numCheck <> Val(DispPanel.txtRec) And nameCheck <> DispPanel.txtNameRec And nameCheck1 <> DispPanel.txtNameRec Then
                'ако няма съвпадения в номер или име на рецепта правим запис
                comIns = "INSERT INTO recepies VALUES(" & RecNew.Code & ",'" & RecNew.Title & "','" & RecNew.Kind & "','" & RecNew.Class _
                   & "','" & RecNew.ClassK & "','" & RecNew.ClassV & "','" & RecNew.ClassH & "','" & RecNew.ClassP & "','" & RecNew.EDM _
                   & "','" & RecNew.Tpour & "','" & RecNew.Tmix _
                   & "','" & RecNew.initIM(1) & "','" & RecNew.kgIM(1) _
                   & "','" & RecNew.initIM(2) & "','" & RecNew.kgIM(2) _
                   & "','" & RecNew.initIM(3) & "','" & RecNew.kgIM(3) _
                   & "','" & RecNew.initIM(4) & "','" & RecNew.kgIM(4) _
                   & "','" & RecNew.initIM(5) & "','" & RecNew.kgIM(5) _
                   & "','" & RecNew.initIM(6) & "','" & RecNew.kgIM(6) _
                   & "','" & RecNew.initScr(1) & "','" & RecNew.kgScr(1) _
                   & "','" & RecNew.initScr(2) & "','" & RecNew.kgScr(2) _
                   & "','" & RecNew.initScr(3) & "','" & RecNew.kgScr(3) _
                   & "','" & RecNew.initScr(4) & "','" & RecNew.kgScr(4) _
                   & "','" & RecNew.initWat(1) & "','" & RecNew.kgWat(1) _
                   & "','" & RecNew.initWat(2) & "','" & RecNew.kgWat(2) _
                   & "','" & RecNew.initChem(1) & "','" & RecNew.kgChem(1) _
                   & "','" & RecNew.initChem(2) & "','" & RecNew.kgChem(2) _
                   & "','" & RecNew.initChem(3) & "','" & RecNew.kgChem(3) _
                   & "','" & RecNew.initChem(4) & "','" & RecNew.kgChem(4) _
                   & "','" & RecNew.initChem(5) & "','" & RecNew.kgChem(5) _
                   & "','" & RecNew.initChem(6) & "','" & RecNew.kgChem(6) _
                   & "','" & RecNew.kgTotal & "','true')"
                Set rs = cn.Execute(comIns)
            ElseIf numCheck = Val(DispPanel.txtRec) And nameCheck1 <> DispPanel.txtNameRec Then
                'ако има такъв номер на рецепта но въведеното име е друго - питаме за редакция по номер
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE recepies SET r_name = '" & RecNew.Title & "',r_type = '" & RecNew.Kind & "',r_class = '" & RecNew.Class _
                       & "',r_classk = '" & RecNew.ClassK & "',r_classv = '" & RecNew.ClassV & "',r_classh = '" & RecNew.ClassH _
                       & "',r_classp = '" & RecNew.ClassP & "',r_edm = '" & RecNew.EDM _
                       & "',r_tpour = '" & RecNew.Tpour & "',r_tmix = '" & RecNew.Tmix _
                       & "',init_im1 = '" & RecNew.initIM(1) & "',kg_im1 = '" & RecNew.kgIM(1) _
                       & "',init_im2 = '" & RecNew.initIM(2) & "',kg_im2 = '" & RecNew.kgIM(2) _
                       & "',init_im3 = '" & RecNew.initIM(3) & "',kg_im3 = '" & RecNew.kgIM(3) _
                       & "',init_im4 = '" & RecNew.initIM(4) & "',kg_im4 = '" & RecNew.kgIM(4) _
                       & "',init_im5 = '" & RecNew.initIM(5) & "',kg_im5 = '" & RecNew.kgIM(5) _
                       & "',init_im6 = '" & RecNew.initIM(6) & "',kg_im6 = '" & RecNew.kgIM(6) _
                       & "',init_scr1 = '" & RecNew.initScr(1) & "',kg_scr1 = '" & RecNew.kgScr(1) _
                       & "',init_scr2 = '" & RecNew.initScr(2) & "',kg_scr2 = '" & RecNew.kgScr(2) _
                       & "',init_scr3 = '" & RecNew.initScr(3) & "',kg_scr3 = '" & RecNew.kgScr(3) _
                       & "',init_scr4 = '" & RecNew.initScr(4) & "',kg_scr4 = '" & RecNew.kgScr(4) _
                       & "',init_wat1 = '" & RecNew.initWat(1) & "',kg_wat1 = '" & RecNew.kgWat(1) _
                       & "',init_wat2 = '" & RecNew.initWat(2) & "',kg_wat2 = '" & RecNew.kgWat(2) _
                       & "',init_chem1 = '" & RecNew.initChem(1) & "',kg_chem1 = '" & RecNew.kgChem(1) _
                       & "',init_chem2 = '" & RecNew.initChem(2) & "',kg_chem2 = '" & RecNew.kgChem(2) _
                       & "',init_chem3 = '" & RecNew.initChem(3) & "',kg_chem3 = '" & RecNew.kgChem(3) _
                       & "',init_chem4 = '" & RecNew.initChem(4) & "',kg_chem4 = '" & RecNew.kgChem(4) _
                       & "',init_chem5 = '" & RecNew.initChem(5) & "',kg_chem5 = '" & RecNew.kgChem(5) _
                       & "',init_chem6 = '" & RecNew.initChem(6) & "',kg_chem6 = '" & RecNew.kgChem(6) _
                       & "',kg_total = '" & RecNew.kgTotal _
                       & "' WHERE r_num =" & RecNew.Code & "" 'корекция по номер
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing
                    DispPanel.txtRec.SetFocus

                    Exit Function

                End If

            ElseIf nameCheck = DispPanel.txtNameRec And numCheck = Val(DispPanel.txtRec) Then
                'ако има такъв номер и името му е същото - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE recepies SET r_type = '" & RecNew.Kind & "',r_class = '" & RecNew.Class _
                       & "',r_classk = '" & RecNew.ClassK & "',r_classv = '" & RecNew.ClassV & "',r_classh = '" & RecNew.ClassH _
                       & "',r_classp = '" & RecNew.ClassP & "',r_edm = '" & RecNew.EDM _
                       & "',r_tpour = '" & RecNew.Tpour & "',r_tmix = '" & RecNew.Tmix _
                       & "',init_im1 = '" & RecNew.initIM(1) & "',kg_im1 = '" & RecNew.kgIM(1) _
                       & "',init_im2 = '" & RecNew.initIM(2) & "',kg_im2 = '" & RecNew.kgIM(2) _
                       & "',init_im3 = '" & RecNew.initIM(3) & "',kg_im3 = '" & RecNew.kgIM(3) _
                       & "',init_im4 = '" & RecNew.initIM(4) & "',kg_im4 = '" & RecNew.kgIM(4) _
                       & "',init_im5 = '" & RecNew.initIM(5) & "',kg_im5 = '" & RecNew.kgIM(5) _
                       & "',init_im6 = '" & RecNew.initIM(6) & "',kg_im6 = '" & RecNew.kgIM(6) _
                       & "',init_scr1 = '" & RecNew.initScr(1) & "',kg_scr1 = '" & RecNew.kgScr(1) _
                       & "',init_scr2 = '" & RecNew.initScr(2) & "',kg_scr2 = '" & RecNew.kgScr(2) _
                       & "',init_scr3 = '" & RecNew.initScr(3) & "',kg_scr3 = '" & RecNew.kgScr(3) _
                       & "',init_scr4 = '" & RecNew.initScr(4) & "',kg_scr4 = '" & RecNew.kgScr(4) _
                       & "',init_wat1 = '" & RecNew.initWat(1) & "',kg_wat1 = '" & RecNew.kgWat(1) _
                       & "',init_wat2 = '" & RecNew.initWat(2) & "',kg_wat2 = '" & RecNew.kgWat(2) _
                       & "',init_chem1 = '" & RecNew.initChem(1) & "',kg_chem1 = '" & RecNew.kgChem(1) _
                       & "',init_chem2 = '" & RecNew.initChem(2) & "',kg_chem2 = '" & RecNew.kgChem(2) _
                       & "',init_chem3 = '" & RecNew.initChem(3) & "',kg_chem3 = '" & RecNew.kgChem(3) _
                       & "',init_chem4 = '" & RecNew.initChem(4) & "',kg_chem4 = '" & RecNew.kgChem(4) _
                       & "',init_chem5 = '" & RecNew.initChem(5) & "',kg_chem5 = '" & RecNew.kgChem(5) _
                       & "',init_chem6 = '" & RecNew.initChem(6) & "',kg_chem6 = '" & RecNew.kgChem(6) _
                       & "',kg_total = '" & RecNew.kgTotal _
                       & "' WHERE r_name = '" & RecNew.Title & "';" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing
                    DispPanel.txtRec.SetFocus

                    Exit Function

                End If

            ElseIf numCheck <> Val(DispPanel.txtRec) And nameCheck1 = DispPanel.txtNameRec Then
                'ако няма такъв номер, но има такова име на рецепта - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE recepies SET r_num = " & RecNew.Code & ",r_type = '" & RecNew.Kind & "',r_class = '" & RecNew.Class _
                       & "',r_classk = '" & RecNew.ClassK & "',r_classv = '" & RecNew.ClassV & "',r_classh = '" & RecNew.ClassH _
                       & "',r_classp = '" & RecNew.ClassP & "',r_edm = '" & RecNew.EDM _
                       & "',r_tpour = '" & RecNew.Tpour & "',r_tmix = '" & RecNew.Tmix _
                       & "',init_im1 = '" & RecNew.initIM(1) & "',kg_im1 = '" & RecNew.kgIM(1) _
                       & "',init_im2 = '" & RecNew.initIM(2) & "',kg_im2 = '" & RecNew.kgIM(2) _
                       & "',init_im3 = '" & RecNew.initIM(3) & "',kg_im3 = '" & RecNew.kgIM(3) _
                       & "',init_im4 = '" & RecNew.initIM(4) & "',kg_im4 = '" & RecNew.kgIM(4) _
                       & "',init_im5 = '" & RecNew.initIM(5) & "',kg_im5 = '" & RecNew.kgIM(5) _
                       & "',init_im6 = '" & RecNew.initIM(6) & "',kg_im6 = '" & RecNew.kgIM(6) _
                       & "',init_scr1 = '" & RecNew.initScr(1) & "',kg_scr1 = '" & RecNew.kgScr(1) _
                       & "',init_scr2 = '" & RecNew.initScr(2) & "',kg_scr2 = '" & RecNew.kgScr(2) _
                       & "',init_scr3 = '" & RecNew.initScr(3) & "',kg_scr3 = '" & RecNew.kgScr(3) _
                       & "',init_scr4 = '" & RecNew.initScr(4) & "',kg_scr4 = '" & RecNew.kgScr(4) _
                       & "',init_wat1 = '" & RecNew.initWat(1) & "',kg_wat1 = '" & RecNew.kgWat(1) _
                       & "',init_wat2 = '" & RecNew.initWat(2) & "',kg_wat2 = '" & RecNew.kgWat(2) _
                       & "',init_chem1 = '" & RecNew.initChem(1) & "',kg_chem1 = '" & RecNew.kgChem(1) _
                       & "',init_chem2 = '" & RecNew.initChem(2) & "',kg_chem2 = '" & RecNew.kgChem(2) _
                       & "',init_chem3 = '" & RecNew.initChem(3) & "',kg_chem3 = '" & RecNew.kgChem(3) _
                       & "',init_chem4 = '" & RecNew.initChem(4) & "',kg_chem4 = '" & RecNew.kgChem(4) _
                       & "',init_chem5 = '" & RecNew.initChem(5) & "',kg_chem5 = '" & RecNew.kgChem(5) _
                       & "',init_chem6 = '" & RecNew.initChem(6) & "',kg_chem6 = '" & RecNew.kgChem(6) _
                       & "',kg_total = '" & RecNew.kgTotal _
                       & "' WHERE r_name = '" & RecNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing
                    DispPanel.txtRec.SetFocus

                    Exit Function

                End If

            ElseIf numCheck = Val(DispPanel.txtRec) And nameCheck1 = DispPanel.txtNameRec And nameCheck1 <> nameCheck Then
                'ако има такъв номер, но има и такова име на рецепта под друг номер
                'извеждаме съобщение за избор на ново име
                MousePointer = vbDefault
                MsgBox MsgNewName, vbOKOnly Or vbCritical, MsgErrBx
                
                'затваряме базата данни и прекратяваме функцията
                rs.Close
                Set rs = Nothing
                cn.Close
                Set cn = Nothing

                Exit Function

            End If

            'обновяваме ListView с рецептите
            DispPanel.lstRec.ListItems.Clear
            
            Set rs = cn.Execute("SELECT r_num FROM recepies ORDER BY r_num ASC;;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
            Do While Not rs.EOF
                RecCount = RecCount + 1

                If Val(rs!r_num) <> RecCount And frCheck = False Then
                    DispPanel.txtRec.Text = Format(RecCount, "0000")
                    frCheck = True
                Else
                    DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
                End If

                rs.MoveNext
            Loop
                
            Set rs = cn.Execute("SELECT * FROM recepies WHERE r_show = '1' ORDER BY r_num ASC;;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
            Do While Not rs.EOF
                RecCh.Code = Val(rs!r_num)
                RecCh.Title = rs!r_name
                RecCh.Kind = rs!r_type
                RecCh.Class = rs!r_class
                RecCh.ClassK = rs!r_classk
                RecCh.ClassV = rs!r_classv
                RecCh.ClassH = rs!r_classh
                RecCh.ClassP = rs!r_classp
                RecCh.EDM = Val(rs!r_edm)
                RecCh.Tpour = Val(rs!r_tpour)
                RecCh.Tmix = Val(rs!r_tmix)
                RecCh.initIM(1) = Val(rs!init_im1)
                RecCh.kgIM(1) = Val(rs!kg_im1)
                RecCh.initIM(2) = Val(rs!init_im2)
                RecCh.kgIM(2) = Val(rs!kg_im2)
                RecCh.initIM(3) = Val(rs!init_im3)
                RecCh.kgIM(3) = Val(rs!kg_im3)
                RecCh.initIM(4) = Val(rs!init_im4)
                RecCh.kgIM(4) = Val(rs!kg_im4)
                RecCh.initIM(5) = Val(rs!init_im5)
                RecCh.kgIM(5) = Val(rs!kg_im5)
                RecCh.initIM(6) = Val(rs!init_im6)
                RecCh.kgIM(6) = Val(rs!kg_im6)
                RecCh.initScr(1) = Val(rs!init_scr1)
                RecCh.kgScr(1) = Val(rs!kg_scr1)
                RecCh.initScr(2) = Val(rs!init_scr2)
                RecCh.kgScr(2) = Val(rs!kg_scr2)
                RecCh.initScr(3) = Val(rs!init_scr3)
                RecCh.kgScr(3) = Val(rs!kg_scr3)
                RecCh.initScr(4) = Val(rs!init_scr4)
                RecCh.kgScr(4) = Val(rs!kg_scr4)
                RecCh.initWat(1) = Val(rs!init_wat1)
                RecCh.kgWat(1) = Val(rs!kg_wat1)
                RecCh.initWat(2) = Val(rs!init_wat2)
                RecCh.kgWat(2) = Val(rs!kg_wat2)
                RecCh.initChem(1) = Val(rs!init_chem1)
                RecCh.kgChem(1) = CSng(rDs(rs!kg_chem1))
                RecCh.initChem(2) = Val(rs!init_chem2)
                RecCh.kgChem(2) = CSng(rDs(rs!kg_chem2))
                RecCh.initChem(3) = Val(rs!init_chem3)
                RecCh.kgChem(3) = CSng(rDs(rs!kg_chem3))
                RecCh.initChem(4) = Val(rs!init_chem4)
                RecCh.kgChem(4) = CSng(rDs(rs!kg_chem4))
                RecCh.initChem(5) = Val(rs!init_chem5)
                RecCh.kgChem(5) = CSng(rDs(rs!kg_chem5))
                RecCh.initChem(6) = Val(rs!init_chem6)
                RecCh.kgChem(6) = CSng(rDs(rs!kg_chem6))
        
                Set itmX = DispPanel.lstRec.ListItems.Add(1, , Format(RecCh.Code, "0000"))
                itmX.SubItems(1) = RecCh.Title
                itmX.SubItems(2) = RecCh.Kind
                itmX.SubItems(3) = RecCh.Class
                itmX.SubItems(4) = RecCh.ClassK
                itmX.SubItems(5) = RecCh.ClassV
                itmX.SubItems(6) = RecCh.ClassH
                itmX.SubItems(7) = RecCh.ClassP
                itmX.SubItems(8) = RecCh.EDM
                itmX.SubItems(9) = RecCh.Tpour
                itmX.SubItems(10) = RecCh.Tmix
        
                For n = 0 To ns1 - 1
                    itmX.SubItems(11 + n) = RecCh.AllkgIM(n + 1)
                Next n
        
                For n = 0 To ns3 - 1
                    itmX.SubItems(11 + ns1 + n) = RecCh.AllkgScr(n + 1)
                Next n
                
                For n = 0 To ns2 - 1
                    itmX.SubItems(11 + ns1 + ns3 + n) = RecCh.AllkgWat(n + 1)
                Next n
                
                For n = 0 To ns4 - 1
                    itmX.SubItems(11 + ns1 + ns3 + ns2 + n) = RecCh.AllkgChem(n + 1)
                Next n
        
                itmX.SubItems(11 + ns1 + ns3 + ns2 + ns4) = rDs(rs!kg_total)
                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------
        
            If DispPanel.lstRec.ListItems.count > 0 Then
                AutoColW DispPanel.lstRec
            Else
                DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
            End If
            
            DispPanel.txtRec.SetFocus
            DispPanel.txtNameRec.Text = ""
            DispPanel.txtTypeRec.Text = ""
            DispPanel.txtClassRec.Text = ""
            DispPanel.txtClassRecK.Text = ""
            DispPanel.txtClassRecV.Text = ""
            DispPanel.txtClassRecH.Text = ""
            DispPanel.txtClassRecP.Text = ""
            DispPanel.txtEDMRec.Text = 0
            DispPanel.txtTimePourRec.Text = TPd
            DispPanel.txtTimeMixRec.Text = TMd
            DispPanel.txtTotalKg.Text = 0
        
            For i = 0 To ns1 - 1
                DispPanel.txtRec1(i) = "0"
                DispPanel.cmbRec1(i).ListIndex = -1
                DispPanel.txtRec1(i).Refresh
            Next i
            
            For i = 0 To ns3 - 1
                DispPanel.txtRec3(i) = "0"
                DispPanel.cmbRec3(i).ListIndex = -1
                DispPanel.txtRec3(i).Refresh
            Next i
        
            For i = 0 To ns2 - 1
                DispPanel.txtRec2(i) = "0"
                DispPanel.cmbRec2(i).ListIndex = -1
                DispPanel.txtRec2(i).Refresh
            Next i
            
            For i = 0 To ns4 - 1
                DispPanel.txtRec4(i) = "0"
                DispPanel.cmbRec4(i).ListIndex = -1
                DispPanel.txtRec4(i).Refresh
            Next i

            MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave

            Exit Function

        Else
            MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function DelRecBut()
    'функция за изтриване на рецепта
    
    If DispPanel.frRecepies.Enabled = True And DispPanel.frRecepies.Visible = True Then
        If Len(DispPanel.txtRec.Text) > 0 And Len(DispPanel.txtNameRec.Text) > 0 And DispPanel.lstRec.ListItems.count > 0 Then

            If DispPanel.lstRec.SelectedItem.Text = DispPanel.txtRec.Text Then
                response = MsgBox(MsgConfDel, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then
            
                    '------------------------------Start PostgreSQL----------------------------------
                    Dim cn       As ADODB.Connection

                    Dim rs       As Recordset

                    Dim frCheck  As Boolean

                    Dim RecCount As Long

                    Dim RecD     As Recipe

                    Set RecD = New Recipe
    
                    frCheck = False
                    RecCount = 0
    
                    Set cn = New ADODB.Connection
                    cn.ConnectionTimeout = 10
                    cn.Open ConStr
                    MousePointer = vbHourglass
                
                    'зареждане на заявките активни към днешния ден
                    comm = "SELECT order_rec FROM orders WHERE stamp_date >= '" & DayToday & "';"
                    Set rs = cn.Execute(comm)
    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
                    Do While Not rs.EOF

                        If Val(DispPanel.txtRec.Text) = rs!order_rec Then
                            MousePointer = vbDefault
                            MsgBox MsgCantDelRec, vbOKOnly Or vbCritical, MsgErrBx
                            MousePointer = vbDefault
                            rs.Close
                            Set rs = Nothing
                            cn.Close
                            Set cn = Nothing

                            Exit Function

                        Else
                            rs.MoveNext
                        End If

                    Loop
                
                    Set rs = cn.Execute("DELETE FROM recepies WHERE r_num = " & Val(DispPanel.txtRec.Text) & ";") 'изтриване на запис
    
                    DispPanel.lstRec.ListItems.Clear
    
                    Set rs = cn.Execute("SELECT r_num FROM recepies ORDER BY r_num ASC;;")
    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                    Do While Not rs.EOF
                        RecCount = RecCount + 1

                        If Val(rs!r_num) <> RecCount And frCheck = False Then
                            DispPanel.txtRec.Text = Format(RecCount, "0000")
                            frCheck = True
                        Else
                            DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
                        End If

                        rs.MoveNext
                    Loop
                
                    Set rs = cn.Execute("SELECT * FROM recepies WHERE r_show = '1' ORDER BY r_num ASC;;")
    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                    Do While Not rs.EOF
                        RecD.Code = Val(rs!r_num)
                        RecD.Title = rs!r_name
                        RecD.Kind = rs!r_type
                        RecD.Class = rs!r_class
                        RecD.ClassK = rs!r_classk
                        RecD.ClassV = rs!r_classv
                        RecD.ClassH = rs!r_classh
                        RecD.ClassP = rs!r_classp
                        RecD.EDM = Val(rs!r_edm)
                        RecD.Tpour = Val(rs!r_tpour)
                        RecD.Tmix = Val(rs!r_tmix)
                        RecD.initIM(1) = Val(rs!init_im1)
                        RecD.kgIM(1) = Val(rs!kg_im1)
                        RecD.initIM(2) = Val(rs!init_im2)
                        RecD.kgIM(2) = Val(rs!kg_im2)
                        RecD.initIM(3) = Val(rs!init_im3)
                        RecD.kgIM(3) = Val(rs!kg_im3)
                        RecD.initIM(4) = Val(rs!init_im4)
                        RecD.kgIM(4) = Val(rs!kg_im4)
                        RecD.initIM(5) = Val(rs!init_im5)
                        RecD.kgIM(5) = Val(rs!kg_im5)
                        RecD.initIM(6) = Val(rs!init_im6)
                        RecD.kgIM(6) = Val(rs!kg_im6)
                        RecD.initScr(1) = Val(rs!init_scr1)
                        RecD.kgScr(1) = Val(rs!kg_scr1)
                        RecD.initScr(2) = Val(rs!init_scr2)
                        RecD.kgScr(2) = Val(rs!kg_scr2)
                        RecD.initScr(3) = Val(rs!init_scr3)
                        RecD.kgScr(3) = Val(rs!kg_scr3)
                        RecD.initScr(4) = Val(rs!init_scr4)
                        RecD.kgScr(4) = Val(rs!kg_scr4)
                        RecD.initWat(1) = Val(rs!init_wat1)
                        RecD.kgWat(1) = Val(rs!kg_wat1)
                        RecD.initWat(2) = Val(rs!init_wat2)
                        RecD.kgWat(2) = Val(rs!kg_wat2)
                        RecD.initChem(1) = Val(rs!init_chem1)
                        RecD.kgChem(1) = CSng(rDs(rs!kg_chem1))
                        RecD.initChem(2) = Val(rs!init_chem2)
                        RecD.kgChem(2) = CSng(rDs(rs!kg_chem2))
                        RecD.initChem(3) = Val(rs!init_chem3)
                        RecD.kgChem(3) = CSng(rDs(rs!kg_chem3))
                        RecD.initChem(4) = Val(rs!init_chem4)
                        RecD.kgChem(4) = CSng(rDs(rs!kg_chem4))
                        RecD.initChem(5) = Val(rs!init_chem5)
                        RecD.kgChem(5) = CSng(rDs(rs!kg_chem5))
                        RecD.initChem(6) = Val(rs!init_chem6)
                        RecD.kgChem(6) = CSng(rDs(rs!kg_chem6))
        
                        Set itmX = DispPanel.lstRec.ListItems.Add(1, , Format(RecD.Code, "0000"))
                        itmX.SubItems(1) = RecD.Title
                        itmX.SubItems(2) = RecD.Kind
                        itmX.SubItems(3) = RecD.Class
                        itmX.SubItems(4) = RecD.ClassK
                        itmX.SubItems(5) = RecD.ClassV
                        itmX.SubItems(6) = RecD.ClassH
                        itmX.SubItems(7) = RecD.ClassP
                        itmX.SubItems(8) = RecD.EDM
                        itmX.SubItems(9) = RecD.Tpour
                        itmX.SubItems(10) = RecD.Tmix
        
                        For n = 0 To ns1 - 1
                            itmX.SubItems(11 + n) = RecD.AllkgIM(n + 1)
                        Next n
        
                        For n = 0 To ns3 - 1
                            itmX.SubItems(11 + ns1 + n) = RecD.AllkgScr(n + 1)
                        Next n
                    
                        For n = 0 To ns2 - 1
                            itmX.SubItems(11 + ns1 + ns3 + n) = RecD.AllkgWat(n + 1)
                        Next n
                    
                        For n = 0 To ns4 - 1
                            itmX.SubItems(11 + ns1 + ns3 + ns2 + n) = RecD.AllkgChem(n + 1)
                        Next n
        
                        itmX.SubItems(11 + ns1 + ns3 + ns2 + ns4) = rDs(rs!kg_total)
                        rs.MoveNext
                    Loop
                
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    MousePointer = vbDefault
                    Set cn = Nothing
                    '--------------------------End PostgreSQL----------------------------------------
                
                    DispPanel.txtRec.SetFocus
                    DispPanel.txtNameRec.Text = ""
                    DispPanel.txtTypeRec.Text = ""
                    DispPanel.txtClassRec.Text = ""
                    DispPanel.txtClassRecK.Text = ""
                    DispPanel.txtClassRecV.Text = ""
                    DispPanel.txtClassRecH.Text = ""
                    DispPanel.txtClassRecP.Text = ""
                    DispPanel.txtEDMRec.Text = 0
                    DispPanel.txtTimePourRec.Text = TPd
                    DispPanel.txtTimeMixRec.Text = TMd
                    DispPanel.txtTotalKg.Text = 0
        
                    For i = 0 To ns1 - 1
                        DispPanel.txtRec1(i) = "0"
                        DispPanel.cmbRec1(i).ListIndex = -1
                    Next i
                
                    For i = 0 To ns3 - 1
                        DispPanel.txtRec3(i) = "0"
                        DispPanel.cmbRec3(i).ListIndex = -1
                    Next i
        
                    For i = 0 To ns2 - 1
                        DispPanel.txtRec2(i) = "0"
                        DispPanel.cmbRec2(i).ListIndex = -1
                    Next i
                
                    For i = 0 To ns4 - 1
                        DispPanel.txtRec4(i) = "0"
                        DispPanel.cmbRec4(i).ListIndex = -1
                    Next i

                    MsgBox MsgDelSuccess, vbOKOnly Or vbInformation, MsgDelBx
                
                    If DispPanel.lstRec.ListItems.count > 0 Then
                        AutoColW DispPanel.lstRec
                    Else
                        DispPanel.txtRec.Text = Format(RecCount + RecMin, "0000")
                    End If

                Else

                    Exit Function

                End If

            Else
                MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
            End If

        Else
            MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function ListRecClick()
    'функция за зареждане на данни при маркиране на запис от таблицата
    
    If DispPanel.frRecepies.Enabled = True And DispPanel.frRecepies.Visible = True Then
        If DispPanel.lstRec.ListItems.count > 0 Then
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn    As ADODB.Connection

            Dim rs    As Recordset
            
            Dim RecSh As Recipe

            Set RecSh = New Recipe
            
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT * FROM recepies WHERE r_num = " & Val(DispPanel.lstRec.ListItems(DispPanel.lstRec.SelectedItem.Index).Text) & ";")
    
            Do While Not rs.EOF
                RecSh.Code = Val(rs!r_num)
                RecSh.Title = rs!r_name
                RecSh.Kind = rs!r_type
                RecSh.Class = rs!r_class
                RecSh.ClassK = rs!r_classk
                RecSh.ClassV = rs!r_classv
                RecSh.ClassH = rs!r_classh
                RecSh.ClassP = rs!r_classp
                RecSh.EDM = Val(rs!r_edm)
                RecSh.Tpour = Val(rs!r_tpour)
                RecSh.Tmix = Val(rs!r_tmix)
                RecSh.initIM(1) = Val(rs!init_im1)
                RecSh.kgIM(1) = Val(rs!kg_im1)
                RecSh.initIM(2) = Val(rs!init_im2)
                RecSh.kgIM(2) = Val(rs!kg_im2)
                RecSh.initIM(3) = Val(rs!init_im3)
                RecSh.kgIM(3) = Val(rs!kg_im3)
                RecSh.initIM(4) = Val(rs!init_im4)
                RecSh.kgIM(4) = Val(rs!kg_im4)
                RecSh.initIM(5) = Val(rs!init_im5)
                RecSh.kgIM(5) = Val(rs!kg_im5)
                RecSh.initIM(6) = Val(rs!init_im6)
                RecSh.kgIM(6) = Val(rs!kg_im6)
                RecSh.initScr(1) = Val(rs!init_scr1)
                RecSh.kgScr(1) = Val(rs!kg_scr1)
                RecSh.initScr(2) = Val(rs!init_scr2)
                RecSh.kgScr(2) = Val(rs!kg_scr2)
                RecSh.initScr(3) = Val(rs!init_scr3)
                RecSh.kgScr(3) = Val(rs!kg_scr3)
                RecSh.initScr(4) = Val(rs!init_scr4)
                RecSh.kgScr(4) = Val(rs!kg_scr4)
                RecSh.initWat(1) = Val(rs!init_wat1)
                RecSh.kgWat(1) = Val(rs!kg_wat1)
                RecSh.initWat(2) = Val(rs!init_wat2)
                RecSh.kgWat(2) = Val(rs!kg_wat2)
                RecSh.initChem(1) = Val(rs!init_chem1)
                RecSh.kgChem(1) = CSng(rDs(rs!kg_chem1))
                RecSh.initChem(2) = Val(rs!init_chem2)
                RecSh.kgChem(2) = CSng(rDs(rs!kg_chem2))
                RecSh.initChem(3) = Val(rs!init_chem3)
                RecSh.kgChem(3) = CSng(rDs(rs!kg_chem3))
                RecSh.initChem(4) = Val(rs!init_chem4)
                RecSh.kgChem(4) = CSng(rDs(rs!kg_chem4))
                RecSh.initChem(5) = Val(rs!init_chem5)
                RecSh.kgChem(5) = CSng(rDs(rs!kg_chem5))
                RecSh.initChem(6) = Val(rs!init_chem6)
                RecSh.kgChem(6) = CSng(rDs(rs!kg_chem6))
                RecSh.kgTotal = CSng(rDs(rs!kg_total))
                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------
                    
            DispPanel.txtRec.Text = Format(RecSh.Code, "0000")
            DispPanel.txtNameRec.Text = RecSh.Title
            DispPanel.txtTypeRec.Text = RecSh.Kind
            DispPanel.txtClassRec.Text = RecSh.Class
            DispPanel.txtClassRecK.Text = RecSh.ClassK
            DispPanel.txtClassRecV.Text = RecSh.ClassV
            DispPanel.txtClassRecH.Text = RecSh.ClassH
            DispPanel.txtClassRecP.Text = RecSh.ClassP
            DispPanel.txtEDMRec.Text = RecSh.EDM
            DispPanel.txtTimePourRec.Text = RecSh.Tpour
            DispPanel.txtTimeMixRec.Text = RecSh.Tmix
            DispPanel.txtTotalKg.Text = RecSh.kgTotal
            
            For i = 0 To ns1 - 1
                DispPanel.txtRec1(i) = RecSh.kgIM(i + 1)
            Next i
                
            For n = 0 To ns1 - 1
                For i = 0 To ns1 - 1

                    If RecSh.initIM(n + 1) = BaseIM + i Then
                        If i + 1 < DispPanel.cmbRec1(n).listCount Then
                            DispPanel.cmbRec1(n).ListIndex = i + 1
                        Else
                            MsgBox MsgRecErrIM, vbOKOnly Or vbCritical, MsgErrBx

                            Exit Function

                        End If
                    End If

                    If RecSh.initIM(n + 1) = 0 Then
                        DispPanel.cmbRec1(n).ListIndex = 0
                    End If

                Next i
            Next n
                
            For i = 0 To ns3 - 1
                DispPanel.txtRec3(i) = RecSh.kgScr(i + 1)
            Next i
                
            For n = 0 To ns3 - 1
                For i = 0 To ns3 - 1

                    If RecSh.initScr(n + 1) = BaseScr + i Then
                        If i + 1 < DispPanel.cmbRec3(n).listCount Then
                            DispPanel.cmbRec3(n).ListIndex = i + 1
                        Else
                            MsgBox MsgRecErrCem, vbOKOnly Or vbCritical, MsgErrBx

                            Exit Function

                        End If
                    End If

                    If RecSh.initScr(n + 1) = 0 Then
                        DispPanel.cmbRec3(n).ListIndex = 0
                    End If

                Next i
            Next n
                
            For i = 0 To ns2 - 1
                DispPanel.txtRec2(i) = RecSh.kgWat(i + 1)
            Next i
                
            For n = 0 To ns2 - 1
                For i = 0 To ns2 - 1

                    If RecSh.initWat(n + 1) = BaseWat + i Then
                        If i + 1 < DispPanel.cmbRec2(n).listCount Then
                            DispPanel.cmbRec2(n).ListIndex = i + 1
                        Else
                            MsgBox MsgRecErrWat, vbOKOnly Or vbCritical, MsgErrBx

                            Exit Function

                        End If
                    End If

                    If RecSh.initWat(n + 1) = 0 Then
                        DispPanel.cmbRec2(n).ListIndex = 0
                    End If

                Next i
            Next n
                        
            For i = 0 To ns4 - 1
                DispPanel.txtRec4(i) = RecSh.kgChem(i + 1)
            Next i
                
            For n = 0 To ns4 - 1
                For i = 0 To ns4 - 1

                    If RecSh.initChem(n + 1) = BaseChem + i Then
                        If i + 1 < DispPanel.cmbRec4(n).listCount Then
                            DispPanel.cmbRec4(n).ListIndex = i + 1
                        Else
                            MsgBox MsgRecErrChem, vbOKOnly Or vbCritical, MsgErrBx

                            Exit Function

                        End If
                    End If

                    If RecSh.initChem(n + 1) = 0 Then
                        DispPanel.cmbRec4(n).ListIndex = 0
                    End If

                Next i
            Next n

        Else
        End If

    Else
    End If

End Function

Public Function OpenClients()
    'зареждане на меню клиенти
    
    Dim itmX As ListItem
    
    DispPanel.txtClnt.SetFocus
    DispPanel.txtClnt.TabIndex = 0
    DispPanel.txtNameClnt.TabIndex = 1
    DispPanel.txtBGClnt.TabIndex = 2
    DispPanel.txtMOLClnt.TabIndex = 3
    DispPanel.txtAddClnt.TabIndex = 4
    DispPanel.txtTelClnt.TabIndex = 5
    DispPanel.btnClearClnt.TabIndex = 6
    DispPanel.btnSvNwClnt.TabIndex = 7
    DispPanel.btnDelClnt.TabIndex = 8
    DispPanel.btnShowClnt.TabIndex = 9
    DispPanel.btnObjects.TabIndex = 10
    DispPanel.btnDelObj.TabIndex = 11
    DispPanel.btnShowObj.TabIndex = 12
    DispPanel.btnDisp.TabIndex = 13
    DispPanel.btnOrders.TabIndex = 14
    DispPanel.btnRecepies.TabIndex = 15
    DispPanel.btnClients.TabIndex = 16
    DispPanel.btnDrivers.TabIndex = 17
    DispPanel.btnSuppliers.TabIndex = 18
    DispPanel.btnMaterials.TabIndex = 19
    DispPanel.btnNotes.TabIndex = 20
    DispPanel.btnAdminPanel.TabIndex = 21
    DispPanel.btnExit.TabIndex = 22
    
    DispPanel.frClients.Left = 120
    DispPanel.frClients.Top = DispPanel.Height \ 2 - FrmPlace
    
    DispPanel.lstClnt.ColumnHeaders.Clear
    DispPanel.lstClnt.ListItems.Clear
    DispPanel.lstObj.ColumnHeaders.Clear
    DispPanel.lstObj.ListItems.Clear
    
    Set colx = DispPanel.lstClnt.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 800
    
    Set colx = DispPanel.lstClnt.ColumnHeaders.Add()
    colx.Text = uniFirm
    colx.Width = 2000
    
    Set colx = DispPanel.lstClnt.ColumnHeaders.Add()
    colx.Text = uniBG
    colx.Width = 1400
    
    Set colx = DispPanel.lstClnt.ColumnHeaders.Add()
    colx.Text = uniMOL
    colx.Width = 2000

    Set colx = DispPanel.lstClnt.ColumnHeaders.Add()
    colx.Text = uniAdd
    colx.Width = 2000

    Set colx = DispPanel.lstClnt.ColumnHeaders.Add()
    colx.Text = uniTel
    colx.Width = 1300
    
    Set colx = DispPanel.lstObj.ColumnHeaders.Add()
    colx.Text = uniObj
    colx.Width = 3000

    Set colx = DispPanel.lstObj.ColumnHeaders.Add()
    colx.Text = uniKm
    colx.Width = 800
    
    DispPanel.txtClnt.MaxLength = 4
    DispPanel.txtClnt.Text = ""
    DispPanel.txtNameClnt.MaxLength = 70
    DispPanel.txtNameClnt.Text = ""
    DispPanel.txtBGClnt.MaxLength = 15
    DispPanel.txtBGClnt.Text = ""
    DispPanel.txtMOLClnt.MaxLength = 60
    DispPanel.txtMOLClnt.Text = ""
    DispPanel.txtAddClnt.MaxLength = 100
    DispPanel.txtAddClnt.Text = ""
    DispPanel.txtTelClnt.MaxLength = 15
    DispPanel.txtTelClnt.Text = ""
    
    '------------------------------Start PostgreSQL----------------------------------
    Dim cn        As ADODB.Connection

    Dim rs        As Recordset

    Dim ClntCount As Long

    Dim frCheck   As Boolean
    
    ClntCount = 0
    frCheck = False
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT c_num FROM clients ORDER BY c_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        ClntCount = ClntCount + 1

        If Val(rs!c_num) <> ClntCount And frCheck = False Then
            DispPanel.txtClnt.Text = Format(ClntCount, "0000")
            frCheck = True
        Else
            DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
        End If

        rs.MoveNext
    Loop
        
    Set rs = cn.Execute("SELECT * FROM clients WHERE c_show = '1' ORDER BY c_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Set itmX = DispPanel.lstClnt.ListItems.Add(1, , Format(rs!c_num, "0000"))
        itmX.SubItems(1) = rs!c_name
        itmX.SubItems(2) = rs!c_bg
        itmX.SubItems(3) = rs!c_mol
        itmX.SubItems(4) = rs!c_add
        itmX.SubItems(5) = rs!c_tel
        rs.MoveNext
    Loop
    
    Set rs = cn.Execute("SELECT w_name, w_km FROM worksites WHERE w_cnum = '" & Val(DispPanel.txtClnt.Text) & "' ORDER BY w_name;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Set itmX = DispPanel.lstObj.ListItems.Add(1, , rs!w_name)
        itmX.SubItems(1) = rs!w_km
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
    '--------------------------End PostgreSQL------------------------------------------
    
    DispPanel.btnObjects.Enabled = False
    
    If DispPanel.lstClnt.ListItems.count > 0 Then
        AutoColW DispPanel.lstClnt
    Else
        DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
    End If

End Function

Public Function ClearClntBut()
    'функция за почистване на клетките на клиента за въвеждане на нова

    If DispPanel.frClients.Enabled = True And DispPanel.frClients.Visible = True Then
        If DispPanel.lstClnt.ListItems.count > 0 Then
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn        As ADODB.Connection

            Dim rs        As Recordset

            Dim ClntCount As Long

            Dim frCheck   As Boolean
    
            ClntCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT c_num FROM clients ORDER BY c_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                ClntCount = ClntCount + 1

                If Val(rs!c_num) <> ClntCount And frCheck = False Then
                    DispPanel.txtClnt.Text = Format(ClntCount, "0000")
                    frCheck = True
                Else
                    DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
                End If

                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------

        Else
            DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
        End If

        DispPanel.txtNameClnt.Text = ""
        DispPanel.txtBGClnt.Text = ""
        DispPanel.txtMOLClnt.Text = ""
        DispPanel.txtAddClnt.Text = ""
        DispPanel.txtTelClnt.Text = ""
        DispPanel.lstObj.ListItems.Clear
        DispPanel.btnObjects.Enabled = False
    Else
    End If

End Function

Public Function SvNwClntBut()
    'функция за запис на клиент

    If DispPanel.frClients.Enabled = True And DispPanel.frClients.Visible = True Then
        If Val(DispPanel.txtClnt.Text) = 0 Then
            MsgBox MsgCodeZero, vbOKOnly Or vbCritical, MsgErrBx

            Exit Function

        Else
        End If
        
        Dim ClntNew As Client

        Set ClntNew = New Client
        
        Dim response As Integer

        If Len(DispPanel.txtClnt.Text) > 0 And Len(DispPanel.txtNameClnt.Text) > 0 Then
            ClntNew.Code = DispPanel.txtClnt.Text
            ClntNew.Title = DispPanel.txtNameClnt.Text
            ClntNew.Ident = DispPanel.txtBGClnt.Text
            ClntNew.MOL = DispPanel.txtMOLClnt.Text
            ClntNew.Address = DispPanel.txtAddClnt.Text
            ClntNew.Phone = DispPanel.txtTelClnt.Text
            
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn         As ADODB.Connection

            Dim rs         As Recordset

            Dim rs1        As Recordset

            Dim comIns     As String

            Dim comEdit    As String

            Dim numCheck   As Long

            Dim nameCheck  As String

            Dim nameCheck1 As String

            Dim ClntCount  As Long

            Dim frCheck    As Boolean
    
            ClntCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT c_num, c_name FROM clients WHERE c_num = " & Val(DispPanel.txtClnt) & ";") 'клиент код ако има

            If Not rs.BOF And Not rs.EOF Then
                numCheck = rs!c_num
                nameCheck = rs!c_name
            End If
            
            Set rs1 = cn.Execute("SELECT c_name FROM clients WHERE c_name = '" & DispPanel.txtNameClnt & "';") 'клиент име ако има

            If Not rs1.BOF And Not rs1.EOF Then nameCheck1 = rs1!c_name
            
            If numCheck <> Val(DispPanel.txtClnt) And nameCheck <> DispPanel.txtNameClnt And nameCheck1 <> DispPanel.txtNameClnt Then
                'ако няма съвпадения в номер или име на клиент правим запис
                comIns = "INSERT INTO clients VALUES(" & ClntNew.Code & ",'" & ClntNew.Title & "','" & ClntNew.Ident & "','" & ClntNew.MOL & "','" & ClntNew.Address & "','" & ClntNew.Phone & "', 'true')"
                Set rs = cn.Execute(comIns)
            ElseIf numCheck = Val(DispPanel.txtClnt) And nameCheck1 <> DispPanel.txtNameClnt Then
                'ако има такъв номер на клиент, но въведеното име е друго - питаме за редакция по номер
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE clients SET c_name = '" & ClntNew.Title & "',c_bg = '" & ClntNew.Ident & "',c_mol = '" & ClntNew.MOL & "',c_add = '" & ClntNew.Address & "',c_tel = '" & ClntNew.Phone & "' WHERE c_num =" & ClntNew.Code & "" 'корекция по номер
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    rs1.Close
                    Set rs1 = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf nameCheck = DispPanel.txtNameClnt And numCheck = Val(DispPanel.txtClnt) Then
                'ако има такъв номер и името му е същото - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE clients SET c_bg = '" & ClntNew.Ident & "',c_mol = '" & ClntNew.MOL & "',c_add = '" & ClntNew.Address & "',c_tel = '" & ClntNew.Phone & "' WHERE c_name ='" & ClntNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    rs1.Close
                    Set rs1 = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf numCheck <> Val(DispPanel.txtClnt) And nameCheck1 = DispPanel.txtNameClnt Then
                'ако няма такъв номер, но има такова име на клиент - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    
                    Dim rs2 As Recordset

                    Dim rs3 As Recordset

                    Dim wCh As String
                    
                    Set rs2 = cn.Execute("SELECT c_num FROM clients WHERE c_name = '" & DispPanel.txtNameClnt & "';")

                    If Not rs2.BOF Or Not rs2.EOF Then wCh = rs2!c_num
                    
                    comEdit = "UPDATE clients SET c_num = " & ClntNew.Code & ",c_bg = '" & ClntNew.Ident & "',c_mol = '" & ClntNew.MOL & "',c_add = '" & ClntNew.Address & "',c_tel = '" & ClntNew.Phone & "' WHERE c_name ='" & ClntNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                    
                    comEdit = "UPDATE worksites SET w_cnum = '" & ClntNew.Code & "' WHERE w_cnum = '" & wCh & "'"
                    Set rs3 = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    rs1.Close
                    Set rs1 = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf numCheck = Val(DispPanel.txtDrv) And nameCheck1 = DispPanel.txtNameDrv And nameCheck1 <> nameCheck Then
                'ако има такъв номер, но има и такова име на водач под друг номер
                'извеждаме съобщение за избор на ново име
                MousePointer = vbDefault
                MsgBox MsgNewName, vbOKOnly Or vbCritical, MsgErrBx
                
                'затваряме базата данни и прекратяваме функцията
                rs.Close
                Set rs = Nothing
                rs1.Close
                Set rs1 = Nothing
                cn.Close
                Set cn = Nothing
                DispPanel.txtDrv.SetFocus

                Exit Function

            End If
            
            MousePointer = vbDefault
            MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
    
            DispPanel.lstClnt.ListItems.Clear
            DispPanel.lstObj.ListItems.Clear
            DispPanel.btnObjects.Enabled = False
    
            Set rs = cn.Execute("SELECT c_num FROM clients ORDER BY c_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                ClntCount = ClntCount + 1

                If Val(rs!c_num) <> ClntCount And frCheck = False Then
                    DispPanel.txtClnt.Text = Format(ClntCount, "0000")
                    frCheck = True
                Else
                    DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
                End If

                rs.MoveNext
            Loop
                
            Set rs = cn.Execute("SELECT * FROM clients WHERE c_show = '1' ORDER BY c_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                Set itmX = DispPanel.lstClnt.ListItems.Add(1, , Format(rs!c_num, "0000"))
                itmX.SubItems(1) = rs!c_name
                itmX.SubItems(2) = rs!c_bg
                itmX.SubItems(3) = rs!c_mol
                itmX.SubItems(4) = rs!c_add
                itmX.SubItems(5) = rs!c_tel
                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------
            
            If DispPanel.lstClnt.ListItems.count > 0 Then
                AutoColW DispPanel.lstClnt
            Else
                DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
            End If

        Else
            MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function DelClntBut()
    'функция за изтриване на клиент

    If DispPanel.frClients.Enabled = True And DispPanel.frClients.Visible = True Then
        If Len(DispPanel.txtClnt.Text) > 0 And Len(DispPanel.txtNameClnt.Text) > 0 And DispPanel.lstClnt.ListItems.count > 0 Then

            If DispPanel.lstClnt.SelectedItem.Text = DispPanel.txtClnt.Text Then
                response = MsgBox(MsgConfDel, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then
            
                    '------------------------------Start PostgreSQL----------------------------------
                    Dim cn        As ADODB.Connection

                    Dim rs        As Recordset

                    Dim ClntCount As Long

                    Dim frCheck   As Boolean
    
                    ClntCount = 0
                    frCheck = False
    
                    Set cn = New ADODB.Connection
                    cn.ConnectionTimeout = 10
                    cn.Open ConStr
                    MousePointer = vbHourglass
                
                    'зареждане на заявките активни към днешния ден
                    comm = "SELECT order_clnt FROM orders WHERE stamp_date >= '" & DayToday & "';"
                    Set rs = cn.Execute(comm)
    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
                    Do While Not rs.EOF

                        If Val(DispPanel.txtClnt.Text) = rs!order_clnt Then
                            MousePointer = vbDefault
                            MsgBox MsgCantDelClnt, vbOKOnly Or vbCritical, MsgErrBx
                            MousePointer = vbDefault
                            rs.Close
                            Set rs = Nothing
                            cn.Close
                            Set cn = Nothing

                            Exit Function

                        Else
                            rs.MoveNext
                        End If

                    Loop
                
                    Set rs = cn.Execute("DELETE FROM clients WHERE c_num = " & Val(DispPanel.txtClnt.Text) & ";") 'изтриване на клиента
                    Set rs = cn.Execute("DELETE FROM worksites WHERE w_cnum = '" & Val(DispPanel.txtClnt.Text) & "';") 'изтриване на обектите му
    
                    DispPanel.lstClnt.ListItems.Clear
                    DispPanel.lstObj.ListItems.Clear
    
                    Set rs = cn.Execute("SELECT c_num FROM clients ORDER BY c_num ASC;")
    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                    Do While Not rs.EOF
                        ClntCount = ClntCount + 1

                        If Val(rs!c_num) <> ClntCount And frCheck = False Then
                            DispPanel.txtClnt.Text = Format(ClntCount, "0000")
                            frCheck = True
                        Else
                            DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
                        End If

                        rs.MoveNext
                    Loop
                    
                    Set rs = cn.Execute("SELECT * FROM clients WHERE c_show = '1' ORDER BY c_num ASC;")
    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                    Do While Not rs.EOF
                        Set itmX = DispPanel.lstClnt.ListItems.Add(1, , Format(rs!c_num, "000000"))
                        itmX.SubItems(1) = rs!c_name
                        itmX.SubItems(2) = rs!c_bg
                        itmX.SubItems(3) = rs!c_mol
                        itmX.SubItems(4) = rs!c_add
                        itmX.SubItems(5) = rs!c_tel
                        rs.MoveNext
                    Loop
                
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    MousePointer = vbDefault
                    Set cn = Nothing
                    '--------------------------End PostgreSQL------------------------------------------
            
                    DispPanel.txtNameClnt.Text = ""
                    DispPanel.txtBGClnt.Text = ""
                    DispPanel.txtMOLClnt.Text = ""
                    DispPanel.txtAddClnt.Text = ""
                    DispPanel.txtTelClnt.Text = ""
            
                    If DispPanel.lstClnt.ListItems.count > 0 Then
                        AutoColW DispPanel.lstClnt
                    Else
                        DispPanel.txtClnt.Text = Format(ClntCount + 1, "0000")
                    End If

                Else

                    Exit Function

                End If

            Else
                MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
            End If

        Else
            MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function ListClntClick()
    'функция за зареждане на данни при маркиране на запис от таблицата
    
    If DispPanel.frClients.Enabled = True And DispPanel.frClients.Visible = True Then
        If DispPanel.lstClnt.ListItems.count > 0 Then
            DispPanel.txtClnt.Text = Format(Val(DispPanel.lstClnt.ListItems(DispPanel.lstClnt.SelectedItem.Index).Text), "0000")
            DispPanel.txtNameClnt.Text = DispPanel.lstClnt.ListItems(DispPanel.lstClnt.SelectedItem.Index).ListSubItems(1).Text
            DispPanel.txtBGClnt.Text = DispPanel.lstClnt.ListItems(DispPanel.lstClnt.SelectedItem.Index).ListSubItems(2).Text
            DispPanel.txtMOLClnt.Text = DispPanel.lstClnt.ListItems(DispPanel.lstClnt.SelectedItem.Index).ListSubItems(3).Text
            DispPanel.txtAddClnt.Text = DispPanel.lstClnt.ListItems(DispPanel.lstClnt.SelectedItem.Index).ListSubItems(4).Text
            DispPanel.txtTelClnt.Text = DispPanel.lstClnt.ListItems(DispPanel.lstClnt.SelectedItem.Index).ListSubItems(5).Text
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn As ADODB.Connection

            Dim rs As Recordset
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
    
            DispPanel.lstObj.ListItems.Clear
    
            Set rs = cn.Execute("SELECT w_name, w_km FROM worksites WHERE w_cnum = '" & Val(DispPanel.txtClnt.Text) & "' AND w_show = 'true' ORDER BY w_name;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
            Do While Not rs.EOF
                Set itmX = DispPanel.lstObj.ListItems.Add(1, , rs!w_name)
                itmX.SubItems(1) = rs!w_km
                rs.MoveNext
            Loop
    
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------

            AutoColW DispPanel.lstObj
            DispPanel.btnObjects.Enabled = True
        End If

    Else
    End If

End Function

Public Function DelObjBut()

    'функция за изтриване на Обект
    Dim response As Integer
    
    If DispPanel.frClients.Enabled = True And DispPanel.frClients.Visible = True And DispPanel.lstObj.ListItems.count > 0 Then
        If Len(DispPanel.txtClnt.Text) > 0 And Len(DispPanel.txtNameClnt.Text) > 0 And DispPanel.lstObj.SelectedItem.Text <> "" Then
            response = MsgBox(MsgConfDel, vbYesNo Or vbQuestion, MsgEditBx)

            If response = vbYes Then
            
                '------------------------------Start PostgreSQL----------------------------------
                Dim cn        As ADODB.Connection

                Dim rs        As Recordset

                Dim comm      As String
                
                Set cn = New ADODB.Connection
                cn.ConnectionTimeout = 10
                cn.Open ConStr
                MousePointer = vbHourglass
                
                'зареждане на заявките активни към днешния ден
                comm = "SELECT order_clnt, order_clnt_obj FROM orders WHERE stamp_date >= '" & DayToday & "';"
                Set rs = cn.Execute(comm)
    
                If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
                Do While Not rs.EOF

                    If Val(DispPanel.txtClnt.Text) = rs!order_clnt And DispPanel.lstObj.SelectedItem.Text = rs!order_clnt_obj Then
                        MousePointer = vbDefault
                        MsgBox MsgCantDelClnt, vbOKOnly Or vbCritical, MsgErrBx
                        MousePointer = vbDefault
                        rs.Close
                        Set rs = Nothing
                        cn.Close
                        Set cn = Nothing

                        Exit Function

                    Else
                        rs.MoveNext
                    End If

                Loop
                
                Set rs = cn.Execute("DELETE FROM worksites WHERE w_cnum = '" & Val(DispPanel.txtClnt.Text) & "' AND w_name = '" & DispPanel.lstObj.SelectedItem.Text & "';") 'изтриване на обектите му
    
                DispPanel.lstObj.ListItems.Clear
    
                Set rs = cn.Execute("SELECT w_name, w_km FROM worksites WHERE w_cnum = '" & Val(DispPanel.txtClnt.Text) & "' ORDER BY w_name ASC;")
    
                If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                Do While Not rs.EOF
                    Set itmX = DispPanel.lstObj.ListItems.Add(1, , rs!w_name)
                    itmX.SubItems(1) = rs!w_km
                    rs.MoveNext
                Loop
                
                rs.Close
                Set rs = Nothing
                cn.Close
                MousePointer = vbDefault
                Set cn = Nothing
                '--------------------------End PostgreSQL------------------------------------------
            
                If DispPanel.lstObj.ListItems.count > 0 Then AutoColW DispPanel.lstObj

            Else

                Exit Function

            End If

        Else
            MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function OpenDrivers()
    'зареждане на меню водачи

    Dim itmX As ListItem
    
    DispPanel.txtDrv.SetFocus
    DispPanel.txtDrv.TabIndex = 0
    DispPanel.txtNameDrv.TabIndex = 1
    DispPanel.txtRegDrv.TabIndex = 2
    DispPanel.txtModDrv.TabIndex = 3
    DispPanel.txtCapDrv.TabIndex = 4
    DispPanel.txtTelDrv.TabIndex = 5
    DispPanel.txtNoteDrv.TabIndex = 6
    DispPanel.btnClearDrv.TabIndex = 7
    DispPanel.btnSvNwDrv.TabIndex = 8
    DispPanel.btnDelDrv.TabIndex = 9
    DispPanel.btnShowDrv.TabIndex = 10
    DispPanel.btnDisp.TabIndex = 11
    DispPanel.btnOrders.TabIndex = 12
    DispPanel.btnRecepies.TabIndex = 13
    DispPanel.btnClients.TabIndex = 14
    DispPanel.btnDrivers.TabIndex = 15
    DispPanel.btnSuppliers.TabIndex = 16
    DispPanel.btnMaterials.TabIndex = 17
    DispPanel.btnNotes.TabIndex = 18
    DispPanel.btnAdminPanel.TabIndex = 19
    DispPanel.btnExit.TabIndex = 20
    
    DispPanel.frDrivers.Left = 120
    DispPanel.frDrivers.Top = DispPanel.Height \ 2 - FrmPlace
    
    DispPanel.lstDrv.ColumnHeaders.Clear
    DispPanel.lstDrv.ListItems.Clear
    
    Set colx = DispPanel.lstDrv.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 700
    
    Set colx = DispPanel.lstDrv.ColumnHeaders.Add()
    colx.Text = uniNm
    colx.Width = 4000
    
    Set colx = DispPanel.lstDrv.ColumnHeaders.Add()
    colx.Text = uniDrvReg
    colx.Width = 1200

    Set colx = DispPanel.lstDrv.ColumnHeaders.Add()
    colx.Text = uniCapacity
    colx.Width = 1000

    Set colx = DispPanel.lstDrv.ColumnHeaders.Add()
    colx.Text = uniMod
    colx.Width = 1500
    
    Set colx = DispPanel.lstDrv.ColumnHeaders.Add()
    colx.Text = uniTel
    colx.Width = 1300
    
    Set colx = DispPanel.lstDrv.ColumnHeaders.Add()
    colx.Text = uniNote
    colx.Width = 3000

    DispPanel.txtDrv.MaxLength = 4
    DispPanel.txtDrv.Text = ""
    DispPanel.txtNameDrv.MaxLength = 60
    DispPanel.txtNameDrv.Text = ""
    DispPanel.txtRegDrv.MaxLength = 10
    DispPanel.txtRegDrv.Text = ""
    DispPanel.txtCapDrv.MaxLength = 5
    DispPanel.txtCapDrv.Text = 0
    DispPanel.txtModDrv.MaxLength = 30
    DispPanel.txtModDrv.Text = ""
    DispPanel.txtTelDrv.MaxLength = 15
    DispPanel.txtTelDrv.Text = ""
    DispPanel.txtNoteDrv.MaxLength = 100
    DispPanel.txtNoteDrv.Text = ""
    
    '------------------------------Start PostgreSQL----------------------------------
    Dim cn       As ADODB.Connection

    Dim rs       As Recordset

    Dim DrvCount As Long

    Dim frCheck  As Boolean
    
    DrvCount = 0
    frCheck = False
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT d_num FROM drivers ORDER BY d_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        DrvCount = DrvCount + 1

        If Val(rs!d_num) <> DrvCount And frCheck = False Then
            DispPanel.txtDrv.Text = Format(DrvCount, "0000")
            frCheck = True
        Else
            DispPanel.txtDrv.Text = Format(DrvCount + 1, "0000")
        End If

        rs.MoveNext
    Loop
    
    Set rs = cn.Execute("SELECT * FROM drivers WHERE d_show = '1' ORDER BY d_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Set itmX = DispPanel.lstDrv.ListItems.Add(1, , Format(rs!d_num, "0000"))
        itmX.SubItems(1) = rs!d_name
        itmX.SubItems(2) = rs!d_reg
        itmX.SubItems(3) = CSng(rDs(rs!d_cap))
        itmX.SubItems(4) = rs!d_mod
        itmX.SubItems(5) = rs!d_tel
        itmX.SubItems(6) = rs!d_note
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
    '--------------------------End PostgreSQL------------------------------------------

    If DispPanel.lstDrv.ListItems.count > 0 Then
        AutoColW DispPanel.lstDrv
    Else
        DispPanel.txtDrv.Text = Format(DrvCount + 1, "0000")
    End If

End Function

Public Function ClearDrvBut()
    'функция за почистване на клетките на водача за въвеждане на нова

    If DispPanel.frDrivers.Enabled = True And DispPanel.frDrivers.Visible = True Then
        If DispPanel.lstDrv.ListItems.count > 0 Then
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn       As ADODB.Connection

            Dim rs       As Recordset

            Dim DrvCount As Long

            Dim frCheck  As Boolean
    
            DrvCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT d_num FROM drivers ORDER BY d_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                DrvCount = DrvCount + 1

                If Val(rs!d_num) <> DrvCount And frCheck = False Then
                    DispPanel.txtDrv.Text = Format(DrvCount, "0000")
                    frCheck = True
                Else
                    DispPanel.txtDrv.Text = Format(DrvCount + 1, "0000")
                End If

                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------

        Else
            DispPanel.txtDrv.Text = Format(DrvCount + 1, "0000")
        End If

        DispPanel.txtNameDrv.Text = ""
        DispPanel.txtRegDrv.Text = ""
        DispPanel.txtCapDrv.Text = 0
        DispPanel.txtModDrv.Text = ""
        DispPanel.txtTelDrv.Text = ""
        DispPanel.txtNoteDrv.Text = ""
    Else
    End If

End Function

Public Function SvNwDrvBut()
    'функция за запис на водач

    If DispPanel.frDrivers.Enabled = True And DispPanel.frDrivers.Visible = True Then
        If Val(DispPanel.txtDrv.Text) = 0 Then
            MsgBox MsgCodeZero, vbOKOnly Or vbCritical, MsgErrBx

            Exit Function

        Else
        End If
        
        Dim DrvNew As Driver

        Set DrvNew = New Driver
        
        Dim response As Integer

        If Len(DispPanel.txtDrv.Text) > 0 And Len(DispPanel.txtNameDrv.Text) > 0 And Len(DispPanel.txtRegDrv.Text) > 0 And Len(DispPanel.txtCapDrv.Text) > 0 Then
            DrvNew.Code = Val(DispPanel.txtDrv.Text)
            DrvNew.Title = DispPanel.txtNameDrv.Text
            DrvNew.CarNum = DispPanel.txtRegDrv.Text
            DrvNew.Capacity = CSng(rDs(DispPanel.txtCapDrv.Text))
            DrvNew.CarModel = DispPanel.txtModDrv.Text
            DrvNew.Phone = DispPanel.txtTelDrv.Text
            DrvNew.Note = DispPanel.txtNoteDrv.Text
            
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn         As ADODB.Connection

            Dim rs         As Recordset

            Dim rs1        As Recordset

            Dim comIns     As String

            Dim comEdit    As String

            Dim numCheck   As Long

            Dim nameCheck  As String

            Dim nameCheck1 As String

            Dim DrvCount   As Long

            Dim frCheck    As Boolean
    
            DrvCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT d_num, d_name FROM drivers WHERE d_num = " & Val(DispPanel.txtDrv) & ";") 'водач код ако има

            If Not rs.BOF And Not rs.EOF Then
                numCheck = rs!d_num
                nameCheck = rs!d_name
            End If
            
            Set rs1 = cn.Execute("SELECT d_name FROM drivers WHERE d_name = '" & DispPanel.txtNameDrv & "';") 'водач име ако има

            If Not rs1.BOF And Not rs1.EOF Then nameCheck1 = rs1!d_name
            
            If numCheck <> Val(DispPanel.txtDrv) And nameCheck <> DispPanel.txtNameDrv And nameCheck1 <> DispPanel.txtNameDrv Then
                'ако няма съвпадения в номер или име на водач правим запис
                comIns = "INSERT INTO drivers VALUES(" & DrvNew.Code & ",'" & DrvNew.Title & "','" & DrvNew.CarNum & "','" & DrvNew.Capacity & "','" & DrvNew.CarModel & "','" & DrvNew.Phone & "','" & DrvNew.Note & "', 'true')"
                Set rs = cn.Execute(comIns)
            ElseIf numCheck = Val(DispPanel.txtDrv) And nameCheck1 <> DispPanel.txtNameDrv Then
                'ако има такъв номер на водач, но въведеното име е друго - питаме за редакция по номер
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE drivers SET d_name = '" & DrvNew.Title & "',d_reg = '" & DrvNew.CarNum & "',d_cap = '" & DrvNew.Capacity & "',d_mod = '" & DrvNew.CarModel & "',d_tel = '" & DrvNew.Phone & "',d_note = '" & DrvNew.Note & "' WHERE d_num =" & DrvNew.Code & "" 'корекция по номер
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf nameCheck = DispPanel.txtNameDrv And numCheck = Val(DispPanel.txtDrv) Then
                'ако има такъв номер и името му е същото - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE drivers SET d_reg = '" & DrvNew.CarNum & "',d_cap = '" & DrvNew.Capacity & "',d_mod = '" & DrvNew.CarModel & "',d_tel = '" & DrvNew.Phone & "',d_note = '" & DrvNew.Note & "' WHERE d_name ='" & DrvNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf numCheck <> Val(DispPanel.txtDrv) And nameCheck1 = DispPanel.txtNameDrv Then
                'ако няма такъв номер, но има такова име на водач - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE drivers SET d_num = " & DrvNew.Code & ",d_reg = '" & DrvNew.CarNum & "',d_cap = '" & DrvNew.Capacity & "',d_mod = '" & DrvNew.CarModel & "',d_tel = '" & DrvNew.Phone & "',d_note = '" & DrvNew.Note & "' WHERE d_name ='" & DrvNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf numCheck = Val(DispPanel.txtDrv) And nameCheck1 = DispPanel.txtNameDrv And nameCheck1 <> nameCheck Then
                'ако има такъв номер, но има и такова име на водач под друг номер
                'извеждаме съобщение за избор на ново име
                MousePointer = vbDefault
                MsgBox MsgNewName, vbOKOnly Or vbCritical, MsgErrBx
                
                'затваряме базата данни и прекратяваме функцията
                rs.Close
                Set rs = Nothing
                cn.Close
                Set cn = Nothing
                DispPanel.txtDrv.SetFocus

                Exit Function

            End If
            
            DispPanel.lstDrv.ListItems.Clear
    
            Set rs = cn.Execute("SELECT d_num FROM drivers ORDER BY d_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                DrvCount = DrvCount + 1

                If Val(rs!d_num) <> DrvCount And frCheck = False Then
                    DispPanel.txtDrv.Text = Format(DrvCount, "0000")
                    frCheck = True
                Else
                    DispPanel.txtDrv.Text = Format(DrvCount + 1, "0000")
                End If

                rs.MoveNext
            Loop
                
            Set rs = cn.Execute("SELECT * FROM drivers WHERE d_show = '1' ORDER BY d_num ASC;")
            
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
            Do While Not rs.EOF
                Set itmX = DispPanel.lstDrv.ListItems.Add(1, , Format(rs!d_num, "0000"))
                itmX.SubItems(1) = rs!d_name
                itmX.SubItems(2) = rs!d_reg
                itmX.SubItems(3) = CSng(rDs(rs!d_cap))
                itmX.SubItems(4) = rs!d_mod
                itmX.SubItems(5) = rs!d_tel
                itmX.SubItems(6) = rs!d_note
                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------

            DispPanel.txtNameDrv.Text = ""
            DispPanel.txtRegDrv.Text = ""
            DispPanel.txtCapDrv.Text = 0
            DispPanel.txtModDrv.Text = ""
            DispPanel.txtTelDrv.Text = ""
            DispPanel.txtNoteDrv.Text = ""
            
            If DispPanel.lstDrv.ListItems.count > 0 Then
                AutoColW DispPanel.lstDrv
            Else
                DispPanel.txtDrv.Text = Format(DrvCount + 1, "0000")
            End If
            
            MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
    
        Else
            MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function DelDrvBut()
    'функция за изтриване на водач
    
    If DispPanel.frDrivers.Enabled = True And DispPanel.frDrivers.Visible = True Then
        If Len(DispPanel.txtDrv.Text) > 0 And Len(DispPanel.txtNameDrv.Text) > 0 And Len(DispPanel.txtRegDrv.Text) > 0 And Len(DispPanel.txtCapDrv.Text) > 0 And DispPanel.lstDrv.ListItems.count > 0 Then

            If DispPanel.lstDrv.SelectedItem.Text = DispPanel.txtDrv.Text Then
                response = MsgBox(MsgConfDel, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then
            
                    '------------------------------Start PostgreSQL----------------------------------
                    Dim cn       As ADODB.Connection

                    Dim rs       As Recordset

                    Dim DrvCount As Long

                    Dim frCheck  As Boolean
    
                    DrvCount = 0
                    frCheck = False
    
                    Set cn = New ADODB.Connection
                    cn.ConnectionTimeout = 10
                    cn.Open ConStr
                    MousePointer = vbHourglass
                
                    Set rs = cn.Execute("DELETE FROM drivers WHERE d_num = " & Val(DispPanel.txtDrv.Text) & ";") 'изтриване на запис
    
                    DispPanel.lstDrv.ListItems.Clear
    
                    Set rs = cn.Execute("SELECT d_num FROM drivers ORDER BY d_num ASC;")
    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                    Do While Not rs.EOF
                        DrvCount = DrvCount + 1

                        If Val(rs!d_num) <> DrvCount And frCheck = False Then
                            DispPanel.txtDrv.Text = Format(DrvCount, "0000")
                            frCheck = True
                        Else
                            DispPanel.txtDrv.Text = Format(ClntCount + 1, "0000")
                        End If

                        rs.MoveNext
                    Loop
                                        
                    Set rs = cn.Execute("SELECT * FROM drivers WHERE d_show = '1' ORDER BY d_num ASC;")
                
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                    Do While Not rs.EOF
                        Set itmX = DispPanel.lstDrv.ListItems.Add(1, , Format(rs!d_num, "00000"))
                        itmX.SubItems(1) = rs!d_name
                        itmX.SubItems(2) = rs!d_reg
                        itmX.SubItems(3) = CSng(rDs(rs!d_cap))
                        itmX.SubItems(4) = rs!d_mod
                        itmX.SubItems(5) = rs!d_tel
                        itmX.SubItems(6) = rs!d_note
                        rs.MoveNext
                    Loop
                
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    MousePointer = vbDefault
                    Set cn = Nothing
                    '--------------------------End PostgreSQL------------------------------------------

                    DispPanel.txtNameDrv.Text = ""
                    DispPanel.txtRegDrv.Text = ""
                    DispPanel.txtCapDrv.Text = 0
                    DispPanel.txtModDrv.Text = ""
                    DispPanel.txtTelDrv.Text = ""
                    DispPanel.txtNoteDrv.Text = ""
                
                    If DispPanel.lstDrv.ListItems.count > 0 Then
                        AutoColW DispPanel.lstDrv
                    Else
                        DispPanel.txtDrv.Text = Format(DrvCount + 1, "0000")
                    End If
                
                    MsgBox MsgDelSuccess, vbOKOnly Or vbInformation, MsgDelBx
                Else

                    Exit Function

                End If

            Else
                MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
            End If

        Else
            MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function ListDrvClick()
    'функция за зареждане на данни при маркиране на запис от таблицата
    
    If DispPanel.frDrivers.Enabled = True And DispPanel.frDrivers.Visible = True Then
        If DispPanel.lstDrv.ListItems.count > 0 Then
            DispPanel.txtDrv.Text = Format(Val(DispPanel.lstDrv.ListItems(DispPanel.lstDrv.SelectedItem.Index).Text), "0000")
            DispPanel.txtNameDrv.Text = DispPanel.lstDrv.ListItems(DispPanel.lstDrv.SelectedItem.Index).ListSubItems(1).Text
            DispPanel.txtRegDrv.Text = DispPanel.lstDrv.ListItems(DispPanel.lstDrv.SelectedItem.Index).ListSubItems(2).Text
            DispPanel.txtCapDrv.Text = DispPanel.lstDrv.ListItems(DispPanel.lstDrv.SelectedItem.Index).ListSubItems(3).Text
            DispPanel.txtModDrv.Text = DispPanel.lstDrv.ListItems(DispPanel.lstDrv.SelectedItem.Index).ListSubItems(4).Text
            DispPanel.txtTelDrv.Text = DispPanel.lstDrv.ListItems(DispPanel.lstDrv.SelectedItem.Index).ListSubItems(5).Text
            DispPanel.txtNoteDrv.Text = DispPanel.lstDrv.ListItems(DispPanel.lstDrv.SelectedItem.Index).ListSubItems(6).Text
        End If

    Else
    End If

End Function

Public Function OpenSuppliers()
    'зареждане на меню доставчици
    
    Dim itmX As ListItem
    
    DispPanel.txtSup.SetFocus
    DispPanel.txtSup.TabIndex = 0
    DispPanel.txtNameSup.TabIndex = 1
    DispPanel.txtBGSup.TabIndex = 2
    DispPanel.txtMOLSup.TabIndex = 3
    DispPanel.txtAddSup.TabIndex = 4
    DispPanel.txtTelSup.TabIndex = 5
    DispPanel.txtNoteSup.TabIndex = 6
    DispPanel.btnClearSup.TabIndex = 7
    DispPanel.btnSvNwSup.TabIndex = 8
    DispPanel.btnDelSup.TabIndex = 9
    DispPanel.btnShowSup.TabIndex = 10
    DispPanel.btnDisp.TabIndex = 11
    DispPanel.btnOrders.TabIndex = 12
    DispPanel.btnRecepies.TabIndex = 13
    DispPanel.btnClients.TabIndex = 14
    DispPanel.btnDrivers.TabIndex = 15
    DispPanel.btnSuppliers.TabIndex = 16
    DispPanel.btnMaterials.TabIndex = 17
    DispPanel.btnNotes.TabIndex = 18
    DispPanel.btnAdminPanel.TabIndex = 19
    DispPanel.btnExit.TabIndex = 20
    
    DispPanel.frSuppliers.Left = 120
    DispPanel.frSuppliers.Top = DispPanel.Height \ 2 - FrmPlace
    
    DispPanel.lstSup.ColumnHeaders.Clear
    DispPanel.lstSup.ListItems.Clear
    
    Set colx = DispPanel.lstSup.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 700
    
    Set colx = DispPanel.lstSup.ColumnHeaders.Add()
    colx.Text = uniFirm
    colx.Width = 3000
    
    Set colx = DispPanel.lstSup.ColumnHeaders.Add()
    colx.Text = uniBG
    colx.Width = 1400
    
    Set colx = DispPanel.lstSup.ColumnHeaders.Add()
    colx.Text = uniMOL
    colx.Width = 2500

    Set colx = DispPanel.lstSup.ColumnHeaders.Add()
    colx.Text = uniAdd
    colx.Width = 2500

    Set colx = DispPanel.lstSup.ColumnHeaders.Add()
    colx.Text = uniTel
    colx.Width = 1300
    
    Set colx = DispPanel.lstSup.ColumnHeaders.Add()
    colx.Text = uniNote
    colx.Width = 2200
    
    DispPanel.txtSup.MaxLength = 3
    DispPanel.txtSup.Text = ""
    DispPanel.txtNameSup.MaxLength = 60
    DispPanel.txtNameSup.Text = ""
    DispPanel.txtBGSup.MaxLength = 15
    DispPanel.txtBGSup.Text = ""
    DispPanel.txtMOLSup.MaxLength = 60
    DispPanel.txtMOLSup.Text = ""
    DispPanel.txtAddSup.MaxLength = 100
    DispPanel.txtAddSup.Text = ""
    DispPanel.txtTelSup.MaxLength = 15
    DispPanel.txtTelSup.Text = ""
    DispPanel.txtNoteSup.MaxLength = 100
    DispPanel.txtNoteSup.Text = ""
    
    '------------------------------Start PostgreSQL----------------------------------
    Dim cn       As ADODB.Connection

    Dim rs       As Recordset

    Dim SupCount As Long

    Dim frCheck  As Boolean
    
    SupCount = 0
    frCheck = False
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT s_num FROM suppliers ORDER BY s_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        SupCount = SupCount + 1

        If Val(rs!s_num) <> SupCount And frCheck = False Then
            DispPanel.txtSup.Text = Format(SupCount, "000")
            frCheck = True
        Else
            DispPanel.txtSup.Text = Format(SupCount + 1, "000")
        End If

        rs.MoveNext
    Loop
    
    Set rs = cn.Execute("SELECT * FROM suppliers WHERE s_show = '1' ORDER BY s_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        Set itmX = DispPanel.lstSup.ListItems.Add(1, , Format(rs!s_num, "000"))
        itmX.SubItems(1) = rs!s_name
        itmX.SubItems(2) = rs!s_bg
        itmX.SubItems(3) = rs!s_mol
        itmX.SubItems(4) = rs!s_add
        itmX.SubItems(5) = rs!s_tel
        itmX.SubItems(6) = rs!s_note
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
    '--------------------------End PostgreSQL------------------------------------------

    If DispPanel.lstSup.ListItems.count > 0 Then
        AutoColW DispPanel.lstSup
    Else
        DispPanel.txtSup.Text = Format(SupCount + 1, "000")
    End If

End Function

Public Function ClearSupBut()
    'функция за почистване на клетките на доставчик за въвеждане на нова
    
    If DispPanel.frSuppliers.Enabled = True And DispPanel.frSuppliers.Visible = True Then
        If DispPanel.lstSup.ListItems.count > 0 Then
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn       As ADODB.Connection

            Dim rs       As Recordset

            Dim SupCount As Long

            Dim frCheck  As Boolean
    
            SupCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT s_num FROM suppliers ORDER BY s_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                SupCount = SupCount + 1

                If Val(rs!s_num) <> SupCount And frCheck = False Then
                    DispPanel.txtSup.Text = Format(SupCount, "000")
                    frCheck = True
                Else
                    DispPanel.txtSup.Text = Format(SupCount + 1, "000")
                End If

                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------

        Else
            DispPanel.txtSup.Text = Format(SupCount + 1, "000")
        End If

        DispPanel.txtNameSup.Text = ""
        DispPanel.txtBGSup.Text = ""
        DispPanel.txtMOLSup.Text = ""
        DispPanel.txtAddSup.Text = ""
        DispPanel.txtTelSup.Text = ""
        DispPanel.txtNoteSup.Text = ""
    Else
    End If

End Function

Public Function SvNwSupBut()
    'функция за запис на доставчик
    
    If DispPanel.frSuppliers.Enabled = True And DispPanel.frSuppliers.Visible = True Then
        If Val(DispPanel.txtSup.Text) = 0 Then
            MsgBox MsgCodeZero, vbOKOnly Or vbCritical, MsgErrBx

            Exit Function

        Else
        End If
        
        Dim SupNew As Supplier

        Set SupNew = New Supplier
        
        Dim response As Integer

        If Len(DispPanel.txtSup.Text) > 0 And Len(DispPanel.txtNameSup.Text) > 0 And Len(DispPanel.txtBGSup.Text) > 0 And Len(DispPanel.txtMOLSup.Text) > 0 Then
            SupNew.Code = DispPanel.txtSup.Text
            SupNew.Title = DispPanel.txtNameSup.Text
            SupNew.Ident = DispPanel.txtBGSup.Text
            SupNew.MOL = DispPanel.txtMOLSup.Text
            SupNew.Address = DispPanel.txtAddSup.Text
            SupNew.Phone = DispPanel.txtTelSup.Text
            SupNew.Note = DispPanel.txtNoteSup.Text
            
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn       As ADODB.Connection

            Dim rs       As Recordset

            Dim comIns   As String

            Dim comEdit  As String

            Dim numCheck As Integer

            Dim SupCount As Long

            Dim frCheck  As Boolean
    
            SupCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT s_num, s_name FROM suppliers WHERE s_num = " & Val(DispPanel.txtSup) & ";") 'доставчик код ако има

            If Not rs.BOF And Not rs.EOF Then
                numCheck = rs!s_num
                nameCheck = rs!s_name
            End If
            
            Set rs1 = cn.Execute("SELECT s_name FROM suppliers WHERE s_name = '" & DispPanel.txtNameSup & "';") 'доставчик име ако има

            If Not rs1.BOF And Not rs1.EOF Then nameCheck1 = rs1!s_name
            
            If numCheck <> Val(DispPanel.txtSup) And nameCheck <> DispPanel.txtNameSup And nameCheck1 <> DispPanel.txtNameSup Then
                'ако няма съвпадения в номер или име на доставчик правим запис
                comIns = "INSERT INTO suppliers VALUES(" & SupNew.Code & ",'" & SupNew.Title & "','" & SupNew.Ident & "','" & SupNew.MOL & "','" & SupNew.Address & "','" & SupNew.Phone & "','" & SupNew.Note & "','true')"
                Set rs = cn.Execute(comIns)
            ElseIf numCheck = Val(DispPanel.txtSup) And nameCheck1 <> DispPanel.txtNameSup Then
                'ако има такъв номер на доставчик, но въведеното име е друго - питаме за редакция по номер
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE suppliers SET s_name = '" & SupNew.Title & "',s_bg = '" & SupNew.Ident & "',s_mol = '" & SupNew.MOL & "',s_add = '" & SupNew.Address & "',s_tel = '" & SupNew.Phone & "',s_note = '" & SupNew.Note & "' WHERE s_num =" & SupNew.Code & "" 'корекция по номер
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf nameCheck = DispPanel.txtNameSup And numCheck = Val(DispPanel.txtSup) Then
                'ако има такъв номер и името му е същото - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE suppliers SET s_bg = '" & SupNew.Ident & "',s_mol = '" & SupNew.MOL & "',s_add = '" & SupNew.Address & "',s_tel = '" & SupNew.Phone & "',s_note = '" & SupNew.Note & "' WHERE s_name = '" & SupNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf numCheck <> Val(DispPanel.txtSup) And nameCheck1 = DispPanel.txtNameSup Then
                'ако няма такъв номер, но има такова име на доставчик - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни
                    comEdit = "UPDATE suppliers SET s_num = " & SupNew.Code & ",s_bg = '" & SupNew.Ident & "',s_mol = '" & SupNew.MOL & "',s_add = '" & SupNew.Address & "',s_tel = '" & SupNew.Phone & "',s_note = '" & SupNew.Note & "' WHERE s_name = '" & SupNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    Set rs = Nothing
                    cn.Close
                    Set cn = Nothing

                    Exit Function

                End If

            ElseIf numCheck = Val(DispPanel.txtDrv) And nameCheck1 = DispPanel.txtNameDrv And nameCheck1 <> nameCheck Then
                'ако има такъв номер, но има и такова име на доставчик под друг номер
                'извеждаме съобщение за избор на ново име
                MousePointer = vbDefault
                MsgBox MsgNewName, vbOKOnly Or vbCritical, MsgErrBx
                
                'затваряме базата данни и прекратяваме функцията
                rs.Close
                Set rs = Nothing
                cn.Close
                Set cn = Nothing

                Exit Function

            End If
            
            DispPanel.lstSup.ListItems.Clear
    
            Set rs = cn.Execute("SELECT s_num FROM suppliers ORDER BY s_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                SupCount = SupCount + 1

                If Val(rs!s_num) <> SupCount And frCheck = False Then
                    DispPanel.txtSup.Text = Format(SupCount, "000")
                    frCheck = True
                Else
                    DispPanel.txtSup.Text = Format(SupCount + 1, "000")
                End If

                rs.MoveNext
            Loop
                
            Set rs = cn.Execute("SELECT * FROM suppliers WHERE s_show = '1' ORDER BY s_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                Set itmX = DispPanel.lstSup.ListItems.Add(1, , Format(rs!s_num, "000"))
                itmX.SubItems(1) = rs!s_name
                itmX.SubItems(2) = rs!s_bg
                itmX.SubItems(3) = rs!s_mol
                itmX.SubItems(4) = rs!s_add
                itmX.SubItems(5) = rs!s_tel
                itmX.SubItems(6) = rs!s_note
                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------
            
            DispPanel.txtNameSup.Text = ""
            DispPanel.txtBGSup.Text = ""
            DispPanel.txtMOLSup.Text = ""
            DispPanel.txtAddSup.Text = ""
            DispPanel.txtTelSup.Text = ""
            DispPanel.txtNoteSup.Text = ""
        
            If DispPanel.lstSup.ListItems.count > 0 Then
                AutoColW DispPanel.lstSup
            Else
                DispPanel.txtSup.Text = Format(SupCount + 1, "000")
            End If
            
            MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
    
        Else
            MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function DelSupBut()
    'функция за изтриване на доставчик
    
    If DispPanel.frSuppliers.Enabled = True And DispPanel.frSuppliers.Visible = True Then
        If Len(DispPanel.txtSup.Text) > 0 And Len(DispPanel.txtNameSup.Text) > 0 And Len(DispPanel.txtBGSup.Text) > 0 And Len(DispPanel.txtMOLSup.Text) > 0 Then
            response = MsgBox(MsgConfDel, vbYesNo Or vbQuestion, MsgEditBx)

            If response = vbYes Then
            
                '------------------------------Start PostgreSQL----------------------------------
                Dim cn       As ADODB.Connection

                Dim rs       As Recordset

                Dim SupCount As Long

                Dim frCheck  As Boolean
    
                SupCount = 0
                frCheck = False
    
                Set cn = New ADODB.Connection
                cn.ConnectionTimeout = 10
                cn.Open ConStr
                MousePointer = vbHourglass
                
                Set rs = cn.Execute("DELETE FROM suppliers WHERE s_num = " & Val(DispPanel.txtSup.Text) & ";") 'изтриване на запис
    
                DispPanel.lstSup.ListItems.Clear
    
                Set rs = cn.Execute("SELECT s_num FROM suppliers ORDER BY s_num ASC;")
    
                If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                Do While Not rs.EOF
                    SupCount = SupCount + 1

                    If Val(rs!s_num) <> SupCount And frCheck = False Then
                        DispPanel.txtSup.Text = Format(SupCount, "000")
                        frCheck = True
                    Else
                        DispPanel.txtSup.Text = Format(SupCount + 1, "000")
                    End If

                    rs.MoveNext
                Loop
                
                Set rs = cn.Execute("SELECT * FROM suppliers WHERE s_show = '1' ORDER BY s_num ASC;")
    
                If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                
                Do While Not rs.EOF
                    Set itmX = DispPanel.lstSup.ListItems.Add(1, , Format(rs!s_num, "000"))
                    itmX.SubItems(1) = rs!s_name
                    itmX.SubItems(2) = rs!s_bg
                    itmX.SubItems(3) = rs!s_mol
                    itmX.SubItems(4) = rs!s_add
                    itmX.SubItems(5) = rs!s_tel
                    itmX.SubItems(6) = rs!s_note
                    rs.MoveNext
                Loop
                
                rs.Close
                Set rs = Nothing
                cn.Close
                MousePointer = vbDefault
                Set cn = Nothing
                '--------------------------End PostgreSQL------------------------------------------

                DispPanel.txtNameSup.Text = ""
                DispPanel.txtBGSup.Text = ""
                DispPanel.txtMOLSup.Text = ""
                DispPanel.txtAddSup.Text = ""
                DispPanel.txtTelSup.Text = ""
                DispPanel.txtNoteSup.Text = ""
                
                If DispPanel.lstSup.ListItems.count > 0 Then
                    AutoColW DispPanel.lstSup
                Else
                    DispPanel.txtSup.Text = Format(SupCount + 1, "000")
                End If
                
                MsgBox MsgDelSuccess, vbOKOnly Or vbInformation, MsgDelBx
            Else

                Exit Function

            End If

        Else
            MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function ListSupClick()
    'функция за зареждане на данни при маркиране на запис от таблицата
    
    If DispPanel.frSuppliers.Enabled = True And DispPanel.frSuppliers.Visible = True Then
        If DispPanel.lstSup.ListItems.count > 0 Then
            DispPanel.txtSup.Text = Format(Val(DispPanel.lstSup.ListItems(DispPanel.lstSup.SelectedItem.Index).Text), "000")
            DispPanel.txtNameSup.Text = DispPanel.lstSup.ListItems(DispPanel.lstSup.SelectedItem.Index).ListSubItems(1).Text
            DispPanel.txtBGSup.Text = DispPanel.lstSup.ListItems(DispPanel.lstSup.SelectedItem.Index).ListSubItems(2).Text
            DispPanel.txtMOLSup.Text = DispPanel.lstSup.ListItems(DispPanel.lstSup.SelectedItem.Index).ListSubItems(3).Text
            DispPanel.txtAddSup.Text = DispPanel.lstSup.ListItems(DispPanel.lstSup.SelectedItem.Index).ListSubItems(4).Text
            DispPanel.txtTelSup.Text = DispPanel.lstSup.ListItems(DispPanel.lstSup.SelectedItem.Index).ListSubItems(5).Text
            DispPanel.txtNoteSup.Text = DispPanel.lstSup.ListItems(DispPanel.lstSup.SelectedItem.Index).ListSubItems(6).Text
        End If

    Else
    End If

End Function

Public Function OpenMaterials()
    'зареждане на меню материали
    
    Dim itmX          As ListItem

    Dim types(0 To 4) As String

    DispPanel.txtMatName.SetFocus
    DispPanel.txtMatName.TabIndex = 0
    DispPanel.cmbMatType.TabIndex = 1
    DispPanel.s1(0).TabIndex = 2
    DispPanel.s1(1).TabIndex = 3
    DispPanel.s1(2).TabIndex = 4
    DispPanel.s1(3).TabIndex = 5
    DispPanel.s1(4).TabIndex = 6
    DispPanel.s1(5).TabIndex = 7
    DispPanel.btnClearMat.TabIndex = 8
    DispPanel.btnSvNwMat.TabIndex = 9
    DispPanel.btnDelMat.TabIndex = 10
    DispPanel.btnAddMatDlvr.TabIndex = 11
    DispPanel.btnSvExp.TabIndex = 12
    DispPanel.btnDisp.TabIndex = 13
    DispPanel.btnOrders.TabIndex = 14
    DispPanel.btnRecepies.TabIndex = 15
    DispPanel.btnClients.TabIndex = 16
    DispPanel.btnDrivers.TabIndex = 17
    DispPanel.btnSuppliers.TabIndex = 18
    DispPanel.btnMaterials.TabIndex = 19
    DispPanel.btnNotes.TabIndex = 20
    DispPanel.btnAdminPanel.TabIndex = 21
    DispPanel.btnExit.TabIndex = 22

    DispPanel.frMaterials.Left = 120
    DispPanel.frMaterials.Top = DispPanel.Height \ 2 - FrmPlace
    
    DispPanel.cmbMatType.Locked = False
    DispPanel.cmbMatType.TabStop = True
    
    DispPanel.lstMat.ColumnHeaders.Clear
    DispPanel.lstMat.ListItems.Clear
    
    DispPanel.btnSvExp.Visible = False
    DispPanel.lblMatLoad.Visible = False
    
    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniCode
    colx.Width = 800
    
    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniNm
    colx.Width = 3000
    
    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniType
    colx.Width = 2000
    
    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniLoad & " Машина 1"
    colx.Width = 2000
    
    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniLoad & " Машина 2"
    colx.Width = 2000

    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniDelivered
    colx.Width = 1500

    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniSold & " Машина 1"
    colx.Width = 1500
    
    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniSold & " Машина 2"
    colx.Width = 1500
    
    Set colx = DispPanel.lstMat.ColumnHeaders.Add()
    colx.Text = uniHave
    colx.Width = 1500
    
    DispPanel.cmbMatType.Clear

    types(0) = uniIM

    types(1) = uniConMat

    types(2) = uniWat

    types(3) = uniChem

    types(4) = uniOther

    For i = 0 To 4
        DispPanel.cmbMatType.AddItem types(i)
    Next i

    DispPanel.txtMat.MaxLength = 3
    DispPanel.txtMat.Text = ""
    DispPanel.txtMatName.MaxLength = 60
    DispPanel.txtMatName.Text = ""
    DispPanel.cmbMatType.ListIndex = -1

    For i = 0 To 5
        DispPanel.s1(i).Visible = False
        DispPanel.s1(i) = 0
    Next i
    
    '------------------------------Start PostgreSQL----------------------------------
    Dim cn       As ADODB.Connection

    Dim rs       As Recordset
    
    Dim rsOther  As Recordset

    Dim MatCount As Integer

    Dim frCheck  As Boolean
    
    Dim MachineOther As Integer
    
    MatCount = 0
    frCheck = False
    
    If MachineNumber = 1 Then MachineOther = 2
    If MachineNumber = 2 Then MachineOther = 1
    
    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 10
    cn.Open ConStr
    MousePointer = vbHourglass
    
    Set rs = cn.Execute("SELECT * FROM materials_bc" & MachineNumber & " ORDER BY m_num ASC;")
    Set rsOther = cn.Execute("SELECT * FROM materials_bc" & MachineOther & " ORDER BY m_num ASC;")
    
    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
    If Not rsOther.EOF And Not rsOther.BOF Then rsOther.MoveFirst
    
    Do While Not rs.EOF Or Not rsOther.EOF
        MatCount = MatCount + 1

        If Val(rs!m_num) <> MatCount And frCheck = False Then
            DispPanel.txtMat.Text = Format(MatCount, "000")
            frCheck = True
        Else
            DispPanel.txtMat.Text = Format(MatCount + 1, "000")
        End If
        
        Set itmX = DispPanel.lstMat.ListItems.Add(1, , Format(rs!m_num, "000"))
        itmX.SubItems(1) = rs!m_name
        
        If MachineNumber = 1 Then
            If Val(rs!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rs!m_type))
            itmX.SubItems(3) = rs!m_load
            If Val(rsOther!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsOther!m_type))
            itmX.SubItems(4) = rsOther!m_load
        ElseIf MachineNumber = 2 Then
            If Val(rsOther!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsOther!m_type))
            itmX.SubItems(3) = rsOther!m_load
            If Val(rs!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rs!m_type))
            itmX.SubItems(4) = rs!m_load
        End If
        
        If Val(rs!m_type) <> 2 Then
            itmX.SubItems(5) = rDs(rs!m_del)
        Else
            itmX.SubItems(5) = "---------"
        End If
        
        If MachineNumber = 1 Then
            itmX.SubItems(6) = rDs(rs!m_sold)
            itmX.SubItems(7) = rDs(rsOther!m_sold)
        ElseIf MachineNumber = 2 Then
            itmX.SubItems(6) = rDs(rsOther!m_sold)
            itmX.SubItems(7) = rDs(rs!m_sold)
        End If
        
        If Val(rs!m_type) <> 2 Then
            itmX.SubItems(8) = CSng(rDs(rs!m_del)) - CSng(rDs(rs!m_sold)) - CSng(rDs(rsOther!m_sold))
        Else
            itmX.SubItems(8) = "---------"
        End If

        rs.MoveNext
        rsOther.MoveNext
    Loop
    
    rs.Close
    rsOther.Close
    Set rs = Nothing
    Set rsOther = Nothing
    cn.Close
    MousePointer = vbDefault
    Set cn = Nothing
    '--------------------------End PostgreSQL------------------------------------------
    
    If DispPanel.lstMat.ListItems.count > 0 Then
        AutoColW DispPanel.lstMat
    Else
        DispPanel.txtMat.Text = Format(MatCount + 1, "000")
    End If

End Function

Public Function ClearMatBut()
    'функция за почистване на клетките на материал за въвеждане на нова
    
    If DispPanel.frMaterials.Enabled = True And DispPanel.frMaterials.Visible = True Then
        DispPanel.cmbMatType.Locked = False
        DispPanel.cmbMatType.TabStop = True

        If DispPanel.lstMat.ListItems.count > 0 Then
        
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn       As ADODB.Connection

            Dim rs       As Recordset

            Dim MatCount As Integer

            Dim frCheck  As Boolean
    
            MatCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            Set rs = cn.Execute("SELECT m_num FROM materials_bc" & MachineNumber & " ORDER BY m_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            
            Do While Not rs.EOF
                MatCount = MatCount + 1

                If Val(rs!m_num) <> MatCount And frCheck = False Then
                    DispPanel.txtMat.Text = Format(MatCount, "000")
                    frCheck = True
                Else
                    DispPanel.txtMat.Text = Format(MatCount + 1, "000")
                End If

                rs.MoveNext
            Loop

            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------

        Else
            DispPanel.txtMat.Text = Format(DispPanel.lstMat.ListItems.count + 1, "000")
        End If

        DispPanel.txtMatName.Text = ""
        DispPanel.cmbMatType.ListIndex = -1

        For i = 0 To 5
            DispPanel.s1(i).Visible = False
            DispPanel.s1(i) = 0
        Next i

    Else
    End If

End Function

Public Function SvNwMatBut()
    'функция за запис на материал
    
    If DispPanel.frMaterials.Enabled = True And DispPanel.frMaterials.Visible = True Then
        
        Dim MatNew As Material

        Set MatNew = New Material
        
        Dim MatNotL As String
        
        Dim response      As Integer

        Dim types(0 To 4) As String

        DispPanel.cmbMatType.Locked = False
        DispPanel.cmbMatType.TabStop = True
        
        types(0) = uniIM

        types(1) = uniConMat

        types(2) = uniWat

        types(3) = uniChem

        types(4) = uniOther
        
        If Len(DispPanel.txtMat.Text) > 0 And Len(DispPanel.txtMatName.Text) > 0 And Len(DispPanel.cmbMatType.Text) > 0 Then
            MatNew.Code = DispPanel.txtMat.Text
            MatNew.Title = DispPanel.txtMatName.Text
            MatNew.Kind = DispPanel.cmbMatType.ListIndex

            If MatNew.Kind = 0 Then 'им

                For i = 0 To 5
                    MatNew.Loaded = MatNew.Loaded & str(DispPanel.s1(i))
                Next i
                
                MatNotL = " 0 0 0 0 0 0"
                
            End If

            If MatNew.Kind = 1 Then 'цимент

                For i = 0 To 3
                    MatNew.Loaded = MatNew.Loaded & str(DispPanel.s1(i))
                Next i
                
                MatNotL = " 0 0 0 0"
                
            End If

            If MatNew.Kind = 2 Then 'вода

                For i = 0 To 1
                    MatNew.Loaded = MatNew.Loaded & str(DispPanel.s1(i))
                Next i
                
                MatNotL = " 0 0"
                
            End If

            If MatNew.Kind = 3 Then 'хд

                For i = 0 To 5
                    MatNew.Loaded = MatNew.Loaded & str(DispPanel.s1(i))
                Next i
                
                MatNotL = " 0 0 0 0 0 0"

            End If

            MatNew.Delivered = 0
            MatNew.Sold = 0
            
            '------------------------------Start PostgreSQL----------------------------------
            Dim cn         As ADODB.Connection

            Dim rs         As Recordset

            Dim rs1        As Recordset
            
            Dim rsOther    As Recordset

            Dim comIns     As String

            Dim comEdit    As String

            Dim numCheck   As Long

            Dim nameCheck  As String

            Dim nameCheck1 As String

            Dim MatCount   As Integer

            Dim frCheck    As Boolean
            
            Dim MachineOther As Integer
            
            If MachineNumber = 1 Then MachineOther = 2
            If MachineNumber = 2 Then MachineOther = 1
            
            MatCount = 0
            frCheck = False
    
            Set cn = New ADODB.Connection
            cn.ConnectionTimeout = 10
            cn.Open ConStr
            MousePointer = vbHourglass
            
            'проверка дали посочената течка е свободна
            Set rs = cn.Execute("SELECT m_name, m_type, m_load FROM materials_bc" & MachineNumber & " WHERE m_type = '" & MatNew.Kind & "';") 'материал тип ако има
            
            If Not rs.BOF And Not rs.EOF Then rs.MoveFirst
            
            Do While Not rs.EOF

                If Val(rs!m_type) = MatNew.Kind And rs!m_name <> MatNew.Title Then
                    If (Val(Mid$(rs!m_load, 1, 2)) = 1 And Val(Mid$(MatNew.Loaded, 1, 2)) = 1) Or (Val(Mid$(rs!m_load, 3, 2)) = 1 And Val(Mid$(MatNew.Loaded, 3, 2)) = 1) Or (Val(Mid$(rs!m_load, 5, 2)) = 1 And Val(Mid$(MatNew.Loaded, 5, 2)) = 1) Or (Val(Mid$(rs!m_load, 7, 2)) = 1 And Val(Mid$(MatNew.Loaded, 7, 2)) = 1) Or (Val(Mid$(rs!m_load, 9, 2)) = 1 And Val(Mid$(MatNew.Loaded, 9, 2)) = 1) Or (Val(Mid$(rs!m_load, 11, 2)) = 1 And Val(Mid$(MatNew.Loaded, 11, 2)) = 1) Then
                        MousePointer = vbDefault
                        MsgBox MsgBusyFlow, vbOKOnly Or vbCritical, MsgErrBx
                        rs.Close
                        Set rs = Nothing
                        cn.Close
                        Set cn = Nothing

                        For i = 0 To 5
                            DispPanel.s1(i) = 0
                        Next i

                        Exit Function

                    Else
                    End If

                Else
                End If

                rs.MoveNext
            Loop
            
            'проверка дали силоз е прикачен към другата машина ако работим с материал - свързващо вещество
            If MatNew.Kind = 1 Then
                Set rsOther = cn.Execute("SELECT m_name, m_load FROM materials_bc" & MachineOther & " WHERE m_type = '" & MatNew.Kind & "';") 'материал тип 1 ако има
                
                If Not rsOther.BOF And Not rsOther.EOF Then rsOther.MoveFirst
                Do While Not rsOther.EOF
                    If rsOther!m_name = MatNew.Title And Val(rsOther!m_load) <> 0 Then
                        MousePointer = vbDefault
                        MsgBox "Силозът е свързан с другата машина!", vbOKOnly Or vbCritical, MsgErrBx
                        rsOther.Close
                        Set rsOther = Nothing
                        cn.Close
                        Set cn = Nothing
                        
                        For i = 0 To 5
                            DispPanel.s1(i) = 0
                        Next i

                        Exit Function
                    
                    Else
                    End If
                    
                    rsOther.MoveNext
                Loop
            End If
            
            'проверки за вида на записите
            Set rs = cn.Execute("SELECT m_num, m_name FROM materials_bc" & MachineNumber & " WHERE m_num = " & Val(DispPanel.txtMat) & ";") 'материал код ако има

            If Not rs.BOF And Not rs.EOF Then
                numCheck = rs!m_num
                nameCheck = rs!m_name
            End If
            
            Set rs1 = cn.Execute("SELECT m_name FROM materials_bc" & MachineNumber & " WHERE m_name = '" & DispPanel.txtMatName & "';") 'материал име ако има

            If Not rs1.BOF And Not rs1.EOF Then nameCheck1 = rs1!m_name
            
            If numCheck <> Val(DispPanel.txtMat) And nameCheck <> DispPanel.txtMatName And nameCheck1 <> DispPanel.txtMatName Then
                'ако няма съвпадения в номер или име на материал правим запис и в двете таблици за материали
                comIns = "INSERT INTO materials_bc" & MachineNumber & " VALUES(" & MatNew.Code & ",'" & MatNew.Title & "','" & MatNew.Kind & "','" & MatNew.Loaded & "','" & MatNew.Delivered & "','" & MatNew.Sold & "')"
                Set rs = cn.Execute(comIns)
                comIns = "INSERT INTO materials_bc" & MachineOther & " VALUES(" & MatNew.Code & ",'" & MatNew.Title & "','" & MatNew.Kind & "','" & MatNotL & "','" & MatNew.Delivered & "','" & MatNew.Sold & "')"
                Set rsOther = cn.Execute(comIns)
            ElseIf numCheck = Val(DispPanel.txtMat) And nameCheck1 <> DispPanel.txtMatName Then
                'ако има такъв номер на материал, но въведеното име е друго - питаме за редакция по номер
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни и в двете таблици без наличностите
                    comEdit = "UPDATE materials_bc" & MachineNumber & " SET m_name = '" & MatNew.Title & "',m_type = '" & MatNew.Kind & "',m_load = '" & MatNew.Loaded & "' WHERE m_num =" & MatNew.Code & "" 'корекция по номер
                    Set rs = cn.Execute(comEdit)
                    comEdit = "UPDATE materials_bc" & MachineOther & " SET m_name = '" & MatNew.Title & "',m_type = '" & MatNew.Kind & "' WHERE m_num =" & MatNew.Code & "" 'корекция по номер
                    Set rsOther = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    rsOther.Close
                    Set rs = Nothing
                    Set rsOther = Nothing
                    cn.Close
                    Set cn = Nothing

                    For i = 0 To 5
                        DispPanel.s1(i) = 0
                    Next i

                    Exit Function

                End If

            ElseIf nameCheck = DispPanel.txtMatName And numCheck = Val(DispPanel.txtMat) Then
                'ако има такъв номер и името му е същото - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни без наличностите
                    comEdit = "UPDATE materials_bc" & MachineNumber & " SET m_type = '" & MatNew.Kind & "',m_load = '" & MatNew.Loaded & "' WHERE m_name ='" & MatNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                    comEdit = "UPDATE materials_bc" & MachineOther & " SET m_type = '" & MatNew.Kind & "' WHERE m_name ='" & MatNew.Title & "'" 'корекция по име
                    Set rsOther = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    rsOther.Close
                    Set rs = Nothing
                    Set rsOther = Nothing
                    cn.Close
                    Set cn = Nothing

                    For i = 0 To 5
                        DispPanel.s1(i) = 0
                    Next i

                    Exit Function

                End If

            ElseIf numCheck <> Val(DispPanel.txtMat) And nameCheck1 = DispPanel.txtMatName Then
                'ако няма такъв номер, но има такова име на материал - питаме за редакция по име
                MousePointer = vbDefault
                response = MsgBox(MsgConfEdit, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then 'при потвърдена редакция правим корекция в базата данни без наличностите
                    comEdit = "UPDATE materials_bc" & MachineNumber & " SET m_num = " & MatNew.Code & ",m_type = '" & MatNew.Kind & "',m_load = '" & MatNew.Loaded & "' WHERE m_name ='" & MatNew.Title & "'" 'корекция по име
                    Set rs = cn.Execute(comEdit)
                    comEdit = "UPDATE materials_bc" & MachineOther & " SET m_num = " & MatNew.Code & ",m_type = '" & MatNew.Kind & "' WHERE m_name ='" & MatNew.Title & "'" 'корекция по име
                    Set rsOther = cn.Execute(comEdit)
                Else 'при отказ от редактиране затваряме базата данни и прекратяваме функцията
                    rs.Close
                    rsOther.Close
                    Set rs = Nothing
                    Set rsOther = Nothing
                    cn.Close
                    Set cn = Nothing

                    For i = 0 To 5
                        DispPanel.s1(i) = 0
                    Next i

                    Exit Function

                End If

            ElseIf numCheck = Val(DispPanel.txtMat) And nameCheck1 = DispPanel.txtMatName And nameCheck1 <> nameCheck Then
                'ако има такъв номер, но има и такова име на материал под друг номер
                'извеждаме съобщение за избор на ново име
                MousePointer = vbDefault
                MsgBox MsgNewName, vbOKOnly Or vbCritical, MsgErrBx
                
                'затваряме базата данни и прекратяваме функцията
                rs.Close
                rsOther.Close
                Set rs = Nothing
                Set rsOther = Nothing
                cn.Close
                Set cn = Nothing

                For i = 0 To 5
                    DispPanel.s1(i) = 0
                Next i

                Exit Function

            End If
            
            DispPanel.lstMat.ListItems.Clear
    
            Set rs = cn.Execute("SELECT * FROM materials_bc" & MachineNumber & " ORDER BY m_num ASC;")
            Set rsOther = cn.Execute("SELECT * FROM materials_bc" & MachineOther & " ORDER BY m_num ASC;")
    
            If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
            If Not rsOther.EOF And Not rsOther.BOF Then rsOther.MoveFirst
            
            Do While Not rs.EOF Or Not rsOther.EOF
                MatCount = MatCount + 1

                If Val(rs!m_num) <> MatCount And frCheck = False Then
                    DispPanel.txtMat.Text = Format(MatCount, "000")
                    frCheck = True
                Else
                    DispPanel.txtMat.Text = Format(MatCount + 1, "000")
                End If
                
                Set itmX = DispPanel.lstMat.ListItems.Add(1, , Format(rs!m_num, "000"))
                itmX.SubItems(1) = rs!m_name
                
                If MachineNumber = 1 Then
                    If Val(rs!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rs!m_type))
                    itmX.SubItems(3) = rs!m_load
                    If Val(rsOther!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsOther!m_type))
                    itmX.SubItems(4) = rsOther!m_load
                ElseIf MachineNumber = 2 Then
                    If Val(rsOther!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsOther!m_type))
                    itmX.SubItems(3) = rsOther!m_load
                    If Val(rs!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rs!m_type))
                    itmX.SubItems(4) = rs!m_load
                End If
                
                If Val(rs!m_type) <> 2 Then
                    itmX.SubItems(5) = rDs(rs!m_del)
                Else
                    itmX.SubItems(5) = "---------"
                End If
                
                If MachineNumber = 1 Then
                    itmX.SubItems(6) = rDs(rs!m_sold)
                    itmX.SubItems(7) = rDs(rsOther!m_sold)
                ElseIf MachineNumber = 2 Then
                    itmX.SubItems(6) = rDs(rsOther!m_sold)
                    itmX.SubItems(7) = rDs(rs!m_sold)
                End If
                
                If Val(rs!m_type) <> 2 Then
                    itmX.SubItems(8) = CSng(rDs(rs!m_del)) - CSng(rDs(rs!m_sold)) - CSng(rDs(rsOther!m_sold))
                Else
                    itmX.SubItems(8) = "---------"
                End If

                rs.MoveNext
                rsOther.MoveNext
            Loop

            rs.Close
            Set rs = Nothing
            cn.Close
            MousePointer = vbDefault
            Set cn = Nothing
            '--------------------------End PostgreSQL------------------------------------------

            If DispPanel.lstMat.ListItems.count > 0 Then
                AutoColW DispPanel.lstMat
            Else
                DispPanel.txtMat.Text = Format(MatCount + 1, "000")
            End If

            DispPanel.txtMatName.Text = ""
            DispPanel.cmbMatType.ListIndex = -1

            For i = 0 To 5
                DispPanel.s1(i).Visible = False
                DispPanel.s1(i) = 0
            Next i

            MsgBox MsgSaveSuccess, vbOKOnly Or vbInformation, uniSave
            DispPanel.btnSvExp.Visible = False
        Else
            MsgBox MsgFillAll, vbOKOnly Or vbCritical, MsgErrBx

            Exit Function

        End If

    Else
    End If

    Load frmNameSilos
    frmNameSilos.Hide
    Call frmNameSilos.btnSaveSilos_Click
End Function

Public Function DelMatBut()

    'функция за изтриване на материал
    Dim types(0 To 4) As String

    types(0) = uniIM

    types(1) = uniConMat

    types(2) = uniWat

    types(3) = uniChem

    types(4) = uniOther
        
    If DispPanel.frMaterials.Enabled = True And DispPanel.frMaterials.Visible = True Then
        DispPanel.cmbMatType.Locked = False
        DispPanel.cmbMatType.TabStop = True
        
        If Len(DispPanel.txtMat.Text) > 0 And Len(DispPanel.txtMatName.Text) > 0 And Len(DispPanel.cmbMatType.Text) > 0 And DispPanel.lstMat.ListItems.count > 0 Then

            If DispPanel.lstMat.SelectedItem.Text = DispPanel.txtMat.Text Then
                response = MsgBox(MsgConfDel, vbYesNo Or vbQuestion, MsgEditBx)

                If response = vbYes Then
            
                    '------------------------------Start PostgreSQL----------------------------------
                    Dim cn       As ADODB.Connection

                    Dim rs       As Recordset
                    
                    Dim rsOther  As Recordset

                    Dim MatCount As Integer

                    Dim frCheck  As Boolean
                    
                    Dim MachineOther As Integer
                    
                    If MachineNumber = 1 Then MachineOther = 2
                    If MachineNumber = 2 Then MachineOther = 1

                    MatCount = 0
                    frCheck = False
    
                    Set cn = New ADODB.Connection
                    cn.ConnectionTimeout = 10
                    cn.Open ConStr
                    MousePointer = vbHourglass
                
                    Set rs = cn.Execute("SELECT m_load, m_del, m_sold FROM materials_bc" & MachineNumber & " WHERE m_num = " & Val(DispPanel.txtMat.Text) & ";")
                    Set rsOther = cn.Execute("SELECT m_load, m_del, m_sold FROM materials_bc" & MachineOther & " WHERE m_num = " & Val(DispPanel.txtMat.Text) & ";")

                    'проверка дали материалът е зареден - ако е зареден не може да се изтрие
                    If Val(rs!m_load) <> 0 Or Val(rsOther!m_load) <> 0 Then
                        MousePointer = vbDefault
                        MsgBox MsgMatLoaded, vbOKOnly Or vbCritical, MsgErrBx
                        rs.Close
                        rsOther.Close
                        Set rs = Nothing
                        Set rsOther = Nothing
                        cn.Close
                        Set cn = Nothing

                        Exit Function

                    Else
                    End If
                
                    'проверка дали има доставка на материал - ако има не може да се изтрие
                    If CSng(rDs(rs!m_del)) > 0 Or CSng(rDs(rsOther!m_del)) > 0 Then
                        MousePointer = vbDefault
                        MsgBox MsgMatDelivery, vbOKOnly Or vbCritical, MsgErrBx
                        rs.Close
                        rsOther.Close
                        Set rs = Nothing
                        Set rsOther = Nothing
                        cn.Close
                        Set cn = Nothing

                        Exit Function

                    Else
                    End If
                
                    'проверка дали има разход на материал - ако има не може да се изтрие
                    If CSng(rDs(rs!m_sold)) <> 0 Or CSng(rDs(rsOther!m_sold)) <> 0 Then
                        MousePointer = vbDefault
                        MsgBox MsgMatSold, vbOKOnly Or vbCritical, MsgErrBx
                        rs.Close
                        rsOther.Close
                        Set rs = Nothing
                        Set rsOther = Nothing
                        cn.Close
                        Set cn = Nothing

                        Exit Function

                    Else
                    End If
                
                    Set rs = cn.Execute("DELETE FROM materials_bc" & MachineNumber & " WHERE m_num = " & Val(DispPanel.txtMat.Text) & ";") 'изтриване на запис
                    Set rsOther = cn.Execute("DELETE FROM materials_bc" & MachineOther & " WHERE m_num = " & Val(DispPanel.txtMat.Text) & ";") 'изтриване на запис
    
                    DispPanel.lstMat.ListItems.Clear
    
                    Set rs = cn.Execute("SELECT * FROM materials_bc" & MachineNumber & " ORDER BY m_num ASC;")
                    Set rsOther = cn.Execute("SELECT * FROM materials_bc" & MachineOther & " ORDER BY m_num ASC;")
                    
                    If Not rs.EOF And Not rs.BOF Then rs.MoveFirst
                    If Not rsOther.EOF And Not rsOther.BOF Then rsOther.MoveFirst
                    
                    Do While Not rs.EOF Or Not rsOther.EOF
                        MatCount = MatCount + 1

                        If Val(rs!m_num) <> MatCount And frCheck = False Then
                            DispPanel.txtMat.Text = Format(MatCount, "000")
                            frCheck = True
                        Else
                            DispPanel.txtMat.Text = Format(MatCount + 1, "000")
                        End If
                    
                        Set itmX = DispPanel.lstMat.ListItems.Add(1, , Format(rs!m_num, "000"))
                        itmX.SubItems(1) = rs!m_name
                        
                        If MachineNumber = 1 Then
                            If Val(rs!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rs!m_type))
                            itmX.SubItems(3) = rs!m_load
                            If Val(rsOther!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsOther!m_type))
                            itmX.SubItems(4) = rsOther!m_load
                        ElseIf MachineNumber = 2 Then
                            If Val(rsOther!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rsOther!m_type))
                            itmX.SubItems(3) = rsOther!m_load
                            If Val(rs!m_type) >= 0 Then itmX.SubItems(2) = types(Val(rs!m_type))
                            itmX.SubItems(4) = rs!m_load
                        End If
                        
                        If Val(rs!m_type) <> 2 Then
                            itmX.SubItems(5) = rDs(rs!m_del)
                        Else
                            itmX.SubItems(5) = "---------"
                        End If
                        
                        If MachineNumber = 1 Then
                            itmX.SubItems(6) = rDs(rs!m_sold)
                            itmX.SubItems(7) = rDs(rsOther!m_sold)
                        ElseIf MachineNumber = 2 Then
                            itmX.SubItems(6) = rDs(rsOther!m_sold)
                            itmX.SubItems(7) = rDs(rs!m_sold)
                        End If
                        
                        If Val(rs!m_type) <> 2 Then
                            itmX.SubItems(8) = CSng(rDs(rs!m_del)) - CSng(rDs(rs!m_sold)) - CSng(rDs(rsOther!m_sold))
                        Else
                            itmX.SubItems(8) = "---------"
                        End If

                        rs.MoveNext
                        rsOther.MoveNext
                    Loop
                
                    rs.Close
                    rsOther.Close
                    Set rs = Nothing
                    Set rsOther = Nothing
                    cn.Close
                    MousePointer = vbDefault
                    Set cn = Nothing
                    '--------------------------End PostgreSQL------------------------------------------

                    DispPanel.btnSvExp.Visible = False
                    MsgBox MsgDelSuccess, vbOKOnly Or vbInformation, MsgDelBx
                    DispPanel.txtMatName.Text = ""
                    DispPanel.cmbMatType.ListIndex = -1
                
                    If DispPanel.lstMat.ListItems.count > 0 Then
                        AutoColW DispPanel.lstMat
                    Else
                        DispPanel.txtMat.Text = Format(MatCount + 1, "000")
                    End If

                Else

                    Exit Function

                End If

            Else
                MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
            End If

        Else
            MsgBox MsgNoSelection, vbOKOnly Or vbCritical, MsgErrBx
        End If

    Else
    End If

End Function

Public Function ListMatClick()
    'функция за зареждане на данни при маркиране на запис от таблицата
    
    If DispPanel.frMaterials.Enabled = True And DispPanel.frMaterials.Visible = True Then

        For i = 0 To 5
            DispPanel.s1(i) = 0
        Next i
        
        If DispPanel.lstMat.ListItems.count > 0 Then
            DispPanel.txtMat.Text = Format(Val(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).Text), "000")
            DispPanel.txtMatName.Text = DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(1).Text
            DispPanel.cmbMatType.Text = DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(2).Text

            If DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(2).Text = uniIM Then

                For i = 0 To ns1 - 1
                    If MachineNumber = 1 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(3).Text, 1 + 2 * i, 2))
                    ElseIf MachineNumber = 2 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(4).Text, 1 + 2 * i, 2))
                    End If
                Next i

            End If

            If DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(2).Text = uniConMat Then

                For i = 0 To ns3 - 1
                    If MachineNumber = 1 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(3).Text, 1 + 2 * i, 2))
                    ElseIf MachineNumber = 2 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(4).Text, 1 + 2 * i, 2))
                    End If
                Next i

            End If

            If DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(2).Text = uniWat Then

                For i = 0 To ns2 - 1
                    If MachineNumber = 1 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(3).Text, 1 + 2 * i, 2))
                    ElseIf MachineNumber = 2 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(4).Text, 1 + 2 * i, 2))
                    End If
                Next i

            End If

            If DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(2).Text = uniChem Then

                For i = 0 To ns4 - 1
                    If MachineNumber = 1 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(3).Text, 1 + 2 * i, 2))
                    ElseIf MachineNumber = 2 Then
                        DispPanel.s1(i) = Val(Mid$(DispPanel.lstMat.ListItems(DispPanel.lstMat.SelectedItem.Index).ListSubItems(4).Text, 1 + 2 * i, 2))
                    End If
                Next i

            End If
        End If

    Else
    End If

    DispPanel.cmbMatType.Locked = True
End Function

Public Function LoadMat()
    'функция за зареждане на комбото с видове материали
    
    DispPanel.btnSvExp.Visible = False
    
    For i = 0 To 5
        DispPanel.s1(i).Visible = False
    Next i
    
    If DispPanel.cmbMatType.ListIndex > -1 Then
        If DispPanel.cmbMatType.Text = uniIM Then

            For i = 0 To ns1 - 1
                DispPanel.s1(i).Visible = True
                DispPanel.s1(i).Caption = uniFlow & i + 1
            Next i

            DispPanel.lblMatLoad.Visible = True
            DispPanel.btnSvExp.Visible = False
        Else
        End If
        
        If DispPanel.cmbMatType.Text = uniConMat Then

            For i = 0 To ns3 - 1
                DispPanel.s1(i).Visible = True
                DispPanel.s1(i).Caption = uniCem & i + 1
            Next i

            DispPanel.lblMatLoad.Visible = True
            DispPanel.btnSvExp.Visible = False
        Else
        End If
        
        If DispPanel.cmbMatType.Text = uniWat Then

            For i = 0 To ns2 - 1
                DispPanel.s1(i).Visible = True
                DispPanel.s1(i).Caption = uniTank & i + 1
            Next i

            DispPanel.lblMatLoad.Visible = True
            DispPanel.btnSvExp.Visible = False
        Else
        End If
        
        If DispPanel.cmbMatType.Text = uniChem Then

            For i = 0 To ns4 - 1
                DispPanel.s1(i).Visible = True
                DispPanel.s1(i).Caption = uniPump & i + 1
            Next i

            DispPanel.lblMatLoad.Visible = True
            DispPanel.btnSvExp.Visible = False
        Else
        End If
        
        If DispPanel.cmbMatType.Text = uniOther Then

            For i = 0 To 5
                DispPanel.s1(i).Visible = False
            Next i

            DispPanel.lblMatLoad.Visible = False
            DispPanel.btnSvExp.Visible = True
        Else
        End If

    Else
    End If

End Function

Public Function LoadDlvrSup()
    'функция за избор на доставчик от екран "доставки..." и зареждане на данните
    
    If frmAddDlvr.cmbDlvrSupName.ListIndex > -1 Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
    
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        Set rs = cn.Execute("SELECT s_num, s_bg FROM suppliers WHERE s_name = '" & frmAddDlvr.cmbDlvrSupName.Text & "';")
    
        frmAddDlvr.txtDlvrSup.Text = Format(rs!s_num, "0000")
        frmAddDlvr.txtDlvrSupBG.Text = rs!s_bg
    
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------

    Else
    End If

End Function

Public Function LoadDlvrSupBG()
    'функция за търсене на доставчик по булстат от екран "доставки..." и зареждане на данните
    
    If Len(frmAddDlvr.txtDlvrSupBG.Text) > 0 Then
    
        '------------------------------Start PostgreSQL----------------------------------
        Dim cn As ADODB.Connection

        Dim rs As Recordset
    
        Set cn = New ADODB.Connection
        cn.ConnectionTimeout = 10
        cn.Open ConStr
        MousePointer = vbHourglass
        
        Set rs = cn.Execute("SELECT s_num, s_name FROM suppliers WHERE s_bg = '" & frmAddDlvr.txtDlvrSupBG.Text & "';")
    
        frmAddDlvr.txtDlvrSup.Text = Format(rs!s_num, "0000")
        frmAddDlvr.cmbDlvrSupName.Text = rs!s_name
    
        rs.Close
        Set rs = Nothing
        cn.Close
        MousePointer = vbDefault
        Set cn = Nothing
        '--------------------------End PostgreSQL------------------------------------------
    
    Else
    End If

End Function

Public Function BtnFillForm1()
    'попълване на форма 1 от старите експедиции
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

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
        prntForm1btn.txtOrdVol.Text = rDs(rsNew1!ord_q) 'общо количество по заявката
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
            ret = r + 1
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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet   As Boolean

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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    'показване на редактора за бележки ако е включена опцията
    If ShowEditor = True Then
        prntForm1btn.Show
    Else
        Call PrintBtnForm1(prntForm1btn)
    End If
    
End Function

Public Sub PrintBtnForm1(frm As Form)
    'принтиране на форма 1
    
    MousePointer = vbHourglass
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Dim ctr As Control
     
'    Printer.ScaleMode = 1
'    frmPrint.ScaleMode = 1
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    Printer.PrintQuality = 600
    
    frmPrint.pbPrint.Width = Printer.Width
    frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 1.41
    
    For i = 1 To numSheetsForm1
        frmPrint.pbPrint.Line (50, 50)-(frmPrint.pbPrint.Width - 200, 50)
        frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height - 100)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
        frmPrint.pbPrint.Line (50, 50)-(50, frmPrint.pbPrint.Height - 100)
        frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height / 2)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height / 2)
        frmPrint.pbPrint.Line (frmPrint.pbPrint.Width - 200, 50)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    
        If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

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
    
        If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
        PrintThePicture frmPrint, frmPrint.pbPrint, 96, 350, 300
        
        If i < numSheetsForm1 Then
            Printer.NewPage
        Else
        End If
    Next i
    
    MousePointer = vbDefault
    
    Printer.EndDoc
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Max
    Unload frmPrint
    Unload frm
End Sub

Public Function BtnFillForm2()
    'попълване на форма 2 от старите експедиции
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    Dim ResForm2 As Result

    Set ResForm2 = New Result
    
    Dim TotalIMKGzfff(0 To 5)   As Single

    Dim TotalIMKGifff(0 To 5)   As Single

    Dim TotalCemKGzfff(0 To 3)  As Single

    Dim TotalCemKGifff(0 To 3)  As Single

    Dim TotalWatKGzfff(0 To 1)  As Single

    Dim TotalWatKGifff(0 To 1)  As Single

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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

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
    
'    ns1 = CInt(frmParam.txtNumIMSilos.Text)
'    ns2 = CInt(frmParam.txtNumWaterSilos.Text)
'    ns3 = CInt(frmParam.txtNumCementSilos.Text)
'    ns4 = CInt(frmParam.txtNumChemSilos.Text)

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

    For i = 0 To 5
        TotalIMKGzfff(i) = 0
        TotalIMKGifff(i) = 0
    Next i

    For i = 0 To 3
        TotalCemKGzfff(i) = 0
        TotalCemKGifff(i) = 0
    Next i

    For i = 0 To 1
        TotalWatKGzfff(i) = 0
        TotalWatKGifff(i) = 0
    Next i

    For i = 0 To 5
        TotalChemKGzfff(i) = 0
        TotalChemKGifff(i) = 0
    Next i

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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet   As Boolean

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
        
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    'показване на редактора за бележки ако е включена опцията
    If ShowEditor = True Then
        prntForm2btn.Show
    Else
        Call PrintBtnForm2(prntForm2btn)
    End If
End Function

Public Sub PrintBtnForm2(frm As Form)
    'принтиране на форма 2
    
    MousePointer = vbHourglass
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrint.pbPrint.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    
    frmPrint.pbPrint.Width = Printer.Width
    frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 1.41
    
    For i = 1 To numSheetsForm2
        frmPrint.pbPrint.Line (50, 50)-(frmPrint.pbPrint.Width - 200, 50)
        frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height - 100)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
        frmPrint.pbPrint.Line (50, 50)-(50, frmPrint.pbPrint.Height - 100)
        frmPrint.pbPrint.Line (frmPrint.pbPrint.Width - 200, 50)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    
        If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

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
    
        If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
        PrintThePicture frmPrint, frmPrint.pbPrint, 96, 350, 300
        
        If i < numSheetsForm2 Then
            Printer.NewPage
        Else
        End If
    Next i
    
    MousePointer = vbDefault
    
    Printer.EndDoc
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Max
    Unload frmPrint
    Unload frm
End Sub

Public Function BtnFillForm3()
    'попълване на форма 3 от старите експедиции
    
    frmPrint.Show
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Min
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

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

    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    MousePointer = vbHourglass
    
    'зареждане от регистъра на разрешението за визуализация на реалното количество произведен бетон върху експедиционната бележка
    Dim PrevSet   As Boolean

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
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    'показване на редактора за бележки ако е включена опцията
    If ShowEditor = True Then
        prntForm3btn.Show
    Else
        Call PrintBtnForm3(prntForm3btn)
    End If
End Function

Public Sub PrintBtnForm3(frm As Form)
    'принтиране на форма 3
    
    MousePointer = vbHourglass
    
    If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
    Dim ctr As Control
     
    frmPrint.pbPrint.ScaleMode = 1
    
    Printer.Orientation = 1
    Printer.PaperSize = vbPRPSA4
    
    frmPrint.pbPrint.Width = Printer.Width
    frmPrint.pbPrint.Height = frmPrint.pbPrint.Width * 1.41
    
    For i = 1 To numSheetsForm3
        frmPrint.pbPrint.Line (50, 50)-(frmPrint.pbPrint.Width - 200, 50)
        frmPrint.pbPrint.Line (50, frmPrint.pbPrint.Height - 100)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
        frmPrint.pbPrint.Line (50, 50)-(50, frmPrint.pbPrint.Height - 100)
        frmPrint.pbPrint.Line (frmPrint.pbPrint.Width - 200, 50)-(frmPrint.pbPrint.Width - 200, frmPrint.pbPrint.Height - 100)
    
        If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

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
    
        If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10
    
        Call PrintRTF(prntForm3btn.Confirmity, 850, 7300, 800, 300)
    
        If frmPrint.barPrint.Value < frmPrint.barPrint.Max Then frmPrint.barPrint.Value = frmPrint.barPrint.Value + 10

        PrintThePicture frmPrint, frmPrint.pbPrint, 96, 350, 300
        
        If i < numSheetsForm3 Then
            Printer.NewPage
        Else
        End If
    Next i
    
    MousePointer = vbDefault
    
    Printer.EndDoc
    
    frmPrint.barPrint.Value = frmPrint.barPrint.Max
    Unload frmPrint
    Unload frm
End Sub

Public Function OpenAbout()
    'зареждане на меню За програмата...
    
    DispPanel.frAbout.Left = 120
    DispPanel.frAbout.Top = DispPanel.Height \ 2 - FrmPlace
End Function
