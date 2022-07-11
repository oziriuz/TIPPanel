Attribute VB_Name = "Module1"
    Public Const MachineNumber = 2

    Public I As Integer
    Dim n As Integer
    Public rfornreq As Integer
    Public streq As Integer
    Public finreq As Integer
    Public mixmix As Integer
    Public btnopen As Integer
    Public iauto As Integer
    Public iman As Integer
    Public iavar As Integer
    Public iskd As Integer
    Public iskf As Integer
    Public iskw As Integer
    Public isku As Integer
    Public btnskpause As Integer
    Public iopened As Integer
    Public iavaria As Integer
    Public avstop As Integer
    Public Stat As String
    Public stint As Long
    Public stv1 As String
    Public stv2 As String
    Public stv3 As String
    Public stv4 As String
    Public stLentaa As Integer
    Public stVoda As Integer
    Public stCiment As Integer
    Public stHimiq As Integer
    
    Public NowMixing As Boolean

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'---------------------------------


Public Function DecToBin(ByVal DeciValue As Single, Optional NoOfBits As Integer = 8, Optional chem As Boolean = False) As String
    Dim I As Integer
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
    
    For I = 0 To (NoOfBits - 1)

        If (DeciValue < 2 ^ I) Then
            DecToBin = "0" & DecToBin
        Else
            bay = newDec And 2 ^ I
            DecToBin = CStr(bay / 2 ^ I) & DecToBin
        End If

    Next I
End Function

Public Function BinToDec(Binary As String) As Long
    Dim n As Long
    Dim s As Integer

    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ _
            (s - 1)))
    Next s

    BinToDec = n
End Function

Public Function Mantisse(Binary As String) As Single
    Dim M1, M2, M3, M4, M5 As Single

    M1 = ((((Mid$(Binary, 23, 1) / 2 + Mid$(Binary, 22, 1)) / 2 + Mid$(Binary, 21, 1)) / 2 + Mid$(Binary, 20, 1)) / 2 + Mid$(Binary, 19, 1)) / 2
    M2 = (((((M1 + Mid(Binary, 18, 1)) / 2 + Mid$(Binary, 17, 1)) / 2 + Mid$(Binary, 16, 1)) / 2 + Mid$(Binary, 15, 1)) / 2 + Mid$(Binary, 14, 1)) / 2
    M3 = (((((M2 + Mid$(Binary, 13, 1)) / 2 + Mid$(Binary, 12, 1)) / 2 + Mid$(Binary, 11, 1)) / 2 + Mid$(Binary, 10, 1)) / 2 + Mid$(Binary, 9, 1)) / 2
    M4 = (((((M3 + Mid$(Binary, 8, 1)) / 2 + Mid$(Binary, 7, 1)) / 2 + Mid$(Binary, 6, 1)) / 2 + Mid$(Binary, 5, 1)) / 2 + Mid$(Binary, 4, 1)) / 2
    M5 = (((M4 + Mid$(Binary, 3, 1)) / 2 + Mid$(Binary, 2, 1)) / 2 + Mid$(Binary, 1, 1)) / 2
    Mantisse = M5
End Function

Public Function IEEE754(BCD As Long) As Variant
    Dim Bin As String
    Dim SignBin As String
    Dim ExpoBin As String
    Dim MantBin As String
    Dim Sign As Integer
    Dim Expo As Variant
    Dim Mant As Single
    
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
    Dim M(1 To 23) As Integer
    Dim ExpoDec As Single
    Dim ToIEEE754Bin As String
    Dim MantCalc As Single
    Dim SignBit As Long
    Dim BinPrep As String
    Dim ExpoBit As Long
    Dim ExpoReady As Long
    Dim ExpoDiv As Single
    Dim MantPrep As Single
    Dim MantisseReady As String
    Dim MantReduct As Single
    
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
        I = 0
        Do While Mid$(BinPrep, 1 + I, 1) = 0
            counter = counter + 1
            I = I + 1
            If I = 32 Then
                counter = counter - 1
                GoTo Jump
            Else
            End If
        Loop
    Else
        I = 0
        Do While 2 ^ counter > realNum
            counter = counter - 1
            I = I + 1
            If I = 32 Then
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
    I = 1
    For I = 1 To 23
        If (MantCalc * 2) > 1 Then
            M(I) = "1"
            MantCalc = (2 * MantCalc) - MantReduct
            GoTo FlagOut
        Else
        End If
        
        If (MantCalc * 2) = 1 Then
            M(I) = "1"
            MantCalc = 0
            GoTo FlagOut
        Else
        End If
        
        If (MantCalc * 2) < 1 And MantCalc > 0 Then
            M(I) = "0"
            MantCalc = 2 * MantCalc
            GoTo FlagOut
        Else
        End If
        
        If MantCalc < 0 Then
            M(I) = "0"
            GoTo FlagOut
        Else
        End If
FlagOut:
    Next I
    MantisseReady = M(1) & M(2) & M(3) & M(4) & M(5) & M(6) & M(7) & M(8) & M(9) & M(10) & M(11) & M(12) & M(13) & M(14) & M(15) & M(16) & M(17) & M(18) & M(19) & M(20) & M(21) & M(22) & M(23)
    ToIEEE754Bin = SignBit & ExpoReady & MantisseReady
    ToIEEE754 = BinToDec(ToIEEE754Bin)
End Function

Public Function ARound(ByVal MyNumber, ByVal Deci)
      ARound = Int(MyNumber * 10 ^ Deci + 1 / 2) / 10 ^ Deci
End Function

Public Function GetStat()
    
    Stat = avstop & iavaria & iskf & iopened & btnskpause & isku & iskw & iskd & iavar & iman & iauto & btnopen & mixmix & finreq & streq & rfornreq
    stint = BinToDec(Stat)
    frmSim.status.ItemValue = stint

    Stat = DecToBin(stint, 16)
    
    rfornreq = Mid$(Stat, 16, 1)
    streq = Mid$(Stat, 15, 1)
    finreq = Mid$(Stat, 14, 1)
    mixmix = Mid$(Stat, 13, 1)
    btnopen = Mid$(Stat, 12, 1)
    iauto = Mid$(Stat, 11, 1)
    iman = Mid$(Stat, 10, 1)
    iavar = Mid$(Stat, 9, 1)
    iskd = Mid$(Stat, 8, 1)
    iskw = Mid$(Stat, 7, 1)
    isku = Mid$(Stat, 6, 1)
    btnskpause = Mid$(Stat, 5, 1)
    iopened = Mid$(Stat, 4, 1)
    iskf = Mid$(Stat, 3, 1)
    iavaria = Mid$(Stat, 2, 1)
    avstop = Mid$(Stat, 1, 1)
    
    stv1 = 0 & 0 & 0 & 0 & stLentaa & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0
    frmSim.cio1001.ItemValue = BinToDec(stv1)
    
    stv2 = 0 & 0 & 0 & 0 & stVoda & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0
    frmSim.cio1002.ItemValue = BinToDec(stv2)
    
    stv3 = 0 & 0 & 0 & 0 & stCiment & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0
    frmSim.cio1003.ItemValue = BinToDec(stv3)
    
    stv4 = 0 & 0 & 0 & 0 & stHimiq & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0 & 0
    frmSim.cio1004.ItemValue = BinToDec(stv4)
    
For I = 0 To 16
    If iauto = 1 Then
        frmSim.indautomode.Caption = "автоматичен режим"
        frmSim.indautomode.Refresh
        iman = 0
        iavar = 0
    Else
        frmSim.indautomode.Caption = ""
        frmSim.indautomode.Refresh
    End If
    
    If iman = 1 Then
        frmSim.indmanualmode.Caption = "ръчен режим"
        frmSim.indmanualmode.Refresh
        iauto = 0
        iavar = 0
    Else
        frmSim.indmanualmode.Caption = ""
        frmSim.indmanualmode.Refresh
    End If
    
    If iavar = 1 Then
        frmSim.indavariamode.Caption = "авариен режим"
        frmSim.indavariamode.Refresh
        iauto = 0
        iman = 0
    Else
        frmSim.indavariamode.Caption = ""
        frmSim.indavariamode.Refresh
    End If

    If iavaria = 1 Then
        frmSim.indavaria.Caption = "авария"
        mixmix = 0
        rfornreq = 0
    Else
        frmSim.indavaria.Caption = ""
    End If
    
    If avstop = 1 Then
        frmSim.indemgstop.Caption = "авариен стоп"
        mixmix = 0
        rfornreq = 0
        btnskpause = 0
    Else
        frmSim.indemgstop.Caption = ""
    End If

    If streq = 1 Then
        frmSim.okreadnewreq.Caption = "заявка стартирана"
        rfornreq = 0
        frmSim.dorec.Visible = False
    Else
        frmSim.okreadnewreq.Caption = ""
    End If
    
    If mixmix = 1 Then
        frmSim.mixermix.Caption = "миксер включен"
        frmSim.turnmix.Caption = "изключи миксера"
    ElseIf mixmix = 0 Then
        frmSim.mixermix.Caption = "миксер изключен"
        frmSim.turnmix.Caption = "включи миксера"
    End If
    
    If mixmix = 1 And iauto = 1 And streq = 0 Then
        frmSim.dorec.Visible = True
        If iskd = 1 And avstop = 0 And iavaria = 0 Then
            rfornreq = 1
        Else
        End If
    Else
        frmSim.dorec.Visible = False
        rfornreq = 0
    End If
    
    If rfornreq = 1 Then
        frmSim.readyfornew.Caption = "готов за заявка"
    Else
        frmSim.readyfornew.Caption = ""
    End If
    
    If iskd = 1 Then
        frmSim.indskipdown.Caption = "скип долу"
        iskw = 0
        isku = 0
    Else
        frmSim.indskipdown.Caption = ""
        rfornereq = 0
    End If
    
    If iskf = 1 Then
        frmSim.indskipfull.Caption = "количка пълна"
        rfornreq = 0
    Else
        frmSim.indskipfull.Caption = ""
    End If
    
    If iskw = 1 Then
        frmSim.indskipwait.Caption = "скип чака"
        iskd = 0
        isku = 0
        rfornreq = 0
    Else
        frmSim.indskipwait.Caption = ""
    End If

    If btnskpause = 1 Then
        frmSim.skipwait.Caption = "скип пауза"
        iskd = 0
        isku = 0
        iskw = 0
        rfornreq = 0
    Else
        frmSim.skipwait.Caption = ""
    End If

    If isku = 1 Then
        frmSim.indskipup.Caption = "скип горе"
        iskd = 0
        iskw = 0
        btnskpause = 0
        rfornreq = 0
    Else
        frmSim.indskipup.Caption = ""
    End If

    If iopened = 1 Then
        frmSim.indmixopened.Caption = "клапа отворена"
    Else
        frmSim.indmixopened.Caption = "клапа затворена"
    End If
    
    If finreq = 1 Then
        frmSim.finishedreq.Caption = "заявка завършена"
    Else
        frmSim.finishedreq.Caption = ""
    End If
Next I
    Stat = avstop & iavaria & iskf & iopened & btnskpause & isku & iskw & iskd & iavar & iman & iauto & btnopen & mixmix & finreq & streq & rfornreq
    stint = BinToDec(Stat)
    
    frmSim.status.ItemValue = stint
End Function

Public Function OpenMix()
    Call GetStat
    If btnopen = 0 Then
        btnopen = 1
    Else
        btnopen = 0
    End If
    Call GetStat
    Sleep 3000
    If btnopen = 0 Then iopened = 0
    If btnopen = 1 Then iopened = 1
    Call GetStat
End Function

Public Function AutoStart()
    Dim ttt As String
    Dim hexdm(0 To 49) As String
    Dim chem(0 To 6) As Single
    Dim counter As Integer
    Dim tim As Integer
    Dim timm As Integer
    Dim cycle As Integer
    Dim watprob As Long
    Dim watprob1 As Long
    Dim nss1 As Integer
    Dim nss3 As Integer
    Dim nss4 As Integer
    NowMixing = True
    tim = 3
    timm = 1
    frmSim.recmix.Text = frmSim.dm1000(0).ItemValue
    
    frmSim.recmix.Refresh
    frmSim.dm1000(0).ItemValue = 0
    frmSim.dm500.ItemValue = 0

    frmSim.resreadymix.Text = 0
    frmSim.finished.Caption = ""
    frmSim.watcc.ItemValue = 0
    
    For b = 1 To 49
        hexdm(b) = Hex$(frmSim.dm1000(b).ItemValue)
        Select Case Len(hexdm(b))
            Case 3
                hexdm(b) = "0" & hexdm(b)
            Case 2
                hexdm(b) = "00" & hexdm(b)
            Case 1
                hexdm(b) = "000" & hexdm(b)
            Case 0
                hexdm(b) = "0000"
        End Select
    Next b
    
    streq = 1
    rfornreq = 0
    frmSim.dorec.Visible = False
    frmSim.dorec.Refresh
    Call GetStat
    For cycle = 1 To CInt(frmSim.recmix.Text)
        For t = 0 To 4
            frmSim.resaggr1(t).Text = "0"
            frmSim.resaggr1(t).Refresh
        Next t
        frmSim.reswat.Text = "0"
        frmSim.reswat.Refresh
        For t = 0 To 3
            frmSim.rescem1(t).Text = "0"
            frmSim.rescem1(t).Refresh
        Next t
        For t = 0 To 5
            frmSim.reschem1(t).Text = "0"
            frmSim.reschem1(t).Refresh
        Next t
        
        frmSim.recim1.Text = "0"
        frmSim.recim1.Refresh
        frmSim.recim2.Text = "0"
        frmSim.recim2.Refresh
        frmSim.recim3.Text = "0"
        frmSim.recim3.Refresh
        frmSim.recim4.Text = "0"
        frmSim.recim4.Refresh
        frmSim.recim5.Text = "0"
        frmSim.recim5.Refresh
        frmSim.recwat.Text = "0"
        frmSim.recwat.Refresh
        frmSim.reccem1.Text = "0"
        frmSim.reccem1.Refresh
        frmSim.reccem2.Text = "0"
        frmSim.reccem2.Refresh
        frmSim.reccem3.Text = "0"
        frmSim.reccem3.Refresh
        frmSim.reccem4.Text = "0"
        frmSim.reccem4.Refresh
        frmSim.recchem1.Text = "0"
        frmSim.recchem1.Refresh
        frmSim.recchem2.Text = "0"
        frmSim.recchem2.Refresh
        frmSim.recchem3.Text = "0"
        frmSim.recchem3.Refresh
        frmSim.recchem4.Text = "0"
        frmSim.recchem4.Refresh
        frmSim.recchem5.Text = "0"
        frmSim.recchem5.Refresh
        frmSim.recchem6.Text = "0"
        frmSim.recchem6.Refresh
        
        counter = 0
        nc = 1
        For v = 1 To 16
            DoEvents
            ttt = (hexdm(nc + 2) & hexdm(nc + 1))
            Select Case frmSim.dm1000(nc).ItemValue
                Case frmSim.idim(0).ItemValue
                    frmSim.recim1.Text = Val(frmSim.recim1.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.recim1.Refresh
                    Do While Val(frmSim.resaggr1(0).Text) < Val(frmSim.recim1.Text) + Int((40 * Rnd) - 10)
                        DoEvents
                        frmSim.resaggr1(0).Text = Val(frmSim.resaggr1(0).Text) + 1
                        frmSim.resaggr1(0).Refresh
                        frmSim.resaggrall.Text = Val(frmSim.resaggrall.Text) + 1
                        frmSim.resaggrall.Refresh
                        frmSim.BCDAggr.ItemValue = ToIEEE754(Val(frmSim.resaggrall.Text))
                        Sleep timm
                    Loop
                Case frmSim.idim(1).ItemValue
                    frmSim.recim2.Text = Val(frmSim.recim2.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.recim2.Refresh
                    Do While Val(frmSim.resaggr1(1).Text) < Val(frmSim.recim2.Text) + Int((40 * Rnd) - 10)
                        DoEvents
                        frmSim.resaggr1(1).Text = Val(frmSim.resaggr1(1).Text) + 1
                        frmSim.resaggr1(1).Refresh
                        frmSim.resaggrall.Text = Val(frmSim.resaggrall.Text) + 1
                        frmSim.resaggrall.Refresh
                        frmSim.BCDAggr.ItemValue = ToIEEE754(Val(frmSim.resaggrall.Text))
                        Sleep timm
                    Loop
                Case frmSim.idim(2).ItemValue
                    frmSim.recim3.Text = Val(frmSim.recim3.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.recim3.Refresh
                    Do While Val(frmSim.resaggr1(2).Text) < Val(frmSim.recim3.Text) + Int((40 * Rnd) - 10)
                        DoEvents
                        frmSim.resaggr1(2).Text = Val(frmSim.resaggr1(2).Text) + 1
                        frmSim.resaggr1(2).Refresh
                        frmSim.resaggrall.Text = Val(frmSim.resaggrall.Text) + 1
                        frmSim.resaggrall.Refresh
                        frmSim.BCDAggr.ItemValue = ToIEEE754(Val(frmSim.resaggrall.Text))
                        Sleep timm
                    Loop
                Case frmSim.idim(3).ItemValue
                    frmSim.recim4.Text = Val(frmSim.recim4.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.recim4.Refresh
                    Do While Val(frmSim.resaggr1(3).Text) < Val(frmSim.recim4.Text) + Int((40 * Rnd) - 10)
                        DoEvents
                        frmSim.resaggr1(3).Text = Val(frmSim.resaggr1(3).Text) + 1
                        frmSim.resaggr1(3).Refresh
                        frmSim.resaggrall.Text = Val(frmSim.resaggrall.Text) + 1
                        frmSim.resaggrall.Refresh
                        frmSim.BCDAggr.ItemValue = ToIEEE754(Val(frmSim.resaggrall.Text))
                        Sleep timm
                    Loop
                Case frmSim.idim(4).ItemValue
                    frmSim.recim5.Text = Val(frmSim.recim5.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.recim5.Refresh
                    Do While Val(frmSim.resaggr1(4).Text) < Val(frmSim.recim5.Text) + Int((40 * Rnd) - 10)
                        DoEvents
                        frmSim.resaggr1(4).Text = Val(frmSim.resaggr1(4).Text) + 1
                        frmSim.resaggr1(4).Refresh
                        frmSim.resaggrall.Text = Val(frmSim.resaggrall.Text) + 1
                        frmSim.resaggrall.Refresh
                        frmSim.BCDAggr.ItemValue = ToIEEE754(Val(frmSim.resaggrall.Text))
                        Sleep timm
                    Loop
                Case frmSim.idwat.ItemValue
                    If cycle = 1 Then
                        Do While watprob1 <> CLng("&H" & ttt)
                            DoEvents
                            frmSim.watcc.ItemValue = CLng("&H" & ttt)
                            watprob1 = frmSim.watcc.ItemValue
                        Loop
                    End If
                    Sleep 777
                    watprob = IEEE754(frmSim.watcc.ItemValue)
                    frmSim.recwat.Text = ARound(watprob, 0)
                    frmSim.recwat.Refresh
                    Do While Val(frmSim.reswat.Text) < ARound(IEEE754(frmSim.watcc.ItemValue), 0) + Int((30 * Rnd) - 7)
                        DoEvents
                        frmSim.reswat.Text = Val(frmSim.reswat.Text) + 1
                        frmSim.reswat.Refresh
                        frmSim.reswatall.Text = Val(frmSim.reswatall.Text) + 1
                        frmSim.reswatall.Refresh
                        frmSim.BCDWat.ItemValue = ToIEEE754(Val(frmSim.reswatall.Text))
                        Sleep tim
                    Loop
                Case frmSim.idcem(0).ItemValue
                    frmSim.reccem1.Text = Val(frmSim.reccem1.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.reccem1.Refresh
                    Do While Val(frmSim.rescem1(0).Text) < Val(frmSim.reccem1.Text) + Int((30 * Rnd) - 5)
                        DoEvents
                        frmSim.rescem1(0).Text = Val(frmSim.rescem1(0).Text) + 1
                        frmSim.rescem1(0).Refresh
                        frmSim.rescemall.Text = Val(frmSim.rescemall.Text) + 1
                        frmSim.rescemall.Refresh
                        frmSim.BCDCem.ItemValue = ToIEEE754(Val(frmSim.rescemall.Text))
                        Sleep tim
                    Loop
                Case frmSim.idcem(1).ItemValue
                    frmSim.reccem2.Text = Val(frmSim.reccem2.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.reccem2.Refresh
                    Do While Val(frmSim.rescem1(1).Text) < Val(frmSim.reccem2.Text) + Int((30 * Rnd) - 5)
                        DoEvents
                        frmSim.rescem1(1).Text = Val(frmSim.rescem1(1).Text) + 1
                        frmSim.rescem1(1).Refresh
                        frmSim.rescemall.Text = Val(frmSim.rescemall.Text) + 1
                        frmSim.rescemall.Refresh
                        frmSim.BCDCem.ItemValue = ToIEEE754(Val(frmSim.rescemall.Text))
                        Sleep tim
                    Loop
                Case frmSim.idcem(2).ItemValue
                    frmSim.reccem3.Text = Val(frmSim.reccem3.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.reccem3.Refresh
                    Do While Val(frmSim.rescem1(2).Text) < Val(frmSim.reccem3.Text) + Int((30 * Rnd) - 5)
                        DoEvents
                        frmSim.rescem1(2).Text = Val(frmSim.rescem1(2).Text) + 1
                        frmSim.rescem1(2).Refresh
                        frmSim.rescemall.Text = Val(frmSim.rescemall.Text) + 1
                        frmSim.rescemall.Refresh
                        frmSim.BCDCem.ItemValue = ToIEEE754(Val(frmSim.rescemall.Text))
                        Sleep tim
                    Loop
                Case frmSim.idcem(3).ItemValue
                    frmSim.reccem4.Text = Val(frmSim.reccem4.Text) + ARound(IEEE754(CLng("&H" & ttt)), 0)
                    frmSim.reccem4.Refresh
                    Do While Val(frmSim.rescem1(3).Text) < Val(frmSim.reccem4.Text) + Int((30 * Rnd) - 5)
                        DoEvents
                        frmSim.rescem1(3).Text = Val(frmSim.rescem1(3).Text) + 1
                        frmSim.rescem1(3).Refresh
                        frmSim.rescemall.Text = Val(frmSim.rescemall.Text) + 1
                        frmSim.rescemall.Refresh
                        frmSim.BCDCem.ItemValue = ToIEEE754(Val(frmSim.rescemall.Text))
                        Sleep tim
                    Loop
                Case frmSim.idchem(0).ItemValue
                    frmSim.recchem1.Text = CSng(frmSim.recchem1.Text) + ARound(IEEE754(CLng("&H" & ttt)), 2)
                    frmSim.recchem1.Refresh
                    chem(1) = 0
                    Do While chem(1) < ARound(IEEE754(CLng("&H" & ttt)) + ((0.3 * Rnd) - 0.1), 2)
                        DoEvents
                        chem(1) = ARound(chem(1) + 0.01, 2)
                        frmSim.reschem1(0).Text = ARound(chem(1), 2)
                        frmSim.reschem1(0).Refresh
                        chem(0) = ARound(chem(0) + 0.01, 2)
                        frmSim.reschemall.Text = chem(0)
                        frmSim.reschemall.Refresh
                        frmSim.BCDChem.ItemValue = ToIEEE754(chem(0))
                        Sleep tim
                    Loop
                Case frmSim.idchem(1).ItemValue
                    frmSim.recchem2.Text = CSng(frmSim.recchem2.Text) + ARound(IEEE754(CLng("&H" & ttt)), 2)
                    frmSim.recchem2.Refresh
                    chem(2) = 0
                    Do While chem(2) < ARound(IEEE754(CLng("&H" & ttt)) + ((0.3 * Rnd) - 0.1), 2)
                        DoEvents
                        chem(2) = ARound(chem(2) + 0.01, 2)
                        frmSim.reschem1(1).Text = ARound(chem(2), 2)
                        frmSim.reschem1(1).Refresh
                        chem(0) = ARound(chem(0) + 0.01, 2)
                        frmSim.reschemall.Text = chem(0)
                        frmSim.reschemall.Refresh
                        frmSim.BCDChem.ItemValue = ToIEEE754(chem(0))
                        Sleep tim
                    Loop
                Case frmSim.idchem(2).ItemValue
                    frmSim.recchem3.Text = CSng(frmSim.recchem3.Text) + ARound(IEEE754(CLng("&H" & ttt)), 2)
                    frmSim.recchem3.Refresh
                    chem(3) = 0
                    Do While chem(3) < ARound(IEEE754(CLng("&H" & ttt)) + ((0.3 * Rnd) - 0.1), 2)
                        DoEvents
                        chem(3) = ARound(chem(3) + 0.01, 2)
                        frmSim.reschem1(2).Text = ARound(chem(3), 2)
                        frmSim.reschem1(2).Refresh
                        chem(0) = ARound(chem(0) + 0.01, 2)
                        frmSim.reschemall.Text = chem(0)
                        frmSim.reschemall.Refresh
                        frmSim.BCDChem.ItemValue = ToIEEE754(chem(0))
                        Sleep tim
                    Loop
                Case frmSim.idchem(3).ItemValue
                    frmSim.recchem4.Text = CSng(frmSim.recchem4.Text) + ARound(IEEE754(CLng("&H" & ttt)), 2)
                    frmSim.recchem4.Refresh
                    chem(4) = 0
                    Do While chem(4) < ARound(IEEE754(CLng("&H" & ttt)), 2) + ((0.6 * Rnd) - 0.2)
                        DoEvents
                        chem(4) = ARound(chem(4) + 0.01, 2)
                        frmSim.reschem1(3).Text = ARound(chem(4), 2)
                        frmSim.reschem1(3).Refresh
                        chem(0) = ARound(chem(0) + 0.01, 2)
                        frmSim.reschemall.Text = chem(0)
                        frmSim.reschemall.Refresh
                        frmSim.BCDChem.ItemValue = ToIEEE754(chem(0))
                        Sleep tim
                    Loop
                Case frmSim.idchem(4).ItemValue
                    frmSim.recchem5.Text = CSng(frmSim.recchem5.Text) + ARound(IEEE754(CLng("&H" & ttt)), 2)
                    frmSim.recchem5.Refresh
                    chem(5) = 0
                    Do While chem(5) < ARound(IEEE754(CLng("&H" & ttt)), 2) + ((0.6 * Rnd) - 0.2)
                        DoEvents
                        chem(5) = ARound(chem(5) + 0.01, 2)
                        frmSim.reschem1(4).Text = ARound(chem(5), 2)
                        frmSim.reschem1(4).Refresh
                        chem(0) = ARound(chem(0) + 0.01, 2)
                        frmSim.reschemall.Text = chem(0)
                        frmSim.reschemall.Refresh
                        frmSim.BCDChem.ItemValue = ToIEEE754(chem(0))
                        Sleep tim
                    Loop
                Case frmSim.idchem(5).ItemValue
                    frmSim.recchem6.Text = CSng(frmSim.recchem6.Text) + ARound(IEEE754(CLng("&H" & ttt)), 2)
                    frmSim.recchem6.Refresh
                    chem(6) = 0
                    Do While chem(6) < ARound(IEEE754(CLng("&H" & ttt)), 2) + ((0.6 * Rnd) - 0.2)
                        DoEvents
                        chem(6) = ARound(chem(6) + 0.01, 2)
                        frmSim.reschem1(5).Text = ARound(chem(6), 2)
                        frmSim.reschem1(5).Refresh
                        chem(0) = ARound(chem(0) + 0.01, 2)
                        frmSim.reschemall.Text = chem(0)
                        frmSim.reschemall.Refresh
                        frmSim.BCDChem.ItemValue = ToIEEE754(chem(0))
                        Sleep tim
                    Loop
                Case "0"
                    If nc <= 39 Then
                        If frmSim.dm1000(nc + 3).ItemValue = 0 And frmSim.dm1000(nc + 6).ItemValue = 0 Then
                            frmSim.rectpour.Text = Hex(frmSim.dm1000(nc + 9).ItemValue) / 10
                            frmSim.rectpour.Refresh
                            frmSim.rectmix.Text = Hex(frmSim.dm1000(nc + 10).ItemValue) / 10
                            frmSim.rectmix.Refresh
                            GoTo done
                            Call frmSim.Timer1_Timer
                        Else
                        End If
                    Else
                    End If
            End Select
            nc = nc + 3
            DoEvents
        Next v
done:
        
        For k = 0 To 51
            frmSim.dm1100(k).ItemValue = 0
        Next k

'изпразване на везните
        stLentaa = 1
        Call GetStat
        Do While Val(frmSim.resaggrall.Text) > 0
            DoEvents
            frmSim.resaggrall.Text = Val(frmSim.resaggrall.Text) - 1
            frmSim.resaggrall.Refresh
            frmSim.BCDAggr.ItemValue = ToIEEE754(Val(frmSim.resaggrall.Text))
            Sleep timm
        Loop
        iskf = 1
        stLentaa = 0
        Call GetStat
        Sleep 300
        iskd = 0
        iskw = 1
        isku = 0
        Call GetStat
        Sleep 100
        iskw = 0
        isku = 1
        stCiment = 0
        Call GetStat
        Sleep 50
        Do While Val(frmSim.rescemall.Text) > 0
            DoEvents
            frmSim.rescemall.Text = Val(frmSim.rescemall.Text) - 1
            frmSim.rescemall.Refresh
            frmSim.BCDCem.ItemValue = ToIEEE754(Val(frmSim.rescemall.Text))
            Sleep tim
        Loop
        stCiment = 1
        stVoda = 0
        Call GetStat
        Do While Val(frmSim.reswatall.Text) > 0
            DoEvents
            frmSim.reswatall.Text = Val(frmSim.reswatall.Text) - 1
            frmSim.reswatall.Refresh
            frmSim.BCDWat.ItemValue = ToIEEE754(Val(frmSim.reswatall.Text))
            Sleep tim
        Loop
        stVoda = 1
        stHimiq = 0
        Call GetStat
        Do While chem(0) > 0
            DoEvents
            chem(0) = ARound(chem(0) - 0.01, 2)
            frmSim.reschemall.Text = chem(0)
            frmSim.reschemall.Refresh
            frmSim.BCDChem.ItemValue = ToIEEE754(chem(0))
            Sleep tim
        Loop
        stHimiq = 1
        Call GetStat
        Sleep (Val(frmSim.rectmix.Text) * 1000)
Report:
        Last = 1
        For z = 0 To Val(frmSim.NumIMSilos.ItemValue) - 1
AgrrRes:
            DoEvents
            frmSim.dm1100(Last).ItemValue = frmSim.idim(z).ItemValue
            kgval = ToIEEE754(frmSim.resaggr1(z).Text)
            If kgval <> 0 Then
                hexVal = Hex(kgval)
                hexval1 = Mid$(hexVal, 5, 8)
                hexval2 = Mid$(hexVal, 1, 4)
                KgVal1 = CInt("&H" & hexval1)
                KgVal2 = CInt("&H" & hexval2)
                frmSim.dm1100(Last + 1).ItemValue = KgVal1
                frmSim.dm1100(Last + 2).ItemValue = KgVal2
            End If
            If Val(frmSim.resaggr1(z).Text) <> 0 Then
            Do Until frmSim.dm1100(Last + 1).ItemValue <> 0 Or frmSim.dm1100(Last + 2).ItemValue <> 0
                frmSim.dm1100(Last + 1).ItemValue = KgVal1
                frmSim.dm1100(Last + 2).ItemValue = KgVal2
            Loop
            End If
            Last = Last + 3
        Next z
    
        frmSim.dm1100(Last).ItemValue = frmSim.idwat.ItemValue
        kgval = ToIEEE754(frmSim.reswat.Text)
        If kgval <> 0 Then
            hexVal = Hex(kgval)
            hexval1 = Mid$(hexVal, 5, 8)
            hexval2 = Mid$(hexVal, 1, 4)
            KgVal1 = CInt("&H" & hexval1)
            KgVal2 = CInt("&H" & hexval2)
            frmSim.dm1100(Last + 1).ItemValue = KgVal1
            frmSim.dm1100(Last + 2).ItemValue = KgVal2
        End If
        If Val(frmSim.reswat.Text) <> 0 Then
        Do Until frmSim.dm1100(Last + 1).ItemValue <> 0 Or frmSim.dm1100(Last + 2).ItemValue <> 0
            frmSim.dm1100(Last + 1).ItemValue = KgVal1
            frmSim.dm1100(Last + 2).ItemValue = KgVal2
        Loop
        End If
        Last = Last + 3
    
        For z = 0 To Val(frmSim.NumCementSilos.ItemValue) - 1
CemRes:
            DoEvents
            frmSim.dm1100(Last).ItemValue = frmSim.idcem(z).ItemValue
            kgval = ToIEEE754(frmSim.rescem1(z).Text)
            If kgval <> 0 Then
                hexVal = Hex(kgval)
                hexval1 = Mid$(hexVal, 5, 8)
                hexval2 = Mid$(hexVal, 1, 4)
                KgVal1 = CInt("&H" & hexval1)
                KgVal2 = CInt("&H" & hexval2)
                frmSim.dm1100(Last + 1).ItemValue = KgVal1
                frmSim.dm1100(Last + 2).ItemValue = KgVal2
            End If
            If Val(frmSim.rescem1(z).Text) <> 0 Then
            Do Until frmSim.dm1100(Last + 1).ItemValue <> 0 Or frmSim.dm1100(Last + 2).ItemValue <> 0
                frmSim.dm1100(Last + 1).ItemValue = KgVal1
                frmSim.dm1100(Last + 2).ItemValue = KgVal2
            Loop
            End If
            Last = Last + 3
        Next z
    
        For z = 0 To Val(frmSim.NumChemSilos.ItemValue) - 1
ChemRes:
            DoEvents
            frmSim.dm1100(Last).ItemValue = frmSim.idchem(z).ItemValue
            kgval = ToIEEE754(frmSim.reschem1(z).Text)
            If kgval <> 0 Then
                hexVal = Hex(kgval)
                hexval1 = Mid$(hexVal, 5, 8)
                hexval2 = Mid$(hexVal, 1, 4)
                KgVal1 = CInt("&H" & hexval1)
                KgVal2 = CInt("&H" & hexval2)
                frmSim.dm1100(Last + 1).ItemValue = KgVal1
                frmSim.dm1100(Last + 2).ItemValue = KgVal2
            End If
            If CSng(frmSim.reschem1(z).Text) <> 0 Then
            Do Until frmSim.dm1100(Last + 1).ItemValue <> 0 Or frmSim.dm1100(Last + 2).ItemValue <> 0
                frmSim.dm1100(Last + 1).ItemValue = KgVal1
                frmSim.dm1100(Last + 2).ItemValue = KgVal2
            Loop
            End If
            Last = Last + 3
        Next z
        
            
        For t = 0 To 4
            frmSim.resaggr1(t).Text = "0"
            frmSim.resaggr1(t).Refresh
        Next t
        frmSim.reswat.Text = "0"
        frmSim.reswat.Refresh
        For t = 0 To 3
            frmSim.rescem1(t).Text = "0"
            frmSim.rescem1(t).Refresh
        Next t
        For t = 0 To 5
            frmSim.reschem1(t).Text = "0"
            frmSim.reschem1(t).Refresh
        Next t
        
        Sleep 100
        
        frmSim.dm1100(0).ItemValue = CInt(cycle)
        frmSim.dm500.ItemValue = CInt(cycle)
        frmSim.dm501.ItemValue = CInt(cycle)
        
        Call OpenMix
        frmSim.resreadymix.Text = Val(frmSim.resreadymix.Text) + 1
        frmSim.resreadymix.Refresh
        Sleep (Val(frmSim.rectpour.Text) * 1000)
        Call OpenMix
        iskf = 0
        isku = 0
        iskw = 1
        iskd = 0
        Call GetStat
        Sleep 50
        iskd = 1
        iskw = 0
        Call GetStat
        Sleep 50
        DoEvents
    Next cycle
    streq = 0
    rfornreq = 1
    Call GetStat
    
'    frmSim.dm1000(0).ItemValue = 0
    frmSim.finished.Caption = "Експедиция готова!"
    frmSim.dorec.Visible = True
    frmSim.dorec.Refresh
    Sleep 1000
    frmSim.dm500.ItemValue = 0
End Function
