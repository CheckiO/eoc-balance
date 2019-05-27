Const resLevelChanged = 3
Const resLevelChanged2 = 6


Sub Button1_Click()
    Application.CalculateFull
End Sub
''''''
' BASE
''''''
' =IF(B4,cio2StepsChangesX(G3,A4,Conf!O$26,Conf!O$27,Conf!P$25,Conf!P$26,Conf!P$27),G3)

Function getConf(path As String) As Long
    getConf = ThisWorkbook.Worksheets("Conf").Range(path).Value
End Function

Function baseBetterNumbers(n As Long) As Long
    If n > 10000000 Then
        baseBetterNumbers = Round(n / 100000) * 100000
    ElseIf n > 1000000 Then
        baseBetterNumbers = Round(n / 10000) * 10000
    ElseIf n > 100000 Then
        baseBetterNumbers = Round(n / 1000) * 1000
    ElseIf n > 10000 Then
        baseBetterNumbers = Round(n / 100) * 100
    ElseIf n > 1000 Then
        baseBetterNumbers = Round(n / 10) * 10
    Else
        baseBetterNumbers = n
    End If
End Function

Function baseSimpleHalf(prev As Long, level As Long) As Long
    baseSimpleHalf = baseBetterNumbers(prev * 1.5)
End Function

Function baseSimpleThird(prev As Long, level As Long) As Long
    baseSimpleThird = baseBetterNumbers(prev * 1.3)
End Function

Function baseSimpleTwice(prev As Long, level As Long) As Long
    baseSimpleTwice = baseBetterNumbers(prev * 2)
End Function

Function baseTwoStepsChanges(prev As Long, level As Long, Optional step1 As Long = 0, Optional step2 As Long = 0) As Long
    If Not step1 Then
        step1 = resLevelChanged
    End If
    If Not step2 Then
        step2 = resLevelChanged2
    End If

    If level > step2 Then
        baseTwoStepsChanges = baseSimpleThird(prev, level)
    ElseIf level > step1 Then
        baseTwoStepsChanges = baseSimpleHalf(prev, level)
    Else
        baseTwoStepsChanges = baseSimpleTwice(prev, level)
    End If
End Function

Function baseTwoStepsChangesConf(prev As Long, level As Long, step1path As String, step2path As String) As Long
    Dim step1 As Long
    Dim step2 As Long

    step1 = getConf(step1path)
    step2 = getConf(step2path)

    If level > step2 Then
        baseTwoStepsChangesConf = baseSimpleThird(prev, level)
    ElseIf level > step1 Then
        baseTwoStepsChangesConf = baseSimpleHalf(prev, level)
    Else
        baseTwoStepsChangesConf = baseSimpleTwice(prev, level)
    End If
End Function

Function cio3StepsChangesX(prev As Long, level As Long, step1 As Long, step2 As Long, step3 As Long, coof0 As Double, coof1 As Double, coof2 As Double, coof3 As Double) As Long
    If level >= step3 Then
        baseTwoStepsChangesConf = baseBetterNumbers(prev * coof3)
    ElseIf level >= step2 Then
        baseTwoStepsChangesConf = baseBetterNumbers(prev * coof2)
    ElseIf level >= step1 Then
        baseTwoStepsChangesConf = baseBetterNumbers(prev * coof1)
    Else
        baseTwoStepsChangesConf = baseBetterNumbers(prev * coof0)
    End If
End Function

Function cio2StepsChangesX(prev As Long, level As Long, step1 As Long, step2 As Long, coof0 As Double, coof1 As Double, coof2 As Double) As Long
    If level >= step2 Then
        cio2StepsChangesX = baseBetterNumbers(prev * coof2)
    ElseIf level >= step1 Then
        cio2StepsChangesX = baseBetterNumbers(prev * coof1)
    Else
        cio2StepsChangesX = baseBetterNumbers(prev * coof0)
    End If
End Function

Function cioUpTime(coof As Double, ccLevel As Long, adPrice As Long) As Long
    cioUpTime = adPrice * coof
End Function

Function secToTime(sec As Long) As String
    Dim min As Long
    Dim hour As Long
    Dim days As Long
    min = WorksheetFunction.RoundDown(sec / 60, 0)
    sec = sec - min * 60
    hour = WorksheetFunction.RoundDown(min / 60, 0)
    min = min - hour * 60
    days = WorksheetFunction.RoundDown(hour / 24, 0)
    hour = hour - days * 24
    If days > 0 Then
        secToTime = CStr(days) & "d " & CStr(hour) & "h " & CStr(min) & "m"
    Else
        secToTime = CStr(hour) & "h " & CStr(min) & "m"
    End If

End Function


'''''''
'
' crystaliteFarm
'
'''''''

Function cioCFEnergy(prev As Long, level As Long) As Long
    cioCFEnergy = baseSimpleThird(prev, level)
End Function

Function cioCFFarm(prev As Long, level As Long) As Long
    cioCFFarm = baseTwoStepsChangesConf(prev, level, "G4", "G5")
End Function

Function cioCFFarmMax(prev As Long, level As Long) As Long
    cioCFFarmMax = baseTwoStepsChangesConf(prev, level, "H4", "H5")
End Function

Function cioCFXPGain(prev As Long, level As Long) As Long
    cioCFXPGain = baseSimpleThird(prev, level)
End Function

Function cioCFUpTime(prev As Long, level As Long) As Long
    cioCFUpTime = baseSimpleHalf(prev, level)
End Function

Function cioCFAdPrice(prev As Long, level As Long) As Long
    cioCFAdPrice = baseTwoStepsChangesConf(prev, level, "C10", "C11")
End Function

Function cioCFTiPrice(prev As Long, level As Long) As Long
    cioCFTiPrice = baseTwoStepsChangesConf(prev, level, "D10", "D11")
End Function

'''''
' adamantiteMine
'''''

Function cioAMEnergy(prev As Long, level As Long) As Long
    cioAMEnergy = baseSimpleThird(prev, level)
End Function

Function cioAMFarm(prev As Long, level As Long) As Long
    cioAMFarm = baseTwoStepsChangesConf(prev, level, "D4", "D5")
End Function

Function cioAMFarmMax(prev As Long, level As Long) As Long
    cioAMFarmMax = baseTwoStepsChangesConf(prev, level, "E4", "E5")
End Function

Function cioAMXPGain(prev As Long, level As Long) As Long
    cioAMXPGain = baseSimpleThird(prev, level)
End Function

Function cioAMUpTime(prev As Long, level As Long) As Long
    cioAMUpTime = baseSimpleHalf(prev, level)
End Function

Function cioAMAdPrice(prev As Long, level As Long) As Long
    cioAMAdPrice = baseSimpleHalf(prev, level)
End Function

Function cioAMTiPrice(prev As Long, level As Long) As Long
    cioAMTiPrice = baseSimpleHalf(prev, level)
End Function

''''''
'
' adamantiteStorage
'
''''''

Function cioASEnergy(prev As Long, level As Long) As Long
    cioASEnergy = baseSimpleThird(prev, level)
End Function

Function cioASSize(prev As Long, level As Long) As Long
    cioASSize = baseTwoStepsChangesConf(prev, level, "C4", "C5")
End Function

Function cioASXPGain(prev As Long, level As Long) As Long
    cioASXPGain = baseSimpleThird(prev, level)
End Function

Function cioASUpTime(prev As Long, level As Long) As Long
    cioASUpTime = baseSimpleHalf(prev, level)
End Function

Function cioASAdPrice(prev As Long, level As Long) As Long
    cioASAdPrice = baseSimpleHalf(prev, level)
End Function

Function cioASTiPrice(prev As Long, level As Long) As Long
    cioASTiPrice = baseSimpleHalf(prev, level)
End Function

''''''
'
' crystaliteSilo
'
''''''

Function cioCSEnergy(prev As Long, level As Long) As Long
    cioCSEnergy = baseSimpleThird(prev, level)
End Function

Function cioCSSize(prev As Long, level As Long) As Long
    cioCSSize = baseTwoStepsChangesConf(prev, level, "F4", "F5")
End Function

Function cioCSXPGain(prev As Long, level As Long) As Long
    cioCSXPGain = baseSimpleThird(prev, level)
End Function

Function cioCSUpTime(prev As Long, level As Long) As Long
    cioCSUpTime = baseSimpleHalf(prev, level)
End Function

Function cioCSAdPrice(prev As Long, level As Long) As Long
    cioCSAdPrice = baseTwoStepsChangesConf(prev, level, "E10", "E11")
End Function

Function cioCSTiPrice(prev As Long, level As Long) As Long
    cioCSTiPrice = baseTwoStepsChangesConf(prev, level, "F10", "F11")
End Function

'''''''
'
' titaniumLab
'
'''''''

Function cioTLEnergy(prev As Long, level As Long) As Long
    cioTLEnergy = baseSimpleThird(prev, level)
End Function

Function cioTLFarm(prev As Long, level As Long) As Long
    cioTLFarm = baseTwoStepsChangesConf(prev, level, "J4", "J5")
End Function

Function cioTLFarmMax(prev As Long, level As Long) As Long
    cioTLFarmMax = baseTwoStepsChangesConf(prev, level, "K4", "K5")
End Function

Function cioTLXPGain(prev As Long, level As Long) As Long
    cioTLXPGain = baseSimpleThird(prev, level)
End Function

Function cioTLUpTime(prev As Long, level As Long) As Long
    cioTLUpTime = baseSimpleHalf(prev, level)
End Function

Function cioTLAdPrice(prev As Long, level As Long) As Long
    cioTLAdPrice = baseSimpleHalf(prev, level)
End Function

Function cioTLTiPrice(prev As Long, level As Long) As Long
    cioTLTiPrice = baseSimpleHalf(prev, level)
End Function




''''''
'
' titaniumStorage
'
''''''

Function cioTSEnergy(prev As Long, level As Long) As Long
    cioTSEnergy = baseSimpleThird(prev, level)
End Function

Function cioTSSize(prev As Long, level As Long) As Long
    cioTSSize = baseTwoStepsChangesConf(prev, level, "I4", "I5")
End Function

Function cioTSXPGain(prev As Long, level As Long) As Long
    cioTSXPGain = baseSimpleThird(prev, level)
End Function

Function cioTSUpTime(prev As Long, level As Long) As Long
    cioTSUpTime = baseSimpleHalf(prev, level)
End Function

Function cioTSAdPrice(prev As Long, level As Long) As Long
    cioTSAdPrice = baseSimpleHalf(prev, level)
End Function

Function cioTSTiPrice(prev As Long, level As Long) As Long
    cioTSTiPrice = baseSimpleHalf(prev, level)
End Function


