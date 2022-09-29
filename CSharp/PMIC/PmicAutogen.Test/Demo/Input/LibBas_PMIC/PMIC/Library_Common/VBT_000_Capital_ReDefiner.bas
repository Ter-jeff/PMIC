Attribute VB_Name = "VBT_000_Capital_ReDefiner"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

'''' Capital ReDefiner

'''' Step1: Uncoment Dim blocks below
'''' Step2: Recoment Dim blocks below
'''' Done. It will redefine all Capital letters and resolve the SVN issues
'
'
'Dim Result, ResultVal, Mode, TheHdw, TheExec, Row, High, LowDim
'Dim PowerUp, ACORE_PowerUp, SubTestMode, SubTestCondition, Block, Current, Voltage, TestNumber, WaitTime, TrimCode
'Dim NumBits, Range, Unit, TwkTrimDelta, TwkTrimDeltaVal, ForceUnit, RunDsp, iLink
'Dim KHz, Index, II, x, y, s, i, m, w, Target, Measure_pin, Measure
'Dim pV, uV, mV, V, pA, uA, mA, A, Hz, MHz, GHz
'Dim Sample, SampleSize, Temp, CurrentLimit, VoltageLimit
'Dim TestName, VDDval, nBits, PatName, CapV, LowLimit, HiLimit, RelayStr, MeasVdiff, Pat, Pat1, Pat2, Pat3, Pat4, Pat5
'Dim ErrHandler, VForce, Pin_Ary, Pin_Cnt, InstanceName, Cnt, VIH, VIL, VOH, VOL, HexStr, TmpStr
'Dim Interval, ForceI, iTarget, iLoad, LoLimit, HighLimit, Meas_Pin, ForceResults, Output, Active, PinName, SampleRate
'Dim RelayOn, RelayOff, Count, State, Sites, hiVal, lowVal, Compare, ShiftBits, Continue, Error, Addr, Low_Side_Pin, High_Side_Pin, BitLength, Ref, CounterValue
'Dim Pin, PinToTest, MeasResult, PinGroup, PatCnt, Status, TNum, TName, Interpose, customUnit, DataArray, DataWidth, TestValue, Acore, Reset, Patt_Reset, Val, Done
'Dim Period, PinNum, StartVolt, Yes, No, StepVolt, SiteStatus, Bypass, LeftShiftBits, DeltaWave, CodeSweep, NumOfBits, DeltaDsp, IRange
'Dim lowCompareSign, highCompareSign, formatStr, Acore_control_registers, SheetName, Frequency, Time, Str, WriteToFile, TF, myStr
'Dim DSP
'Dim Code, TName_TrimLink, START_CODE, STOP_CODE
'Dim Exit_Function, TName_None, PostBurn, Name, Globals, CodeStep, FinalTrim
'Dim Header, Path
'Dim k
'Dim TESTMODE, bIsTRIM, rtnData, eREGWRITE, eREGREAD, BitWiseAnd
'Dim Address, pintype, ItemName, tlUtilBitOff, StartTime, tlUtilBitOn, Validating_, TestNum
'Dim funcName, CreateWaveDefinition
'
'
