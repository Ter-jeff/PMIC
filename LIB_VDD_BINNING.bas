Attribute VB_Name = "LIB_VDD_BINNING"
Option Explicit

Public Function decide_test_type(test_type As testType, inst_name As String)
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = LCase(inst_name)
    
    If strTemp Like "*cpu*bist*" Or strTemp Like "*gfx*bist*" Or strTemp Like "*soc*bist*" Or strTemp Like "*gpu*bist*" Then
        test_type = testType.Mbist
    ElseIf strTemp Like "*elb*" Or strTemp Like "*ilb*" Or strTemp Like "*tmps*" _
    Or strTemp Like "*gfxtd*" Or strTemp Like "*gputd*" Or strTemp Like "*cputd*" Or strTemp Like "*soctd*" _
    Or strTemp Like "*gfxsa*" Or strTemp Like "*gpusa*" Or strTemp Like "*cpusa*" Or strTemp Like "*socsa*" Then
        test_type = testType.TD
    ElseIf strTemp Like "*spi*" Then
        test_type = testType.SPI
    ElseIf strTemp Like "*rtos*" Then
        test_type = testType.RTOS
    Else
        TheExec.Datalog.WriteComment "Test instance:" & TheExec.DataManager.instanceName & ", it doesn't have the correct keyword to decide TestType. Error!!!"
        TheExec.ErrorLogMessage "Test instance:" & TheExec.DataManager.instanceName & ", it doesn't have the correct keyword to decide TestType. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of decide_test_type"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210831: Modified to print the info if Find_IDS_ZONE_per_site can't find the IDS_Zone for the performance mode.
'20210830: Modified to initialize the siteVariable find_ids_zone_flag.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210331: Modified to print the full performance mode for Find_IDS_ZONE_per_site.
'20200914: Modified to merge the redundant branches.
'20200317: Modified for SearchByPmode.
'20191127: Modified for the revised InitVddBinTable.
'20190716: Modified to unify the unit for IDS. ids_current with unit mA.
'20190507: Modified to add "Cdec" for IDS to avoid double format accuracy issues.
'20190422: Modified to define the bin number for DUT with IDS on the IDS_limit.
Public Function Find_IDS_ZONE_per_site(ids_current As SiteDouble, p_mode As Integer)
    Dim site As Variant
    Dim test_type As testType
    Dim ids_zone_num As Long
    Dim find_ids_zone_flag As New SiteBoolean
    Dim str_flag_BinOut As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Compare each IDS Range to find out the IDS ZONE by site.
'''2. If the first step passbin of the IDS ZONE is greater than current passbin, use the passbin from the IDS ZONE to be current passbin.
'''3. It updates flags "F_IDS_Binx" and "F_IDS_Biny" for Bin_Table.
'''//==================================================================================================================================================================================//'''
    '''//init.
    '''//The default Testtype is TD.
    test_type = testType.TD
    '''20210830: Modified to initialize the siteVariable find_ids_zone_flag.
    find_ids_zone_flag = False
    
    For Each site In TheExec.sites
        For ids_zone_num = 0 To Max_IDS_Zone - 1
            If VBIN_IDS_ZONE(p_mode).Used = True Then
                '''//If ids_current >= ids_range, use the next zone.
                '''//IDS calculation uses the scale and the unit in "mA".
                If EnableWord_Vddbin_PTE_Debug = True Then
                    VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = ids_zone_num
                    find_ids_zone_flag(site) = True
                    Exit For
                Else
                    If CDec(ids_current(site)) >= CDec(VBIN_IDS_ZONE(p_mode).Ids_range(ids_zone_num, test_type)) _
                    And CDec(ids_current(site)) < CDec(VBIN_IDS_ZONE(p_mode).Ids_range(ids_zone_num + 1, test_type)) Then
                        VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = ids_zone_num
                        find_ids_zone_flag(site) = True
                        
                        '''//PE asked us to add the flag to distinguish the BinX parts with IDS > Bin1E1 IDS limit from the BinX parts with IDS < Bin1E1 IDS limit.
                        If VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER <> 0 And CDec(ids_current(site)) > CDec(BinCut(p_mode, 1).IDS_CP_LIMIT(0)) Then
                            str_flag_BinOut = "F_" & UCase(VddBinName(p_mode)) & "_IDS"
                            TheExec.sites.Item(site).FlagState(str_flag_BinOut) = logicTrue
                        End If
                        
                        '''//Check if CurrentPassBinCutNum is same as VBIN_IDS_ZONE(p_mode).passBinCut.
                        If CurrentPassBinCutNum(site) < VBIN_IDS_ZONE(p_mode).passBinCut(ids_zone_num, 0) Then
                            CurrentPassBinCutNum(site) = VBIN_IDS_ZONE(p_mode).passBinCut(ids_zone_num, 0)
                        End If
                        
                        '''//Update FlagState of "F_IDS_BinX" and "F_IDS_BinY" by CurrentPassBinCutNum(Site).
                        If CurrentPassBinCutNum(site) = 2 Then
                            TheExec.sites.Item(site).FlagState("F_IDS_BinX") = logicTrue '''for Binx IDS binning
                        ElseIf CurrentPassBinCutNum = 3 Then
                            TheExec.sites.Item(site).FlagState("F_IDS_BinY") = logicTrue '''for Biny IDS binning
                        End If
                        
                        '''//Exit ids_zone-loop
                        Exit For
                    End If
                End If
            End If
        Next ids_zone_num
        
        '''//If it can't find any IDS zone...
        If find_ids_zone_flag(site) = False Then
            '''20210831: Modified to print the info if Find_IDS_ZONE_per_site can't find the IDS_Zone for the performance mode.
            TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & ", Find_IDS_ZONE_per_site can't find the IDS_Zone for the performance mode. Error!!!"
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Find_IDS_ZONE_per_site"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210810: Modified to skip printing the info about the step-adjusted voltage for p_mode, requested by C651 Si and TSMC ZYLINI.
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210611: Modified to bin out the failed site in Adjust_Multi_PassBinCut_Per_Site.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210319: Modified to print bincutNum(site).
'20210317: Modified to revise the format of string about the adjusted BinCut voltage for Adjust_Multi_PassBinCut_Per_Site.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210120: Modified to merge the branches.
'20200317: Modified for SearchByPmode.
'20191127: Modified for the revised InitVddBinTable.
'20181004: If bincutNum>VBIN_RESULT(P_mode).PASSBINCUT, we adjust the bin number.
Public Function Adjust_Multi_PassBinCut_Per_Site(p_mode As Integer, site As Variant, bincutNum As Long, Optional bool_SkipPrintingVoltage As Boolean = False)
    Dim idx_step As Long
    Dim find_out_flag As Boolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''If the CurrentPassBinCutNum is Bin2 but the result of current performance mode is still Bin1.
'''We will find the fisrt step which is Bin2 in the IDS zone to adjust the grade and gradevdd to Bin2.
'''Ex: original grade is Bin1 EQ4, adjust to Bin2
'''                  C                              EQ                             PassBinCut
'''     step0  step1  step2  step3      step0  step1  step2  step3        step0  step1  step2  step3
'''     700    720    780    800          4      3      2      1            1      1      2      2
'''                    V                                V                                 V
'''//==================================================================================================================================================================================//'''
    '''//init
    find_out_flag = False
    
    '''//If bincutNum>VBIN_RESULT(P_mode).PASSBINCUT, we adjust the bin number.
    If VBIN_RESULT(p_mode).passBinCut < bincutNum Then
        For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
            '''//Find out the step in Dynamic_IDS_ZONE to match the current BinCut number of DUT.
            '''//Then update step, BinCut voltage(Grade) and Efuse product voltage(GradeVDD).
            If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) = bincutNum Then
                VBIN_RESULT(p_mode).passBinCut = bincutNum
                VBIN_RESULT(p_mode).step_in_IDS_Zone = idx_step
                VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1
                VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                find_out_flag = True
                Exit For
            End If
        Next idx_step
        
        '''//Check if p_mode has the matched step in DYNAMIC_IDS_Zone for current PassBin.
        If find_out_flag = True Then
            If bool_SkipPrintingVoltage = False Then
                '''20210810: Modified to skip printing the info about the step-adjusted voltage for p_mode, requested by C651 Si and TSMC ZYLINI.
                TheExec.Datalog.WriteComment "site:" & TheExec.sites.SiteNumber & "," & VddBinName(p_mode) & "," & "bin=" & bincutNum & "," & _
                                                "Adjust_Multi_PassBinCut_Per_Site changes " & AllBinCut(p_mode).powerPin & "=" & VBIN_RESULT(p_mode).GRADE
            End If
        Else '''If find_out_flag = False
            TheExec.Datalog.WriteComment "site:" & TheExec.sites.SiteNumber & "," & VddBinName(p_mode) & "," & "bin=" & bincutNum & "," & _
                                            "it can't find out the correct step for Adjust_Multi_PassBinCut_Per_Site. Error!!!"
            
            '''20210611: Modified to bin out the failed site in Adjust_Multi_PassBinCut_Per_Site.
            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
            TheExec.sites.Item(site).SortNumber = 9801
            TheExec.sites.Item(site).binNumber = 5
            TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
            '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
            TheExec.sites.Item(site).result = tlResultFail
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Adjust_Multi_PassBinCut_Per_Site"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210623: Modified to update step_Start for the vbt functon find_start_voltage if EnableWord "Vddbin_PTE_Debug" is enabled.
'20210611: Modified to bin out the failed site in find_start_voltage.
'20210526: Modified to remove Monotonicity_Offset check from find_start_voltage because C651 Si revised the check rules.
'20210518: Modified to update inst_info.is_Monotonicity_Offset_triggered(site).
'20210507: Modified to remove the redundant site-loop and use DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(PassBin,1) for "Vddbin_PTE_Debug".
'20210503: Modified to check if VBIN_RESULT(p_mode).GradeVDD < VBIN_RESULT(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode)).GradeVDD for Monotonicity_Offset.
'20210429: Modified to replace voltage_CalculatedFromIds with DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_Start).
'20210427: Modified to add "Monotonicity_Offset" for GradeVDD check of p_mode.
'20210420: C651 Si did internal syncup and confirmed that Montonicitiy Check should use product voltage(PV) only.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210419: Modified replace "AllBinCut(inst_info.p_mode).Allow_Equal <> cntVddbinPmode + 1" with "AllBinCut(inst_info.p_mode).Allow_Equal <> 0".
'20210419: Modified to set step_stop = step_inherit because voltage heritance between p_mode and previous mode.
'20210408: Modified to overwrite step_inherit and VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone if p_mode is interpolated.
'20210407: Modified to revise the vbt code for the new Interpolation method proposed by C651 Toby.
'20210226: Modified to use step_Start and step_Stop to get startVoltage and StopVoltage.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for find_start_voltage.
'20200502: Modified to replace variable name "IdsVoltage" with "voltage_CalculatedFromIds".
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20200317: Modified for SearchByPmode.
'20200120: Modified to print the information about voltage adjustment.
'20191127: Modified for the revised InitVddBinTable.
'20190716: Modified to unify the unit for IDS.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
Public Function find_start_voltage(inst_info As Instance_Info)
    Dim site As Variant
    Dim step_Start As New SiteLong '''Bin (1[highest V]-6) decided by IDS_Distribution table
    Dim step_inherit As New SiteLong
    Dim EQ_Num As Long
    Dim exit_while_flag As Boolean
    Dim bincutNum As New SiteLong
    Dim i As Integer
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//Find the start voltage of searching for each test Instance, The consider condition as following:
'''1. If the Current Pass Bin Number has been Bin2, adjust the start Voltage to first Bin2 step in the IDS ZONE Number.
'''2. If the previous performance mode had been tested, we need to consider the inheritance by comparing the product value.
'''3. If the current performance mode had been tested, we need to consider the inheritance the step in IDS Zone.
'''4. If the current performance mode was not tested, we need to follow the IDS_START_EQ to find the start step in IDS ZONE.
'''20210420: C651 Si did internal syncup and confirmed that Montonicitiy Check should use product voltage(PV) only.
'''//==================================================================================================================================================================================//'''
'=============================================================================================================================================================
' A. The variable "step_inherit" is the step in the IDS Zone which inherit from the PREVIOUS Performance Mode and Current Performance Mode.
'    1. If the current performance mode did not been tested, we will use last EQ Number (step 0 in ids zone) to be the "step_inherit".
'    2. If the current performance mode had been tested, we will use the result(p_mode).step_in_ids_zone to inherit the previous test items to be the "step_inherit".
'    3. If the PREVIOUS Performance Mode had been tested, we will compare the efuse value (product value) to inherit the step in ids zone to be the "step_inherit".
'    4. If the current pass bincut number is Bin2, we will adjust the current performance mode to the Bin2 last EQ number in the ids zone.
'
' B. The variable "Step_start" is the step in the IDS Zone which come from the start EQ number in the IDS distribution Table.
'    1. If the current performance mode had not been tested,
'       ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
'
'    2. If the current performance mode had been tested but the test type is "SPI",
'       ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
'
'    3. If the current performance mode had been tested but the test type is not "SPI",
'       ==> we don't need to consider the IDS Distribution Table. And set the "Step_start" to 0(Just let the "step_inherit" to be the final step in ids zone).
'
' C. Compare the "step_inherit" and "Step_start" to decide the final step in the ids zone, And use the step to calculate the start voltage.
'=============================================================================================================================================================
    For Each site In TheExec.sites
        '''//Get PassBin of the performance mode.
        bincutNum(site) = CurrentPassBinCutNum(site)
        
        '''//If the passbincut doesn't match CurrentPassBinCutNum, adjust the correct step to align passbincut of P_mode with CurrentPassBinCutNum, and update Grade/GradeVDD.
        If VBIN_RESULT(inst_info.p_mode).passBinCut <> bincutNum(site) Then
            Adjust_Multi_PassBinCut_Per_Site inst_info.p_mode, site, bincutNum(site)
        End If
        
'''//Start of determining step_inherit//'''
        '''=============================================================================================================================================================
        ''' A. Just consider the step for inherit from previous performance mode and current performance mode to define the step in ids zone, step, Grade and GradeVDD,
        '''   ==> The final step needs to compare with "Step_start".
        '''=============================================================================================================================================================
        '''//Decide PassBin by considering previous test result of the same performance mode and previous lower performance mode (for voltage inheritance).
        If VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = True Then  'if previous mode has been Tested, the flag is ture; if not, bypass.
            '''**************************************************************************************************************************************************************************'''
            '''Judge whether current performance mode has been tested or not, or ECID fail, because of ECID fail will become 0,
            '''If the grade had adjusted to Bin2, the grade is not 0. (if the result had become 0 and the current performance mode has been tested, we can not assign the value to it.)
            '''**************************************************************************************************************************************************************************'''
            If VBIN_RESULT(inst_info.p_mode).tested = False Or VBIN_RESULT(inst_info.p_mode).GRADE > 0 Then
                If EnableWord_Vddbin_PTE_Debug = True Then
                    '''//"Vddbin_PTE_Debug" should use step of Bin1 EQN1 in Dynamic_IDS_Zone.
                    If VBIN_RESULT(inst_info.p_mode).passBinCut = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                    Else
                        TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ", it should use Bin1 DUT for Vddbin_PTE_Debug. Error!!!"
                    End If
                ElseIf VBIN_RESULT(inst_info.p_mode).tested = False And bincutNum(site) = 1 Then
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    ''' 1. If the current performance mode has not been tested ==> we will use last EQ Number (step 0 in ids zone) to be the "step_inherit".
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '''//Since p_mode is interpolated but not tested, the revised Interpolated method changes IDS_Start_step Dynamic_IDS_ZONE, it needs to revise step_inherit.
                    If DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).interpolated = True Then
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Interpolated_Start
                    Else
                        step_inherit = 0 '''but EcidVddExecuted(P_mode)(Site) might be true, just keep the same level.
                    End If
                    
                    VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = 0
                Else '''If VBIN_RESULT(inst_info.p_mode).tested =True or bincutNum(site) <> 1
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    ''' 2. If the current performance mode had been tested  ==> we will use the result(p_mode).step_in_ids_zone to inherit the previous test items to be the "step_inherit".
                    ''' 3. If the PREVIOUS Performance Mode had been tested ==> we will compare the efuse value (product value) to inherit the step in ids zone to be the "step_inherit".
                    ''' 4. If the current pass bincut number is Bin2        ==> we will adjust the current performance mode to the bin2 last EQ number in the ids zone.
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    step_inherit = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                End If
                
                If VBIN_RESULT(inst_info.p_mode).tested = False Then
                    VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_inherit)     '''use the ids zone step
                    EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(step_inherit)
                    '''//PRODUCT = CP LVCC + CPGB.
                    VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_inherit)
                End If
                
                '''//The current performance-mode's efuse voltage should be greater than lower performance mode's efuse voltage.
                exit_while_flag = False
                
                '''//Check if Allow_Equal and previous performance mode of p_mode are tested.
                If AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode = AllBinCut(inst_info.p_mode).Allow_Equal And AllBinCut(inst_info.p_mode).Allow_Equal <> 0 Then
                    '''//If Voltage of current performance mode (Grade and GradeVDD) are lower than previous performance mode.
                    '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                    While (CDec(VBIN_RESULT(inst_info.p_mode).GRADEVDD) < CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADEVDD) And exit_while_flag = False) _
                    'Or (CDec(VBIN_RESULT(inst_info.p_mode).GRADE) < CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADE) And exit_while_flag = False)
                        step_inherit = step_inherit + 1 'it will be increased
                        
                        If step_inherit > DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 Then
                            TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",find_start_voltage failed. Error!!!"
                            
                            '''//If no step avaiable, exit the while...
                            exit_while_flag = True
                            '''20210611: Modified to bin out the failed site in find_start_voltage.
                            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                            TheExec.sites.Item(site).SortNumber = 9801
                            TheExec.sites.Item(site).binNumber = 5
                            TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                            '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                            TheExec.sites.Item(site).result = tlResultFail
                        Else
                            '''//If Voltage of current performance mode (Grade and GradeVDD) are lower than previous performance mode.
                            VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_inherit)
                            EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(step_inherit)
                            '''//PRODUCT = CP LVCC + CPGB.
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_inherit)
                        End If
                    Wend
                Else '''//If p_mode has no Allow_Equal.
                    '''//If Voltage of current performance mode is lower than previous performance mode, step-looping until gradevdd is greater than gradevdd of previous mode.
                    '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                    While ((CDec(VBIN_RESULT(inst_info.p_mode).GRADEVDD) <= CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADEVDD)) And exit_while_flag = False) _
                    'Or (CDec(VBIN_RESULT(inst_info.p_mode).GRADE) <= CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADE) And exit_while_flag = False)
                        step_inherit = step_inherit + 1
                        
                        TheExec.Datalog.WriteComment "site:" & site & ", BinCut voltage for pmode:" & VddBinName(inst_info.p_mode) & " is equal or smaller than the previous pmode:" _
                                                        & VddBinName(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode) & ". So that adjust the BinCut step for find_start_voltage."
                        
                        If step_inherit > DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 Then
                            TheExec.Datalog.WriteComment "site:" & site & ", " & VddBinName(inst_info.p_mode) & ", it has the incorrect step. Error!!!"
                            
                            '''//If no step avaiable, exit the while...
                            exit_while_flag = True
                            '''20210611: Modified to bin out the failed site in find_start_voltage.
                            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                            TheExec.sites.Item(site).SortNumber = 9801
                            TheExec.sites.Item(site).binNumber = 5
                            TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                            '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                            TheExec.sites.Item(site).result = tlResultFail
                        Else
                            VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_inherit) '''until it is greater than gradeVDD of previous mode.
                            EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(step_inherit)
                            '''//PRODUCT = CP LVCC + CPGB.
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_inherit)
                        End If
                    Wend
                End If

                VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = step_inherit
                EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = EQ_Num - 1
            End If
        Else '''If VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = False
            If EnableWord_Vddbin_PTE_Debug = True Then
                '''//"Vddbin_PTE_Debug" should use step of Bin1 EQN1 in Dynamic_IDS_Zone.
                If VBIN_RESULT(inst_info.p_mode).passBinCut = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then
                    step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                Else
                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ", it should use Bin1 DUT for Vddbin_PTE_Debug. Error!!!"
                End If
            Else
                '''=======================================================================================================================================================
                ''' If the PREVIOUS Performance Mode has not been tested,
                ''' we do not need to consider the inherit step from PREVIOUS Performance Mode and just use the step in ids zone to be the "step_inherit".
                '''=======================================================================================================================================================
                step_inherit = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
            End If
        End If
'''//End of determining step_inherit//'''
    Next site

    For Each site In TheExec.sites
        '''//Update step_inherit to DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Lowest.
        DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_inherit(site) = step_inherit(site)
        
        '''=============================================================================================================================================================
        ''' B. consider the IDS start EQ for current performance mode to define the step in ids zone, step, Grade and GradeVDD.
        '''=============================================================================================================================================================
        '''//Find start point according to IDS
        If VBIN_RESULT(inst_info.p_mode).tested = False _
        Or (VBIN_RESULT(inst_info.p_mode).tested = True And LCase(TheExec.DataManager.instanceName) Like "*spi*") _
        Or (VBIN_RESULT(inst_info.p_mode).tested = True And LCase(TheExec.DataManager.instanceName) Like "*rtos*") Then 'use IDS for only the first instance of every performance mode
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            ''' 1. If the current performance mode had not been tested.
            '''   ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
            ''' 2. If the current performance mode had been tested but the test type is "SPI".
            '''   ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            If EnableWord_Vddbin_PTE_Debug = True Then
                '''//"Vddbin_PTE_Debug" should use step of Bin1 EQN1 in Dynamic_IDS_Zone.
                If VBIN_RESULT(inst_info.p_mode).passBinCut = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then
                    step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                    step_Start = step_inherit
                Else
                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ", it should use Bin1 DUT for Vddbin_PTE_Debug. Error!!!"
                End If
            Else
                '''from step number to get the LVCC search voltage (refer to "IDS_Distribution table")
                step_Start = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).IDS_START_STEP(inst_info.test_type) '''Base on ids zone number and column of "Start Bin" to get bin number
            End If '''then from bin number to know which step no. is corresponding
        Else '''test instance is not "spi" or "rtos".
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '''3. If the current performance mode had been tested but the test type is not "SPI".
            '''   ==> we don't need to consider the IDS Distribution Table. And set the "Step_start" to 0. (Just let the "step_inherit" to be the final step in ids zone.)
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            step_Start(site) = 0
        End If
        
        If bincutNum > 1 Then  '''IDS distribution has no start point prediction for non-Bin1 DUT.
            step_Start = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
        End If
        
        '''//Decide start point by considering previous performance result, previous lower performance and IDS.
        '''=============================================================================================================================================================
        ''' A. algorithm = IDS condition:
        '''   1. the current performance mode and previous performance mode did not be tested and the ids_start_step is not 0.
        '''
        ''' B. algorithm = Linear condition:
        '''   1. the current performance mode had been tested.
        '''   2. the ids_start_step is 0.
        '''   3. the previous performance mode had been tested and the step had been adjusted to grater or equal to the ids_start_step.
        ''' C. compare the "Step_start" and "step_inherit" to define the step in ids zone, and calculate the start voltage.
        '''    If the "Step_start is greater than "step_inherit", the Start Voltage = voltage(Step_start) and the Stop Voltage = Voltage(step_inherit).
        '''=============================================================================================================================================================
        If VBIN_RESULT(inst_info.p_mode).tested = False And VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = False Then
            If step_Start(site) <> 0 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(step_Start) = 1 And EnableWord_Vddbin_PTE_Debug = False Then
                inst_info.step_Start(site) = step_Start
                '''20210419: Modified to set step_stop = step_inherit because voltage heritance between p_mode and previous mode.
                inst_info.step_Stop(site) = step_inherit '''lowest available step
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.IDS '''IDS search
            Else '''Start from EQN-based voltage with the lowest step, ex: Eqn7
                inst_info.step_Start(site) = step_Start
                inst_info.step_Stop(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 '''highest available step
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear '''Linear search
            End If
        Else '''If VBIN_RESULT(inst_info.p_mode).tested = True Or VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = True
            If step_Start(site) <= step_inherit Then
                step_Start(site) = step_inherit
                inst_info.step_Start(site) = step_Start
                inst_info.step_Stop(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear ' Linear search
            Else '''for SPI test
                inst_info.step_Start(site) = step_Start
                inst_info.step_Stop(site) = step_inherit
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.IDS ' IDS search
            End If
        End If
        
        '''//Adjust step according to Pass step_inherit Category.
        inst_info.step_Current(site) = step_Start '''If step_start is defined, the IndexLevelPerSite can directly use step_start.
        
        '''//IDS calculation uses the scale and the unit in "mA".
        TheExec.Datalog.WriteComment VddBinName(inst_info.p_mode) & "," & site & "," & _
                            "Alg=" & inst_info.gradeAlg & "," & _
                            DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.step_Start) & "mV," & _
                            DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.step_Stop) & "mV," & _
                            Format(inst_info.ids_current(site), ".0") & "mA," & _
                            DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_Start) & "mV," & _
                            VBIN_RESULT(inst_info.p_mode).GRADE & "mV," & _
                            VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADE & "mV," & _
                            "bin" & bincutNum(site) & "," & _
                            "ids_zone " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).IDS_ZONE_NUMBER
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of find_start_voltage"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "stepcountMax As Long" as "maxStep As New SiteLong" for Public Type Instance_Info.
'20210805: Modified to remove the redundant vbt function initialize_step_control since it initialized step_control.All_Site_Mask = 0 in the vbt function decide_binSearch_and_start_voltage.
'20210805: Modified to check if inst_info.is_BinSearch=True for the vbt function decide_binSearch_and_start_voltage.
'20210803: Modified to update inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real in the vbt function initialize_inst_info and remove the redundant vbt function set_IDS_current.
'20210706: Modified to replace is_BinCutJob_for_StepSearch with AllBinCut(p_mode).is_for_BinSearch...
'20210126: Modified to revise the vbt code for DevChar.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for ResetPmodePowerforBincut, set_IDS_current, and find_start_voltage.
'20201208: Modified to use "initialize_step_control".
'20201207: Created to decide the flag is binSearch and find start_voltage.
'20201203: Modified to revise the vbt code for the undefined testJobs.
Public Function decide_binSearch_and_start_voltage(inst_info As Instance_Info, FuncTestOnly As Boolean)
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
'''//inst_info.is_BinSearch =True if testCondition for PowerDomain of the binning p_mode contains the keyword "*Evaluate*Bin*".
'''inst_info.is_BinSearch = True    : find the start voltage, stop voltage, algorithm(Linear or IDS distribution).
'''inst_info.is_BinSearch = False   : only do functional test.
'''//AllBinCut(inst_info.p_mode).is_for_BinSearch = True is defined if testCondition from BinCut flow table(sheet "Non_Binning_Rail") with the keyword "*Evaluate*Bin*".
'''//==================================================================================================================================================================================//'''
    '''//inst_info.is_BinSearch=True is determined in the vbt function initialize_inst_info if testCondition for powerDomain of the binning p_mode contains the keyword "*evaluate*bin*".
    If inst_info.is_BinSearch = True Then
        If inst_info.is_DevChar_Running = True Then '''for DevChar.
            inst_info.maxStep = 0
        Else
            '''//Initialize flags of Grade_Found and AnySiteGradeFound.
            '''//If EnableWord "Vddbin_DoAll_DebugCollection" is enabled, initial performance mode result for Char. BinCut search voltage(Grade) and efuse product voltage(GradeVdd) as 0.
            Call ResetPmodePowerforBincut(inst_info)
        
            '''//Get the IDS value of PowerDomain for the binning p_mode.
            '''//IDS calculation uses the scale and the unit in "mA".
            '''20210803: Modified to update inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real in the vbt function initialize_inst_info and remove the redundant vbt function set_IDS_current.
            If inst_info.powerDomain <> "" Then
                inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(inst_info.powerDomain)).Real '''unit: mA
                
                '''//Check if IDS values of the binning powerDomain > 0.
                For Each site In TheExec.sites
                    If inst_info.ids_current(site) <= 0 Then
                        TheExec.Datalog.WriteComment "site:" & site & ",Instance:" & inst_info.inst_name & ", it doesn't get the correct IDS value for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance and IDS values from DC and Efuse. Error!!!"
                        TheExec.ErrorLogMessage "site:" & site & ",Instance:" & inst_info.inst_name & ", it can't get the correct powerDomain for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance and IDS values from DC and Efuse. Error!!!"
                    End If
                Next site
            Else
                TheExec.Datalog.WriteComment "Instance:" & inst_info.inst_name & ", it can't get the correct powerDomain for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance. Error!!!"
                TheExec.ErrorLogMessage "Instance:" & inst_info.inst_name & ", it can't get the correct powerDomain for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance. Error!!!"
            End If
            
            '''//Find the start voltage, stop voltage, algorithm.
            '''===========================================================================================================================================
            '''Start serach the LVCC based on IDS Zone, the VDD_BIN_ALL(P_mode).MODE_STEP means the maximum EQ count for all BinCut Tables.
            '''Before the searching start, we had based on the IDS current to find out the IDS Zone number, search algorithm and start step in the IDS Zone.
            '''We use the stop voltage to be the stop EQN and when the Grade_Found_Mask = all site, we exit the loop to judge the PF for all sites.
            '''===========================================================================================================================================
            find_start_voltage inst_info
            
            '''//Decide max step for stepcount-loop.
            inst_info.maxStep = AllBinCut(inst_info.p_mode).Mode_Step
        End If
    ElseIf inst_info.is_BinSearch = False Or FuncTestOnly = True Then '''Only do functional test.
        '''//Decide max step for stepcount-loop.
        inst_info.maxStep = 0
    Else '''For the undefined BinCut settings...
        inst_info.maxStep = -1
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of decide_binSearch_and_start_voltage"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of decide_binSearch_and_start_voltage"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
Public Function IsExcludedVddBin(p_mode As Integer) As Boolean
On Error GoTo errHandler
    If BinCut(p_mode, CurrentPassBinCutNum).ExcludedPmode = True Then
        IsExcludedVddBin = True '''The Performance mode doesn't exist in the CurrentPassBinCutNum.
    Else
        IsExcludedVddBin = False '''The Performance mode exists in the CurrentPassBinCutNum.
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of IsExcludedVddBin"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
Public Function SkipTestBin2Site(p_mode As Integer, Active_site_count As Long)
    Dim site As Variant
    Dim EnableSites As New SiteBoolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''PassBinCutNum is global variable and will be decided if this site already run into Bin2.
'''If this site has been in Bin2, and Bin2 doesn't have this performance mode in BinCut voltage table, this site will be disabled.
'''//==================================================================================================================================================================================//'''
    RestoredSites = TheExec.sites.Selected
    EnableSites = TheExec.sites.Selected
    Active_site_count = 0
    '''**********************************************************************************************************************************************************'''
    '''If the CurrentPassBinCutNum is Bin2 but the current performance mode does not exist in the Bin2 Table, set th site disable and skip in this Test Instance.
    '''**********************************************************************************************************************************************************'''
    For Each site In TheExec.sites
        If CurrentPassBinCutNum(site) = 2 Then '''Do not test the performance mode which do not exist in the Bin2 table if it is Bin2 device.
            EnableSites(site) = False
        Else
            Active_site_count = Active_site_count + 1
        End If
    Next site
    TheExec.sites.Selected = EnableSites
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of SkipTestBin2Site"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
Public Function RestoreSkipTestBin2Site(p_mode As Integer)
On Error GoTo errHandler
    'if IsExcludedVddBin(P_mode) Then
        TheExec.sites.Selected = RestoredSites
    'End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of RestoreSkipTestBin2Site"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
'20191126: Modified to use the dictionary to store powerPin and pmode.
'20190516: Modified for temporary use due the incosistent pin name in the BinCut flow and EFUSE_BitDef_Table.
Public Function VddBinStr2Enum(performance_mode As String) As Integer
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = UCase(performance_mode)

    If VddbinPmodeDict.Exists(strTemp) Then
        VddBinStr2Enum = CInt(VddbinPmodeDict.Item(strTemp))
    Else
        VddBinStr2Enum = cntVddbinPmode + 1
        TheExec.Datalog.WriteComment performance_mode & " doesn't have the matched definition of Enum p_mode in VddBinStr2Enum. Error!!!"
        'TheExec.ErrorLogMessage performance_mode & " doesn't have the matched definition of Enum p_mode in VddBinStr2Enum. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of VddBinStr2Enum"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210719: Modified to check AllBinCut(p_mode).Mode_Step with TotalStepPerMode, as requested by ZYLINI and ZQLIN.
'20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
'20210312: Modified to check BinCutList in Vdd_Binning_Def tables.
'20210106: Modified to support the new format with "+" of BinCutList in Vdd_Binning_Def tables, requested by AutoGen team.
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20200703: Modiifed to use "check_Sheet_Range".
'20200528: Modified to check header of the table.
'20200423: Modified to remove the unused argument "col_lvcc as Long".
'20200421: Modified to remove "Init AllBinCut(p_mode).allow_equal".
'20200421: Modified to check the column of "CPIDSMax".
'20200415: Modified to check "col_soft_bin".
'20191127: Modified for the revised InitVddBinTable.
'20191126: Modified to use the dictionary to store powerDomain and pmode.
'20191125: Modified to check if the items of the header exist in the table.
'20190426: Modified to use the function "Find_Sheet".
'20190321: Modified to add the utility to check if the sheet "Vdd_Binning_Def_appA_1" and bincutlist exist.
Public Function initVddBinTable()
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim col_binned As Integer
    Dim col_domain As Integer
    Dim col_mode As Integer
    Dim col_cpids As Integer
    Dim col_sort As Integer
    Dim str_PassBinCut As String
    Dim bincutNum As Variant
    Dim p_mode As Integer
    Dim strAry_PassBinCut() As String
    Dim passBinCut As Variant
    Dim i As Long
    Dim row As Long, col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim Row_of_BasicInfo As Integer
    Dim Row_of_AdditionalInfo As Integer
    Dim row_of_title As Integer
    Dim powerDomain As String
    Dim pmodeName As String
    Dim str_pmode_FullName As String
    Dim PinTEMP As String
    Dim pmodeTemp As String
    Dim pmodeAllTemp As String
    Dim split_array0() As String
    Dim split_array1() As String
    Dim idxArray As Long
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Sheets "Vdd_Binning_Def" (for CorePower) and "Other_Rail" (for OtherRail) are merged into sheets "Vdd_Binning_Def" (for CorePower and OtherRail).
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "Vdd_Binning_Def_appA_1"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        Version_Vdd_Binning_Def = ""
        str_PassBinCut = ""
        BV_StepVoltage = 0
        VddbinningBaseVoltage = 0
        Total_Bincut_Num = 0  'initialize the variable
        i = 0
        pmodeName = ""
        str_pmode_FullName = ""
        PinTEMP = ""
        pmodeTemp = ""
        pmodeAllTemp = ""
        cntVddbinPin = 0
        cntVddbinPmode = 0
        idxArray = -1
        enableRowParsing = False
        
        For row = 1 To MaxRow
            For col = 1 To maxcol
                If LCase(ws_def.Cells(row, col).Value) Like LCase("Rev*") Then '''//Revision of BinCut tables
                    Version_Vdd_Binning_Def = LCase(ws_def.Cells(row, col + 1).Value)
                    Row_of_BasicInfo = row
                End If
                
                If Row_of_BasicInfo > 0 Then
                    If LCase(ws_def.Cells(Row_of_BasicInfo, col).Value) Like LCase("Bin*Cut*List*") Then '''//Number of BinCut tables
                        str_PassBinCut = LCase(ws_def.Cells(Row_of_BasicInfo, col + 1).Value)
                        
                        '''//Check if the sheet Vdd_Binning_Def_appA_1 contains the correct BinCutList.
                        '''If that, parse BinCutList to get BinCut number.
                        If str_PassBinCut <> "" Then
                            '''20210106: Modified to support the new format with "+" of BinCutList in Vdd_Binning_Def tables, requested by AutoGen team.
                            If str_PassBinCut Like "*,*" Then
                                strAry_PassBinCut = Split(str_PassBinCut, ",")   '//BinCut number
                            Else
                                strAry_PassBinCut = Split(str_PassBinCut, "+")   '//BinCut number
                            End If
                            
                            '''//Parse the string in the cell to decide PassBinCut
                            ReDim PassBinCut_ary(UBound(strAry_PassBinCut)) '''//redefine size of array PassBinCut_ary().
                            
                            For Each bincutNum In strAry_PassBinCut '''put how many bincut tables to the array
                                idxArray = idxArray + 1
                                
                                '''//Check PassBin number from BinCutList in Vdd_Binning_Def tables.
                                If bincutNum = CStr(idxArray + 1) Then
                                    PassBinCut_ary(idxArray) = CLng(bincutNum)
                                Else
                                    TheExec.Datalog.WriteComment "BinCutList:" & str_PassBinCut & " of sheet:" & sheetName & " doesn't have any correct sequence of BinCut passBin numbers. Error!!!"
                                    TheExec.ErrorLogMessage "BinCutList:" & str_PassBinCut & " of sheet:" & sheetName & " doesn't have any correct sequence of BinCut passBin numbers. Error!!!"
                                End If
                            Next bincutNum
                        Else
                            TheExec.Datalog.WriteComment sheetName & " doesn't contain any correct BinCutList. Error!!!"
                            TheExec.ErrorLogMessage sheetName & " doesn't contain any correct BinCutList. Error!!!"
                        End If
                    ElseIf LCase(ws_def.Cells(Row_of_BasicInfo, col).Value) Like LCase("col_soft_bin*") Then '''//Stat column of sort bin
                        col_sort = CLng(ws_def.Cells(Row_of_BasicInfo, col + 1).Value)
                    End If
                End If
            Next col
                
            If Row_of_BasicInfo > 0 Then
                If idxArray > -1 And col_sort > 0 Then
                    '''Do nothing...
                Else
                    Row_of_BasicInfo = 0
                    If idxArray = -1 Then
                        TheExec.Datalog.WriteComment "Column Bin Cut List doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column Bin Cut List doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    If col_sort = 0 Then
                        TheExec.Datalog.WriteComment "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                End If
                
                Exit For
            End If
        Next row

        If Row_of_BasicInfo > 0 Then
            For row = Row_of_BasicInfo + 1 To MaxRow
                For col = 1 To maxcol
                    If LCase(ws_def.Cells(row, col).Value) Like LCase("Base*Voltage*") Then '''//Base Voltage
                        VddbinningBaseVoltage = CDbl(ws_def.Cells(row, col + 1).Value)
                        Row_of_AdditionalInfo = row
                    End If

                    If Row_of_AdditionalInfo > 0 Then
                        If LCase(ws_def.Cells(Row_of_AdditionalInfo, col).Value) Like LCase("Step*Size*") Then '''//Step Size
                            BV_StepVoltage = CDbl(ws_def.Cells(Row_of_AdditionalInfo, col + 1).Value)
                            
                            '''//Check if StepVoltage from Vdd_Binning_Def matches the definition of "gC_StepVoltage" in globalVariable.
                            If BV_StepVoltage = gC_StepVoltage Then
                                'Jeff TheExec.Datalog.WriteComment "Step Size Voltage and gC_StepVoltage are " & BV_StepVoltage
                            Else
                                BV_StepVoltage = 0
                                TheExec.Datalog.WriteComment "The Step Size Voltage in BinCut voltage table = " & BV_StepVoltage & ", The Gc_Stepvoltage = " & gC_StepVoltage & ". Error!!!"
                                TheExec.ErrorLogMessage "The Step Size Voltage in BinCut voltage table = " & BV_StepVoltage & ", The Gc_Stepvoltage = " & gC_StepVoltage & ". Error!!!"
                            End If
                        End If
                    End If
                Next col
                    
                '''//Row of the header
                If Row_of_AdditionalInfo > 0 Then
                    If BV_StepVoltage > 0 Then
                        '''Do nothing...
                    Else
                        Row_of_AdditionalInfo = 0
                        TheExec.Datalog.WriteComment "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!"
                    End If
                    
                    Exit For
                End If
            Next row
        Else
            TheExec.Datalog.WriteComment "Column Rev doesn't exist in header of " & sheetName & ". Error!!!"
            TheExec.ErrorLogMessage "Column Rev doesn't exist in header of " & sheetName & ". Error!!!"
        End If

        If Row_of_AdditionalInfo > 0 Then
            For row = Row_of_AdditionalInfo + 1 To MaxRow
                For col = 1 To maxcol
                    If LCase(ws_def.Cells(row, col).Value) = LCase("Binned") Then
                        col_binned = col
                        row_of_title = row
                    End If

                    If row_of_title > 0 Then
                        '''//Check if the items of the header exist in the table.
                        If LCase(ws_def.Cells(row_of_title, col).Value) = LCase("Domain") Then '''//Domain
                            col_domain = col
                        ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = LCase("Mode") Then '''//Mode
                            col_mode = col
                        '''//check the column of "CPIDSMax".
                        ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpidsmax" Then '''//CP IDS limit
                            col_cpids = col
                        End If

                        '''//Check col_soft_bin
                        If col_sort > 0 Then
                            If LCase(ws_def.Cells(row_of_title, col_sort - 1).Value) = LCase("comment") Then '''//Comment
                                '''Do nothing
                            Else
                                TheExec.Datalog.WriteComment "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!"
                                TheExec.ErrorLogMessage "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!"
                            End If
                        Else
                            TheExec.Datalog.WriteComment sheetName & " doesn't contain the correct column position of col_soft_bin. Error!!!"
                            TheExec.ErrorLogMessage sheetName & " doesn't contain the correct column position of col_soft_bin. Error!!!"
                        End If
                    End If
                Next col

                '''//If items are found, exit the for-loop.
                If row_of_title > 0 Then
                    If col_domain > 0 And col_mode > 0 And col_sort > 0 And LCase(ws_def.Cells(row_of_title, col_sort - 1).Value) = LCase("comment") Then
                        enableRowParsing = True
                    Else
                        enableRowParsing = False
                        If col_binned = 0 Then
                            TheExec.Datalog.WriteComment "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!"
                        End If
                        
                        If col_domain = 0 Then
                            TheExec.Datalog.WriteComment "Column Domain doesn't exist in header of " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "Column Domain doesn't exist in header of " & sheetName & ". Error!!!"
                        End If
                        
                        If col_mode = 0 Then
                            TheExec.Datalog.WriteComment "Column Mode doesn't exist in header of " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "Column Mode doesn't exist in header of " & sheetName & ". Error!!!"
                        End If
                        
                        If LCase(ws_def.Cells(row_of_title, col_sort - 1).Value) <> LCase("comment") Then
                            TheExec.Datalog.WriteComment "Column Softbin doesn't exist in header of " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "Column Softbin doesn't exist in header of " & sheetName & ". Error!!!"
                        End If
                        
                        If col_cpids = 0 Then
                            TheExec.Datalog.WriteComment "Column CPIDSMax doesn't exist in header of " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "Column CPIDSMax doesn't exist in header of " & sheetName & ". Error!!!"
                        End If
                    End If
                    
                    Exit For
                Else
                    TheExec.Datalog.WriteComment "Column Binned doesn't exist in header of " & sheetName & ". Error!!!"
                    TheExec.ErrorLogMessage "Column Binned doesn't exist in header of " & sheetName & ". Error!!!"
                End If
            Next row
        Else
            TheExec.Datalog.WriteComment "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!"
            TheExec.ErrorLogMessage "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!"
        End If

        If enableRowParsing = True Then
            '''//Parse the table to enumerate powerDomain(from column Domain) and pmode(from column Mode) for BinCut.
            For row = row_of_title + 1 To MaxRow
                If ws_def.Cells(row, col_domain).Value <> "" And ws_def.Cells(row, col_mode).Value <> "" Then
                    '''//column binned="true" is CorePower, and column binned="false" or "ate" is OtherRail.
                    If LCase(Trim(ws_def.Cells(row, col_binned).Value)) = "true" _
                    Or LCase(Trim(ws_def.Cells(row, col_binned).Value)) = "ate" _
                    Or LCase(Trim(ws_def.Cells(row, col_binned).Value)) = "false" Then
                        '''******************************************************'''
                        '''//Create the dictionary to store powerDomain and pmode
                        '''******************************************************'''
                        '''//powerDomain
                        If UCase(ws_def.Cells(row, col_domain).Value) Like UCase("VDD_*") Then
                            powerDomain = UCase(ws_def.Cells(row, col_domain).Value)
                        Else
                            powerDomain = UCase("VDD_" & ws_def.Cells(row, col_domain).Value)
                        End If
                        
                        '''//Pmode name
                        pmodeName = UCase(UCase(ws_def.Cells(row, col_mode).Value))
                        
                        '''//Full Pmode name
                        str_pmode_FullName = powerDomain & "_" & UCase(ws_def.Cells(row, col_mode).Value)
                        
                        '''//Add powerDomain and Pmode to the temporary group
                        If PinTEMP <> "" Then
                            If LCase("*," & PinTEMP & ",*") Like LCase("*," & powerDomain & ",*") Then
                                '''Do nothing
                            Else
                                PinTEMP = PinTEMP & "," & powerDomain
                            End If
                        Else
                            PinTEMP = powerDomain
                        End If
                        
                        '''//Store DomainType (CorePower or OtherRail of PowerDomain into the dictionary "dict_IsCorePower".
                        If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                            '''Do nothing...
                        Else
                            If LCase(Trim(ws_def.Cells(row, col_binned).Value)) = "true" Then
                                dict_IsCorePower.Add UCase(powerDomain), True
                            ElseIf LCase(Trim(ws_def.Cells(row, col_binned).Value)) = "ate" Or LCase(Trim(ws_def.Cells(row, col_binned).Value)) = "false" Then
                                dict_IsCorePower.Add UCase(powerDomain), False
                            End If
                        End If

                        '''//Only pmode of CorePower
                        If LCase(Trim(ws_def.Cells(row, col_binned).Value)) = "true" Then
                            If pmodeTemp <> "" Then
                                If LCase("*," & pmodeTemp & ",*") Like LCase("*," & pmodeName & ",*") Then
                                    '''Do nothing
                                Else
                                    pmodeTemp = pmodeTemp & "," & pmodeName
                                End If
                            Else
                                pmodeTemp = pmodeName
                            End If
                            
                            If pmodeAllTemp <> "" Then
                                If LCase("*," & pmodeAllTemp & ",*") Like LCase("*," & str_pmode_FullName & ",*") Then
                                    '''Do nothing
                                Else
                                    pmodeAllTemp = pmodeAllTemp & "," & str_pmode_FullName
                                End If
                            Else
                                pmodeAllTemp = str_pmode_FullName
                            End If
                        End If
                    End If
                End If
            Next row
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't contain the correct columns of Binned, Domain, and Mode. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't contain the correct columns of Binned, Domain, and Mode. Error!!!"
        End If
        
        '''//Create the dictionary for pmode (replace str2enum)
        '''We also fill the array "VddBinName" (replace enum2str).
        If PinTEMP <> "" And pmodeTemp <> "" And pmodeAllTemp <> "" Then
            split_array0 = Split(PinTEMP, ",")
            For i = 0 To UBound(split_array0)
                If Not (VddbinPinDict.Exists(split_array0(i))) Then
                    cntVddbinPin = cntVddbinPin + 1
                    VddbinPinDict.Add split_array0(i), cntVddbinPin
                End If
            Next i
            
            pmodeTemp = PinTEMP & "," & pmodeTemp
            pmodeAllTemp = PinTEMP & "," & pmodeAllTemp
            split_array0 = Split(pmodeTemp, ",")
            split_array1 = Split(pmodeAllTemp, ",")
            
            If UBound(split_array0) = UBound(split_array1) Then
                For i = 0 To UBound(split_array1)
                    If Not (VddbinPmodeDict.Exists(split_array1(i))) Then
                        cntVddbinPmode = cntVddbinPmode + 1
                        VddbinPmodeDict.Add split_array1(i), cntVddbinPmode
                        
                        ReDim Preserve VddBinName(cntVddbinPmode)
                        VddBinName(cntVddbinPmode) = split_array1(i)
                        
                        If Not (VddbinPmodeDict.Exists(split_array0(i))) Then
                            VddbinPmodeDict.Add split_array0(i), cntVddbinPmode
                        End If
                    End If
                Next i
            Else
                TheExec.Datalog.WriteComment sheetName & " doesn't contain the correct format of Domain and Mode. Error!!!"
                TheExec.ErrorLogMessage sheetName & " doesn't contain the correct format of Domain and Mode. Error!!!"
            End If
        End If

        '''MaxBincutPowerdomainCount
        If cntVddbinPin > MaxBincutPowerdomainCount Then
            TheExec.Datalog.WriteComment "GlobalVariable MaxBincutPowerdomainCount: " & MaxBincutPowerdomainCount & " doesn't match the number of BinCut powerDomain " & cntVddbinPin & " in the sheet " & sheetName & ". Error!!!"
            TheExec.ErrorLogMessage "GlobalVariable MaxBincutPowerdomainCount: " & MaxBincutPowerdomainCount & " doesn't match the number of BinCut powerDomain " & cntVddbinPin & " in the sheet " & sheetName & ". Error!!!"
        End If
        
        If cntVddbinPmode > MaxPerformanceModeCount Then
            TheExec.Datalog.WriteComment "GlobalVariable MaxPerformanceModeCount: " & MaxPerformanceModeCount & " doesn't match (the number of Pmode)+1 :" & cntVddbinPmode + 1 & " in the sheet " & sheetName & ". Error!!!"
            TheExec.ErrorLogMessage "GlobalVariable MaxPerformanceModeCount: " & MaxPerformanceModeCount & "  doesn't match (the number of Pmode)+1 :" & cntVddbinPmode + 1 & " in the sheet " & sheetName & ". Error!!!"
        End If
        
        '''initilize the MODE_STEP for AllBinCut
        For p_mode = 0 To MaxPerformanceModeCount - 1
            AllBinCut(p_mode).Mode_Step = 0
            AllBinCut(p_mode).is_for_BinSearch = False
        Next p_mode
        
        For Each passBinCut In PassBinCut_ary
            '''Parsing Vdd_Binning_Def sheets for CorePower
            initVddBinTableOneMod CLng(passBinCut), col_cpids, col_sort
            
            '''Parsing Vdd_Binning_Def sheets for OtherRail
            initVddotherrailOneMod CLng(passBinCut)
            
            For p_mode = 0 To MaxPerformanceModeCount - 1
                '''20210719: Modified to check BinCut(p_mode, passBinCut).Mode_Step with TotalStepPerMode, as requested by ZYLINI and ZQLIN.
                If BinCut(p_mode, passBinCut).Mode_Step > TotalStepPerMode Then
                    TheExec.Datalog.WriteComment "bin" & CLng(passBinCut) & "," & VddBinName(p_mode) & ", it has steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", but it is greater than BinCut globalVariable TotalStepPerMode=" & TotalStepPerMode & ", please check check tables Vdd_Binning_Def and update TotalStepPerMode. Error!!!"
                    TheExec.ErrorLogMessage "bin" & CLng(passBinCut) & "," & VddBinName(p_mode) & ", it has steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", but it is greater than BinCut globalVariable TotalStepPerMode=" & TotalStepPerMode & ", please check check tables Vdd_Binning_Def and update TotalStepPerMode. Error!!!"
                ElseIf BinCut(p_mode, passBinCut).Mode_Step >= 0 Then '''if the bincut exists, add to AllBinCut
                    AllBinCut(p_mode).Mode_Step = AllBinCut(p_mode).Mode_Step + BinCut(p_mode, passBinCut).Mode_Step + 1
                End If
            Next p_mode
        Next passBinCut
        
        '''Corrected for max step of per mode
        For p_mode = 0 To MaxPerformanceModeCount - 1
            AllBinCut(p_mode).Mode_Step = AllBinCut(p_mode).Mode_Step - 1
            
            '''//Check if max step of p_mode is greater than TotalStepPerMode.
            '''20210719: Modified to check AllBinCut(p_mode).Mode_Step with TotalStepPerMode, as requested by ZYLINI and ZQLIN.
            If AllBinCut(p_mode).Mode_Step > Max_IDS_Step Then
                TheExec.Datalog.WriteComment VddBinName(p_mode) & ", all BinCut voltage tables have steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", it is greater than BinCut globalVariable Max_IDS_Step=" & Max_IDS_Step & ", please check check tables Vdd_Binning_Def and update Max_IDS_Step. Error!!!"
                TheExec.ErrorLogMessage VddBinName(p_mode) & ", all BinCut voltage tables have steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", it is greater than BinCut globalVariable Max_IDS_Step=" & Max_IDS_Step & ", please check check tables Vdd_Binning_Def and update Max_IDS_Step. Error!!!"
            End If
        Next p_mode
        
        '''//Check if p_mode is ExcludedPmode.
        For Each passBinCut In PassBinCut_ary
            For p_mode = 0 To MaxPerformanceModeCount - 1
                If BinCut(p_mode, passBinCut).ExcludedPmode = True And ExcludedPmode(p_mode) = False Then
                    TheExec.Datalog.WriteComment "Test performance mode " & VddBinName(p_mode) & " doesn't exist in BinCut " & passBinCut - 1 & ". Error!!!"
                    TheExec.ErrorLogMessage "Test performance mode " & VddBinName(p_mode) & " doesn't exist in BinCut " & passBinCut - 1 & ". Error!!!"
                End If
            Next p_mode
        Next passBinCut
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVddBinTable"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVddBinTable"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191227: Modified to remove checking AllBinCut(pmode).Used=true.
'20191127: Modified for the revised InitVddBinTable.
Public Function InitVddBinInherit(Power_Seq() As String)
    Dim i As Integer
On Error GoTo errHandler
    For i = 0 To UBound(Power_Seq)
        If i = 0 Then
            AllBinCut(VddBinStr2Enum(Power_Seq(i))).PREVIOUS_Performance_Mode = cntVddbinPmode + 1
        Else
            AllBinCut(VddBinStr2Enum(Power_Seq(i))).PREVIOUS_Performance_Mode = VddBinStr2Enum(Power_Seq(i - 1))
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of InitVddBinInherit"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of InitVddBinInherit"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210530: Modified to replace typo "Multisftp_Binout" with "MultiFstp_NoBinout".
'20210514: Modified to overwrite failStop if failflag "MultiFstp_NoBinout" is enabled for MultiFSTP.
'20191127: Modified for the revised InitVddBinTable.
'20190422: Modified to check if alarmFail(site) is triggered or not.
Public Function judge_PF_func(p_mode As Integer, test_type As testType, patt_result As SiteBoolean)
    Dim site As Variant
    Dim inst_name As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210514: Modified to overwrite failStop if failflag "MultiFstp_NoBinout" is enabled for MultiFSTP.
'''We modified the vbt function to mask the failed site by theExec.sites.Selected and trig the failFlag if theExec.sites.item(site).FlagState("MultiFstp_NoBinout") = logicTrue for MultiFSTP instance.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        If patt_result(site) = False Or alarmFail(site) = True Then
            '''//If the flag "MultiFstp_NoBinout" = True, skip SortNumber and fail-stop for MultiFSTP instances.
'''ToDo: Please check if the failFlag ""MultiFstp_NoBinout"" exists in the flow table!!!
            If TheExec.sites.Item(site).FlagState("MultiFstp_NoBinout") = logicTrue Then 'MultiFstp without Binout
                inst_name = TheExec.DataManager.instanceName
                TheExec.Datalog.WriteComment "Site:" & site & "," & inst_name & ", test failed, but MultiFSTP bypassed BinOut!"
            Else
                '''//Check if alarmFail(site) is triggered or not.
                If alarmFail(site) = True Then
                    TheExec.Datalog.WriteComment "Site:" & site & ", alarmFail!!!"
                End If
                
                '''//Bin out the failed DUT with SoftBin and HardBin defined in Vdd_Binning_Def tables.
                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                TheExec.sites.Item(site).SortNumber = BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).SBIN_LVCC_FAIL(0, test_type)
                TheExec.sites.Item(site).binNumber = BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).HBIN_LVCC_FAIL(0, test_type)
                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                TheExec.sites.Item(site).result = tlResultFail
            End If
        End If
        TheExec.sites.Item(site).IncrementTestNumber
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of judge_PF_func"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210831: Modified to update HarvestBinningFlag for Harvest in BinCut.
'20210830: As per discussion with TSMC ZYLINI, we decided to use step inherited from Judge_stored_IDS and updated by CurrentPassBinCutNum for HarvestBinningFlag.
'20210830: Modified to revised the vbt code for Harvest in BinCut, as requested by C651 Toby.
'20210820: Modified to remove the redundant GB_delta from the vbt function judge_PF.
'20210803: Modified to use GB_delta for search in non-cp1.
'20210729: Modified to replace "powerDomain = AllBinCut(inst_info.p_mode).powerPin" with inst_info.powerDomain.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210629: Modified to print the message about skipping Sort Number / Bin Number / fail-stop.
'20210531: Modified to adjust the format of "EQN" items in Judge_PF.
'20210530: Modified to update theExec.sites.Selected for MultiFSTP before exiting Judge_PF.
'20210530: Modified to replace typo "Multisftp_Binout" with "MultiFstp_NoBinout".
'20210529: Modified to check if inst_info.Pattern_Pmode and inst_info.By_Mode exist for MultiFSTP.
'20210529: Modified to unifiy the naming rule of failFlags for MultiFSTP with prefix "F_Multifstp_".
'20210528: Modified to assemble the FailFlag of CP1 MultiFSTP for the failed site.
'20210528: Modified to replace testLimit with theExec.Datalog.WriteParametricResult because testLimit latched FailFlag in flow table incorrectly.
'20210528: Modified to update TheExec.sites.Selected.
'20210526: Modified to remove Monotonicity_Offset check from find_start_voltage because C651 Si revised the check rules.
'20210525: Modified to update siteMask for MultiFSTP in CP1.
'20210514: Modified to check if Montonicity_Offset is triggered.
'20210401: Modified to separate "step_1stPass_in_IDS_Zone" into "Bin_1stPass" and "EQN_1stPass".
'20210325: Modified to print info about Bin and EQN for "COF_StepInheritance" and "Vddbin_DoAll_DebugCollection" if grade_found=False.
'20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
'20210325: Modified to merge branches of "COF_StepInheritance" and "Vddbin_DoAll_DebugCollection".
'20210324: Modified to use step_mapping(passBin, EQN1) as step_IDS_Zone for COF_StepInheritance and overwrite step_1stPass_in_IDS_Zone.
'20210322: Modified to decide Flag_Vddbin_COF_StepInheritance by checking TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance").
'20210317: Modified to use VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone to decide PassBinCut.
'20210315: Modified to overwrite VBIN_RESULT(p_mode) for the new COF method requested by C651 Si Li if TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance") = True.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for judge_PF.
'20201102: Modified to use "Public Type Instance_Info".
'20201020: Modified to add the variable "COFInstance" and "PerEqnLog" for COFInstance.
'20201016: Modified to use "print_info_for_COFInstance".
'20201015: Modified to print the summary for "COFInstance".
'20201015: Modified to add the argument "is_COFInstance_enabled".
'20200811: Modified to align the naming rule of failFlag in BinTable with Ellis and JC-Chop.
'20200615: Modified to get dynamic_offset type from the argument "offsetTestTypeIdx As Integer" for judge_PF.
'20200429: Modified to print info while "Vddbin_DoAll_DebugCollection".
'20200212: Modified to print DSSC_Dec when DSSC_Dec=-1.
'20200102: Modified to print the message for C651 PE and checkscript.
'20191219: Modified to add the EnableWord "Vddbin_DoAll_DebugCollection" for Bincut_DoAll_debug.
'20191127: Modified for the revised InitVddBinTable.
'20190722: Modified to printout the scale and the unit for BinCut voltages and IDS values.
'20190716: Modified to unify the unit for IDS. ids_current with unit mA.
'20190417: Modified to rename the output string "DSSCDEC" with "SELSRAM_DSSC".
'20190226: Modified the calculation for dynamic offset.
'20180821: Modified for BinCut testjob mapping.
Public Function judge_PF(inst_info As Instance_Info, passBinCut As SiteLong)
    Dim site As Variant
    Dim ids_step As Long
    Dim lvcc_step As New SiteLong
    Dim strChannel As String
    Dim voltage_Temp As Double
    Dim str_MultiFSTP_FailFlag As String
    Dim PassBinNum As Long
    Dim dbl_BV_lo_limit As Double
    Dim dbl_BV_hi_limit As Double
    Dim str_testJob_Keyword As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly, 20210528.
'''We modified the vbt function to mask the failed site by theExec.sites.Selected and trig the failFlag if theExec.Flow.EnableWord("Multifstp_Datacollection") = True for MultiFSTP instance.
'''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
'''2. C651 Toby updated the rules of step voltage calculation. It should not use GB_delta, 20210728.
'''3. C651 Toby said that HarvestBinningFlag is for BinCur search only, 20210831.
'''//==================================================================================================================================================================================//'''
    '''****************************************************************************************************************'''
    ''' Judge IDS Fail first
    ''' If the grade is not found and the EQ number is not 1 then set Binning Fail Bin
    '''****************************************************************************************************************'''
    '''IDS Pass/Fail Check
    For Each site In TheExec.sites
        If inst_info.grade_found = True Then
            ids_step = VBIN_RESULT(inst_info.p_mode).step_in_BinCut
        Else
            ids_step = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) - 1 '''STEP = EQ -1. If the grade is not found, use the last step to identify IDS or LVCC fail.
        End If
        
        '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Datalog.WriteParametricResult should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
        If inst_info.grade_found = True Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    0, inst_info.ids_current(site) / 1000, BinCut(inst_info.p_mode, passBinCut).IDS_CP_LIMIT(ids_step) / 1000, _
                                                    unitAmp, 0, unitAmp, 0, , , "IDS", scaleMilli
            
        '''****************************************************************************************************************'''
        ''' <LVCC Fail>
        ''' If the C and M of the last step in the IDS Zone are the same with the EQ1 in the BinCut. We fail belong LVCC Fail
        '''****************************************************************************************************************'''
        ElseIf inst_info.grade_found = False And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).c(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) = BinCut(inst_info.p_mode, DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1)).c(0) _
        And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).M(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) = BinCut(inst_info.p_mode, DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1)).M(0) Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    0, inst_info.ids_current(site) / 1000, BinCut(inst_info.p_mode, passBinCut).IDS_CP_LIMIT(0) / 1000, _
                                                    unitAmp, 0, unitAmp, 0, , , "IDS", scaleMilli

        '''****************************************************************************************************************'''
        ''' <IDS fail>
        ''' If the C and M of the last step in the IDS Zone are different from the EQ1 in the BinCut.
        ''' It means some EQ numbers can not be tested in this IDS zone becuase the IDS current is over CPIDSMAX spec. It belongs to Binning Fail.
        '''****************************************************************************************************************'''
        Else
            '''//use ids limit of one step less (less current limit)
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    0, inst_info.ids_current(site) / 1000, BinCut(inst_info.p_mode, passBinCut).IDS_CP_LIMIT(ids_step - 1) / 1000, _
                                                    unitAmp, 0, unitAmp, 0, , , "IDS", scaleMilli
        End If
        TheExec.sites.Item(site).IncrementTestNumber
    Next site

    '''****************************************************************************************************************'''
    ''' If IDS PASS then judge the LVCC.
    ''' If the grade is not found and the IDS is not fail then set LVCC Fail Bin.
    '''****************************************************************************************************************'''
    '''LVCC Pass/Fail Check
    For Each site In TheExec.sites
        If inst_info.grade_found = True Then
            lvcc_step = VBIN_RESULT(inst_info.p_mode).step_in_BinCut
            PassBinNum = VBIN_RESULT(inst_info.p_mode).passBinCut
        Else
            lvcc_step = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) - 1
            PassBinNum = 1
        End If
        
        '''//Select the keyword about the current testJob for TName of testLimit.
        Select Case LCase(bincutJobName)
            Case "cp1": str_testJob_Keyword = "CP1"
            Case "cp2": str_testJob_Keyword = "CP2"
            Case "ft_room": str_testJob_Keyword = "FT1"
            Case "ft_hot": str_testJob_Keyword = "FT2"
            Case "qa": str_testJob_Keyword = "QA"
            Case Else: str_testJob_Keyword = "CP1"
                TheExec.Datalog.WriteComment "site:" & site & ", " & bincutJobName & " is the incorrect BinCut testJob for judge_PF. Error!!!"
                'TheExec.ErrorLogMessage "site:" & site & ", " & bincutJobName & " is the incorrect BinCut testJob for judge_PF. Error!!!"
        End Select
        
        '''//Calculate BV_lo_limit and BV_hi_limit to check BinCut voltage(Grade).
        dbl_BV_lo_limit = BinCut(inst_info.p_mode, PassBinNum).CP_Vmin(lvcc_step)
        dbl_BV_hi_limit = BinCut(inst_info.p_mode, PassBinNum).CP_Vmax(lvcc_step)
        
        '''****************************************************************************************************************'''
        ''' BinCut voltage (Grade)
        '''****************************************************************************************************************'''
        '''//BinCut voltage calculation uses the scale and the unit in "mV", but TheExec.Flow.TestLimit should convert voltage value into "V" with settings "unit:=unitVolt" and "scaleMilli".
        If inst_info.grade_found = True Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                    dbl_BV_lo_limit / 1000, VBIN_RESULT(inst_info.p_mode).GRADE / 1000, dbl_BV_hi_limit / 1000, unitVolt, 0, unitVolt, 0, , , str_testJob_Keyword, scaleMilli, "%.4f"
        Else
        
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                    dbl_BV_lo_limit / 1000, 0, dbl_BV_hi_limit / 1000, unitVolt, 0, unitVolt, 0, , , str_testJob_Keyword, scaleMilli, "%.4f"
        End If
        
        '''//Align testNumber.
        TheExec.sites.Item(site).IncrementTestNumber
    Next site
    
    '''****************************************************************************************************************'''
    ''' Dynamic Offset
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        '''//Dynamic offset is not related to product voltage or efuse, so it doesn't need take the least multiple of stepVoltage.
        '''BinCut voltage calculation uses the scale and the unit in "mV", but TheExec.Flow.TestLimit should convert voltage value into "V" with settings "unit:=unitVolt" and "scaleMilli".
        voltage_Temp = BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).DYNAMIC_OFFSET(inst_info.jobIdx, inst_info.offsetTestTypeIdx)
        
        TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                   -0.1, voltage_Temp / 1000, 0.1, unitVolt, 0, unitVolt, 0, , , "OFFSET", scaleMilli, "%.4f"
                                                   
        TheExec.sites.Item(site).IncrementTestNumber
    Next site
    
    '''****************************************************************************************************************'''
    ''' EQN Result
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        '''********************************************************************************************************************************************************'''
        '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
        '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
        '''********************************************************************************************************************************************************'''
        If inst_info.grade_found = True Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).Mode_Step + 1, _
                                                    unitNone, 0, unitNone, 0, , , "EQN", scaleNoScaling, "%.0f"
        Else
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, 0, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).Mode_Step + 1, _
                                                    unitNone, 0, unitNone, 0, , , "EQN", scaleNoScaling, "%.0f"
        End If
        
        TheExec.sites.Item(site).IncrementTestNumber
    Next site
    
    '''****************************************************************************************************************'''
    ''' BinCut PassBinNumber Result
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        If inst_info.grade_found = True Then
            '''//BinCut PASSBIN values.
            '''********************************************************************************************************************************************************'''
            '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
            '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
            '''********************************************************************************************************************************************************'''
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, VBIN_RESULT(inst_info.p_mode).passBinCut, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , "PASSBIN", scaleNoScaling, "%.0f"
            
            '''****************************************************************************************************************'''
            ''' Record the test type and performance mode for Bin2 or Bin3 binning
            '''****************************************************************************************************************'''
            If VBIN_RESULT(inst_info.p_mode).passBinCut = 2 And Binx_fail_flag(site) = False Then
                If inst_info.test_type = testType.TD Then
                    Binx_fail_power(site) = "TD_" & UCase(VddBinName(inst_info.p_mode))
                    Binx_fail_flag(site) = True
                ElseIf inst_info.test_type = testType.Mbist Then
                    Binx_fail_power(site) = "MBIST_" & UCase(VddBinName(inst_info.p_mode))
                    Binx_fail_flag(site) = True
                ElseIf inst_info.test_type = testType.SPI Then
                    Binx_fail_power(site) = "SPI_" & UCase(VddBinName(inst_info.p_mode))
                    Binx_fail_flag(site) = True
                ElseIf inst_info.test_type = testType.RTOS Then
                    Binx_fail_power(site) = "RTOS_" & UCase(VddBinName(inst_info.p_mode))
                    Binx_fail_flag(site) = True
                End If
            ElseIf VBIN_RESULT(inst_info.p_mode).passBinCut = 3 And Biny_fail_flag(site) = False Then
                If inst_info.test_type = testType.TD Then
                    Biny_fail_power(site) = "TD_" & UCase(VddBinName(inst_info.p_mode))
                    Biny_fail_flag(site) = True
                ElseIf inst_info.test_type = testType.Mbist Then
                    Biny_fail_power(site) = "MBIST_" & UCase(VddBinName(inst_info.p_mode))
                    Biny_fail_flag(site) = True
                ElseIf inst_info.test_type = testType.SPI Then
                    Biny_fail_power(site) = "SPI_" & UCase(VddBinName(inst_info.p_mode))
                    Biny_fail_flag(site) = True
                ElseIf inst_info.test_type = testType.RTOS Then
                    Biny_fail_power(site) = "RTOS_" & UCase(VddBinName(inst_info.p_mode))
                    Biny_fail_flag(site) = True
                End If
            End If
        Else
            '''********************************************************************************************************************************************************'''
            '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
            '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
            '''********************************************************************************************************************************************************'''
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, 0, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , "PASSBIN", scaleNoScaling, "%.0f"
        End If
        TheExec.sites.Item(site).IncrementTestNumber
    Next site
    
    '''****************************************************************************************************************'''
    ''' SELSRM DSSC DigiSrc (converted into decimal).
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        If VBIN_RESULT(inst_info.p_mode).DSSC_Dec <> -1 Then
            '''********************************************************************************************************************************************************'''
            '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
            '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
            '''********************************************************************************************************************************************************'''
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    VBIN_RESULT(inst_info.p_mode).DSSC_Dec, VBIN_RESULT(inst_info.p_mode).DSSC_Dec, VBIN_RESULT(inst_info.p_mode).DSSC_Dec, _
                                                    unitNone, 0, unitNone, 0, , , "SELSRAM_DSSC", scaleNoScaling, "%.0f"
        Else
            VBIN_RESULT(inst_info.p_mode).DSSC_Dec = -1
        End If
    Next site
    
    '''****************************************************************************************************************'''
    '''//Print the summary about COFInstance with PTR format into STDF. Requested by C651 Si Li.
    '''****************************************************************************************************************'''
    If inst_info.enable_PerEqnLog = True Then
        Call print_info_for_COFInstance(inst_info)
    End If
    
    '''****************************************************************************************************************'''
    ''' Decide SortNumber/BinNumber, and bin out the failed DUT.
    '''//SortNumber and BinNumber are generated from "BinNumberConfig" sheet by Tautogen into "Vdd_Binning_Def".
    '''****************************************************************************************************************'''
    '''ToDo: Maybe we can merge the vbt code of this part to the vbt function judge_PF_func...
    For Each site In TheExec.sites
        If inst_info.grade_found = False Then
            If Flag_Vddbin_COF_StepInheritance = True Then
                '''*******************************************************************************************************************************************************************'''
                '''20210315: Modified to overwrite VBIN_RESULT(p_mode) for the new COF method requested by C651 Si Li if TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance") = True.
                '''//If p_mode has found PassBin and EQN in previous instances, it will overwrite VBIN_RESULT with 1st_Pass_Step of p_mode for COF_StepInstance.
                '''//If p_mode didn't find any PassBin and EQN in previous instances, it will use step from DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(CurrentPassBinCutNum, EQN1)
                '''*******************************************************************************************************************************************************************'''
                VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False
                
                If (VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone > -1) = True Then '''If p_mode was tested and has found Grade in previous instances...
                    VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone
                    VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                Else '''If It didn't find any 1st_Pass_Step in previous instances...
                    '''*******************************************************************************************************************************************************************'''
                    '''//step_control.step_Start(site) stores start step inherited from previous and current p_mode.
                    '''It can get passBin from step_control.step_Start(site), then it can set step_mapping(passBin, EQN1) as step_IDS_Zone.
                    '''*******************************************************************************************************************************************************************'''
                    VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Start(site))
                    VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                    VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                End If
                
                '''//Update PassBin, BinCut voltage(Grade), and Efuse product voltage(GradeVDD).
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                CurrentPassBinCutNum(site) = VBIN_RESULT(inst_info.p_mode).passBinCut
                
                '''//Print info about bin, Eqn, Grade, GradeVDD.
                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",Vddbin_COF_StepInheritance overwrites test result of the failed DUT" & _
                                                ",bin=" & VBIN_RESULT(inst_info.p_mode).passBinCut & _
                                                ",EQN=" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                ",Grade=" & VBIN_RESULT(inst_info.p_mode).GRADE & ",GradeVDD=" & VBIN_RESULT(inst_info.p_mode).GRADEVDD
                
            ElseIf Flag_Vddbin_DoAll_DebugCollection = True Then
                TheExec.Datalog.WriteComment "site:" & CStr(site) & "," & VddBinName(inst_info.p_mode) & ",is forced to Bin1 EQN1 CPVmax because Vddbin_DoAll_DebugCollection is enabled, and this test item couldn't find any grade-search result."
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = 0
                VBIN_RESULT(inst_info.p_mode).passBinCut = 1
                
                '''//VddBin_DoAll_DebugCollection should forces CP1 BinCut voltage to Bin1 EQN1 CPVmax.
                VBIN_RESULT(inst_info.p_mode).GRADE = BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_Vmax(VBIN_RESULT(inst_info.p_mode).step_in_BinCut)
                '''//Efuse product voltage(GradeVDD) = BinningVoltage(Grade) + binning_GuardBand.
                VBIN_RESULT(inst_info.p_mode).GRADEVDD = BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_Vmax(VBIN_RESULT(inst_info.p_mode).step_in_BinCut) + BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_GB(VBIN_RESULT(inst_info.p_mode).step_in_BinCut)
                
                '''//Print info about bin, Eqn, Grade, GradeVDD.
                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",Vddbin_DoAll_DebugCollection overwrites test result of the failed DUT" & _
                                                ",bin=" & VBIN_RESULT(inst_info.p_mode).passBinCut & _
                                                ",EQN=" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                ",Grade=" & VBIN_RESULT(inst_info.p_mode).GRADE & ",GradeVDD=" & VBIN_RESULT(inst_info.p_mode).GRADEVDD
            
            '''//If the flag "MultiFstp_NoBinout" = True, skip SortNumber and fail-stop for MultiFSTP instances.
            '''ToDo: Please check if the failFlag ""MultiFstp_NoBinout"" exists in the flow table!!!
            ElseIf TheExec.sites.Item(site).FlagState("MultiFstp_NoBinout") = logicTrue Then 'MultiFstp without Binout
                '''//Mask the failed site for MultiFSTP instances.
                gb_siteMask_current(site) = False
                
                '''//Print info about overwriting fail-stop(BinOut) for MultiFSTP.
                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ", test failed, but overwrite siteMask not to directly bin out the failed site for MultiFSTP."
                    
                '''//It should assemble and update the failFalg of MultiFSTP instances for the failed site.
                '''*******************************************************************************************************************************************************************'''
                '''//Unify the naming rule of failFlags for MultiFSTP with prefix "F_Multifstp_", ex: F_Multifstp_MGX001_X4_BV, F_Multifstp_MGX003_X6_BV, F_Multifstp_MGX008_X10_BV.
                '''Warning: Remember to check if all related FailFlags for MultiFSTP exist in BinCut test flow tables!!!
                '''*******************************************************************************************************************************************************************'''
                '''//Check if keywords of Pattern_Pmode and By_Mode for MultiFSTP (Harvest Core DSSC) are available in the instance name.
                If inst_info.Pattern_Pmode <> "" And inst_info.By_Mode <> "" Then
                    '''//Update test result to the FailFlag assembled with prefix "F_Multifstp_" for MultiFSTP instances.
                    str_MultiFSTP_FailFlag = "F_Multifstp_" & inst_info.Pattern_Pmode & "_" & inst_info.By_Mode & "_BV"
                    TheExec.sites.Item(site).FlagState(str_MultiFSTP_FailFlag) = logicTrue
                Else
                    '''//Print the message about skipping Sort Number / Bin Number / fail-stop.
                    TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but flag MultiFstp_NoBinout is enabled to overwrite Judge_PF."
                    'TheExec.ErrorLogMessage "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but flag MultiFstp_NoBinout is enabled to overwrite Judge_PF."
                End If
                
            '''//If HarvestBinningFlag is not empty, skip updating sortNumber, and skip fail-stop.
            '''20210830: Modified to revised the vbt code for Harvest in BinCut, as requested by C651 Toby.
            ElseIf inst_info.HarvestBinningFlag <> "" Then '''HarvestBinning
                VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False
                
                '''20210831: Modified to update HarvestBinningFlag for Harvest in BinCut.
                TheExec.sites.Item(site).FlagState(inst_info.HarvestBinningFlag) = logicTrue
                TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, it has to update Harvest Core failFlag:" & inst_info.HarvestBinningFlag & "=True."
                
                '''//Update the flagstate of strGlb_Flag_HarvestBinningFlag_AllCorePass because one Core fails.
                If strGlb_Flag_HarvestBinningFlag_AllCorePass <> "" Then
                    If TheExec.sites.Item(site).FlagState(strGlb_Flag_HarvestBinningFlag_AllCorePass) = logicTrue Then
                        TheExec.sites.Item(site).FlagState(strGlb_Flag_HarvestBinningFlag_AllCorePass) = logicFalse
                        TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, so that update Harvest AllCorePass failFlag:" & strGlb_Flag_HarvestBinningFlag_AllCorePass & "=False."
                        
                        If (VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone > -1) = True Then '''If p_mode was tested and has found Grade in previous instances...
                            VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone
                        Else '''If It didn't find any 1st_Pass_Step in previous instances...
                            '''*******************************************************************************************************************************************************************'''
                            '''//Note:
                            '''20210830: As per discussion with TSMC ZYLINI, we decided to use step inherited from Judge_stored_IDS and updated by CurrentPassBinCutNum for HarvestBinningFlag.
                            '''*******************************************************************************************************************************************************************'''
                            VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = inst_info.step_Start(site)
                            '''//Overwrite step_1stPass_in_IDS_zone.
                            VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                        End If
                        
                        '''//Update PassBin, BinCut voltage(Grade), and Efuse product voltage(GradeVDD).
                        VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                        VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                        VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                        VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                        CurrentPassBinCutNum(site) = VBIN_RESULT(inst_info.p_mode).passBinCut
                        
                        '''//Print info about HarvestBinningFlag, bin, Eqn, Grade, GradeVDD.
                        TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",HarvestBinningFlag:" & inst_info.HarvestBinningFlag & ", it overwrites test result of the failed DUT" & _
                                                        ",bin=" & VBIN_RESULT(inst_info.p_mode).passBinCut & _
                                                        ",EQN=" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                        ",Grade=" & VBIN_RESULT(inst_info.p_mode).GRADE & ",GradeVDD=" & VBIN_RESULT(inst_info.p_mode).GRADEVDD
                    Else '''If more than two Cores fails, bin out the failed DUT...
                        TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", Harvest Core failFlag:" & inst_info.HarvestBinningFlag & ", more than one Harvest Core failed, so that bin out the failed DUT."
                    
                        '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                        TheExec.sites.Item(site).SortNumber = BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).SBIN_LVCC_FAIL(lvcc_step, inst_info.test_type)
                        TheExec.sites.Item(site).binNumber = BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).HBIN_LVCC_FAIL(lvcc_step, inst_info.test_type)
                        TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                        '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                        TheExec.sites.Item(site).result = tlResultFail
                    End If
                Else
                    TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but Judge_PF can't update the failFlag of AllCorePass because failFlag:" & strGlb_Flag_HarvestBinningFlag_AllCorePass & " doesn't exist. Error!!!"
                    TheExec.ErrorLogMessage "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but Judge_PF can't update the failFlag of AllCorePass because failFlag:" & strGlb_Flag_HarvestBinningFlag_AllCorePass & " doesn't exist. Error!!!"
                End If
            Else
                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                TheExec.sites.Item(site).SortNumber = BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).SBIN_LVCC_FAIL(lvcc_step, inst_info.test_type)
                TheExec.sites.Item(site).binNumber = BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).HBIN_LVCC_FAIL(lvcc_step, inst_info.test_type)
                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                TheExec.sites.Item(site).result = tlResultFail
            End If
        End If
    Next site
    
    '''//Update theExec.sites.Selected for MultiFSTP before exiting Judge_PF.
    '''Warning!!! It can update theExec.sites.Selected outside site-loop only.
    '''ToDo: Please check if EnableWord("Multifstp_Datacollection") exists in the flow table!!!
    '''20210530: Modified to update theExec.sites.Selected for MultiFSTP before exiting Judge_PF.
    If EnableWord_Multifstp_Datacollection Then
        TheExec.sites.Selected = gb_siteMask_current
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of judge_PF"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210126: Modified to revise the vbt code for DevChar.
'20201211: Created to align testNumber, then do judge_PF for binSearch and judge_PF_func for functional test.
'20201127: Modified to remove the redundant argument "IfStoreData As SiteBoolean".
'20201125: As suggestion from Chihome, modified to clear capture Memory (CMEM) after PostTestIPF.
Public Function update_sort_result(inst_info As Instance_Info, pattPass As SiteBoolean, Org_Test_Number As Long, Optional failpins As String, Optional CollectOnEachStep As Boolean)
    Dim site As Variant
    Dim Site_Align As Long
On Error GoTo errHandler
    '''//Only BinSearch can use CMEM.
    If inst_info.enable_CMEM_collection = True And CollectOnEachStep = False Then
        Call PostTestIPF(inst_info.performance_mode, failpins, inst_info.PrintSize, inst_info.BC_CMEM_StoreData)
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles.
    End If
    
    If inst_info.is_DevChar_Running = False Then
        '''//Update sort number and bin out the failed DUT.
        If inst_info.is_BinSearch = True Then '''BinCut search.
            '''//Update PassBinCut for DUT "grade_found=false".
            Call Update_PassBinCut_for_GradeNotFound(inst_info)
            
            '''************************************************************************************************************************************************'''
            ''' (2) For TestNumber align, calculate the EQs in BinCut Tables and get a guard band to avoid the TestNumber is different from another touch down.
            '''************************************************************************************************************************************************'''
            Site_Align = Org_Test_Number + (inst_info.count_PrePatt_decomposed + inst_info.count_FuncPat_decomposed) * Max_V_Step_per_IDS_Zone + 10
            
            For Each site In TheExec.sites
                TheExec.sites(site).TestNumber = Site_Align
            Next site
            
            '''************************************************************************************************************************************************'''
            ''' Base on the search result, print EQN and BinCut CP voltage to datalog and bin out the failed DUT.
            '''************************************************************************************************************************************************'''
            judge_PF inst_info, CurrentPassBinCutNum
            
            '''//For the performance mode that does not exist in the bincut table
            RestoreSkipTestBin2Site inst_info.p_mode
        Else '''BinCut check.
            judge_PF_func inst_info.p_mode, inst_info.test_type, pattPass
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of update_sort_result"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of update_sort_result"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to use "Public Type Instance_Info".
'20201022: Modified to fix the vbt code for PatternBurst result issues. Requested by TSMC PCLINZG.
'20200827: Modified to remove the redundant site-loop.
'20200319: Modified to switch off save_core_power_vddbinning and restore_core_power_vddbinning if Flag_Enable_Rail_Switch = True.
'20200203: Modified to use the function "print_bincut_power".
'20200113: Modified for pattern bursted without decomposing pattern.
'20191127: Modified for the revised InitVddBinTable.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerPins.
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerPins.
Public Function run_prepatt(PrePatt As Pattern, inst_name As String, p_mode As Integer, PrePattPass As SiteBoolean, result_mode As tlResultMode, Optional special_voltage_setup As Boolean)
    Dim site As Variant
    Dim CorePowerStored() As New SiteDouble
    Dim i As Integer
    Dim siteResult As New SiteBoolean
    Dim inst_info As Instance_Info
On Error GoTo errHandler
    If PrePatt.Value <> "" Then
        '''init
        inst_info.is_BV_Safe_Voltage_printed = False
        inst_info.is_BV_Payload_Voltage_printed = False
        inst_info.inst_name = inst_name
        inst_info.p_mode = p_mode
        inst_info.special_voltage_setup = special_voltage_setup
        inst_info.PrePatt = PrePatt
        inst_info.step_Current = -1
    
        '''//siteDouble "CorePowerStored()" is used to save/restore voltages for BinCut powerDomains.
        ReDim CorePowerStored(UBound(pinGroup_BinCut))
        
        For i = 0 To UBound(pinGroup_BinCut)
            CorePowerStored(i) = 0
        Next i
        
        '''//Save payload voltages of CorePower and OtherRail powerPins before init pattern.
        If Flag_noRestoreVoltageForPrepatt = False Then
            save_core_power_vddbinning CorePowerStored
        End If
                
        '''//Set to nominal voltage (NV).
'''ToDo: If initial voltages and safe voltage(init voltage) use the same DC category, we will skip "set_core_power_vddbinning_VT" after initial voltages...
        set_core_power_vddbinning_VT VddBinName(p_mode), "NV"
        TheHdw.Wait 0.0001
        
        '''//Print safe voltages(init voltages) for PrePatt(init patt).
        print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages
                
        '''//Set "result_mode = tlResultModeModule" (return a unique pass/fail result for each module and time domain) if pattern bursted without decomposing pattern.
        '''20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
        result_mode = tlResultModeDomain
        
        '''//Run the pattern
        Call TheHdw.Patterns(PrePatt).Test(pfAlways, 0, result_mode)
        DebugPrintFunc PrePatt.Value
        
        '''//Check pattern Pass/Fail.
        '''//Warning!!! currently "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" doesn't support "result_mode=tlResultModeModule" with PatternBurst=Yes and DecomposePatt=No.
        '''20201022: Modified to fix the vbt code for PatternBurst result issues. Requested by TSMC PCLINZG.
        PrePattPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        
        '''//Restore the BinCut voltages for payload patterns after PrePatt.
        If Flag_noRestoreVoltageForPrepatt = False Then
            restore_core_power_vddbinning CorePowerStored
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of run_prepatt"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210803: Modified to update inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real in the vbt function initialize_inst_info and remove the redundant vbt function set_IDS_current.
'20200923: Modified to remove "clear_after_patt".
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20190706: Modified for the new datatype of power_seq.
Public Function sort_power_seqence(power_list As String, Power_Seq() As String) As Long
    Dim strAry_Performance_Mode() As String
    Dim idxA As Integer
    Dim idxB As Integer
    Dim idxC As Integer
    Dim replace_flag As Boolean
    Dim strTemp As String
    Dim strReplaced As String
On Error GoTo errHandler
    '''init
    idxA = 0
    idxB = 0
    idxC = 0
    
    strAry_Performance_Mode = Split(power_list, ",")
    ReDim Power_Seq(UBound(strAry_Performance_Mode))
    
    For idxA = 0 To UBound(strAry_Performance_Mode)
        replace_flag = False
        If idxA = 0 Then
            Power_Seq(idxB) = strAry_Performance_Mode(idxA)
            idxB = idxB + 1
        Else
            For idxC = 0 To UBound(Power_Seq)
                If Power_Seq(idxC) <> "" Then
                    If BinCut(VddBinStr2Enum(strAry_Performance_Mode(idxA)), 1).MAX_ID < BinCut(VddBinStr2Enum(Power_Seq(idxC)), 1).MAX_ID Then
                        If replace_flag = False Then
                            replace_flag = True
                            strTemp = Power_Seq(idxC)
                            Power_Seq(idxC) = strAry_Performance_Mode(idxA)
                            idxC = idxC + 1
                        End If
                    End If
                    If replace_flag = True Then
                        strReplaced = Power_Seq(idxC)
                        Power_Seq(idxC) = strTemp
                        strTemp = strReplaced
                    End If
                Else
                   If replace_flag = False Then
                        Power_Seq(idxC) = strAry_Performance_Mode(idxA)
                   Else
                        Power_Seq(idxC) = strTemp
                   End If
                   Exit For
                End If
            Next idxC
        End If
    Next idxA
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of sort_power_seqence"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of sort_power_seqence"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'20210803: Modified to calculate step voltages with BV_StepVoltage.
'20210728: C651 Toby updated the rules of step voltage calculation. It should not use GB_delta.
'20210722: Modified to update VBIN_IDS_ZONE(p_mode).Product_Voltage.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20191127: Modified for the revised InitVddBinTable.
'20190716: Modified to unify the unit for IDS. ids_current with unit mA.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
Public Function Generate_IDS_ZONE_Voltage_Per_Site(ids_current As SiteDouble, p_mode As Integer)
    Dim site As Variant
    Dim test_type As testType
    Dim idx_step As Long
    Dim remainder As Double
    Dim voltage_Temp As Double
    Dim Zone_Num As Integer
    Dim dbl_CPVmax As Double
    Dim dbl_CPVmin As Double
    Dim PassBinNum As Long
    Dim step_in_BinCut As Long
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Calculate the voltage of each step and each zone by IDS values for p_mode.
'''2. C651 Toby updated the rules of step voltage calculation. It should not use GB_delta, 20210728.
'''3. C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand, 20210812.
'''//==================================================================================================================================================================================//'''
    '''//The default Testtype is TD
    test_type = testType.TD
    
    If VBIN_IDS_ZONE(p_mode).Used = True Then
        For Each site In TheExec.sites
            For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)        'loop IDS Range for all IDS Zone
                For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) - 1        'loop the step in the IDS Zone
                    '''************************************************************************************************************************************************'''
                    '''//Formula: CP voltage = C-M*log10(IDS).
                    '''CP voltage with unit: mV. ids_current with unit mA.
                    '''************************************************************************************************************************************************'''
                    voltage_Temp = VBIN_IDS_ZONE(p_mode).c(Zone_Num, idx_step) - VBIN_IDS_ZONE(p_mode).M(Zone_Num, idx_step) * (Log(ids_current) / Log(10))
                    
                    '''//For LVCC, floor the value by step_voltage defined in the header of sheet "Vdd_Binning_Def".
                    '''//Floor step voltages of each step in Dynamic_IDS_Zone by BV_StepVoltage.
                    remainder = Floor(voltage_Temp / BV_StepVoltage)
                    voltage_Temp = remainder * BV_StepVoltage
                    
                    '''//Update PassBinNum and step_in_BinCut for each step in IDS_Zone.
                    PassBinNum = VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, idx_step)
                    step_in_BinCut = VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, idx_step) - 1
                    
                    '''//Check if voltage of each step in IDS_Zone is between CPVmax and CPVmin.
                    dbl_CPVmax = BinCut(p_mode, PassBinNum).CP_Vmax(step_in_BinCut)
                    dbl_CPVmin = BinCut(p_mode, PassBinNum).CP_Vmin(step_in_BinCut)
                    
                    If CDec(voltage_Temp) > CDec(dbl_CPVmax) Then
                        voltage_Temp = dbl_CPVmax
                    ElseIf CDec(voltage_Temp) < CDec(dbl_CPVmin) Then
                        voltage_Temp = dbl_CPVmin
                    End If
                    
                    '''//Calculate GradeVDD for each step.
                    '''Efuse product voltage(GradeVDD) = BinCut voltage(Grade) + binning_GuardBand.
                    '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                    VBIN_IDS_ZONE(p_mode).Voltage(Zone_Num, idx_step) = voltage_Temp
                    VBIN_IDS_ZONE(p_mode).Product_Voltage(Zone_Num, idx_step) = voltage_Temp + BinCut(p_mode, PassBinNum).CP_GB(step_in_BinCut)
                Next idx_step
            Next Zone_Num
        Next site
    Else
        TheExec.Datalog.WriteComment VddBinName(p_mode) & ", it doesn't have any correct IDS ZONE for Generate_IDS_ZONE_Voltage_Per_Site. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Generate_IDS_ZONE_Voltage_Per_Site"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20181101: Modified for the format of .CSV file.
'20181030: Modified to integrate all IDS distribution related functions.
'20181026: Modified for IDS, by MSLi.
Public Function Print_IDS_ZONE_Table_to_sheet()
    Dim site As Variant
    Dim wb As Workbook
    Dim test_type As testType
    Dim p_mode As Integer
    Dim i As Long, k As Long, L As Long
    Dim ids_range_step(MaxPerformanceModeCount) As Long
    Dim IDS_current_Max(MaxPerformanceModeCount) As Double
    Dim p_col As Integer, p_row As Integer
    Dim max_print_step As Integer
    Dim SheetCnt As Long
    Dim str_CurPath As String
    Dim SheetExist As Boolean
    Dim str_output As String
    Dim str_header As String
    Dim IsHeaderPrinted As Boolean
On Error GoTo errHandler
    '''init
    str_CurPath = "D:\IDS_ZONE_TABLE.csv"
    Open str_CurPath For Output As #1
    Set wb = Application.ActiveWorkbook
    SheetCnt = ActiveWorkbook.Sheets.Count
    SheetExist = False
    test_type = testType.TD
    p_col = 1
    p_row = 1
    str_output = ""
    str_header = ""
    IsHeaderPrinted = False
    
    '''use the max step count to print the table
    max_print_step = Max_V_Step_per_IDS_Zone
    
    For i = 0 To MaxPerformanceModeCount - 1
        ids_range_step(i) = 0
        IDS_current_Max(i) = 0
    Next i
    
    For p_mode = 0 To MaxPerformanceModeCount - 1
        If VBIN_IDS_ZONE(p_mode).Used = True Then
            IsHeaderPrinted = False
        
            '''//Performance mode//
            Print #1, VddBinName(p_mode)
            
            For k = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                '''//IDS range & Start bin
                For L = 0 To MaxTestType - 1
                    If L = 0 Then
                        str_header = "IDS Range"
                        str_output = VBIN_IDS_ZONE(p_mode).Ids_range(k, L)
                    Else
                        str_header = str_header & "," & "IDS Range"
                        str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).Ids_range(k, L)
                    End If
                    
                    str_header = str_header & "," & "Start Bin"
                    str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(k, L)
                Next L
                
                '''//C
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "C"
                    If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).c(k, i)
                    Else
                        str_output = str_output & "," & " "
                    End If
                Next i
                
                '''//M
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "M"
                     If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).M(k, i)
                    Else
                        str_output = str_output & "," & " "
                    End If
                Next i
                
                '''//EQN
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "EQN"
                    If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).EQ_Num(k, i)
                    Else
                        str_output = str_output & "," & " "
                    End If
                Next i
                
                '''//PASSBINCUT
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "PASSBINCUT"
                     If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).passBinCut(k, i)
                    Else
                        str_output = str_output & "," & " "
                    End If
                Next i
                
                '''//print out the header for each performance mode
                If IsHeaderPrinted = False Then
                    Print #1, str_header
                    str_header = ""
                    IsHeaderPrinted = True
                End If
                
                Print #1, str_output
            Next k
            
            Print #1, ""
        End If
    Next p_mode
    
    '''//Close the csv file//
    Close #1
    
    '''//Import csv to the sheet"IDS_ZONE_TABLE"//
    'Open str_CurPath For Input As #1
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Print_IDS_ZONE_Table_to_sheet"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Print_IDS_ZONE_Table_to_sheet"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20191204: Modified for the revised initVddBinTable.
'20181101: Modified for the format of .CSV file.
'20181031: Modified for the sheet "IDS_ZONE_voltage".
Public Function Print_IDS_ZONE_voltage_to_sheet()
    Dim site As Variant
    Dim test_type As testType
    Dim idx_step As Long
    Dim p_mode As Integer
    Dim i As Long
    Dim ids_range_step(MaxPerformanceModeCount) As Long
    Dim IDS_current_Max(MaxPerformanceModeCount) As Double
    Dim p_col As Integer, p_row As Integer
    Dim Zone_Number As Long
    Dim SheetCnt As Long
    Dim str_CurPath As String
    Dim SheetExist As Boolean
    Dim str_output As String
    Dim str_header As String
    Dim IsHeaderPrinted As Boolean
On Error GoTo errHandler
    '''init
    str_CurPath = "D:\IDS_ZONE_Voltage.csv"
    Open str_CurPath For Output As #1
    p_col = 1
    p_row = 1
    test_type = testType.TD
    SheetCnt = ActiveWorkbook.Sheets.Count
    SheetExist = False
    str_output = ""
    str_header = ""
    IsHeaderPrinted = False
        
    For i = 0 To MaxPerformanceModeCount - 1
        ids_range_step(i) = 0
        IDS_current_Max(i) = 0
    Next i
    
    For p_mode = 0 To MaxPerformanceModeCount - 1
        If VBIN_IDS_ZONE(p_mode).Used = True Then
            IsHeaderPrinted = False
            
            '''//Performance mode//
            Print #1, VddBinName(p_mode)
            
            For Each site In TheExec.sites
                str_header = "Site"
                str_output = site
                
                str_header = str_header & "," & "IDS_ZONE_NUMBER"
                str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER
                
                For Zone_Number = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                    For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Number) - 1
                        str_header = str_header & "," & "V" & "_" & "step" & idx_step
                        str_output = str_output & "," & VBIN_IDS_ZONE(p_mode).Voltage(Zone_Number, idx_step)
                    Next idx_step
                Next Zone_Number
                
                If IsHeaderPrinted = False Then
                    Print #1, str_header
                    str_header = ""
                    IsHeaderPrinted = True
                End If
                
                Print #1, str_output
            Next site
            Print #1, ""
        End If
    Next p_mode
    
    '''//Close the csv file//
    Close #1
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Print_IDS_ZONE_voltage_to_sheet"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Print_IDS_ZONE_voltage_to_sheet"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210405: Modified to remove "PassBinCutList_per_Zone(Max_IDS_Zone) As Long"
'20210303: Modified to replace "Dim steps As Long" with "Dim idx_step As Long".
'20191127: Modified for the revised InitVddBinTable.
Public Function init_IDS_ZONE()
    Dim ids_zone_num As Long
    Dim test_type As Long
    Dim idx_step As Long
    Dim p_mode As Integer
On Error GoTo errHandler
    For p_mode = 0 To MaxPerformanceModeCount - 1
        VBIN_IDS_ZONE(p_mode).Used = False
        
        For idx_step = 0 To Max_IDS_Step
            For ids_zone_num = 0 To Max_IDS_Zone - 1
                VBIN_IDS_ZONE(p_mode).Max_Step(ids_zone_num) = 0
                VBIN_IDS_ZONE(p_mode).c(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).M(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).passBinCut(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).EQ_Num(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).Voltage(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) = 0
                
                For test_type = 0 To MaxTestType - 1
                    VBIN_IDS_ZONE(p_mode).Ids_range(ids_zone_num, test_type) = 0
                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(ids_zone_num, test_type) = 0
                    VBIN_IDS_ZONE(p_mode).IDS_START_STEP(ids_zone_num, test_type) = 0
                    VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type) = 0
                Next test_type
            Next ids_zone_num
        Next idx_step
    Next p_mode
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of init_IDS_ZONE"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210405: Modified to remove "PassBinCutList_per_Zone(Max_IDS_Zone) As Long"
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20200703: Modiifed to use "check_Sheet_Range".
'20200506: Modified to revise the vbt code for parsing the sheet.
'20200505: Modified to check if "IDS_Distribution" exists in the workbook.
'20191127: Modified for the revised InitVddBinTable.
'20160614: Modified by TSMC Jack.
Public Function initIDSTable()
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim sheetName As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim row As Long, col As Long
    Dim p_mode As Integer
    Dim idx_step As Long
    Dim test_type As Long
    Dim col_testTypeSelected As Long
    Dim zoneTemp As Integer
    Dim row_of_title As Integer
    Dim row_of_step0 As Integer
    Dim col_TestType(MaxTestType) As Integer
    Dim strTemp As String
    Dim isSheetFound As Boolean
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "IDS_Distribution"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        Version_IDS_Distribution = ""
        
        '''//Initialize the array
        '''//Please check "Enum TestType" and "MaxTestType" in GlobalVariable.
        For test_type = 0 To MaxTestType - 1
            For p_mode = 0 To MaxPerformanceModeCount - 1
                For zoneTemp = 0 To Max_IDS_Zone - 1
                    IDS_Distribution_Table(p_mode).range(zoneTemp, test_type) = 0
                    IDS_Distribution_Table(p_mode).Start_Bin(zoneTemp, test_type) = 0
                    IDS_Distribution_Table(p_mode).RANGE_COUNT = 0
                    IDS_Distribution_Table(p_mode).Used = False
                Next zoneTemp
            Next p_mode
            
            '''init the array to store the column number of each TestType.
            col_TestType(test_type) = 0
        Next test_type
    
        '''//Find the start point of the header.
        For row = 1 To MaxRow
            For col = 1 To maxcol
                If LCase(ws_def.Cells(row, col).Value) Like "*rev*" And LCase(ws_def.Cells(row + 1, col).Value) Like "td" Then
                    Version_IDS_Distribution = ws_def.Cells(1, 2).Value
                    row_of_title = row + 1
                End If
                
                If row_of_title > 0 And row = row_of_title Then
                    If LCase(ws_def.Cells(row, col).Value) <> "" Then
                        col_TestType(decide_test_type_for_string(ws_def.Cells(row, col).Value)) = col
                    End If
                End If
            Next col
        Next row
        
        If row_of_title > 0 And col_TestType(testType.TD) > 0 Then
            For row = row_of_title + 1 To MaxRow
                If LCase(ws_def.Cells(row, col_TestType(testType.TD)).Value) = "ids range" And LCase(ws_def.Cells(row, col_TestType(testType.TD) + 1).Value) = "start bin" Then '''//Find the row with "IDS Range","Start Bin"
                    p_mode = 0
                    strTemp = ws_def.Cells(row - 1, col_TestType(testType.TD)).Value
                    
                    If VddbinPmodeDict.Exists(strTemp) Then
                        p_mode = VddBinStr2Enum(strTemp)
                        row_of_step0 = row + 1
                    Else
                        p_mode = 0
                        TheExec.Datalog.WriteComment sheetName & " doesn't have any correct Performance_mode in row" & row & ". Error!!!"
                        'TheExec.ErrorLogMessage SheetName & " doesn't have any correct Performance_mode in row" & Row & ". Error!!!"
                    End If
                        
                    If p_mode > 0 Then
                        For test_type = 0 To MaxTestType - 1
                            col_testTypeSelected = col_TestType(test_type)
                            
                            If col_testTypeSelected > 0 Then
                                If ws_def.Cells(row_of_step0 - 2, col_testTypeSelected).Value = strTemp Then
                                    idx_step = 0
                                    row = row_of_step0
                                    
                                    If IsNumeric(ws_def.Cells(row, col_testTypeSelected).Value) And (IsEmpty(ws_def.Cells(row, col_testTypeSelected).Value) = False) Then
                                        While (LCase(ws_def.Cells(row, col_testTypeSelected).Value) <> "end" And (ws_def.Cells(row, col_testTypeSelected).Value) <> "")
                                            IDS_Distribution_Table(p_mode).range(idx_step, test_type) = CDbl(ws_def.Cells(row, col_testTypeSelected).Value)
                                            IDS_Distribution_Table(p_mode).Start_Bin(idx_step, test_type) = CDbl(ws_def.Cells(row, col_testTypeSelected + 1).Value)
                                            IDS_Distribution_Table(p_mode).RANGE_COUNT = idx_step
                                            IDS_Distribution_Table(p_mode).Used = True
                                            idx_step = idx_step + 1
                                            row = row + 1 '''Row Offset
                                        Wend
                                    End If
                                    
                                    row = row_of_step0
                                Else
                                    TheExec.Datalog.WriteComment sheetName & " doesn't have the correct Performance_mode in row" & (row_of_step0 - 2) & ", col" & col_testTypeSelected & " consistent with other TestType columns. Error!!!"
                                    'TheExec.ErrorLogMessage SheetName & " doesn't have the correct Performance_mode in row" & Row & ", col" & col_TestType(TestType.TD) & " consistent with other TestType columns. Error!!!"
                                End If
                            End If
                        Next test_type
                    End If
                End If
            Next row
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't have correct format of the header. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't have correct format of the header. Error!!!"
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initIDSTable"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initIDSTable"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210303: Modified to replace "Dim steps As Long" with "Dim idx_step As Long".
'20191127: Modified for the revised InitVddBinTable.
Public Function init_IDS_ZONE_Voltage()
    Dim ids_zone_num As Long
    Dim idx_step As Long
    Dim p_mode As Integer
On Error GoTo errHandler
    For p_mode = 0 To MaxPerformanceModeCount - 1
        For idx_step = 0 To Max_IDS_Step
            For ids_zone_num = 0 To Max_IDS_Zone - 1
                VBIN_IDS_ZONE(p_mode).Voltage(ids_zone_num, idx_step) = 0 '''siteDouble
                VBIN_IDS_ZONE(p_mode).Product_Voltage(ids_zone_num, idx_step) = 0 '''siteDouble
            Next ids_zone_num
        Next idx_step
    Next p_mode
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of init_IDS_ZONE_Voltage"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210906: Modified to remove the unused variables for the vbt function Generate_IDS_ZONE_RANGE.
'20200512: Modified to merge Generate_IDS_Zone_with_IDS_Distribution_Table and Generate_IDS_Zone_NO_IDS_Distribution_Table into "Generate_IDS_Zone_with_IDS_Distribution_Table".
'20160614: Modified by TSMC Jack.
Public Function Generate_IDS_ZONE_RANGE()
On Error GoTo errHandler
    '''//Generate IDS_Zone for each p_mode with/without IDS_Distribution_Table.
    Call Generate_IDS_Zone_with_IDS_Distribution_Table
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Generate_IDS_ZONE_RANGE"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210303: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20200512: Created to merge "Generate_IDS_Zone_with_IDS_Distribution_Table" and "Generate_IDS_Zone_NO_IDS_Distribution_Table" into "Generate_IDS_Zone_with_IDS_Distribution_Table".
'20190422: Modified to define the bin number for DUT with IDS on the IDS_limit.
Public Function Generate_IDS_Zone_with_IDS_Distribution_Table()
    Dim test_type As Integer
    Dim i As Integer
    Dim j As Integer
    Dim RngNum As Integer
    Dim Bincut_step As Integer
    Dim idx_step As Integer
    Dim Ids_zone_cnt As Integer
    Dim DblTemp As Double
    Dim Zone_Num As Integer
    Dim Duplicate_Flag As Boolean
    Dim bincutNum As Variant
    Dim p_mode As Integer
    Dim Zone As Integer
    Dim Max_mode_step As Integer
    Dim Ids_range As Double
    Dim bincut_max_step As Integer
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Determine how many ids zone would be after considering bincut and ids_distribution.
'''//==================================================================================================================================================================================//'''
    For p_mode = 0 To MaxPerformanceModeCount - 1
        '''***********************************************************************'''
        '''[Step0] Initialize the IDS Zone value and Start-Search bin.
        '''***********************************************************************'''
        For test_type = 0 To MaxTestType - 1
            For Zone_Num = 0 To Max_IDS_Zone
                VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) = 0
                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = 0
                VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Zone_Num, test_type) = -999
                VBIN_IDS_ZONE_Temp(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = -999
            Next Zone_Num
        Next test_type
        
        '''//If IDS distribution table is parsed, and p_mode is defined in the table.
        If IDS_Distribution_Table(p_mode).Used = True Then
            For Each bincutNum In PassBinCut_ary
                If bincutNum = 1 Then
                    '''***********************************************************************'''
                    '''[Step1] For offline simulation mode, we make man-made data for simulation.
                    '''***********************************************************************'''
                    If TheExec.TesterMode = testModeOffline Then
                        '''If last parameter =-1 means bincut ids limit less than ids distribution range
                        '''If last parameter =0 means bincut ids limit is same as ids distribution range
                    End If
                            
                    '''***********************************************************************'''
                    '''[Step2] Copy ids_distribution table into a ids zone array.
                    '''***********************************************************************'''
                    For test_type = 0 To MaxTestType - 1 '''For all test type, TD, MBIST, SPI, RTOS, TMPS and LDCBFD
                        Ids_zone_cnt = 0
                        
                        For RngNum = 0 To IDS_Distribution_Table(p_mode).RANGE_COUNT
                            VBIN_IDS_ZONE_Temp(p_mode).Ids_range(RngNum, test_type) = IDS_Distribution_Table(p_mode).range(RngNum, test_type)
                            Ids_zone_cnt = Ids_zone_cnt + 1
                        Next RngNum
                        
                        VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(test_type) = Ids_zone_cnt
                    Next test_type
                End If '''If BincutNum = 1 Then
                                        
                '''***********************************************************************'''
                '''[Step3] Copy bincut table into ids zone array.
                '''***********************************************************************'''
                For test_type = 0 To MaxTestType - 1 '''For all test type "TD, MBIST, SPI, RTOS, TMPS and LDCBFD"
                    Ids_zone_cnt = VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(test_type)
                    
                    For idx_step = 0 To BinCut(p_mode, bincutNum).Mode_Step
                        VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Ids_zone_cnt, test_type) = BinCut(p_mode, bincutNum).IDS_CP_LIMIT(idx_step)
                        Ids_zone_cnt = Ids_zone_cnt + 1
                    Next idx_step
                    
                    VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(test_type) = Ids_zone_cnt
                Next test_type
            Next bincutNum
                                    
            '''***********************************************************************'''
            '''[Step4] Screen out the duplicate ids limit from vbin_ids_zone().
            '''***********************************************************************'''
            For test_type = 0 To MaxTestType - 1
                Zone_Num = 0
                Ids_zone_cnt = VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(test_type)
                For i = 0 To Ids_zone_cnt - 1
                    Duplicate_Flag = False
                    
                    For j = i + 1 To Ids_zone_cnt
                        If VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, test_type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(j, test_type) Then
                            Duplicate_Flag = True
                            Exit For
                        End If
                    Next j
                    
                    If Duplicate_Flag = False Then
                        VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, test_type)
                        Zone_Num = Zone_Num + 1
                    End If
                Next i
                
                VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type) = Zone_Num - 1 'This is for fit in with original bincut vbt structure.
                VBIN_IDS_ZONE(p_mode).Used = True                               'This auxiliary setting is to ensure no error in check_ids test instance.
            Next test_type

            '''***********************************************************************'''
            '''[Step5] Sorting the merged array in ascending way because IDS Zone is from low to high.
            '''***********************************************************************'''
            For test_type = 0 To MaxTestType - 1
                Ids_zone_cnt = VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                For i = 0 To Ids_zone_cnt - 1
                    For j = i + 1 To Ids_zone_cnt
                        If (VBIN_IDS_ZONE(p_mode).Ids_range(i, test_type) > VBIN_IDS_ZONE(p_mode).Ids_range(j, test_type)) Then
                            DblTemp = VBIN_IDS_ZONE(p_mode).Ids_range(j, test_type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(j, test_type) = VBIN_IDS_ZONE(p_mode).Ids_range(i, test_type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(i, test_type) = DblTemp
                        End If
                    Next j
                Next i
            Next test_type
            
            '''***********************************************************************'''
            '''[Step6] Determine the start-search level for merged bincut ids limits.
            '''***********************************************************************'''
            For Each bincutNum In PassBinCut_ary
                bincut_max_step = BinCut(p_mode, bincutNum).Mode_Step
                
                For test_type = 0 To MaxTestType - 1
                    For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                        Ids_range = VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type)
                        
                        '''//If any ids zone execeed the maximum ids of bincut, its start bin is set to 0.
                        If Ids_range >= BinCut(p_mode, bincutNum).IDS_CP_LIMIT(bincut_max_step) Then
                            VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = 0
                        Else
                            For j = 0 To IDS_Distribution_Table(p_mode).RANGE_COUNT - 1
                                If Ids_range >= IDS_Distribution_Table(p_mode).range(j, test_type) And Ids_range < IDS_Distribution_Table(p_mode).range(j + 1, test_type) Then
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = IDS_Distribution_Table(p_mode).Start_Bin(j, test_type)
                                    Exit For
                                End If
                            Next j
                        End If '''If Ids_range >= VDD_BIN(P_mode, BincutNum).IDS_CP_LIMIT(bincut_max_step) Then
                    Next Zone_Num
                Next test_type
            Next bincutNum
            
        '''//If IDS distribution table is not parsed, or p_mode is not defined in the table.
        Else
            '''//The default Testtype is TD
            test_type = testType.TD
            Zone_Num = 0
        
            For Each bincutNum In PassBinCut_ary
                If BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                    VBIN_IDS_ZONE(p_mode).Used = True '''With this statement, then no argue(error) happens in instance of Check_IDS
                        
                    '''***********************************************************************'''
                    '''[Step2] Copy bincut table into a ids zone array.
                    '''***********************************************************************'''
                    If Zone_Num = 0 Then
                        VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Zone_Num, test_type) = 0 '''Ids zone value starts from 0 based on current structure
                        Zone_Num = Zone_Num + 1
                    End If
                                
                    For idx_step = 1 To BinCut(p_mode, bincutNum).Mode_Step + 1
                        VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Zone_Num, test_type) = BinCut(p_mode, bincutNum).IDS_CP_LIMIT(idx_step - 1)
                        Zone_Num = Zone_Num + 1
                    Next idx_step
                    VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(test_type) = Zone_Num
                End If '''If VDD_BIN(P_mode, BincutNum).ExcludedPmode = False Then
            Next bincutNum

            '''***********************************************************************'''
            '''[Step3] Screen out the duplicate ids limit from vbin_ids_zone().
            '''***********************************************************************'''
            If BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                Zone_Num = 0
                Ids_zone_cnt = VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(test_type)
                
                For i = 0 To Ids_zone_cnt - 1
                    Duplicate_Flag = False
                    
                    For j = i + 1 To Ids_zone_cnt
                        If VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, test_type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(j, test_type) Then
                            Duplicate_Flag = True
                            Exit For
                        End If
                    Next j
                    
                    If Duplicate_Flag = False Then
                        VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, test_type)
                        Zone_Num = Zone_Num + 1
                    End If
                Next i
                
                VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type) = Zone_Num - 1 '''Minus 1 is for fit in with the original vdd-binning code's counting
                
                '''***********************************************************************'''
                '''[Step5] Sorting the merged array in ascending way.
                '''***********************************************************************'''
                '''This is because IDS Zone is from low to high.
                Ids_zone_cnt = VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                For i = 0 To Ids_zone_cnt - 1
                    For j = i + 1 To Ids_zone_cnt
                        If (VBIN_IDS_ZONE(p_mode).Ids_range(i, test_type) > VBIN_IDS_ZONE(p_mode).Ids_range(j, test_type)) Then
                            DblTemp = VBIN_IDS_ZONE(p_mode).Ids_range(j, test_type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(j, test_type) = VBIN_IDS_ZONE(p_mode).Ids_range(i, test_type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(i, test_type) = DblTemp
                        End If
                    Next j
                Next i
                
                '''***********************************************************************'''
                '''[Step6] Copt TD to the remaining test type (MBIST, SPI, ...)
                '''***********************************************************************'''
                For test_type = 1 To MaxTestType - 1
                    VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type) = VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(testType.TD)
                Next test_type
                
                For test_type = 1 To MaxTestType - 1
                    For RngNum = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                        VBIN_IDS_ZONE(p_mode).Ids_range(RngNum, test_type) = VBIN_IDS_ZONE(p_mode).Ids_range(RngNum, testType.TD)
                    Next RngNum
                Next test_type
                                                
                '''***********************************************************************'''
                '''[Step7] Assign the start-search level for each ids zone.
                '''***********************************************************************'''
                '''Because start-search level is totally referred to bincut 1, we don't need to have bincut loop here.
                test_type = testType.TD '''//The default Testtype is TD
                Max_mode_step = BinCut(p_mode, 1).Mode_Step
                
                If BinCut(p_mode, 1).ExcludedPmode = False Then
                    For RngNum = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                        If VBIN_IDS_ZONE(p_mode).Ids_range(RngNum, test_type) < BinCut(p_mode, 1).IDS_CP_LIMIT(Max_mode_step) Then
                            VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, test_type) = Max_mode_step + 1
                        Else
                            VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, test_type) = 0
                        End If
                    Next RngNum
                End If
                
                '''***********************************************************************'''
                '''[Step8] Copy TD to the remaining test type (MBIST, SPI, ...)
                '''***********************************************************************'''
                For test_type = 1 To MaxTestType - 1
                    For RngNum = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                        VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, test_type) = VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, testType.TD)
                    Next RngNum
                Next test_type
            End If '''If VDD_BIN(P_mode, BincutNum).ExcludedPmode = False Then
        End If '''If IDS_Distribution_Table(P_mode).Used = True Then
    Next p_mode
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Generate_IDS_Zone_with_IDS_Distribution_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20191204: Modified for the revised InitVddBinTable.
'20190314: Modified to check the following condition to avoid any C vlaue in Bin2 < Bin1 maxC in each ids zone.
'20181031: Modified to add TotalPmodeRangeCnt for Print_IDS_ZONE_Table_to_sheet.
'20160614: Modified by TSMC Jack.
Public Function Generate_IDS_ZONE_CONTENT()
    Dim test_type As testType
    Dim idx_step As Long
    Dim Srch_Step As Long
    Dim i As Long
    Dim ids_range_step(MaxPerformanceModeCount) As Long
    Dim IDS_current_Max(MaxPerformanceModeCount) As Double
    Dim IdsRng As Long
    Dim Zone_Num As Long
    Dim Mode_Step As Long
    Dim p_mode As Integer
    Dim isPmodeUsed As Boolean
    Dim bincutNum As Variant
    Dim StepCnt As Long
On Error GoTo errHandler
    '''init
    Max_V_Step_per_IDS_Zone = 0
    
    '''//The default Testtype is TD
    test_type = testType.TD
            
    For i = 0 To MaxPerformanceModeCount - 1
        ids_range_step(i) = 0
        IDS_current_Max(i) = 0
    Next i
    
    For p_mode = 0 To MaxPerformanceModeCount - 1
        '''***********************************************************************'''
        '''[Step1] Skip those unused mode from mode select loop.
        '''***********************************************************************'''
        isPmodeUsed = False
        If Flag_IDS_Distribution_enable = True Then '''If IDS_Distribution_Table_table is available
            If IDS_Distribution_Table(p_mode).Used = True Or BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                isPmodeUsed = True
            End If
        Else '''If IDS_Distribution_Table_table is not available
            If BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                isPmodeUsed = True
            End If
        End If
        
        '''***********************************************************************'''
        '''[Step2]
        ''' 2.1: Copy C and M of Bincut to VBIN_IDS_Zone().
        ''' 2.2: Assign Level and bincutnum search sequence to VBIN_IDS_Zone().
        '''***********************************************************************'''
        If isPmodeUsed = True Then
            For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type) - 1
               Srch_Step = VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num)         'This is a count-up step across multiple bincuts
               
               For Each bincutNum In PassBinCut_ary
                   Mode_Step = BinCut(p_mode, bincutNum).Mode_Step          'This is mode step number of different bincut table
                   StepCnt = 0                                              'Count up step of per ids zone
                   For idx_step = 0 To Mode_Step
                       If VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) < BinCut(p_mode, bincutNum).IDS_CP_LIMIT(idx_step) Then
                           If bincutNum = 1 Then
                               VBIN_IDS_ZONE(p_mode).c(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).c(Mode_Step - StepCnt)
                               VBIN_IDS_ZONE(p_mode).M(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).M(Mode_Step - StepCnt)
                               VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, Srch_Step) = (Mode_Step + 1) - StepCnt
                               VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, Srch_Step) = bincutNum
                               Srch_Step = Srch_Step + 1
                               StepCnt = StepCnt + 1
                           Else
                               '''20190314: Modified. The following condition is to avoid any C value in bin2 < bin1 maxC in each ids zone.
                               If BinCut(p_mode, bincutNum).CP_Vmax(Mode_Step - idx_step) >= BinCut(p_mode, bincutNum - 1).CP_Vmax(Mode_Step - idx_step) Then
                                   VBIN_IDS_ZONE(p_mode).c(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).c(Mode_Step - StepCnt)
                                   VBIN_IDS_ZONE(p_mode).M(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).M(Mode_Step - StepCnt)
                                   VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, Srch_Step) = (Mode_Step + 1) - StepCnt
                                   VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, Srch_Step) = bincutNum
                                   Srch_Step = Srch_Step + 1
                              End If
                              
                              StepCnt = StepCnt + 1
                           End If 'If BincutNum = 1 Then
                       End If 'If VBIN_IDS_ZONE(P_mode).Ids_range(Zone_Num, Test_Type)
                   Next idx_step 'For Step = 0 To Mode_Step
               Next bincutNum
               
               VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) = Srch_Step
               
               If VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) > Max_V_Step_per_IDS_Zone Then
                   Max_V_Step_per_IDS_Zone = VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num)
               End If
            Next Zone_Num
        End If 'If IsPmodeUsed = True Then
    Next p_mode
    
    '''***********************************************************************'''
    '''[Step3]
    ''' 3.1: Determine VBIN_IDS_ZONE().IDS_Start_Bin.
    ''' 3.2: While all VBIN_IDS_ZONE().IDS_Start_Bin() are 0 if we only refer to bincut alone.
    '''***********************************************************************'''
    If Flag_IDS_Distribution_enable = True Then
        For p_mode = 0 To MaxPerformanceModeCount - 1
            If BinCut(p_mode, 1).ExcludedPmode = False Then
                Mode_Step = BinCut(p_mode, 1).Mode_Step
                
                For test_type = 0 To MaxTestType - 1
                    '''//Because all test type has same ids range, we select td as representative.
                    For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(test_type)
                        For IdsRng = 0 To IDS_Distribution_Table(p_mode).RANGE_COUNT
                            '''//Zone_Num always start from 0 and it is same as IDS_Distribution_Table table at its idsrng=0.
                            If Zone_Num = 0 Then
                                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = IDS_Distribution_Table(p_mode).Start_Bin(0, test_type)
                                Exit For 'This zone has been given start_bin
                            ElseIf VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) < IDS_Distribution_Table(p_mode).range(IdsRng, test_type) Then
                                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = IDS_Distribution_Table(p_mode).Start_Bin(IdsRng - 1, test_type)
                                If VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) >= BinCut(p_mode, 1).IDS_CP_LIMIT(Mode_Step) Then
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = 0
                                End If
                                Exit For 'This zone has been given start_bin
                            ElseIf Zone_Num > 0 And VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) = IDS_Distribution_Table(p_mode).range(IdsRng, test_type) Then
                                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = IDS_Distribution_Table(p_mode).Start_Bin(IdsRng + 1, test_type)
                                If VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, test_type) < BinCut(p_mode, 1).IDS_CP_LIMIT(Mode_Step) Then
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = IDS_Distribution_Table(p_mode).Start_Bin(IdsRng, test_type)
                                Else
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = 0
                                End If
                                Exit For 'This zone has been given start_bin
                            End If
                        Next IdsRng
                    Next Zone_Num
                Next test_type
            End If '''If VDD_BIN(P_mode, 1).ExcludedPmode = False Then
        Next p_mode
    End If '''If Flag_IDS_Distribution_Table_enable = True Then

    '''***********************************************************************'''
    '''[Step4] Determine the start search step from VBIN_IDS_ZONE().IDS_Start_EQ_Num and VBIN_IDS_ZONE().EQ_Num.
    '''***********************************************************************'''
    If Flag_IDS_Distribution_enable = True Then
        For p_mode = 0 To MaxPerformanceModeCount - 1
            If BinCut(p_mode, 1).ExcludedPmode = False Then
                For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(0)
                    For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) - 1
                        For test_type = 0 To MaxTestType - 1
                            If VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) = 0 Then
                                VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Zone_Num, test_type) = 0
                            Else
                                If VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, idx_step) = VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type) And VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, idx_step) = 1 Then
                                    VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Zone_Num, test_type) = idx_step
                                End If
                            End If
                        Next test_type
                    Next idx_step
                Next Zone_Num
            End If
        Next p_mode
    End If
    
    If Flag_Print_Out_tables_enable = True Then
        Print_IDS_ZONE_Table_to_sheet
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Generate_IDS_ZONE_CONTENT"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetInstrument_BV(PinList As String, site As Variant) As String
    Dim strChannel As String
    Dim strAry_PinName() As String
    Dim NumberPins As Long
    Dim strAry_slot() As String
    Dim slot As Long
On Error GoTo errHandler
    Call TheExec.DataManager.DecomposePinList(PinList, strAry_PinName(), NumberPins)
    Call TheExec.DataManager.GetChannelStringFromPinAndSite(strAry_PinName(0), site, strChannel)
        
    If strChannel = "" Then
        MsgBox ("Please check pin type of " & PinList & " in channel map")
    Else
        strAry_slot = Split(strChannel, ".")
        slot = CLng(strAry_slot(0))
        GetInstrument_BV = TheHdw.config.Slots(slot).Type
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of GetInstrument_BV"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initIDSTable"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210407: Modified to add "interpolated as new SiteBoolean" and "step_Interpolated_Start as new SiteLong" for "Public Type DYNAMIC_VBIN_IDS_ZONE".
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210223: Modified to map DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(BinNum, EQN) to step in DYNAMIC_IDS_Zone.
'20201113: Modified to use "Last_Bin1_Step" to store last Bin1 Step (ex: EQN1) in Dynamic_IDS_ZONE for Interpolation.
'20191127: Modified for the revised InitVddBinTable.
Public Function Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site(p_mode As Integer)
    Dim site As Variant
    Dim test_type As testType
    Dim idx_step As Long
    Dim Zone_Num As Integer
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Calculate the voltage of each step and each zone from IDS_Zone for p_mode.
'''//==================================================================================================================================================================================//'''
    '''init
    test_type = testType.TD

    If VBIN_IDS_ZONE(p_mode).Used = True Then
        For Each site In TheExec.sites
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True
            DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER
            Zone_Num = DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER
            
            '''//Initialize the lowest step in Dynamic_IDS_Zone as 0.
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_inherit(site) = 0
            
            '''//Use "Last_Bin1_Step" to store last Bin1 Step (ex: EQN1) in Dynamic_IDS_ZONE. Initialize it as -1.
            '''For interpolation.
            DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated = False
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Interpolated_Start = -1
            
            For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) - 1    'loop the step in the IDS Zone
                DYNAMIC_VBIN_IDS_ZONE(p_mode).c(idx_step) = VBIN_IDS_ZONE(p_mode).c(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).M(idx_step) = VBIN_IDS_ZONE(p_mode).M(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) = VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) = VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step = VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step) = VBIN_IDS_ZONE(p_mode).Voltage(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step) = VBIN_IDS_ZONE(p_mode).Product_Voltage(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step), DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step)) = idx_step
               
                For test_type = 0 To MaxTestType - 1
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(test_type) = VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, test_type)
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(test_type) = VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Zone_Num, test_type)
                Next test_type
            Next idx_step
        Next site
    Else
        TheExec.Datalog.WriteComment VddBinName(p_mode) & ", it doesn't have any correct IDS_Zone for Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210728: Modified to remove the redundant vbt function Clear_Dynamic_IDS_ZONE_by_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210407: Modified to add "interpolated as new SiteBoolean" and "step_Interpolated_Start as new SiteLong" for "Public Type DYNAMIC_VBIN_IDS_ZONE".
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210223: Modified to map DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(BinNum, EQN) to step in DYNAMIC_IDS_Zone.
'20191127: Modified for the revised InitVddBinTable.
Public Function init_Dynamic_IDS_ZONE()
    Dim test_type As Long
    Dim idx_step As Long
    Dim p_mode As Integer
    Dim idx_binNum As Long
    Dim idx_EQN As Long
On Error GoTo errHandler
    For p_mode = 0 To MaxPerformanceModeCount - 1
        DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = False
        DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = 99
        DYNAMIC_VBIN_IDS_ZONE(p_mode).step_inherit = 0
        '''For interpolation.
        DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated = False
        DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Interpolated_Start = -1
        
        For idx_step = 0 To Max_IDS_Step
            DYNAMIC_VBIN_IDS_ZONE(p_mode).c(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).M(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step) = 0
            
            '''//Initialize array of mapping (binNum, EQN) to step in Dynamic_IDS_Zone.
            For idx_binNum = 0 To MaxPassBinCut
                For idx_EQN = 0 To Max_IDS_Step + 1
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(idx_binNum, idx_EQN) = -1
                Next idx_EQN
            Next idx_binNum
            
            For test_type = 0 To MaxTestType - 1
                 DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(test_type) = 0
                 DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(test_type) = 0
            Next test_type
        Next idx_step
    Next p_mode
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of init_Dynamic_IDS_ZONE"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210831: Modified to remove the vbt code related to CPVmax and CPVmin.
'20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
'20210813: Modified to print interpolation info while SkipTest=True for BinX and BinY.
'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210618: Modified to update SortNumber and binNumber if F_Vddbinning_Interpolation_fail is triggered in the vbt function ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site.
'20210412: Modified to remove the vbt code not to check if the voltage_Calc is in the range between CPVmin and CPVmax, requested by PCLINZG.
'20210409: Modified to separate voltage calculation for interpolation and check steps.
'20210408: Modified to update step for BinX/Y if AllBinCut(p_mode).INTP_SKIPTEST = True.
'20210408: Modified to overwrite step_inherit and VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone if DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).interpolated = True.
'20210407: Modified to revise the vbt code for the new Interpolation method proposed by C651 Toby.
'20210407: Modified to add "interpolated as new SiteBoolean" and "step_Interpolated_Start as new SiteLong" for "Public Type DYNAMIC_VBIN_IDS_ZONE".
'20210322: Modified to initialize Last_Bin1_Step = -1 for Interpolation.
'20210305: Modified to check if enableCalcInterpolation = True.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210223: Modified to map DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(BinNum, EQN) to step in DYNAMIC_IDS_Zone.
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20201229: Modified to check if start_p_mode and end_p_mode were tested or not, requested by Verity.
'20201218: Modified to update "FIRSTPASSBINCUT(p_mode) = VBIN_RESULT(p_mode).passBinCut", requested by Leon Weng.
'20201113: C651 Toby requested the new method to check BinCut voltages montonicity, so that we need to keep all available steps in Dynamic_IDS_Zone.
'20201113: Modified to use "Last_Bin1_Step" to store last Bin1 Step (ex: EQN1) in Dynamic_IDS_ZONE for Interpolation.
'20200502: Modified to replace "VBIN_IDS_ZONE" with "DYNAMIC_VBIN_IDS_ZONE".
'20200501: Modified to replace "BinCut(p_mode, PassBinNum(site)).INTP_MFACTOR(0)" with "AllBinCut(p_mode).INTP_SKIPTEST".
'20200425: Modified to change the output format of the interpolated voltage string.
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20200414: Modified the branches of Intrpolation for Bin1 and non-Bin1.
'20200410: Modified to control "Exit Function" by the siteFlag "enableCalcInterpolation".
'20200406: Modified to revise the format of interpolation output strings.
'20200401: Modified to use IDS to calculate BinCut payload voltage for BinX and BinY.
'20200330: Modified to use "Int_Offset","Int_SkipTest" for interpolation.
'20200319: Modified to find start p_mode, end p_mode, and interpolation factor by site.
'20191127: Modified for the revised InitVddBinTable.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
'20190226: Modified to skip interpolation for start_p_mode and end_p_mode if INT_MF=0.
'20181107: Modified by Oscar. Based on Customer's new Vdd_Binning_Def tables (Add columns of Interpolation factor and Info).
'20181029: Modified for Cebu interpolation, by Oscar.
'20171222: SWLINZA modified CPVmax to use E1.
Public Function ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site(p_mode As Integer, PassBinNum As SiteLong)
    Dim site As Variant
    Dim test_type As testType
    Dim start_p_mode As Integer '''start p_mode of interpolation
    Dim end_p_mode As Integer   '''end p_mode of interpolation
    Dim enableCalcInterpolation As New SiteBoolean
    '''
    Dim remainder As Double
    Dim voltage_INTP_L As Double
    Dim voltage_INTP_H As Double
    Dim voltage_Calc As Double
    Dim dbl_Interpolation_MF As Double
    Dim dbl_Interpolation_Offset As Double
    '''
    Dim gotCorretPmode As Boolean
    Dim step_Calc As Long
    Dim idx_step As Long
    Dim EQ_Num As Long
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. C651 Si added the column "INT_SKIPTEST" of Interpolation in table "Vdd_Binning_Def_appA_1" only, so that we only check this from Bin1 table.
'''2. Only Bin1 DUT uses Interpolation, and Bin1 has the Interpolation factor.
'''3. Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Interpolation_fail in Bin_Table before using this.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        '''init the flag
        enableCalcInterpolation = False
        DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = False
        step_Calc = -1
        gotCorretPmode = False
        
        '''//Check if p_mode is tested.
        If VBIN_RESULT(p_mode).tested = True Then
            gotCorretPmode = False
            '''ToDo: Maybe we can remove the error message...
            TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & " is tested. It should not be interpolated in Interpolatetion instance again. Error!!!"
            TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(p_mode) & " is tested. It should not be interpolated in Interpolatetion instance again. Error!!!"
        Else
            gotCorretPmode = True
        End If
        
        If gotCorretPmode = True Then
            '''//Align PassBinCutNum.
            '''20210408: Modified to update step for BinX/Y if AllBinCut(p_mode).INTP_SKIPTEST = True.
            If VBIN_RESULT(p_mode).passBinCut <> PassBinNum(site) Then
                '''20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
                Adjust_Multi_PassBinCut_Per_Site p_mode, site, PassBinNum(site)
            End If
            
            '''======================================================================================
            '''Only Bin1 DUT uses Interpolation, and Bin1 has the Interpolation factor.
            '''//If the Interpolation factor of p_mode is 0, it will skip interpolation calculation.
            '''======================================================================================
            If PassBinNum(site) = 1 And BinCut(p_mode, PassBinNum(site)).INTP_MFACTOR(0) <> 0 Then '''Bin1
                '''//All Equation in the interpolation item has the start P_mode, end P_mode and factor info. so step can be 0 ~ Bin1 maxstep.
                enableCalcInterpolation = True
                start_p_mode = BinCut(p_mode, PassBinNum(site)).INTP_MODE_L(0) 'only step0
                end_p_mode = BinCut(p_mode, PassBinNum(site)).INTP_MODE_H(0) 'only step0
                dbl_Interpolation_MF = BinCut(p_mode, PassBinNum(site)).INTP_MFACTOR(0)
                dbl_Interpolation_Offset = BinCut(p_mode, PassBinNum(site)).INTP_OFFSET(0)
                
                '''//start_p_mode and end_p_mode should be tested prior to p_mode interpolated.
                If enableCalcInterpolation = True And (VBIN_RESULT(start_p_mode).tested = False Or VBIN_RESULT(end_p_mode).tested = False) Then
                    enableCalcInterpolation = False
                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(start_p_mode) & " or " & VddBinName(end_p_mode) & " wasn't tested or failed. It should not be interpolated. Error!!!"
                    TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(start_p_mode) & " or " & VddBinName(end_p_mode) & " wasn't tested or failed. It should not be interpolated. Error!!!"
                    '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Interpolation_fail in Bin_Table before using this.
                    '''20210618: Modified to update SortNumber and binNumber if F_Vddbinning_Interpolation_fail is triggered in the vbt function ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site.
                    TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Interpolation_fail) = logicTrue
                End If
            Else '''non-Bin1
                enableCalcInterpolation = False
                step_Calc = VBIN_RESULT(p_mode).step_in_IDS_Zone
            End If
        End If
        
        '''//Only Bin1 DUT can calulate voltages for Interpolation.
        If enableCalcInterpolation = True Then
            '''************************************************************************************************************'''
            '''//Check if start P_mode and end P_mode of Interpolation exist (and are used for BinCut).
            '''(Errors occur when one of the interpolated performance mode and its start/end performance mode are not used)
            '''************************************************************************************************************'''
            If (DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True And DYNAMIC_VBIN_IDS_ZONE(start_p_mode).Used = True And DYNAMIC_VBIN_IDS_ZONE(end_p_mode).Used = True) Then
                '''======================================================================================
                '''[Step1] Vi = Vlow + (Vhigh - Vlow) * MF
                '''======================================================================================
                '''//Calcute voltage by Interpolation from start_p_mode and end_p_mode.
                voltage_INTP_L = VBIN_RESULT(start_p_mode).GRADE(site)
                voltage_INTP_H = VBIN_RESULT(end_p_mode).GRADE(site)
                voltage_Calc = voltage_INTP_L + (voltage_INTP_H - voltage_INTP_L) * dbl_Interpolation_MF
                
                '''======================================================================================
                '''[Step2] Vx = Vi + Int_Offset
                '''======================================================================================
                If dbl_Interpolation_Offset <> 0 Then
                    voltage_Calc = voltage_Calc + dbl_Interpolation_Offset
                End If
                
                '''//Ceiling voltage_Calc by BV_StepVoltage defined in BinCut voltage table(sheet "Vdd_Binning_Def").
                '''20210412: Modified to remove the vbt code not to check if the voltage_Calc is in the range between CPVmin and CPVmax, requested by PCLINZG.
                '''20210831: Modified to remove the vbt code related to CPVmax and CPVmin.
                remainder = Ceiling(voltage_Calc / BV_StepVoltage)
                voltage_Calc = remainder * BV_StepVoltage
                
                '''======================================================================================
                '''[Step3] Round Vx up to next bincut equation: Vx_rounded = RoundUpToEQN(Vx).
                ''' 1. Find the new_En_step in IDS Zone(most close to and higher than Vx EBB).
                ''' 2. Find the last_bin1_step in IDS Zone(last step of Bin1 in IDS ZOne).
                '''======================================================================================
                '''//Find the nearest step for interpolated voltage (step_interpolated).
                For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
                    If CDec(DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)) >= CDec(voltage_Calc) And step_Calc = -1 Then
                        step_Calc = idx_step
                        Exit For
                    End If
                Next idx_step
            Else
                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & ", Dynamic_IDS ZONE doesn't exist. it can't do Interpolation. Error!!!"
                TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(p_mode) & ", Dynamic_IDS ZONE doesn't exist. it can't do Interpolation. Error!!!"
            End If
        End If
        
        '''//Check if step_Calc is valid in Dynamic_IDS_Zone.
        If step_Calc <> -1 Then
            DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = True
            
            '''==================================================================================================================================================================
            '''//For SkipTest = True  : If no step in Bin1/X/Y available, do not use step Bin1 EQN1, just bin out the failed DUT with failFlag "F_Vddbinning_Interpolation_fail".
            '''//For SkipTest = False : If step_Calc is greater than step(PassBinNum(site),EQN1), set step(PassBinNum(site),EQN1) as 1st step of Interpolation.
            '''==================================================================================================================================================================
            If AllBinCut(p_mode).INTP_SKIPTEST = True Then
                '''//Maybe step_Calc is in the higher PassBin, it should update VBIN_RESULT(p_mode).step_in_IDS_Zone and CurrentPassBinCutNum(site).
                If CurrentPassBinCutNum(site) <> DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(VBIN_RESULT(p_mode).step_in_IDS_Zone) Then
                    CurrentPassBinCutNum(site) = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(VBIN_RESULT(p_mode).step_in_IDS_Zone)
                    '''20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
                    Adjust_Multi_PassBinCut_Per_Site p_mode, site, CurrentPassBinCutNum(site)
                End If
            Else '''If AllBinCut(p_mode).INTP_SKIPTEST = False
                '''//If step_Calc is greater than step(PassBinNum(site),EQN1), set step(PassBinNum(site),EQN1) as 1st step of Interpolation.
                If step_Calc > DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(PassBinNum(site), 1) Then
                    step_Calc = DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(PassBinNum(site), 1)
                End If
            End If
        Else '''//If step_Calc = -1, it means no matched step for Interpolation, it only overwrite step with step_mapping(Bin1,EQN1) for Bin1 DUT; otherwise, bin out the failed DUT.
            If PassBinNum(site) = 1 Then
                If AllBinCut(p_mode).INTP_SKIPTEST = True Then
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = False
                    TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Interpolation_fail) = logicTrue
                    TheExec.Datalog.WriteComment "site:" & site & ", " & "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", Binning Mode:" & VddBinName(p_mode) & ", no step available in Bin1/X/Y for " & VddBinName(start_p_mode) & ". Bin out the failed DUT."
                    TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                Else '''If AllBinCut(p_mode).INTP_SKIPTEST = False
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = True
                    step_Calc = DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(PassBinNum(site), 1)
                End If
            Else
                DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = False
                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Interpolation_fail) = logicTrue
                TheExec.Datalog.WriteComment "site:" & site & ", " & "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", Binning Mode:" & VddBinName(p_mode) & ", no step available in Bin1/X/Y for " & VddBinName(start_p_mode) & ". Bin out the failed DUT."
                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
            End If
        End If
        
        '''//Update step_Calc to VBIN_RESULT(p_mode).step_in_IDS_Zone.
        If DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = True Then '''If Interpolation is completed, update VBIN_RESULT(p_mode) and print out the related info.
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Interpolated_Start = step_Calc
            VBIN_RESULT(p_mode).step_in_IDS_Zone = step_Calc
            '''20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_inherit = 0
            
            '''==================================================================================================================================================================
            '''//Default: test_type = testType.TD
            '''//DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(test_type) will be used as step_inherit for the vbt function "find_start_voltage".
            '''Remember to check the branches in the vbt function "find_start_voltage"!!!
            '''==================================================================================================================================================================
            '''20210408: Modified to overwrite step_inherit and VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone if DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).interpolated = True.
            For test_type = 0 To MaxTestType - 1
                DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(test_type) = step_Calc
            Next test_type
            
            '''//Print info about result of the interpolation.
            '''ex: Site:1,CurrentPassBinCutNum:1,Binning Mode:VDD_GPU_MG003,The lowest Performance Mode:MG002,The highest Performance Mode:MG007,The MFx:0.19,SkipTest:True,The interpolated Voltage:565.625,The selected Eqn:5,The selected Voltage:565.625
            If CurrentPassBinCutNum(site) = 1 Then
                TheExec.Datalog.WriteComment "site:" & site & ", " & _
                                            "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", " & _
                                            "Binning Mode:" & VddBinName(p_mode) & ", " & _
                                            "The lowest Performance Mode:" & VddBinName(CInt(start_p_mode)) & ", " & _
                                            "The highest Performance Mode:" & VddBinName(CInt(end_p_mode)) & ", " & _
                                            "The MFx:" & dbl_Interpolation_MF & ", " & _
                                            "SkipTest:" & AllBinCut(p_mode).INTP_SKIPTEST & ", " & _
                                            "The interpolated Voltage:" & voltage_Calc & " mV" & ", " & _
                                            "The selected Eqn:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(step_Calc) & ", " & _
                                            "The selected Voltage:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(step_Calc) & " mV"
            ElseIf AllBinCut(p_mode).INTP_SKIPTEST = True Then
                '''20210813: Modified to print interpolation info while SkipTest=True for BinX and BinY.
                TheExec.Datalog.WriteComment "site:" & site & ", " & _
                                            "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", " & _
                                            "Binning Mode:" & VddBinName(p_mode) & ", " & _
                                            "The lowest Performance Mode:" & VddBinName(CInt(start_p_mode)) & ", " & _
                                            "The highest Performance Mode:" & VddBinName(CInt(end_p_mode)) & ", " & _
                                            "The MFx:" & dbl_Interpolation_MF & ", " & _
                                            "SkipTest:" & AllBinCut(p_mode).INTP_SKIPTEST & ", " & _
                                            "The interpolated Voltage:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(step_Calc) & " mV" & ", " & _
                                            "The selected Eqn:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(step_Calc) & ", " & _
                                            "The selected Voltage:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(step_Calc) & " mV"
            End If
            
            '''************************************************************************************************************************************************************'''
            '''//If Interpolation SkipTest="Yes" (AllBinCut(p_mode).INTP_SKIPTEST = True) ==> Just put step0 of "DYNAMIC_VBIN_IDS_ZONE(p_mode)" into "VBIN_RESULT(p_mode)".
            '''************************************************************************************************************************************************************'''
            If AllBinCut(p_mode).INTP_SKIPTEST = True Then
                '''//Update PassBin, Pass step, flag"VBIN_Result(p_mode).tested", and voltage to VBIN_Result by the step in Dynamic_IDS_Zone.
                '''20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
                Call Set_VBinResult_by_Step(site, p_mode, step_Calc)
                
                VBIN_RESULT(p_mode).FLAGFAIL = False
            End If
        End If '''DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated = True
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200211: Modified to replace "cntFlowTestCond" with "cntAdditionalMode".
'20200211: Modified to replace the function name "FlowTestCondStr2Enum" with "AdditionalModeStr2Enum".
'20191202: Modified for the revised initVddBinCondition.
Public Function AdditionalModeStr2Enum(additional_mode As String) As Integer
On Error GoTo errHandler
    additional_mode = UCase(additional_mode)
    
    If AdditionalModeDict.Exists(additional_mode) Then
        AdditionalModeStr2Enum = AdditionalModeDict.Item(additional_mode)
    Else
        AdditionalModeStr2Enum = cntAdditionalMode + 1
        TheExec.Datalog.WriteComment "Enum:" & AdditionalModeStr2Enum & ", TestCond = " & additional_mode & " doesn't exist in enum FlowSheetCondition. Error!!!"
        TheExec.ErrorLogMessage "Enum:" & AdditionalModeStr2Enum & ", TestCond = " & additional_mode & " doesn't exist in enum FlowSheetCondition. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of AdditionalModeStr2Enum"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190502: Modified for OFFSET_FUNC, ex: "MG005_GFXTD_BPL_BV".
Public Function decide_offset_testType(strInput As String) As Integer
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = LCase(strInput)
    
    If strTemp Like "*_td*" Then
        decide_offset_testType = testType.TD
    ElseIf strTemp Like "*_bist*" Then
        decide_offset_testType = testType.Mbist
    ElseIf strTemp Like "*_func*" Then '''added for dynamic offset_Func, 20190502
        decide_offset_testType = testType.Func
    Else
        decide_offset_testType = testType.ldcbfd '''It's a pseudo testType. If we read "Offset_CP1_FUNC" from sheet, set it as test type "LDCBFD".
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of decide_offset_testType"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200817: Modified to added "strTemp=Lcase(inst_name)".
'20191107: Modified to add FUNC.
'20190502: Modified for OFFSET_FUNC, ex: "MG005_GFXTD_BPL_BV".
Public Function decide_offset_testType_byInstName(inst_name As String) As Integer
    Dim strTemp As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''For those instances with special offsets different from TD or Mbist, it should use keywords in the instance names same as special keywords added in column of FUNC, ex: BPL.
'''//==================================================================================================================================================================================//'''
    strTemp = LCase(inst_name)

    If strTemp Like "*cpu*bist*" Or strTemp Like "*gfx*bist*" Or strTemp Like "*soc*bist*" Or strTemp Like "*gpu*bist*" Then
        decide_offset_testType_byInstName = testType.Mbist
    ElseIf strTemp Like "*gfxtd*" Or strTemp Like "*cputd*" Or strTemp Like "*soctd*" Or strTemp Like "*gputd*" Then
        decide_offset_testType_byInstName = testType.TD
    Else
        'decide_offset_testType_byInstName = TestType.ldcbfd 'fake test type, if no match td/mbist/spi..., pretend as test type "LDCBFD"
        decide_offset_testType_byInstName = testType.Func
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of decide_offset_testType_byInstName"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
'20210129: Modified to revise the vbt code for DevChar.
'20201124: Modified to remove "Dim CorePowerStored_Init As New SiteDouble".
'20201118: Modified to use "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" to get siteResult of pattern pass/fail.
'20201029: Modified to remove the argument "result_mode As tlResultMode" and use inst_info.result_mode.
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to remove the argument "Optional idxBlock_Selsrm_singlePatt As Integer".
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to add the argument "IndexLevelPerSite As SiteLong".
'20201027: Modified to use "Public Type Instance_Info".
'20201026: Modified to revise the vbt code for TD pattern burst proposed by C651 Toby.
'20201012: Modified to use "check_patt_Pass_Fail" to check pattern Pass/Fail.
'20200924: Modified to move "select_DCVS_output_for_powerDomain" from GradeSearch_VT to "run_prepatt_decompose_VT".
'20200923: Modified to move the position of "Check_Pattern_NoBurst_NoDecompose".
'20200923: Modified to remove "run_prepatt" and keep "run_prepatt_decompose_VT".
'20200921: Modified to check if "Test_Type = TestType.Mbist".
'20200520: Modified to use Check_Pattern_NoBurst_NoDecompose to show the errorLogMessage if "burst=no" and "Decompose_Pattern=false".
'20200319: Modified to switch off save_core_power_vddbinning and restore_core_power_vddbinning if Flag_Enable_Rail_Switch = True.
'20200203: Modified to use the function "print_bincut_power".
'20200115: Modified to skip applying safe voltages to non-selsram powerpin for project with selsrm_mapping_table.
'20200113: Modified for pattern bursted without decomposing pattern.
'20200106: Modified to add "TheHdw.Alarms.Check".
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20191125: Modified PrePattPass to avoid pseudo pass.
'20191125: Modified to remove IGSIM block.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerPins.
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerPins.
'20190606: Modified to add the argument "DcSpecsCategoryForInitPat as string" for Init patterns with the new test setting DC Specs.
Public Function run_prepatt_decompose_VT(inst_info As Instance_Info, PrePatt As String, ary_PrePatt_decomposed() As String, count_PrePatt_decomposed As Long, PrePattPass As SiteBoolean, Optional DcSpecsCategoryForInitPat As String)
    Dim site As Variant
    Dim CorePowerStored() As New SiteDouble
    Dim indexPatt As Long
    Dim i As Integer
    Dim offsetTestTypeIdx As Integer
    Dim sitePatPass As New SiteBoolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
'''It seemed that theExec.sites.Selected masked the failed site (not siteShutDown). But the site still ran pattern.test without updating test results.
'''Discussed this with Chihome. He saw this is ancient projects, and he suggested us to check if test results were correct.
'''We checked test results, and it seemed no error with PassBin and EQN.
'''//==================================================================================================================================================================================//'''
    If PrePatt <> "" Then
        '''//Update theExec.sites.Selected for MultiFSTP before running PrePatt.
        '''20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
        '''ToDo: Please check if EnableWord("Multifstp_Datacollection") exists in the flow table!!!
        If EnableWord_Multifstp_Datacollection Then
            TheExec.sites.Selected = gb_siteMask_current
        End If
    
        '''//For BIST instance, it saves current DCVS Vmain values into the array "CorePowerStored", and then set safe voltage to DCVS Vmain by DC Category.
        If inst_info.test_type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
            '''init
            '''//siteDouble "CorePowerStored()" is used to save/restore voltages for BinCut powerDomains.
            ReDim CorePowerStored(UBound(pinGroup_BinCut))
            
            For i = 0 To UBound(pinGroup_BinCut)
                CorePowerStored(i) = 0
            Next i
        
            '''//Save payload voltages of CorePower and OtherRail powerPins before init pattern.
            If Flag_noRestoreVoltageForPrepatt = False Then
                save_core_power_vddbinning CorePowerStored
            End If
            
            '''//Get BinCut INIT voltages (safe voltages for SELSRM DSSC), usually set to nominal voltage.
            '''//If initial voltages and safe voltage(init voltage) use the same DC category, it can skip "set_core_power_vddbinning_VT" after initial voltages...
            '''Note: For PTE/TTR, it can use the flag "Flag_Skip_ReApplyInitVolageToDCVS" to skip "set_core_power_vddbinning_VT".
            set_core_power_vddbinning_VT "NV", DcSpecsCategoryForInitPat
            TheHdw.Wait 0.0001
        End If
        
        '''//Print safe voltages(init voltages) for PrePatt(init patt).
        print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.SafeVoltage
        
'**********************************************
'@@PrePatt pattern-loop Start
'**********************************************
        For indexPatt = 0 To count_PrePatt_decomposed - 1
            '''//Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
            Call prepare_DCVS_Output_for_RailSwitch(inst_info, ary_PrePatt_decomposed(indexPatt), inst_info.idxBlock_Selsrm_PrePatt)
            
            '''//Run pattern.
            Call TheHdw.Patterns(inst_info.ary_PrePatt_decomposed(indexPatt)).Test(pfAlways, 0, inst_info.result_mode)
            
            '''//Get siteResult of pattern pass/fail.
            '''//Warning!!! currently "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" doesn't support "result_mode=tlResultModeModule" with PatternBurst=Yes and DecomposePatt=No.
            sitePatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            
            '''//Check alarmFail for pattern.
            Call check_alarmFail_for_pattern(sitePatPass)


            sitePatPass.Value = sitePatPass.LogicalAnd(PrePattPass)
            '''for DevChar.
'            If inst_info.is_DevChar_Running = False Then
'                '''//Update pattern pass/fail to the flag.
'                Call update_Pattern_result_to_PattPass(sitePatPass, PrePattPass)
'            End If
        Next indexPatt
        
        DebugPrintFunc PrePatt
'**********************************************
'@@PrePatt pattern-loop End
'**********************************************
        '''//Check if running Pattern with "burst=no" and "Decompose_Pattern=false".
        Call Check_Pattern_NoBurst_NoDecompose(inst_info.PrePatt, inst_info.count_PrePatt_decomposed, inst_info.enable_DecomposePatt)
        
        '''//Restore the BinCut voltages for payload patterns after init pattern.
        If inst_info.test_type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
            If Flag_noRestoreVoltageForPrepatt = False Then
                restore_core_power_vddbinning CorePowerStored
            End If
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of run_prepatt_decompose_VT"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210823: Modified to revise the vbt code for C651 new rules of testJobs naming, as requested by C651 Toby and TSMC ZYLINI.
'20200827: Modified to replace "If..Else" with "Select Case".
'20200731: Modified to merge MappingBincutJobName and Mapping_TPJobName_to_BincutJobName into Mapping_TestJobName_to_BincutJobName.
'20180704: Created for BinCut testjob mapping.
Public Function Mapping_TestJobName_to_BincutJobName(strTestJob As String) As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//Please discuss TestJobs of test program with C651 project DRI and decide how to mapping testjobs (especially for FT3 and XXX_QA) to 5 BinCut testJobs (CP1, CP2, FT_ROOM, FT_HOT, QA)!!!
'''For example, BinCut doesn't have FT3, but DC_TEST_IDS still uses BinCut IDS limit as IDS HiLimit.
'''//==================================================================================================================================================================================//'''
    '''//This subroutine needs to be kept maintaining depending on individual project// ''''
    Select Case LCase(strTestJob)
        '''//cp1
        Case "cp1": Mapping_TestJobName_to_BincutJobName = "cp1"
        '''//cp2
        Case "cp2": Mapping_TestJobName_to_BincutJobName = "cp2"
        '''//ft_room
        Case "ft_room": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "wlft1": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "ft1": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "ft2_25c": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "ft3": Mapping_TestJobName_to_BincutJobName = "ft_room"
        '''20210823: Modified to revise the vbt code for C651 new rules of testJobs naming, as requested by C651 Toby and TSMC ZYLINI.
        Case "rma_room": Mapping_TestJobName_to_BincutJobName = "ft_room"
        '''//ft_hot
        Case "ft_hot": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        Case "ft2": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        Case "ft2_85c": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        '''20210823: Modified to revise the vbt code for C651 new rules of testJobs naming, as requested by C651 Toby and TSMC ZYLINI.
        Case "rma_hot": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        '''//qa
        Case "qa": Mapping_TestJobName_to_BincutJobName = "qa"
        '''20200827: Modified to mask "wlft1_qa" and "ft2_85c_qa" because these testJob are not the common testJobs in recent projects. These need to be discussed with C651 project DRIs!!!
        'Case "wlft1_qa": Mapping_TestJobName_to_BincutJobName = "qa"
        'Case "ft2_85c_qa": Mapping_TestJobName_to_BincutJobName = "qa"
'''ToDo: Discuss rules of testJobs mapping for "FT1_FQA", "FT2_FQA", "T0TX_ROOM", "T0TX_HOT" with C651 and TSMC...
        '''//others
        Case Else:
                Mapping_TestJobName_to_BincutJobName = ""
                TheExec.Datalog.WriteComment "job:" & strTestJob & ", it doesn't have any matched definition for Mapping_TestJobName_to_BincutJobName. Error!!!"
                TheExec.ErrorLogMessage "job:" & strTestJob & ", it doesn't have any matched definition for Mapping_TestJobName_to_BincutJobName. Error!!!"
    End Select
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Mapping_TestJobName_to_BincutJobName"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Mapping_TestJobName_to_BincutJobName"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
'20210707: Modified to check if ids_name (Efuse category) exists in dict_EfuseCategory2BinCutTestJob.
'20210617: Discussed this with TSMC T-Cre team and C651 Si. C651 Si said that Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201005: Modified to set Lo_Limit=0 for non-CP1.
'20200827: Modified to replace "If..Else" with "Select Case".
'20200106: Modified to remove the ErrorLogMessage.
'20190813: Modified to use different IDS lo_limit by BinCut testjobs.
'20190722: Modified to printout the scale and the unit for BinCut voltages and IDS values.
'20190716: Modified to unify the unit for IDS.
'20190612: Modified to use the new datatype of IDS.
'20180917: Due to data with double format accuracy issue, we follow the suggestion from Microsoft official document to use "Cdec".
'20180209: SWLINZA modified this because we should read efuse to identify BinX/BinY.
'20170810: SWLINZA modified to get Resolution from EFUSE_BitDef_Table and set it as IDS low limit for CorePower.
Public Function judge_IDS(ids_current As SiteDouble, performance_mode As String, site As Variant)
    Dim hi_limit As Double
    Dim lo_limit As Double
    Dim p_mode As Integer
    Dim powerDomain As String
    Dim i As Long
    Dim Mode_Step As Integer
    Dim str_IDS_PowerDomain As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''2. C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs, 20210727.
'''//==================================================================================================================================================================================//'''
    '''//Get p_mode from performance mode
    p_mode = VddBinStr2Enum(performance_mode)
    powerDomain = AllBinCut(p_mode).powerPin
    
    '''//Get IDS name by powerDomain for each site.
    str_IDS_PowerDomain = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site)  'Modify to site variant
        
    '''************************************************************************************************************************************************************'''
    '''//Use Resolution of IDS from EFUSE_BitDef_Table as IDS low limit with the scale and the unit in "mA".
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    '''20210617: Discussed this with TSMC T-Cre team and C651 Si. C651 Si said that Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search.
    '''************************************************************************************************************************************************************'''
    '''//Check if ids_name of PowerDomain is the correct Efuse category in Efuse_BitDef_Table.
    If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True Then
        '''For project with Efuse DSP vbt code.
        lo_limit = 1# * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA

        '''************************************************************************************************************************************************'''
        ''' Loop the BinCut table and use the different CP current limit to print out the datalog and only judge the IDS by AllBinCut(P_mode).IDS_CP_LIMIT
        '''************************************************************************************************************************************************'''
        For i = 0 To UBound(PassBinCut_ary)
            '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
            If i = UBound(PassBinCut_ary) Then
                hi_limit = AllBinCut(p_mode).IDS_CP_LIMIT '''unit: mA
                
                TheExec.Flow.TestLimit ids_current(site) / 1000, lo_limit / 1000, hi_limit / 1000, , tlSignLess, scaleMilli, Unit:=unitAmp, _
                                        PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut" & PassBinCut_ary(i) & " IDS", ForceUnit:=unitAmp
            Else
                Mode_Step = BinCut(p_mode, PassBinCut_ary(i)).Mode_Step
                
                '''//IDS calculation uses the scale and the unit in "mA".
                hi_limit = BinCut(p_mode, PassBinCut_ary(i)).IDS_CP_LIMIT(Mode_Step) '''unit: mA
                
                '''//TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
                '''20180917: Due to data with double format accuracy issue, we follow the suggestion from Microsoft official document to use "Cdec".
                If CDec(ids_current) < CDec(hi_limit) Then
                    TheExec.Flow.TestLimit ids_current(site) / 1000, lo_limit / 1000, hi_limit / 1000, , tlSignLess, scaleMilli, Unit:=unitAmp, _
                                            PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut" & PassBinCut_ary(i) & " IDS", ForceUnit:=unitAmp
                Else
                    TheExec.Flow.TestLimit ids_current(site) / 1000, lo_limit / 1000, hi_limit / 1000, , tlSignLess, scaleMilli, Unit:=unitAmp, _
                                            PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut" & PassBinCut_ary(i) & " IDS", ForceResults:=tlForcePass, ForceUnit:=unitAmp
                End If
            End If
        Next i
    Else
        TheExec.Datalog.WriteComment performance_mode & ",Efuse category:" & str_IDS_PowerDomain & ",it can't use Efuse category to get IDS values for judge_IDS. Error!!!"
        TheExec.ErrorLogMessage performance_mode & ",Efuse category:" & str_IDS_PowerDomain & ",it can't use Efuse category to get IDS values for judge_IDS. Error!!!"
    End If '''If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in the subroutine of judge_IDS"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210907: Modified to check if siteNumber from the argument site is correct.
'20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
'20210707: Modified to check if ids_name (Efuse category) exists in dict_EfuseCategory2BinCutTestJob.
'20210707: Modified to add site-loop to trig the failFlag strGlb_Flag_Vddbinning_IDS_fail for each site.
'20210629: Modified to adjust the sequence of the vbt code in get_I_VDD_values.
'20210507: Modified to replace testLimit with failStop for get_I_VDD_values.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20200815: Modified to check powerPin exists.
'20200812: Modified to check powerDomain and powerPin.
'20200717: Modified to use the globalVariable "Flag_Vddbinning_IDS_fail".
'20200712: Modified to bin out the DUT due to IDS failed, suggested and requested by PCLIN.
'20200430: Modified to print the string about the incorrect IDS.
'20200130: Modified to get 1st powerPin from powerDomain.
'20200114: Modified to check if powerDomain exists in domain2pinDict or pin2domainDict.
'20200106: Modified to remove the ErrorLogMessage.
'20190716: Modified to unify the unit for IDS.
'20190630: Modified to use the real IDS values for non-CP1 tests.
'20190624: Modified to unify the unit of IDS with "A".
'20190615: Modified to align "Efuse Read Write Decimal" with I_VDD_xxx values
'20190523: Modified the argument for the new IDS datatype.
Public Function get_I_VDD_values(site As Variant, powerDomain As String, I_VDD_val As SiteDouble)
    Dim str_IDS_PowerDomain As String
    Dim powerPin As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''2. Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_IDS_fail in Bin_Table before using this.
'''3. As per discussion with TSMC SWLINZA, for powerPin group, it should use 1st powerPin to check IDS limit of powerPin group, 20210707.
'''ex: powerGroup: VDD_FIXED_GRP, and its 1st powerPin: VDD_FIXED, so that compare IDS value of VDD_FIXED with IDS_limit of VDD_FIXED_GRP. It must have Efuse category in Efuse_BitDef_Table to store IDS for VDD_FIXED.
'''4. C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs, 20210727.
'''//==================================================================================================================================================================================//'''
    '''//Get the selected site.
    If CLng(site) < 0 Then
        TheExec.Datalog.WriteComment "Please check the site number of get_I_VDD_values. Error!!!"
        'TheExec.ErrorLogMessage "Please check the site number of get_I_VDD_values. Error!!!"
    End If
    
    '''//Warning!!! Please contact project Efuser owner to see if CFG Read/Write with scale mA or not.
    '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    '''ToDo: Please discuss this with project Efuse owner to see if rules about "Programming Stage" in Efuse_BitDef_Table are changed.
    '''20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
    '''***************************************************************************************************************************************************'''
    '''[CP1] get IDS values from Efuse CFG data structure by IDS name.
    '''***************************************************************************************************************************************************'''
    '''//Step1: Get IDS name of powerDomain for each site.
    str_IDS_PowerDomain = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site)

    '''//Step2: Get IDS value from Efuse CFG by IDS name of powerDomain.
    '''//If IDS_name of PowerDomain is correct, it will get IDS real values from Efuse Read.Decimal or Write.Decimal.
    If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True Then
        '''==========================================================================================================
        '''Note: Read IDS from efuse Read Decimal.
        '''If the IDS from efuse Read Decimal is 0, it can get IDS from efuse Write Decimal. Make sure the IDS is not 0.
        '''==========================================================================================================
        '''For project with Efuse DSP vbt code.
        If CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Read.Decimal(site) * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion <> 0 Then
            I_VDD_val(site) = CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Read.Decimal(site) * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA
        ElseIf CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Write.Decimal(site) * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion <> 0 Then
            I_VDD_val(site) = CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Write.Decimal(site) * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA
        Else
            I_VDD_val(site) = 0
            TheExec.Datalog.WriteComment "site:" & site & ",powerDomain:" & powerDomain & ",Efuse category:" & str_IDS_PowerDomain & ", IDS value from Efuse CFG is 0. Please check DC_TEST_IDS or EFuse. Error!!!"
            '''//Use the globalVariable "Flag_Vddbinning_IDS_fail" to bin out the DUT due to IDS failed, as suggested and requested by PCLIN.
            TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_IDS_fail) = logicTrue
        End If
    Else
        I_VDD_val(site) = 0
        TheExec.Datalog.WriteComment "site:" & site & ",powerDomain:" & powerDomain & ",Efuse category:" & str_IDS_PowerDomain & ", it can't use Efuse category for get_I_VDD_values, Please check Efuse_BitDef_Table. Error!!!"
        'TheExec.ErrorLogMessage "site:" & site & ",powerDomain:" & powerDomain & ",Efuse category:" & str_IDS_PowerDomain & ", it can't use Efuse category for get_I_VDD_values, Please check Efuse_BitDef_Table. Error!!!"
    End If
    'Jeff
    I_VDD_val(site) = 0.5
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Please check the input for get_I_VDD_values!!! Error!!!"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20180807: Modified for BinCut testjob mapping.
Public Function getBinCutJobDefinition(strInput As String) As Integer
    Dim strTemp As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''We use strInput (the keyword about testJob) to decide the jobidx, Enum BinCutJobDefinition in global variable.
'''define CP1 = 0, CP2 = 1, FT1 = 2, FT2 = 3, QA = 4.
'''//==================================================================================================================================================================================//'''
    strTemp = LCase(Trim(strInput))

    If strTemp Like "*cp1*" Or strTemp Like "*binsearch*" Then
        getBinCutJobDefinition = BinCutJobDefinition.CP1
    ElseIf strTemp Like "*cp2*" Then
        getBinCutJobDefinition = BinCutJobDefinition.CP2
    ElseIf strTemp Like "*ft_room*" Or strTemp Like "*ft1" Or strTemp Like "*ft2_25c" Then
        getBinCutJobDefinition = BinCutJobDefinition.FT1
    ElseIf strTemp Like "*ft_hot*" Or strTemp Like "*ft2_85c*" Then
        getBinCutJobDefinition = BinCutJobDefinition.FT2
    ElseIf strTemp Like "*qa*" Then
        getBinCutJobDefinition = BinCutJobDefinition.QA
    Else
        getBinCutJobDefinition = BinCutJobDefinition.COND_ERROR
        TheExec.Datalog.WriteComment "getBinCutJobDefinition = " & strInput & " doesn't exist in enum BinCutJobDefinition. Error!!!"
        'TheExec.ErrorLogMessage "getBinCutJobDefinition = " & strInput & " doesn't exist in enum BinCutJobDefinition. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of getBinCutJobDefinition"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210701: Modified to remove "Exit Function" from the vbt function check_voltageInheritance_for_powerDomain, as requested by TER Verity.
'20210621: Modified to use the argument "powerDomain as string" for the input powerDomain of the vbt function check_voltageInheritance_for_powerDomain.
'20210621: Modified to remove the redundant argument "ids_PowerDomain As SiteDouble" from the vbt function check_voltageInheritance_for_powerDomain.
'20210621: Modified to rename the vbt function check_pmode_for_adjust_VddBinning as check_voltageInheritance_for_powerDomain.
'20210621: Modified to merge the vbt code from the vbt function find_next_bin_eq_interpolation.
'20210303: Modified to remove the redundant "ids_current As SiteDouble" from arguments of the vbt function "find_next_bin_eq_interpolation".
'20201113: Modified to rename the argument "I_VDD_core_power" as "ids_PowerDomain".
'20200825: Modified to remove the redundant branches.
'20191127: Modified for the revised InitVddBinTable.
'20190227: For complete BinCut search steps in DYNAMIC_VBIN_IDS_ZONE, we replace "find_next_bin_eq" with "find_next_bin_eq_interpolation" according to BinCut monthly meeting Dec-2018.
'20181026: Added for interpolation by Oscar.
'20180816: Created this function to simplify adjust_VddBinning.
Public Function check_voltageInheritance_for_powerDomain(powerDomain As String)
    Dim site As Variant
    Dim strAry_Pmode_Seq() As String
    Dim p_mode As Integer
    Dim gradevdd_last As New SiteDouble
    Dim grade_last As New SiteDouble
    Dim i As Long
    '''
    Dim idx_step As Long
    Dim next_bin_flag As Boolean
    Dim exit_while_flag As Boolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Only testjob with keyword "*Evaluate*Bin*" in the testCondition of BinCut flow can check voltage inheritance in Adjust_VddBinning.
'''2. 20210610 C651 Si defined the rule of voltage inheritance check: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset if without Allow_Equal.
'''//==================================================================================================================================================================================//'''
    '''//Check if powerDomain is the BinCut CorePower defined in the header of BinCut flow(sheet "Non_Binning_Rail").
    If dict_IsCorePowerInBinCutFlowSheet.Exists(UCase(powerDomain)) = True Then
        If dict_IsCorePowerInBinCutFlowSheet.Item(UCase(powerDomain)) = True Then
            strAry_Pmode_Seq = BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
            
            If UBound(strAry_Pmode_Seq) >= 1 Then
                '''BinCut powerDomain has more than one p_mode, and it's ready to check voltage inheritance for the Power_Seq of BinCut powerDomain.
            Else
                '''If UBound(strAry_Pmode_Seq)=0, skip voltage inheritance check for BinCut powerDomain.
                Exit Function
            End If
        Else
            TheExec.Datalog.WriteComment "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
            'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
            Exit Function
        End If
    Else
        TheExec.Datalog.WriteComment "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
        'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
        Exit Function
    End If
    
    '''//If UBound(strAry_Pmode_Seq) >= 1, adjust the vdd binning value to make voltage of each p_mode is always greater than its previous performance_mode for BinCut powerDomain.
    For i = 1 To UBound(strAry_Pmode_Seq)
        p_mode = VddBinStr2Enum(strAry_Pmode_Seq(i))
        
        '''//Get BinCut voltage and Efuse product voltage of the previous performance_mode.
        gradevdd_last = VBIN_RESULT(AllBinCut(VddBinStr2Enum(strAry_Pmode_Seq(i))).PREVIOUS_Performance_Mode).GRADEVDD
        grade_last = VBIN_RESULT(AllBinCut(VddBinStr2Enum(strAry_Pmode_Seq(i))).PREVIOUS_Performance_Mode).GRADE
        
        If BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut(site)).ExcludedPmode = True Then
            '''Set Grade and GradeVdd to 0 if the Performance Mode is not enabled in the Test Result (Bin1 or BinX).
            VBIN_RESULT(p_mode).GRADE = 0
            VBIN_RESULT(p_mode).step_in_BinCut = -1
            VBIN_RESULT(p_mode).GRADEVDD = 0
        Else
            '''//Check the voltage heritance between p_mode and the previous performance_mode.
            For Each site In TheExec.sites
                next_bin_flag = False
            
                If CDec(VBIN_RESULT(p_mode).GRADE) > 0 Then
                    idx_step = VBIN_RESULT(p_mode).step_in_IDS_Zone
                    exit_while_flag = False
                    
                    '''//Check if p_mode has Allow_Equal with the previous performance_mode.
                    If AllBinCut(p_mode).PREVIOUS_Performance_Mode = AllBinCut(p_mode).Allow_Equal And AllBinCut(p_mode).Allow_Equal <> 0 Then '''for AllowEqual
                        '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                        '''//Print the status of VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) with PTR format in Adjust_Binning for datalogs.
                        '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
                        If Flag_Get_column_Monotonicity_Offset = True Then
                            If (CDec(VBIN_RESULT(p_mode).GRADEVDD - gradevdd_last(site)) < CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                                VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                            End If
                        End If
                        
                        '''//Check if GradeVDD(p_mode) => GradeVDD(previous_performance_mode) + Monotonicity_Offset(p_mode).
                        While (CDec(VBIN_RESULT(p_mode).GRADEVDD) < CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False) _
                        'Or (CDec(VBIN_RESULT(p_mode).GRADE) < CDec(grade_last) And exit_while_flag = False)
                            idx_step = idx_step + 1
                            
                            If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) <> CurrentPassBinCutNum Then
                                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to adjust the product voltage for voltage inheritance check. Error!!!"
                                exit_while_flag = True
                                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                                TheExec.sites.Item(site).SortNumber = 9801
                                TheExec.sites.Item(site).binNumber = 5
                                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                                '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                                TheExec.sites.Item(site).result = tlResultFail
                            Else
                                VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                                VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                                next_bin_flag = True
                            End If
                        Wend
                    Else
                        '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                        '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
                        '''20210610: Modified for the rule: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset if without Allow_Equal.
                        If Flag_Get_column_Monotonicity_Offset = True Then
                            If (CDec(VBIN_RESULT(p_mode).GRADEVDD - gradevdd_last(site)) <= CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                                VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                            End If
                        End If
                        
                        '''//Check if GradeVDD(p_mode) > GradeVDD(previous_performance_mode) + Monotonicity_Offset(p_mode).
                        While (CDec(VBIN_RESULT(p_mode).GRADEVDD) <= CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False) _
                        'Or (CDec(VBIN_RESULT(p_mode).GRADE) <= CDec(grade_last) And exit_while_flag = False)
                            idx_step = idx_step + 1
                            
                            If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) <> CurrentPassBinCutNum Then
                                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to adjust the product voltage for voltage inheritance check. Error!!!"
                                exit_while_flag = True
                                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                                TheExec.sites.Item(site).SortNumber = 9801
                                TheExec.sites.Item(site).binNumber = 5
                                '''ToDo: Maybe we can use this fail-stop flag to mask the failed DUT in Adjust_VddBinning...
                                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                                '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                                TheExec.sites.Item(site).result = tlResultFail
                            Else
                                VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                                VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                                next_bin_flag = True
                            End If
                        Wend
                    End If
                    
                    If next_bin_flag = True Then
                        VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1
                        VBIN_RESULT(p_mode).step_in_IDS_Zone = idx_step
                    End If
                End If '''If CDec(VBIN_RESULT(p_mode).GRADE) > 0
            Next site
        End If '''If BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).ExcludedPmode = True
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_voltageInheritance_for_powerDomain"
'    TheExec.ErrorLogMessage "Error encountered in VBT Function of check_voltageInheritance_for_powerDomain"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210909: Modified to merge the branches of the vbt function run_patt_offline_simulation.
'20201210: Modified to remove the redundant branch "If siteResult_Offline(site) = False Then".
'20201118: Modified to remove the redundant argument "offline_flag_patallpass As Boolean".
'20200922: Created to run Pattern offline simulation for GradeSearch_HVCC_VT / GradeSearch_postBinCut_VT / run_patt_only_VT.
'20200730: Modified to add the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'20200106: As per discussion with SWLINZA, he suggested us to add this to check any alarm.
Public Function run_patt_offline_simulation(patt_selected As String, result_mode As tlResultMode, siteResult_Offline As SiteBoolean)
    Dim site As Variant
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''It can use the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'''//==================================================================================================================================================================================//'''
    If Flag_VDD_Binning_Offline = True Then '''Offline test.
        '''//Generate offline simulation with random pattern Pass/Fail.
        '''Note: It can use the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
        If EnableWord_VDDBinning_Offline_AllPattPass = True Or EnableWord_Golden_Default = True Then
            siteResult_Offline = True
        Else
            For Each site In TheExec.sites
                siteResult_Offline(site) = IIf(Round(WorksheetFunction.Min(1, 1), 0) = 1, True, False)
                'siteResult_Offline(site) = IIf(Round(WorksheetFunction.Min(1, Rnd * 8), 0) = 1, True, False)
            Next site
        End If
        
        '''//Run the pattern by offline simulation random pattern Pass/Fail.
        Call TheHdw.Patterns(patt_selected).Test(pfNever, 0, result_mode)
                    
        '''20210909: Modified to merge the branches of the vbt function run_patt_offline_simulation.
        For Each site In TheExec.sites
            If siteResult_Offline(site) = False Then
                Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.Item(site).TestNumber, logTestFail)
            Else
                Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.Item(site).TestNumber, logTestPass)
            End If
        Next
    Else
        TheExec.Datalog.WriteComment "Flag_VDD_Binning_Offline is " & CStr(Flag_VDD_Binning_Offline) & ". It is incorrect to use run_patt_offline_simulation. Error!!!"
        TheExec.ErrorLogMessage "Flag_VDD_Binning_Offline is " & CStr(Flag_VDD_Binning_Offline) & ". It is incorrect to use run_patt_offline_simulation. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of run_FuncPat_and_check_PassFail"
'    TheExec.ErrorLogMessage "Error encountered in VBT Function of run_FuncPat_and_check_PassFail"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201125: Modified to replace the argument "pattPass as SiteBoolean" with "siteResult As SiteBoolean".
'20201118: Modified to remove the redundant arguments "Optional is_VddBinning_offline As Boolean = False" and "Optional offline_pat_status As SiteBoolean".
'20201118: Created to check alarmFail for pattern.
Public Function check_alarmFail_for_pattern(siteResult As SiteBoolean)
    Dim site As Variant
On Error GoTo errHandler
    For Each site In TheExec.sites
        '''//Check if alarmFail(site) is triggered or not.
        If alarmFail(site) = True Then
            TheExec.Datalog.WriteComment "site:" & site & ", alarmFail!!!"
            siteResult(site) = False
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_patt_Pass_Fail"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201125: Modified to replace the argument "PrePattPass as SiteBoolean" with "siteResult As SiteBoolean".
'20201015: Modified to rename "check_PrePattPass_for_PattPass" as "update_Pattern_result_to_PattPass".
'20201012: Created to update PrePattPass to PattPass.
Public Function update_Pattern_result_to_PattPass(siteResult As SiteBoolean, pattPass As SiteBoolean)
    Dim site As Variant
On Error GoTo errHandler
    For Each site In TheExec.sites
        If pattPass(site) = False Or siteResult(site) = False Then
            pattPass(site) = False
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of update_Pattern_result_to_PattPass"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for update_control_flag_for_patt_loop.
'20200923: Created to update the status of "AllSiteFailPatt" and "All_Patt_Pass".
'20200923: Modified to merge the vbt blocks of "AllSiteFailPatt" and "All_Patt_Pass".
'20200923: Modified to remove the unused condition from the branch of "AllSiteFailPatt" and "All_Patt_Pass".
'20200922: Modified to update the status of "AllSiteFailPatt".
Public Function update_control_flag_for_patt_loop(inst_info As Instance_Info, pattPass As SiteBoolean)
    Dim site As Variant
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Loop each pattern in the pattern set, and use flag for all site to record if the site had failed.
'''//==================================================================================================================================================================================//'''
'''============================================
''' EQ3 Site0  Site1  Site2
''' loop 5 patterns Start
''' Patt1    F  P  P
''' Patt2    P  F  F
''' Patt3      skip and go to next level
''' Patt4    do not need to test
''' Patt5    do not need to test
'''loop 5 patterns End
'''============================================
    '''//Update the status of "AllSiteFailPatt" and "All_Patt_Pass".
    For Each site In TheExec.sites
        If pattPass(site) = False Then
            '''======================================================================================
            ''' Site 0 pass, Site 1 fail, site 2 fail =>  AllSiteFailPatt = 2^1 or 2^2 = 6.
            ''' Site 0 fail, Site 1 fail, site 2 fail =>  AllSiteFailPatt = 2^0 or 2^1 or 2^2 = 7.
            '''======================================================================================
            inst_info.AllSiteFailPatt = inst_info.AllSiteFailPatt Or 2 ^ site
            
            '=======================================================================================================
            ' If this site had found the grade, but the pattern is failed this step,
            ' we will clear all result and define the site is not found the grade yet, the situation is shmoo hole.
            '  Site0   Site1   Site2
            '  EQ4(F)  EQ4(F)  EQ4(P)
            '  EQ3(F)  EQ3(P)  EQ3(F) => the EX. for site2 EQ3(F)
            '  EQ2(P)  EQ2(P)  EQ2(P)
            '
            '  Grade   Grade   Grade
            '  EQ2     EQ3     EQ2
            '=======================================================================================================
            inst_info.All_Patt_Pass(site) = False
            
            '''//Check if Grade_Found but pattern fails. It means DUT has the shmoo hole.
            '''20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
            If inst_info.grade_found(site) = True And inst_info.gradeAlg = GradeSearchAlgorithm.linear Then
                inst_info.grade_found(site) = False '''Shmoo hole: lower VCC passes, but higher VCC fails.
                VBIN_RESULT(inst_info.p_mode).GRADE = 0
                VBIN_RESULT(inst_info.p_mode).GRADEVDD = 0
            End If
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of update_control_flag_for_patt_loop"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
'20210526: Modified to add "Flag_Get_column_Monotonicity_Offset" for Monotonicity_Offset check because C651 Si revised the check rules.
'20210429: Modified to remove the globalVariable "Flag_Using_Montonicity_Offset".
'20210427: Modified to parse the column of "Monotonicity_Offset".
'20210325: Modified to use the 1-dimension array to store SRAM_Vth.
'20210322: Discussed the vbt code that checked if int_Offset is multiple of StepVoltage, all project BinCut owners decided to remove the vbt code because this was the redundant action.
'20210312: Modified to parse columns of "Softbin" and "HardBin" from column col_sort+1.
'20210305: Modified to set INTP_MODE_L, INTP_MODE_L, and AllowEqual if the cells are empty.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20200824: Modified to check TotalStepPerMode. Revised by Leon Weng.
'20200703: Modiifed to use "check_Sheet_Range".
'20200501: Modified to use "AllBinCut(p_mode).INTP_SKIPTEST".
'20200427: Modified to move "Flag_Interpolation_enable" from "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2" to "initVddBinning".
'20200421: Modified to check the column of "Allow Equal".
'20200415: Modified to check "col_soft_bin".
'20200331: Modified to check if int_Offset is multiple of 3.125.
'20200330: Modified to parse "Int_Offset","Int_SkipTest" for interpolation.
'20200206: Modified to check if CPVmin, CPVmax, and CPGB are multiple of 3.125.
'20191219: Modified to check powerDomain in the vbt function "initDomain2Pin".
'20191127: Modified for the revised InitVddBinTable.
'20191113: Modified to check pmode, allowEqual, and MaxPV/MinPV.
'20191023: Modified to check if "MaxPV(pmode0/pmode1)" is in the column "Comment" or not.
'20191014: Modified to parse the table with different powerPins when different testjobs.
'20191001: Modified for the new header defined by C651, they separated IDS_limit into "CPIDSMax" and "IDSMax_HOT"
'20190706: Modified to check if the powerPin (Domain) is in the pin_group "FullCorePowerinFlowSheet " or not.
'20190426: Modified to use the function "Find_Sheet".
'20190321: Modified for checking if powerpin exists in pinmap and channelmap.
'20190312: Modified for adding powerpins into power_group "FullCorePowerinFlowSheet".
'20180221: Anderson enhanced this for VddBinDef & OtherRail combine together.
Public Function initVddBinTableOneMod(passBinCut As Long, col_ids As Integer, col_sort As Integer)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim site As Variant
    Dim strAry_Temp() As String
    Dim row As Long, col As Long
    Dim main_p_mode As Integer
    Dim idx_step As Long
    Dim test_type As Long
    Dim col_binned As Integer
    Dim col_domain As Integer
    Dim col_mode As Integer
    Dim col_eqn As Integer
    Dim col_id As Integer
    Dim col_c As Integer
    Dim col_m As Integer
    Dim col_cpids As Integer
    Dim col_ftids As Integer
    Dim col_cp_vmax As Integer
    Dim col_cp_vmin As Integer
    Dim col_montonicityoffset As Integer '''Monotonicity_Offset
    Dim col_cpgb As Integer
    Dim col_cp2gb As Integer
    Dim col_ft1gb As Integer
    Dim col_ft2gb As Integer
    Dim col_sltgb As Integer
    Dim col_htol_ro_gb As Integer
    Dim col_htol_ro_gb_room As Integer
    Dim col_htol_ro_gb_hot As Integer
    Dim col_ate_ftqa_gb As Integer
    Dim col_slt_ftqa_gb As Integer
    Dim col_cphv As Integer
    Dim col_fthv As Integer
    Dim col_qahv As Integer
    Dim col_intModeL As Integer '''for start p_mode of interpolation
    Dim col_intModeH As Integer '''for end p_mode of interpolation
    Dim col_intMFactor As Integer '''for factor of interpolation
    Dim col_intOffset As Integer
    Dim col_intSkipTest As Integer
    Dim col_allow_equal As Integer
    Dim col_comment As Integer
    Dim col_sram_vt_spec(1) As Integer '''SRAM_VTH_SPEC(0): for CP1 BV binSearch and postBinCut/OutsideBinCut, SRAM_VTH_SPEC(1): for CP1 HBV and non-CP1 BV/HBV.
    Dim col_dynamic_offset(MaxJobCountInVbt, MaxTestType) As Integer
    Dim strTemp As String
    Dim split_content() As String
    Dim p_mode As Integer
    Dim jobIdx As Integer, testTypeIdx As Integer
    Dim powerDomain As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim row_of_title As Integer
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "Vdd_Binning_Def_appA_" & passBinCut
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        Flag_Adjust_Max_Enable = False
        Flag_Adjust_Min_Enable = False
        Adjust_Power_Max_pmode = ""
        Adjust_Power_Min_pmode = ""
                
        If Total_Bincut_Num < passBinCut Then                   'capture the max BinCut number
            Total_Bincut_Num = passBinCut
        End If
        
        For p_mode = 0 To MaxPerformanceModeCount - 1           'initilize the MODE_STEP and ExcludedPmode
            BinCut(p_mode, passBinCut).Mode_Step = -99
            BinCut(p_mode, passBinCut).ExcludedPmode = True     'If you do not assign to True , the default value is False
            ExcludedPmode(p_mode) = True
        Next p_mode
        
        For row = 1 To MaxRow
            For col = 1 To maxcol
                '''******************************************************************************************************************'''
                '''//If CorePower and OtherRail are in the same table (only Vdd_Binning_Def), 1st column is "Binned".
                '''//If CorePower and OtherRail are in the different tables (Vdd_Binning_Def and Other_Rail), 1st column is "Domain".
                '''******************************************************************************************************************'''
                '''If 1st column 1 of the header is "Binned", split the line and find out the keyword column.
                If LCase(ws_def.Cells(row, col).Value) Like "binned" Then
                    col_binned = col
                    row_of_title = row
                End If
            
                If row_of_title > 0 Then
                    If LCase(ws_def.Cells(row_of_title, col).Value) = "domain" Then
                        col_domain = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "mode" Then
                        col_mode = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "id" Then
                        col_id = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "eqn" Then
                        col_eqn = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "c" Then
                        col_c = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "m" Then
                        col_m = col
                    '''********************************************************'''
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpidsmax" Then
                        col_cpids = col
                    
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "idsmax_hot" Or LCase(ws_def.Cells(row_of_title, col).Value) = "ftids" Then
                        col_ftids = col
                    '''********************************************************'''
                    '''20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpvmax" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("BinningVmax") Then
                        col_cp_vmax = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpvmin" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("BinningVmin") Then
                        col_cp_vmin = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpgb" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("BinningGB") Then
                        col_cpgb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cp2gb" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("CP_GB_HOT") Then
                        col_cp2gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "ft1gb" Or LCase(ws_def.Cells(row, col).Value) = "ft_gb_room" Then
                        col_ft1gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "ft2gb" Or LCase(ws_def.Cells(row, col).Value) = "ft_gb_hot" Then
                        col_ft2gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "sltgb" Then
                        col_sltgb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "htol_ro_gb" Then
                        col_htol_ro_gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "htol_ro_gb_room" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("HTOL_T0TX_GB_ROOM") Then
                        col_htol_ro_gb_room = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "htol_ro_gb_hot" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("HTOL_T0TX_GB_HOT") Then
                        col_htol_ro_gb_hot = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "ate_fqagb" Then
                        col_ate_ftqa_gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "slt_fqa_gb" Then
                        col_slt_ftqa_gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cphv" Then
                        col_cphv = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "fthv" Then
                        col_fthv = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "qahv" Then
                        col_qahv = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) Like "offset_*_*" Then
                        jobIdx = getBinCutJobDefinition(LCase(ws_def.Cells(row_of_title, col).Value))
                        testTypeIdx = decide_offset_testType(LCase(ws_def.Cells(row_of_title, col).Value))
                        col_dynamic_offset(jobIdx, testTypeIdx) = col
                    '''Allow_Equal
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "allow equal" Then
                        col_allow_equal = col
                    '''interpolation
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "int_mode_l" Then
                        col_intModeL = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "int_mode_h" Then
                        col_intModeH = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "int_mf" Then
                        col_intMFactor = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "int_offset" Then
                        col_intOffset = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "int_skiptest" Then
                        col_intSkipTest = col
                    '''Monotonicity
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "monotonicity_offset" Then '''20210427: Modified to parse the column of "Monotonicity_Offset".
                        col_montonicityoffset = col
                        '''20210526: Modified to add "Flag_Get_column_Monotonicity_Offset" for Monotonicity_Offset check because C651 Si revised the check rules.
                        Flag_Get_column_Monotonicity_Offset = True
                    '''SRAM_Vth
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) Like "sramthresh*" Then
                        '''************************************************************************************************************************'''
                        '''column "SRAMthresh_CP1" or "SRAMthresh_BinSearch"==> SRAM_VTH_SPEC(0): for BinCut search.
                        '''column "SRAMthresh_Product"                      ==> SRAM_VTH_SPEC(1): for BinCut check.
                        '''************************************************************************************************************************'''
                        '''20210325: Modified to use the 1-dimension array to store SRAM_Vth.
                        If LCase(ws_def.Cells(row_of_title, col).Value) Like "*cp1" Or LCase(ws_def.Cells(row_of_title, col).Value) Like "*binsearch" Then
                            col_sram_vt_spec(0) = col
                        ElseIf LCase(ws_def.Cells(row_of_title, col).Value) Like "*product" Then
                            col_sram_vt_spec(1) = col
                        Else
                            enableRowParsing = False
                            TheExec.Datalog.WriteComment ws_def.Cells(row_of_title, col).Value & " is the undefined column in the sheet:" & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage ws_def.Cells(row_of_title, col).Value & " is the undefined column in the sheet:" & sheetName & ". Error!!!"
                            Exit For
                        End If
                    ElseIf LCase(ws_def.Cells(row, col).Value) = "comment" Then
                        col_comment = col
                        
                        '''//Check if column of "Softbin" is next to column "Comment".
                        If col_comment <> col_sort - 1 Then
                            col_comment = 0
                            TheExec.Datalog.WriteComment "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!"
                        End If
                    End If
                End If
                
                '''//Check if all columns of the header exist...
                '''Note: col_comment should be checked as the last available column of the table.
                If col_domain > 0 And col_mode > 0 And col_id > 0 And col_eqn > 0 And col_c > 0 And col_m > 0 _
                And col_cpids > 0 And col_ftids > 0 And col_cp_vmax > 0 And col_cp_vmin > 0 _
                And col_cpgb > 0 And col_cp2gb And col_ft1gb > 0 And col_ft2gb > 0 _
                And col_sltgb > 0 And col_ate_ftqa_gb > 0 And col_slt_ftqa_gb > 0 _
                And col_cphv > 0 And col_fthv > 0 And col_qahv > 0 _
                And col_intModeL > 0 And col_intModeH > 0 _
                And col_allow_equal > 0 And col_comment > 0 Then
                    enableRowParsing = True
                End If
            Next col
            
            '''//If all columns of the header are found, skip the loop and start parsing each row.
            If enableRowParsing = True Then
                Exit For
            End If
            
            If row = MaxRow And (col_domain = 0 Or col_mode = 0 Or col_cpids = 0 Or col_ftids = 0 Or col_allow_equal = 0 Or col_comment = 0) Then
                enableRowParsing = False
                TheExec.Datalog.WriteComment "Columns of header in " & sheetName & " are incorrect. Error!!!"
                TheExec.ErrorLogMessage "Columns of header in " & sheetName & " are incorrect. Error!!!"
            End If
        Next row
        
        If enableRowParsing = True And row_of_title + 1 <= MaxRow Then
            For row = row_of_title + 1 To MaxRow
                '''//If first word in the mode column is M(ex: MC601).
                '''//If column "Binned" is "true", it means that performance_mode of powerDomain is the binning mode of CorePower.
                If ws_def.Cells(row, col_mode).Value Like "M*" And LCase(ws_def.Cells(row, col_binned).Value) = "true" Then
                    '''//performance_mode is enumerated into the p_mode dictionary when "initVddBinTable".
                    main_p_mode = VddBinStr2Enum(ws_def.Cells(row, col_mode))
                    
                    If UCase(ws_def.Cells(row, col_domain).Value) Like "VDD*" Then
                        powerDomain = UCase(Trim(ws_def.Cells(row, col_domain)))
                        AllBinCut(main_p_mode).powerPin = powerDomain
                    ElseIf UCase(ws_def.Cells(row, col_domain).Value) <> "" Then
                        powerDomain = "VDD_" & UCase(Trim(ws_def.Cells(row, col_domain)))
                        AllBinCut(main_p_mode).powerPin = powerDomain
                    Else
                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_domain) & " doesn't have the correct Domain cell in sheet " & sheetName & ". Error!!!"
                        TheExec.ErrorLogMessage ws_def.Cells(row, col_domain) & " doesn't have the correct Domain cell in sheet " & sheetName & ". Error!!!"
                    End If
                    
                    '''//Use "dict_IsCorePower" to check if powerDomain is BinCut CorePower/OtherRail listed in Vdd_Binning_Def_appA_1.
                    If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                        If LCase(ws_def.Cells(row, col_binned).Value) = LCase(CStr(dict_IsCorePower.Item(UCase(powerDomain)))) Then
                            '''Do nothing...
                        Else
                            TheExec.Datalog.WriteComment "column Binned of " & ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!"
                            TheExec.ErrorLogMessage "column Binned of " & ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!"
                        End If
                    Else
                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"
                        TheExec.ErrorLogMessage ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"
                    End If
                    
                    '''//Check if the p_mode is excluded Pmode...
                    '''if the bincut total passbin number is more than 1, and if the performance mode doesn't exist in BinCut 1 but exists in Bin2, it has the error.
                    If passBinCut > 1 And BinCut(main_p_mode, passBinCut - 1).ExcludedPmode = True Then
                        BinCut(main_p_mode, passBinCut).ExcludedPmode = False
                        ExcludedPmode(main_p_mode) = False
                        TheExec.Datalog.WriteComment "Test performance Mode " & VddBinName(main_p_mode) & " do not exist in BinCut " & passBinCut - 1
                        TheExec.ErrorLogMessage "Test Performance Mode " & VddBinName(main_p_mode) & " do not exist in BinCut " & passBinCut - 1
                    Else
                        BinCut(main_p_mode, passBinCut).ExcludedPmode = False
                        ExcludedPmode(main_p_mode) = False
                    End If
                    
                    If ws_def.Cells(row, col_eqn).Value Like "E#*" Then '''read the E1 ~ En
                        strAry_Temp = Split(ws_def.Cells(row, col_eqn), "E") '''ex: array(0)=E ; array(1)=1
                        idx_step = CLng(strAry_Temp(1)) - 1 '''step: the address for store the EQ number, ex: BinCut(P_mode,passbinnum).EQ_Num(0)=1, step = 0, EQ = 1
                        BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) = CLng(strAry_Temp(1))

                        '''20200824: Modified to check TotalStepPerMode. Revised by Leon Weng.
                        If BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) > TotalStepPerMode Then
                            TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQ num:" & BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) & " is greater than global variable (TotalStepPerMode:" & CStr(TotalStepPerMode) & "), please check it. Error!!!"
                            TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQ num:" & BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) & " is greater than global variable (TotalStepPerMode:" & CStr(TotalStepPerMode) & "), please check it. Error!!!"
                        End If
                        
                        BinCut(main_p_mode, passBinCut).MAX_ID = CDbl(ws_def.Cells(row, col_id).Value)
                        BinCut(main_p_mode, passBinCut).c(idx_step) = CDbl(ws_def.Cells(row, col_c).Value)
                        BinCut(main_p_mode, passBinCut).M(idx_step) = CDbl(ws_def.Cells(row, col_m).Value)
                        '''*************************************************************************************'''
                        '''//Check if CPVmin, CPVmax, and CPGB are multiple of Step Size voltage.
                        BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) = CDbl(ws_def.Cells(row, col_cp_vmax).Value)
                        If BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) <> (Floor(BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                            TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQN:" & ws_def.Cells(row, col_eqn).Value & ", CPVmax:" & BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) & " should be multiple of 3.125. Error!!!"
                            TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQN:" & ws_def.Cells(row, col_eqn).Value & ", CPVmax:" & BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) & " should be multiple of 3.125. Error!!!"
                        End If
                        
                        BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) = CDbl(ws_def.Cells(row, col_cp_vmin).Value)
                        If BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) <> (Floor(BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                            TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQN:" & ws_def.Cells(row, col_eqn).Value & ", CPVmin:" & BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) & " should be multiple of 3.125. Error!!!"
                            TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQN:" & ws_def.Cells(row, col_eqn).Value & ", CPVmin:" & BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) & " should be multiple of 3.125. Error!!!"
                        End If

                        '''*************************************************************************************'''
                        '''//Parse the column of "Monotonicity_Offset".
                        If col_montonicityoffset <> 0 Then
                            BinCut(main_p_mode, passBinCut).Monotonicity_Offset(idx_step) = CDbl(ws_def.Cells(row, col_montonicityoffset).Value)
                        End If
                        '''*************************************************************************************'''
                        
                        BinCut(main_p_mode, passBinCut).CP_GB(idx_step) = CDbl(ws_def.Cells(row, col_cpgb).Value)

                        If BinCut(main_p_mode, passBinCut).CP_GB(idx_step) <> (Floor(BinCut(main_p_mode, passBinCut).CP_GB(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                            TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQN:" & ws_def.Cells(row, col_eqn).Value & ", CPGB:" & BinCut(main_p_mode, passBinCut).CP_GB(idx_step) & " should be multiple of 3.125. Error!!!"
                            TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).Value & ", EQN:" & ws_def.Cells(row, col_eqn).Value & ", CPGB:" & BinCut(main_p_mode, passBinCut).CP_GB(idx_step) & " should be multiple of 3.125. Error!!!"
                        End If
                        '''*************************************************************************************'''
                        BinCut(main_p_mode, passBinCut).CP2_GB(idx_step) = CDbl(ws_def.Cells(row, col_cp2gb).Value)
                        BinCut(main_p_mode, passBinCut).FT1_GB(idx_step) = CDbl(ws_def.Cells(row, col_ft1gb).Value)
                        BinCut(main_p_mode, passBinCut).FT2_GB(idx_step) = CDbl(ws_def.Cells(row, col_ft2gb).Value)
                        BinCut(main_p_mode, passBinCut).SLT_GB(idx_step) = CDbl(ws_def.Cells(row, col_slt_ftqa_gb).Value)
                        BinCut(main_p_mode, passBinCut).FTQA_GB(idx_step) = CDbl(ws_def.Cells(row, col_ate_ftqa_gb).Value)
                        BinCut(main_p_mode, passBinCut).SLT_FTQA_GB(idx_step) = CDbl(ws_def.Cells(row, col_slt_ftqa_gb).Value)
                        BinCut(main_p_mode, passBinCut).HVCC_CP(idx_step) = CDbl(ws_def.Cells(row, col_cphv).Value)
                        BinCut(main_p_mode, passBinCut).HVCC_FT(idx_step) = CDbl(ws_def.Cells(row, col_fthv).Value)
                        BinCut(main_p_mode, passBinCut).HVCC_QA(idx_step) = CDbl(ws_def.Cells(row, col_qahv).Value)
                        
                        If col_htol_ro_gb > 0 Then
                            BinCut(main_p_mode, passBinCut).HTOL_RO_GB(idx_step) = CDbl(ws_def.Cells(row, col_htol_ro_gb).Value)
                        End If
                        
                        If col_htol_ro_gb_room > 0 Then
                            If TheExec.Flow.EnableWord("HTOL_TX_ROOM") = True Then
                                BinCut(main_p_mode, passBinCut).FT1_GB(idx_step) = CDbl(ws_def.Cells(row, col_htol_ro_gb_room).Value)
                            End If
                            BinCut(main_p_mode, passBinCut).HTOL_RO_GB_ROOM(idx_step) = CDbl(ws_def.Cells(row, col_htol_ro_gb_room).Value)
                        End If
                        
                        If col_htol_ro_gb_hot > 0 Then
                            If TheExec.Flow.EnableWord("HTOL_TX_HOT") = True Then
                                BinCut(main_p_mode, passBinCut).FT2_GB(idx_step) = CDbl(ws_def.Cells(row, col_htol_ro_gb_hot).Value)
                            End If
                            BinCut(main_p_mode, passBinCut).HTOL_RO_GB_HOT(idx_step) = CDbl(ws_def.Cells(row, col_htol_ro_gb_hot).Value)
                        End If
                        '''*************************************************************************************'''
                        If (ws_def.Cells(row, col_allow_equal).Value <> "") Then '''20161228, liki
                            BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step) = VddBinStr2Enum(ws_def.Cells(row, col_allow_equal).Value)
                        Else
                            BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step) = 0
                        End If
                        
                        '''//Check the column of "Allow Equal".
                        If passBinCut = 1 And idx_step = 0 Then
                            AllBinCut(main_p_mode).Allow_Equal = BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step)
                        Else
                            If AllBinCut(main_p_mode).Allow_Equal <> BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step) Then
                                TheExec.Datalog.WriteComment "The Allow Equal from Eqn" & idx_step + 1 & " of " & VddBinName(main_p_mode) & " in sheet " & sheetName & " doesn't match other step or PassBinCut. Error!!!"
                                TheExec.ErrorLogMessage "The Allow Equal from Eqn" & idx_step + 1 & " of " & VddBinName(main_p_mode) & " in sheet " & sheetName & " doesn't match other step or PassBinCut. Error!!!"
                            End If
                        End If
                        
                        If ws_def.Cells(row, col_comment).Value <> "" Then
                            strTemp = LCase(ws_def.Cells(row, col_comment).Value)
                            
                            '''//Check if "Max PV (pmode0/pmode1)" is in the column "Comment" or not. Check check pmode, allowEqual, and MaxPV/MinPV.
                            If LCase(strTemp) Like "max*pv*(*)" Or LCase(strTemp) Like "min*pv*(*)" Then '''//ex: Max PV (MP008/MP009/MP00A/MP105)
                                If ws_def.Cells(row, col_mode).Value <> "" Then
                                    If LCase(strTemp) Like LCase("*" & ws_def.Cells(row, col_mode).Value & "*") _
                                    And LCase(strTemp) Like LCase("*" & ws_def.Cells(row, col_allow_equal).Value & "*") Then
                                        strTemp = UCase(Replace(strTemp, "/", ","))
                                        split_content = Split(strTemp, "(")
                                        split_content = Split(split_content(UBound(split_content)), ")")
                                        
                                        If LCase(strTemp) Like "max*pv*" Then
                                            Flag_Adjust_Max_Enable = Flag_Adjust_Max_Enable Or True
                                            
                                            If Adjust_Power_Max_pmode <> "" Then
                                                If UCase("*+" & Adjust_Power_Max_pmode & "+*") Like UCase("*+" & split_content(0) & "+*") Then
                                                    '''Do nothing...
                                                Else
                                                    Adjust_Power_Max_pmode = Adjust_Power_Max_pmode & "+" & split_content(0)
                                                End If
                                            Else
                                                Adjust_Power_Max_pmode = split_content(0)
                                            End If
                                        ElseIf LCase(strTemp) Like "min*pv*" Then
                                            Flag_Adjust_Min_Enable = Flag_Adjust_Min_Enable Or True
                                            
                                            If Adjust_Power_Min_pmode <> "" Then
                                                If UCase("*+" & Adjust_Power_Min_pmode & "+*") Like UCase("*+" & split_content(0) & "+*") Then
                                                    '''Do nothing...
                                                Else
                                                    Adjust_Power_Min_pmode = Adjust_Power_Min_pmode & "+" & split_content(0)
                                                End If
                                            Else
                                                Adjust_Power_Min_pmode = split_content(0)
                                            End If
                                        End If
                                    Else
                                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_mode).Value & " doesn't have the correct AllowEqual cell with MaxPV or MinPV in sheet " & sheetName & ". Error!!!"
                                        TheExec.ErrorLogMessage ws_def.Cells(row, col_mode).Value & " doesn't have the correct AllowEqual cell with MaxPV or MinPV in sheet " & sheetName & ". Error!!!"
                                    End If
                                Else
                                    TheExec.Datalog.WriteComment ws_def.Cells(row, col_mode).Value & " doesn't have the correct AllowEqual cell in sheet " & sheetName & ". Error!!!"
                                    TheExec.ErrorLogMessage ws_def.Cells(row, col_mode).Value & " doesn't have the correct AllowEqual cell in sheet " & sheetName & ". Error!!!"
                                End If
                            End If
                        End If
                        
                        '''***//Interpolation//***'''
                        '''20210305: Modified to set INTP_MODE_L, INTP_MODE_L, and AllowEqual if the cells are empty.
                        '''//Start performance mode
                        If col_intModeL > 0 Then
                            If (ws_def.Cells(row, col_intModeL).Value <> "") Then '//20180312: modified by Anderson.
                                BinCut(main_p_mode, passBinCut).INTP_MODE_L(idx_step) = VddBinStr2Enum(ws_def.Cells(row, col_intModeL).Value)
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_MODE_L(idx_step) = 0
                            End If
                        End If
                        
                        '''//End performance mode
                        If col_intModeH > 0 Then
                            If (ws_def.Cells(row, col_intModeH).Value <> "") Then '//20180312: modified by Anderson.
                                BinCut(main_p_mode, passBinCut).INTP_MODE_H(idx_step) = VddBinStr2Enum(ws_def.Cells(row, col_intModeH).Value)
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_MODE_H(idx_step) = 0
                            End If
                        End If
                        
                        '''//interpolation factor
                        If col_intMFactor > 0 Then
                            If (ws_def.Cells(row, col_intMFactor).Value <> "") Then '//20180312: modified by Anderson.
                                Flag_Interpolation_enable = True
                                BinCut(main_p_mode, passBinCut).INTP_MFACTOR(idx_step) = CDbl(ws_def.Cells(row, col_intMFactor).Value)
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_MFACTOR(idx_step) = 0
                            End If
                        End If
                        
                        '''//offset of interpolation.
                        If col_intOffset > 0 Then
                            If (ws_def.Cells(row, col_intOffset).Value <> "") Then
                                BinCut(main_p_mode, passBinCut).INTP_OFFSET(idx_step) = CDbl(ws_def.Cells(row, col_intOffset).Value)
                                '''20210322: Discussed the vbt code that checked if int_Offset is multiple of StepVoltage, all project BinCut owners decided to remove the vbt code because this was the redundant action.
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_OFFSET(idx_step) = 0
                            End If
                        End If
                        
                        '''//Check "AllBinCut(p_mode).INTP_SKIPTEST".
                        '''//flag to skip interpolation tests of p_mode.
                        If col_intSkipTest > 0 Then
                            If (ws_def.Cells(row, col_intSkipTest).Value <> "") Then
                                If LCase(ws_def.Cells(row, col_intSkipTest).Value) = "yes" Then
                                    If ws_def.Cells(row, col_intMFactor).Value <> "" Then
                                        BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = True
                                        
                                        If passBinCut = 1 Then
                                            If idx_step = 0 Then
                                                AllBinCut(main_p_mode).INTP_SKIPTEST = True
                                            Else
                                                If BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) <> AllBinCut(main_p_mode).INTP_SKIPTEST Then
                                                    TheExec.Datalog.WriteComment ws_def.Cells(row, col_mode) & " has interpolation skipTest=yes in row " & row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!"
                                                    TheExec.ErrorLogMessage ws_def.Cells(row, col_mode) & " has interpolation skipTest=yes in row " & row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!"
                                                End If
                                            End If
                                        End If
                                    Else
                                        BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = False
                                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_mode) & " has interpolation skipTest=yes in row " & row & ", but interpolation factor doesn't have the correct value in sheet " & sheetName & ". Error!!!"
                                        TheExec.ErrorLogMessage ws_def.Cells(row, col_mode) & " has interpolation skipTest=yes in row " & row & ", but interpolation factor doesn't have the correct value in sheet " & sheetName & ". Error!!!"
                                    End If
                                Else
                                    BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = False
                                End If
                            Else
                                If passBinCut = 1 And idx_step > 0 Then
                                    If AllBinCut(main_p_mode).INTP_SKIPTEST = True Then
                                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_mode) & " has interpolation skipTest=yes in row " & row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!"
                                        TheExec.ErrorLogMessage ws_def.Cells(row, col_mode) & " has interpolation skipTest=yes in row " & row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!"
                                    End If
                                End If
                                BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = False
                            End If
                        End If
                        
                        If idx_step = 0 Then
                            AllBinCut(main_p_mode).IDS_CP_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_FT_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_QA_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_FT2_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_FT2_QA_LIMIT = 0
                            '''****************************************************************************************************************'''
                            '''Use ID to define if the inheritance had been changed or not
                            '''****************************************************************************************************************'''
                            If passBinCut = 1 Then
                                If BinCut(main_p_mode, passBinCut).MAX_ID = 1 Then
                                    If Power_List_All <> "" Then
                                        Power_List_All = Power_List_All & "@" & UCase(ws_def.Cells(row, col_mode))
                                     Else
                                        Power_List_All = UCase(ws_def.Cells(row, col_mode))
                                    End If
                                Else
                                    If Power_List_All <> "" Then
                                        Power_List_All = Power_List_All & "," & UCase(ws_def.Cells(row, col_mode))
                                    Else
                                        Power_List_All = UCase(ws_def.Cells(row, col_mode))
                                    End If
                                End If
                            End If
                        End If
                        
                        BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step) = CDbl(ws_def.Cells(row, col_cpids).Value)
                        BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step) = CDbl(ws_def.Cells(row, col_ftids).Value)
                        If AllBinCut(main_p_mode).IDS_CP_LIMIT < BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step) Then AllBinCut(main_p_mode).IDS_CP_LIMIT = BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step)
                        If AllBinCut(main_p_mode).IDS_FT_LIMIT < BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step) Then AllBinCut(main_p_mode).IDS_FT_LIMIT = BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step)
                        
                        '''20210312: Modified to parse columns of "Softbin" and "HardBin" from column col_sort+1.
                        For test_type = 0 To MaxTestType - 3 '''only use SPI, Mbist, TD first
                            BinCut(main_p_mode, passBinCut).SBIN_BINNING_FAIL(idx_step, test_type) = CLng(ws_def.Cells(row, (col_comment + 1) + 4 * test_type).Value)      'large IDS at certain level
                            BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, test_type) = CLng(ws_def.Cells(row, (col_comment + 1) + 1 + 4 * test_type).Value)       'Can find LVCC
                            BinCut(main_p_mode, passBinCut).HBIN_BINNING_FAIL(idx_step, test_type) = CLng(ws_def.Cells(row, (col_comment + 1) + 2 + 4 * test_type).Value)    'large IDS at certain level
                            BinCut(main_p_mode, passBinCut).HBIN_LVCC_FAIL(idx_step, test_type) = CLng(ws_def.Cells(row, (col_comment + 1) + 3 + 4 * test_type).Value)       'Can find LVCC
                        Next test_type
                        
                        '''************************************************************************************************************************'''
                        '''column "SRAMthresh_CP1" or "SRAMthresh_BinSearch"==> SRAM_VTH_SPEC(0): for BinCut search.
                        '''column "SRAMthresh_Product"                      ==> SRAM_VTH_SPEC(1): for BinCut check.
                        '''************************************************************************************************************************'''
                        '''20210325: Modified to use the 1-dimension array to store SRAM_Vth.
                        If col_sram_vt_spec(0) > 0 Then
                            BinCut(main_p_mode, passBinCut).SRAM_VTH_SPEC(0) = CDbl(ws_def.Cells(row, col_sram_vt_spec(0)).Value)
                        End If
                        
                        If col_sram_vt_spec(1) > 0 Then
                            BinCut(main_p_mode, passBinCut).SRAM_VTH_SPEC(1) = CDbl(ws_def.Cells(row, col_sram_vt_spec(1)).Value)
                        End If
                        
                        For jobIdx = 0 To MaxJobCountInVbt
                            For testTypeIdx = 0 To MaxTestType
                                If (col_dynamic_offset(jobIdx, testTypeIdx) <> 0) Then
                                    BinCut(main_p_mode, passBinCut).DYNAMIC_OFFSET(jobIdx, testTypeIdx) = CDbl(ws_def.Cells(row, col_dynamic_offset(jobIdx, testTypeIdx)).Value)
                                End If
                            Next testTypeIdx
                        Next jobIdx
                        
                        BinCut(main_p_mode, passBinCut).Mode_Step = idx_step
                    End If
                End If
            Next row
        Else
            TheExec.Datalog.WriteComment "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!"
            TheExec.ErrorLogMessage "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!"
        End If
    End If '''If isSheetFound = True
    
    '''set the last Step to the error value
    For p_mode = 0 To MaxPerformanceModeCount - 1
        BinCut(p_mode, passBinCut).EQ_Num(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).c(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).M(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).CP_Vmax(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).CP_Vmin(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).CP_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).FT1_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).CP2_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).FT2_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).SLT_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).FTQA_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).SLT_FTQA_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).HTOL_RO_GB(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).HTOL_RO_GB_ROOM(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).HTOL_RO_GB_HOT(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).IDS_CP_LIMIT(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).IDS_FT_LIMIT(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).IDS_QA_LIMIT(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).IDS_FT2_LIMIT(TotalStepPerMode) = 0
        BinCut(p_mode, passBinCut).IDS_FT2_QA_LIMIT(TotalStepPerMode) = 0
        
        For test_type = 0 To MaxTestType - 1
            BinCut(p_mode, passBinCut).SBIN_BINNING_FAIL(TotalStepPerMode, test_type) = 0
            BinCut(p_mode, passBinCut).SBIN_LVCC_FAIL(TotalStepPerMode, test_type) = 0
            BinCut(p_mode, passBinCut).HBIN_BINNING_FAIL(TotalStepPerMode, test_type) = 0
            BinCut(p_mode, passBinCut).HBIN_LVCC_FAIL(TotalStepPerMode, test_type) = 0
        Next test_type
    Next p_mode
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVddBinTableOneMod"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVddBinTableOneMod"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
'20210720: Modified to revise ids_hi_limit with CPIDSMax because C651 Si asked us to use Efuse processed IDS for search in FT.
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20200703: Modiifed to use "check_Sheet_Range".
'20200415: Modified to use the same parsing method with "initVddotherrailOneMod".
'20191219: Modified to check powerDomain in the vbt function "initDomain2Pin".
'20191127: Modified for the revised InitVddBinTable.
'20191007: Modified to merge the vbt code for SRAM with/without p_mode.
'20191001: Modified for the new header defined by C651, they separated IDS_limit into "CPIDSMax" and "IDSMax_HOT".
'20191001: Modified to check the column number before storing the data from the column.
'20190706: Modified to check if the powerDomain is in the pin_group "ShtOtherRailinFlowSheet" or not.
'20190606: Modified for CPIDS_Spec_OtherRail and FTIDS_Spec_OtherRail.
'20190603: Modified for IDS limit of SRAM with CorePower p_mode.
'20190527: Modified for C651 new string format of OtherRail, ex: "MCS601 CPVmax", "CPVmax".
'20190426: Modified to use the function "Find_Sheet".
'20190319: Modified to parse dynamic_offset of OtherRail (requested by C651).
'20190311: Modified for CP1_GB of OtherRail.
'20190307: Modified the vbt code for compatible with conventional projects (SRAM_*** with CorePower Pmode), ex: MCS601 for VDD_CPU_SRAM.
Public Function initVddotherrailOneMod(passBinCut As Long)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim row As Long, col As Long
    Dim p_mode As Integer
    Dim other_p_mode As Integer
    Dim col_binned As Integer
    Dim col_domain As Integer
    Dim col_mode As Integer
    Dim col_cp_vmax As Integer
    Dim col_cp_vmin As Integer
    Dim col_cpgb As Integer
    Dim col_c As Integer
    Dim col_cp2gb As Integer
    Dim col_ft1gb As Integer
    Dim col_ft2gb As Integer
    Dim col_sltgb As Integer
    Dim col_htol_ro_gb As Integer
    Dim col_htol_ro_gb_room As Integer
    Dim col_htol_ro_gb_hot As Integer
    Dim col_ate_ftqa_gb As Integer
    Dim col_slt_ftqa_gb As Integer
    Dim col_cphv As Integer
    Dim col_fthv As Integer
    Dim col_qahv As Integer
    Dim col_cpids As Integer
    Dim col_ftids As Integer
    Dim col_dynamic_offset(MaxJobCountInVbt, MaxTestType) As Integer '''Added for dynamic_offset of OtherRail, 20190319
    Dim jobIdx As Integer, testTypeIdx As Integer
    Dim powerDomain As String
    Dim i As Long
    Dim row_of_title As Integer
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim start_p_mode As Integer
    Dim stop_p_mode As Integer
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "Vdd_Binning_Def_appA_" & passBinCut
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        start_p_mode = -1
        stop_p_mode = -1
        
        '''//Check the header of the table
        '''Get the columns for the diverse coefficient.
        For row = 1 To MaxRow
            For col = 1 To maxcol
                '''******************************************************************************************************************'''
                '''//If CorePower and OtherRail are in the same table (only Vdd_Binning_Def), 1st column is "Binned".
                '''//If CorePower and OtherRail are in the different tables (Vdd_Binning_Def and Other_Rail), 1st column is "Domain".
                '''******************************************************************************************************************'''
                '''If 1st column 1 of the header is "Binned", split the line and find out the keyword column.
                If LCase(ws_def.Cells(row, col).Value) Like "binned" Then
                    col_binned = col
                    row_of_title = row
                End If
                                    
                If row_of_title > 0 Then
                    If LCase(ws_def.Cells(row_of_title, col).Value) = "domain" Then
                        col_domain = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "mode" Then
                        col_mode = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "c" Then
                        col_c = col
                    '''********************************************************'''
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpidsmax" Then
                        col_cpids = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "idsmax_hot" Or LCase(ws_def.Cells(row_of_title, col).Value) = "ftids" Then
                        col_ftids = col
                    '''********************************************************'''
                    '''20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpvmax" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("BinningVmax") Then
                        col_cp_vmax = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpvmin" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("BinningVmin") Then
                        col_cp_vmin = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cpgb" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("BinningGB") Then
                        col_cpgb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cp2gb" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("CP_GB_HOT") Then
                        col_cp2gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "ft1gb" Or LCase(ws_def.Cells(row, col).Value) = "ft_gb_room" Then
                        col_ft1gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "ft2gb" Or LCase(ws_def.Cells(row, col).Value) = "ft_gb_hot" Then
                        col_ft2gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "sltgb" Then
                        col_sltgb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "htol_ro_gb" Then
                        col_htol_ro_gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "htol_ro_gb_room" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("HTOL_T0TX_GB_ROOM") Then
                        col_htol_ro_gb_room = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "htol_ro_gb_hot" Or LCase(ws_def.Cells(row_of_title, col).Value) = LCase("HTOL_T0TX_GB_HOT") Then
                        col_htol_ro_gb_hot = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "ate_fqagb" Then
                        col_ate_ftqa_gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "slt_fqa_gb" Then
                        col_slt_ftqa_gb = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "cphv" Then
                        col_cphv = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "fthv" Then
                        col_fthv = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) = "qahv" Then
                        col_qahv = col
                    ElseIf LCase(ws_def.Cells(row_of_title, col).Value) Like "offset_*_*" Then
                        jobIdx = getBinCutJobDefinition(LCase(ws_def.Cells(row_of_title, col).Value))
                        testTypeIdx = decide_offset_testType(LCase(ws_def.Cells(row_of_title, col).Value))
                        col_dynamic_offset(jobIdx, testTypeIdx) = col
                    End If
                End If
                
                '''//Check if all columns of the header exist...
                '''Note: col_comment should be checked as the last available column of the table.
                If col_domain > 0 And col_mode > 0 And col_c > 0 And col_cpgb > 0 _
                And col_cp2gb > 0 And col_ft1gb > 0 And col_ft2gb > 0 And col_sltgb > 0 _
                And col_ate_ftqa_gb > 0 And col_slt_ftqa_gb > 0 _
                And col_cphv > 0 And col_fthv > 0 And col_qahv > 0 And col_cp_vmax > 0 _
                And col_cp_vmin > 0 And col_cpids > 0 And col_ftids > 0 Then
                    enableRowParsing = True
                End If
            Next col
            
            '''//If all columns of the header are found, skip the loop and start parsing each row.
            If enableRowParsing = True Then
                Exit For
            End If
            
            If row = MaxRow And (col_binned = 0 Or col_domain = 0 Or col_mode = 0 Or col_cpids = 0) Then
                enableRowParsing = False
                TheExec.Datalog.WriteComment "Columns of header in " & sheetName & " are incorrect. Error!!!"
                TheExec.ErrorLogMessage "Columns of header in " & sheetName & " are incorrect. Error!!!"
            End If
        Next row
        
        If enableRowParsing = True And row_of_title + 1 <= MaxRow Then
            For row = row_of_title + 1 To MaxRow
                '''//If column "Binned" is "false" or "ate", it means that power Domain is for OtherRail, ex: SRAM_*** and fixed and low...
                If (LCase(ws_def.Cells(row, col_binned).Value) = "false" Or LCase(ws_def.Cells(row, col_binned).Value) = "ate") Then
                    '''=====================================================================================
                    '''[Step1] Get OtherRail from Domain column.
                    '''//performance_mode is enumerated into the p_mode dictionary when "initVddBinTable".
                    '''=====================================================================================
                    If UCase(Trim(ws_def.Cells(row, col_domain).Value)) Like "VDD*" Then '''ex: "VDD_CPU_SRAM"
                        powerDomain = UCase(Trim(ws_def.Cells(row, col_domain).Value))
                        other_p_mode = VddBinStr2Enum(powerDomain) '//p_mode
                        AllBinCut(other_p_mode).powerPin = powerDomain
                        
                    ElseIf UCase(ws_def.Cells(row, col_domain).Value) <> "" Then '''ex: "CPU_SRAM"
                        powerDomain = UCase("VDD_" & Trim(ws_def.Cells(row, col_domain).Value))
                        other_p_mode = VddBinStr2Enum(powerDomain) '//p_mode
                        AllBinCut(other_p_mode).powerPin = powerDomain
                    Else
                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_domain) & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"
                        TheExec.ErrorLogMessage ws_def.Cells(row, col_domain) & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"
                    End If
                    
                    '''=====================================================================================
                    '''[Step2] Use "dict_IsCorePower" to check if powerDomain is BinCut CorePower/OtherRail listed in Vdd_Binning_Def_appA_1.
                    '''=====================================================================================
                    If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                        If dict_IsCorePower.Item(UCase(powerDomain)) = False Then
                            '''Do nothing...
                        Else
                            TheExec.Datalog.WriteComment "column Binned of " & ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!"
                            TheExec.ErrorLogMessage "column Binned of " & ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!"
                        End If
                    Else
                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"
                        TheExec.ErrorLogMessage ws_def.Cells(row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"
                    End If
                    
                    '''=====================================================================================
                    '''[Step3] Check if mode column with/without P_mode.
                    '''=====================================================================================
                    If ws_def.Cells(row, col_domain).Value = ws_def.Cells(row, col_mode).Value Then '''//If Mode is same as Domain, ex: "SRAM_CPU"
                        '''//If SRAM without p_mode, col_domain and col_mode are the same.
                        start_p_mode = 0
                        stop_p_mode = MaxPerformanceModeCount - 1
                    ElseIf ws_def.Cells(row, col_mode).Value Like "M*##*" Then '''ex: "MC60A", "MCS60A"
                        '''//If SRAM with p_mode, col_mode has the p_mode related to CorePower, ex: "MC60A".
                        If Len(ws_def.Cells(row, col_mode).Value) = 6 Then
                            start_p_mode = VddBinStr2Enum(Mid(ws_def.Cells(row, col_mode).Value, 1, 2) & Mid(ws_def.Cells(row, col_mode).Value, 4, 3))
                            stop_p_mode = start_p_mode
                        Else
                            start_p_mode = VddBinStr2Enum(ws_def.Cells(row, col_mode).Value)
                            stop_p_mode = start_p_mode
                        End If
                    Else
                        TheExec.Datalog.WriteComment ws_def.Cells(row, col_domain).Value & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"
                        TheExec.ErrorLogMessage ws_def.Cells(row, col_domain).Value & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"
                    End If
                    
                    For p_mode = start_p_mode To stop_p_mode
                        BinCut(p_mode, passBinCut).OTHER_PRODUCT_RAIL(other_p_mode) = CDbl(ws_def.Cells(row, col_c).Value) + CDbl(ws_def.Cells(row, col_cpgb).Value)
                        BinCut(p_mode, passBinCut).OTHER_CP1_RAIL(other_p_mode) = CDbl(ws_def.Cells(row, col_c).Value)
                        BinCut(p_mode, passBinCut).OTHER_CP1_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_cpgb).Value) 'added for CP1_GB, 20190311
                        BinCut(p_mode, passBinCut).OTHER_CP2_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_cp2gb).Value)
                        BinCut(p_mode, passBinCut).OTHER_FT1_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_ft1gb).Value)
                        BinCut(p_mode, passBinCut).OTHER_FT2_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_ft2gb).Value)
                        BinCut(p_mode, passBinCut).OTHER_SLT_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_sltgb).Value)
                        BinCut(p_mode, passBinCut).OTHER_CPIDS(other_p_mode) = CDbl(ws_def.Cells(row, col_cpids).Value)
                        BinCut(p_mode, passBinCut).OTHER_FTIDS(other_p_mode) = CDbl(ws_def.Cells(row, col_ftids).Value)
                        BinCut(p_mode, passBinCut).OTHER_ATE_FQA_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_ate_ftqa_gb).Value)
                        
                        If col_htol_ro_gb > 0 Then
                            BinCut(p_mode, passBinCut).OTHER_HTOL_RO_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_htol_ro_gb).Value)
                        End If
                        
                        If col_htol_ro_gb_room > 0 Then
                            If TheExec.Flow.EnableWord("HTOL_TX_ROOM") = True Then
                                BinCut(p_mode, passBinCut).OTHER_FT1_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_htol_ro_gb_room).Value)
                            End If
                            BinCut(p_mode, passBinCut).OTHER_HTOL_RO_GB_ROOM(other_p_mode) = CDbl(ws_def.Cells(row, col_htol_ro_gb_room).Value)
                        End If
                        
                        If col_htol_ro_gb_hot > 0 Then
                            If TheExec.Flow.EnableWord("HTOL_TX_HOT") = True Then
                                BinCut(p_mode, passBinCut).OTHER_FT2_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_htol_ro_gb_hot).Value)
                            End If
                            BinCut(p_mode, passBinCut).OTHER_HTOL_RO_GB_HOT(other_p_mode) = CDbl(ws_def.Cells(row, col_htol_ro_gb_hot).Value)
                        End If
                        
                        BinCut(p_mode, passBinCut).OTHER_SLT_FQA_GB(other_p_mode) = CDbl(ws_def.Cells(row, col_slt_ftqa_gb).Value)
                        BinCut(p_mode, passBinCut).HVCC_OTHER_CP_RAIL(other_p_mode) = CDbl(ws_def.Cells(row, col_cphv).Value)
                        BinCut(p_mode, passBinCut).HVCC_OTHER_FT_RAIL(other_p_mode) = CDbl(ws_def.Cells(row, col_fthv).Value)
                        BinCut(p_mode, passBinCut).HVCC_OTHER_QA_RAIL(other_p_mode) = CDbl(ws_def.Cells(row, col_qahv).Value)
                        BinCut(p_mode, passBinCut).OTHER_CP_Vmax(other_p_mode) = CDbl(ws_def.Cells(row, col_cp_vmax).Value)
                        BinCut(p_mode, passBinCut).OTHER_CP_Vmin(other_p_mode) = CDbl(ws_def.Cells(row, col_cp_vmin).Value)
                        
                        '''****************************************************************************************'''
                        '''//Parsing dynamic_offset for OtherRail(request by C651).
                        '''//For storing the values, we take the domain column as OtherRail performance mode index.
                        For jobIdx = 0 To MaxJobCountInVbt
                            For testTypeIdx = 0 To MaxTestType - 1
                                If (col_dynamic_offset(jobIdx, testTypeIdx) <> 0) Then
                                    BinCut(other_p_mode, passBinCut).DYNAMIC_OFFSET(jobIdx, testTypeIdx) = CDbl(ws_def.Cells(row, col_dynamic_offset(jobIdx, testTypeIdx)).Value)
                                End If
                            Next testTypeIdx
                        Next jobIdx
                        '''****************************************************************************************'''
                    Next p_mode
                    
                    '''=====================================================================================
                    '''[Step4] According to test jobs and PassBinCut, decide the IDS limit for OtherRail.
                    '''=====================================================================================
                    CPIDS_Spec(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(row, col_cpids).Value)
                    FTIDS_Spec(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(row, col_ftids).Value)
                    
                    '''//Choose IDS hi_limit by BinCut testjobs.
                    '''//IDS calculation uses the scale and the unit in "mA".
                    '''20210720: Modified to revise ids_hi_limit with CPIDSMax because C651 Si asked us to use Efuse processed IDS for search in FT.
                    '''<org>
'                    If bincutJobName = "cp1" Or bincutJobName = "ft_room" Or bincutJobName = "qa" Then '''for testjobs with the normal temperature 25C
'                        ids_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(row, col_cpids).Value)
'                    ElseIf bincutJobName = "cp2" Or bincutJobName = "ft_hot" Then '''for testjobs with the high temperature 85C
'                        ids_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(row, col_ftids).Value)
'                    End If
                    '''<new>
                    gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(row, col_cpids).Value) '''use CPIDSMax as IDS_hi_limit for Efuse processed IDS.
                    
                    If AllBinCut(other_p_mode).IDS_CP_LIMIT = 0 Then
                        AllBinCut(other_p_mode).IDS_CP_LIMIT = CDbl(ws_def.Cells(row, col_cpids).Value)
                    ElseIf AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT < CDbl(ws_def.Cells(row, col_cpids).Value) Then
                        AllBinCut(other_p_mode).IDS_CP_LIMIT = CDbl(ws_def.Cells(row, col_cpids).Value)
                    End If
                    
                    If AllBinCut(other_p_mode).IDS_FT_LIMIT = 0 Then
                        AllBinCut(other_p_mode).IDS_FT_LIMIT = CDbl(ws_def.Cells(row, col_ftids).Value)
                    ElseIf AllBinCut(VddBinStr2Enum(powerDomain)).IDS_FT_LIMIT < CDbl(ws_def.Cells(row, col_ftids).Value) Then
                        AllBinCut(other_p_mode).IDS_FT_LIMIT = CDbl(ws_def.Cells(row, col_ftids).Value)
                    End If
                End If
            Next row
        Else
            TheExec.Datalog.WriteComment "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!"
            TheExec.ErrorLogMessage "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!"
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVddotherrailOneMod"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVddotherrailOneMod"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: As per discussion with Jeff, he suggested us to use the keyword "Binning Domain" for each testJob.
'20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
'20210819: Modified to revise the vbt code for the new format of the header in BinCut flow table, as requested by C651 Toby.
'20210819: Modified to assemble job_keyword by bincutJobName according to the vbt function Mapping_TestJobName_to_BincutJobName.
'20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY) greater than the highest bin number.
'20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
'20210201: Modified to check if testCondtion with performance mode for SRAM_Vth, ex: "640mv (MI003)".
'20210131: Modified to check "UCase(Trim(ws_def.Cells(j, Col).value))".
'20210121: Modified to check the format of testConditions, request by TSMC ZQLIN.
'20201222: Modified to use the dictionary "dict_OutsideBinCut_additionalMode" to check if any duplicate additional mode exists in different Outside BinCut flow tables.
'20201222: Modified to revise the vbt function "initVddBinCondition" for multiple "Non_Binning_Rail_Outside_BinCut" sheets.
'20201215: Modified to check if testCondition contain keyword "*Evaluate Bin*" to decide "is_BinCutJob_for_StepSearch" while isParsingOutsideBinCutFlow = False.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201112: Modified to use the dictionary "dict_IsCorePowerInBinCutFlowSheet".
'20201023: Modified to store all headers into the dictionary "dict_BinCutFlow_Domain2Column" by "col_jobBlock_start" and "col_jobBlock_end"...
'20201023: Modified to initialize the array of BV and HBV testConditions with the empty string "".
'20201022: Modified to reduce the complity of the loop to find row of the header VDD.
'20201022: Modified to modify the parsing method for "FullCorePowerinFlowSheet" and "FullOtherRailinFlowSheet".
'20201021: As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
'20201021: Modified to support multiple columns with "IGNORE COLUMN".
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20201021: Modified to revise the vbt code for parsing "Non_Binning_Rail" and "Non_Binning_Rail_outside_BinCut".
'20201013: Modified to trim string of the testCondition from Non_Binning_Rail.
'20201005: Modified to check "AllBinCut(p_mode).INTP_SKIPTEST = True" to update "AllBinCut(p_mode).used".
'20200827: Modified to check the mapping BinCut testJob.
'20200827: Modified to replace "If..Else" with "Select Case".
'20200711: Modified to check if any testcondition contains "#REF!" or "#NAME?".
'20200703: Modiifed to use "check_Sheet_Range".
'20200611: Modified to check "IGNORE COLUMN".
'20200211: Modified to replace "FlowTestCondName" with "AdditionalModeName".
'20200211: Modified to replace "cntFlowTestCond" with "cntAdditionalMode".
'20191227: Modified to check if allbincut(pmode).used is true.
'20191219: Modified for Domain2Pin and Pin2Domain.
'20191218: Modified to check if trackpower is "N/C" or not.
'20191204: Modified to check if pinGroup_BinCut exists in BV and HBV columns.
'20191129: Modified to check "IGNORE COLUMN".
'20191128: Modified for checking the additional mode (flowTestCondition).
'20191127: Modified for the revised InitVddBinTable.
'20191014: Modified to parse the table with different powerPins when different testjobs.
'20190905: Modified for those projects with bin1/binx in the same block of "Non_Binning_Rail" sheet.
'20190706: Modified to check if the powerPin (Domain) is in the pin_group "ShtCorePowerinFlowSheet" or "ShtOtherRailinFlowSheet".
'20190704: Modified to remove the hard-code "power_list" and "power_seq".
'20190704: As the discussion with SWLINZA, we should check digit1-2 of the p_mode are same as p_modes in the power_list.
'20190521: Modified for the correct column numbers.
'20190426: Modified to use the function "Find_Sheet".
'20190319: Modified for assemblying the BinCut powerpin group "FullBinCutPowerinFlowSheet"
'20190117: Modified for checking if powerpin exists in pinmap and channelmap.
'20181209: Modified for blocking unlisted/untested performance modes.
'20180723: Modified for BinCut testjob mapping.
Public Function initVddBinCondition(sheetName As String)
    Dim sheetName_OutsideBinCut As String
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim row As Long, col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim split_content() As String
    Dim performance_mode As String
    Dim additional_mode As String
    Dim main_p_mode As Integer
    Dim p_mode As Integer
    Dim addi_mode As Integer '''For the additional mode
    Dim i As Long, j As Long, k As Long, L As Long
    Dim s As Long
    Dim other_voltage_start_point As Long
    Dim HVCC_flag As Long
    Dim passBinCut As Long
    '''for testjob mapping
    Dim row_of_title As Long
    Dim row_of_testJob As Long
    Dim job_keyword As String
    Dim row_jobBlock As Long
    '''
    Dim cnt_testJob As Long
    Dim idx_testJob As Long
    Dim col_testJob() As Long
    Dim col_jobBlock_start As Long
    Dim col_jobBlock_end As Long
    '''for Performance Mode
    Dim col_mode As Long
    '''for powerDomain
    Dim got_correct_header As Boolean
    Dim got_CorrectDomain As Boolean
    '''for trackPower
    Dim trackpowerTemp As String
    Dim strAry_trackpower() As String
    '''variables
    Dim corePower As Long
    Dim powerDomain As String
    Dim selected_powerDomain As String
    Dim str_mode_temp As String
    Dim binNumStart As Long
    Dim binNumStop As Long
    Dim isSheetFound As Boolean
    Dim isParsingOutsideBinCutFlow As Boolean
    Dim isIgnoreColumn As Boolean
    Dim testCondition As String
    Dim str_mainColumn_content As String
    Dim strTemp As String
    Dim strSplitted() As String
    Dim bincutNum As Long
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Post_BinCut.
'''2. As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
'''3. Please check keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    'sheetName = "Non_Binning_Rail"
    sheetName_OutsideBinCut = "Non_Binning_Rail_Outside" '''for outsiteBinCut.
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        binNumStart = 1
        binNumStop = Total_Bincut_Num
        HVCC_flag = 0
        other_voltage_start_point = 0
        trackpowerTemp = ""
        row_of_title = 0
        row_of_testJob = 0
        cnt_testJob = 0
        idx_testJob = -1
        got_correct_header = True
        isIgnoreColumn = False
        got_CorrectDomain = False
        str_mainColumn_content = ""
        str_mode_temp = ""
        performance_mode = ""
        additional_mode = ""
        job_keyword = ""
        col_jobBlock_start = 0
        col_jobBlock_end = 0
        col_mode = 0
        
        '''//Check if sheet "Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Post_BinCut"
        If LCase(sheetName) Like LCase("*" & sheetName_OutsideBinCut & "*") Then
            If dict_BinCutFlow_Domain2Column.Count <> 0 Then
                isParsingOutsideBinCutFlow = True '''Parsing outsideBinCutFlow...
            Else
                isParsingOutsideBinCutFlow = False '''Parsing BinCutFlow (sheet "Non_Binning_Rail")
                TheExec.Datalog.WriteComment "Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Outside_BinCut"
                TheExec.ErrorLogMessage "Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Outside_BinCut"
            End If
        Else
            isParsingOutsideBinCutFlow = False '''//Parsing BinCutFlow (sheet "Non_Binning_Rail")
        End If
        '''20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
        
        '''//Get keyword for BinCut testJob mapping.
        '''20210819: Modified to assemble job_keyword by bincutJobName according to the vbt function Mapping_TestJobName_to_BincutJobName.
        job_keyword = LCase("*" & bincutJobName & "*") '''ex: "*cp1*", "*cp2*", "*ft_room*", "*ft_hot*", "*qa*".
    Else
        Exit Function
    End If
    
    '''//Find the keyword of BinCut testjob, and find column of the selected testJob.
    '''20210819: As per discussion with Jeff, he suggested us to use the keyword "Binning Domain" for each testJob.
    If job_keyword <> "" Then
        For row = 1 To MaxRow
            For col = 1 To maxcol
                If LCase(ws_def.Cells(row, col).Value) Like LCase("Binning Domain") Then
                    ReDim Preserve col_testJob(cnt_testJob)
                    col_testJob(cnt_testJob) = col
                    cnt_testJob = cnt_testJob + 1
                    
                    If row_of_testJob = 0 Then
                        row_of_testJob = row
                    End If
                End If
            Next col
            
            If row_of_testJob > 0 Then
                Exit For
            End If
        Next row
    End If
    
    '''//Check if any matched block for IGXL Job.
    If cnt_testJob > 0 Then
        For i = 0 To cnt_testJob - 1
            If LCase(ws_def.Cells(row_of_testJob, col_testJob(i) + 2).Value) Like job_keyword Then
                idx_testJob = i
                other_voltage_start_point = col_testJob(i) + 2
                
                '''//Get start/stop columns for block of the selected testJob.
                col_jobBlock_start = col_testJob(idx_testJob) '''column of Domain
                If idx_testJob = UBound(col_testJob) Then
                    col_jobBlock_end = maxcol
                Else
                    col_jobBlock_end = col_testJob(idx_testJob + 1) - 1
                End If
                
                '''//Get column of "Performance Mode"
                For col = col_jobBlock_start To col_jobBlock_end
                    If LCase(Trim(ws_def.Cells(row_of_testJob, col))) Like LCase("Performance Mode") Then
                        col_mode = col
                        
                        '''//Check if column "Performance mode" is defined in the dictionary dict_BinCutFlow_Domain2Column.
                        If dict_BinCutFlow_Domain2Column.Exists(UCase("Performance mode")) = True Then
                            If col = dict_BinCutFlow_Domain2Column.Item(UCase("Performance mode")) Then
                                '''Do nothing...
                            Else
                                col_mode = 0
                                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", it doesn't have the correct columns of Domain in the header. Error!!!"
                                TheExec.ErrorLogMessage "sheet:" & sheetName & ", it doesn't have the correct columns of Domain in the header. Error!!!"
                            End If
                        Else
                            dict_BinCutFlow_Domain2Column.Add UCase("Performance mode"), col
                        End If
                        
                        Exit For
                    End If
                Next col
                
                Exit For
            End If
        Next i
    End If
    
    '''//Find the column of BinCut 1st powerDomain.
    If idx_testJob > -1 And col_mode > 0 And other_voltage_start_point > 0 Then
        '''Do nothing...
    Else
        other_voltage_start_point = 0
        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", initVddBinCondition doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!"
        TheExec.ErrorLogMessage "sheet:" & sheetName & ", initVddBinCondition doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!"
        Exit Function
    End If
            
    '''====================================================-====================================================================
    '''[Step1] Find row of the Header with column of all BinCut powerDomains and pattern keywords.
    '''====================================================-====================================================================
    '''//If column of the selected testJob is found, start to parse the header VDD.
    For row = row_of_testJob + 1 To MaxRow
        str_mainColumn_content = LCase(ws_def.Cells(row, other_voltage_start_point).Value)
        
        '''//Check if powerDomain is listed in VddbinPinDict (defined by sheet "Vdd_Binning_Def").
        For col = col_jobBlock_start To col_jobBlock_end
            strTemp = UCase(Trim(ws_def.Cells(row, col).Value))
        
            If strTemp <> "" Then
                If col >= other_voltage_start_point And col < other_voltage_start_point + cntVddbinPin Then '''powerDomain
                    '''//Check if the tracking power exists in the column of the header.
                    If strTemp Like "*,*" Then
                        strAry_trackpower = Split(strTemp, ",")
                        powerDomain = UCase(Trim(strAry_trackpower(0)))
                        trackpowerTemp = UCase(Trim(Replace(strTemp, (UCase(strAry_trackpower(0)) & ","), "")))
                    Else
                        powerDomain = strTemp
                        trackpowerTemp = ""
                    End If
                    
                    '''//Check if powerDomain is CorePower or OtherRail shown in BinCut sheet "Vdd_Binning_Def".
                    '''//dict_IsCorePower is dictionary of BinCut CorePower/OtherRail.
                    If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                        '''//Add column of the powerDomain into the dictionary "dict_BinCutFlow_Domain2Column".
                        If dict_BinCutFlow_Domain2Column.Exists(powerDomain) Then
                            If dict_BinCutFlow_Domain2Column.Item(powerDomain) = col Then
                                got_correct_header = got_correct_header And True
                            Else
                                got_correct_header = got_correct_header And False
                                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & row & " of the header VDD. Error!!!"
                                TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & row & " of the header VDD. Error!!!"
                            End If
                        Else
                            got_correct_header = got_correct_header And True
                            dict_BinCutFlow_Domain2Column.Add powerDomain, col
                            dict_BinCutFlow_Column2Domain.Add col, powerDomain
                        End If
                        
                        '''//Parsing BinCutFlow (sheet "Non_Binning_Rail") to add BinCut powerDomain into "FullBinCutPowerinFlowSheet".
                        If isParsingOutsideBinCutFlow = False Then
                            '''//Check if any "IGNORE COLUMN" exists in column of powerDomain..
                            isIgnoreColumn = False
                            For j = 1 To row
                                '''20210131: Modified to check "UCase(Trim(ws_def.Cells(j, Col).value))".
                                If UCase(Trim(ws_def.Cells(j, col).Value)) Like UCase("IGNORE*COLUMN") Then
                                    isIgnoreColumn = True
                                    Exit For
                                End If
                            Next j
                            
                            If isIgnoreColumn = False Then
                                If FullBinCutPowerinFlowSheet <> "" Then
                                    If LCase("*," & FullBinCutPowerinFlowSheet & ",*") Like LCase("*," & powerDomain & ",*") Then
                                        got_correct_header = got_correct_header And False
                                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & row & " of the header VDD. Error!!!"
                                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & row & " of the header VDD. Error!!!"
                                    Else
                                        FullBinCutPowerinFlowSheet = FullBinCutPowerinFlowSheet & "," & powerDomain
                                    End If
                                Else
                                    FullBinCutPowerinFlowSheet = powerDomain
                                End If
                                
                                '''//TrackPower
                                If trackpowerTemp <> "" Then
                                    AllBinCut(VddBinStr2Enum(powerDomain)).TRACKINGPOWER = trackpowerTemp
                                End If
                            End If
                        End If
                    Else '''If dict_IsCorePower.Exists(UCase(powerDomain)) = False
                        got_correct_header = got_correct_header And False
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", the header has the undefined powerDomain:" & strTemp & ", it doesn't show in sheet Vdd_Binning_Def sheet_appA_1. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", the header has the undefined powerDomain:" & strTemp & ", it doesn't show in sheet Vdd_Binning_Def sheet_appA_1. Error!!!"
                    End If '''If dict_IsCorePower.Exists(UCase(powerDomain)) = True
                Else
                    powerDomain = ""
                    trackpowerTemp = ""
                End If
                
                '''//Check columns of non-powerDomain, ex: "All Others".
                If powerDomain = "" Then
                    '''//Add column of the powerDomain into the dictionary "dict_BinCutFlow_Domain2Column".
                    If dict_BinCutFlow_Domain2Column.Exists(strTemp) Then
                        If dict_BinCutFlow_Domain2Column.Item(strTemp) = col Then
                            got_correct_header = got_correct_header And True
                        Else
                            got_correct_header = got_correct_header And False
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & row & " of the header VDD. Error!!!"
                            TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & row & " of the header VDD. Error!!!"
                        End If
                    Else
                        got_correct_header = got_correct_header And True
                        dict_BinCutFlow_Domain2Column.Add strTemp, col
                        dict_BinCutFlow_Column2Domain.Add col, strTemp
                    End If
                End If '''If powerDomain = ""
            End If '''If LCase(ws_def.Cells(row, col).Value) <> ""
        Next col
            
        If got_correct_header = True Then
            row_of_title = row
            Exit For
        End If
    Next row
    
    '''====================================================-====================================================================
    '''[Step2] Parse each row to get testConditions of powerDomain.
    '''====================================================-====================================================================
    If row_of_title > 0 Then '''It means that columns of BinCut powerDomains are found.
        While LCase(ws_def.Cells(row, 1).Value) <> "end"
            '''//Check if any testcondition contains "#REF!" or "#NAME?".
            If IsError(ws_def.Cells(row, other_voltage_start_point).Value) Then
                str_mainColumn_content = ""
                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell (row:" & row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(row, 1).Value & ". The cell contains the incorrect content. Error!!!"
                TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(row, 1).Value & ". The cell contains the incorrect content. Error!!!"
            Else
                str_mainColumn_content = LCase(ws_def.Cells(row, other_voltage_start_point).Value)
                
                '''//HVCC block (for HBV) in "Non_Binning_Rail" sheet
                '''If keyword of "HVCC" exists in the column of other_voltage_start_point or other_voltage_start_point-1, check which bin_number to use...
                If str_mainColumn_content Like "*hvcc*" Or str_mainColumn_content Like "*hbv*" Then '''"Bin1 - HVCC CP1 @ 25'C, mV"
                    HVCC_flag = 1
        
                    If str_mainColumn_content Like LCase("*bin1*binx*") Or LCase(ws_def.Cells(row, other_voltage_start_point - 2).Value) Like LCase("*bin1*binx*") Then
                        binNumStart = 1
                        binNumStop = 2
                    ElseIf str_mainColumn_content Like LCase("*bin1*") Or LCase(ws_def.Cells(row, other_voltage_start_point - 2).Value) Like LCase("*bin1*") Then
                        binNumStart = 1
                        binNumStop = 1
                    ElseIf str_mainColumn_content Like LCase("*binx*") Or LCase(ws_def.Cells(row, other_voltage_start_point - 2).Value) Like LCase("*binx*") Then
                        binNumStart = 2
                        binNumStop = 2
                    ElseIf str_mainColumn_content Like LCase("*biny*") Or LCase(ws_def.Cells(row, other_voltage_start_point - 2).Value) Like LCase("*biny*") Then
                        binNumStart = 3
                        binNumStop = 3
                    ElseIf LCase(ws_def.Cells(row, other_voltage_start_point - 2).Value) Like "" Then
                        binNumStart = 1
                        binNumStop = 3
                    Else
                        TheExec.Datalog.WriteComment "The Content of HVCC in the sheet " & sheetName & " is wrong. Error!!!"
                        TheExec.ErrorLogMessage "The Content of HVCC in the sheet " & sheetName & " is wrong. Error!!!"
                    End If
        
                    If binNumStop < binNumStart Then
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", row:" & row & ", the Content of HVCC bin number doesn't match Total_Bincut_Num of Vdd_Binning_Def sheets. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", row:" & row & ", the Content of HVCC bin number doesn't match Total_Bincut_Num of Vdd_Binning_Def sheets. Error!!!"
                    End If
                            
                '''//Check if columns of HBV BinCut powerDomains match columns of BV BinCut powerDomains.
                ElseIf str_mainColumn_content Like "vdd_*" Then '''ex: "VDD_PCPU".
                    For col = other_voltage_start_point To col_jobBlock_end
                        strTemp = UCase(Trim(ws_def.Cells(row, col).Value))
                        
                        '''//Check if the tracking power exists in the column of the header.
                        If strTemp Like "*,*" Then
                            strAry_trackpower = Split(strTemp, ",")
                            powerDomain = UCase(Trim(strAry_trackpower(0)))
                            trackpowerTemp = UCase(Trim(Replace(strTemp, (UCase(strAry_trackpower(0)) & ","), "")))
                        Else
                            powerDomain = strTemp
                            trackpowerTemp = ""
                        End If
                        
                        If dict_BinCutFlow_Domain2Column.Item(powerDomain) = col Then
                            If dict_IsCorePower.Exists(powerDomain) Then
                                If trackpowerTemp = AllBinCut(VddBinStr2Enum(powerDomain)).TRACKINGPOWER Then
                                    '''Do nothing
                                Else
                                    TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell (row:" & row & ",column:" & col & "), content:" & strTemp & ", trackpower of powerDomain is different from cell (row:" & row_of_title & ", col:" & dict_BinCutFlow_Domain2Column.Item(powerDomain) & ") in the header. Error!!!"
                                    TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & row & ",column:" & col & "), content:" & strTemp & ", trackpower of powerDomain is different from cell (row:" & row_of_title & ", col:" & dict_BinCutFlow_Domain2Column.Item(powerDomain) & ") in the header. Error!!!"
                                End If
                            End If
                        Else
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell (row:" & row & ",column:" & col & "), content:" & strTemp & ", it is different from the header VDD in row" & row_of_title & ". Error!!!"
                            TheExec.ErrorLogMessage strTemp & " in row" & row & " col" & col & " of sheet " & sheetName & " is different from the header VDD in row" & row_of_title & ". Error!!!"
                        End If
                    Next col
                    
                    If HVCC_flag = 1 Then
                        strTemp = UCase(Trim(ws_def.Cells(row, other_voltage_start_point - 1).Value))
                        
                        If dict_BinCutFlow_Domain2Column.Item(strTemp) = other_voltage_start_point - 1 Then
                            '''Do nothing
                        Else
                            TheExec.Datalog.WriteComment strTemp & " in row" & row & " col" & other_voltage_start_point - 1 & " of sheet " & sheetName & " is different from the header VDD in row" & row_of_title & ". Error!!!"
                            TheExec.ErrorLogMessage strTemp & " in row" & row & " col" & other_voltage_start_point - 1 & " of sheet " & sheetName & " is different from the header VDD in row" & row_of_title & ". Error!!!"
                        End If
                    End If
                   
                '''//Check if any performance_mode exists in column "Performance Mode".
                ElseIf LCase(ws_def.Cells(row, dict_BinCutFlow_Domain2Column.Item(UCase("Performance Mode")))) Like "m*" Then '''ex: "MS001", "MS001_GPU".
                    '''//Get performance_mode
                    str_mode_temp = UCase(ws_def.Cells(row, dict_BinCutFlow_Domain2Column.Item(UCase("Performance Mode"))).Value)
                    split_content = Split(str_mode_temp, "_")
                    performance_mode = UCase(split_content(0))
                    
                    '''//If with all empty conditions, it means performance_mode without any additional_mode.
                    If UBound(split_content) > 0 Then
                        additional_mode = UCase(Replace(UCase(str_mode_temp), (performance_mode & "_"), ""))
                    Else
                        additional_mode = ""
                    End If
                    
                    '''//Check if the main performance_mode exists in the dictionary "VddbinPmodeDict".
                    If VddbinPmodeDict.Exists(performance_mode) Then
                        main_p_mode = VddBinStr2Enum(performance_mode)
                        powerDomain = AllBinCut(main_p_mode).powerPin
                        
                        If gb_bincut_power_list(VddBinStr2Enum(powerDomain)) <> "" Then
                            '''//Check if performance_mode exists in gb_bincut_power_list(VddBinStr2Enum(powerDomain).
                            If UCase("*," & gb_bincut_power_list(VddBinStr2Enum(powerDomain)) & ",*") Like UCase("*," & performance_mode & ",*") Then
                                '''pmode exists in the list, so that do nothing...
                            Else
                                '''=============================================================================================='''
                                '''//Check digit2-3 of the p_mode are same as p_modes in the power_list, ex: "MC" of "MC601" and "MC602".
                                '''20190704: As the discussion with SWLINZA, we should check digit1-2 of the p_mode are same as p_modes in the power_list.
                                '''=============================================================================================='''
                                split_content = Split(gb_bincut_power_list(VddBinStr2Enum(powerDomain)), ",")
                                
                                If Mid(UCase(ws_def.Cells(row, 2).Value), 1, 2) = Mid(UCase(split_content(0)), 1, 2) Then
                                    gb_bincut_power_list(VddBinStr2Enum(powerDomain)) = gb_bincut_power_list(VddBinStr2Enum(powerDomain)) & "," & performance_mode
                                Else
                                    TheExec.Datalog.WriteComment "sheet:" & sheetName & "," & performance_mode & " is incosistent with " & powerDomain & " power_seq " & split_content(0) & ". Please check Domain and Mode columns in Vdd_Binning_Def. initVddBinCondition has the incorrect keyword. Error!!!"
                                    TheExec.ErrorLogMessage "sheet:" & sheetName & "," & performance_mode & " is incosistent with " & powerDomain & " power_seq " & split_content(0) & ". Please check Domain and Mode columns in Vdd_Binning_Def. initVddBinCondition has the incorrect keyword. Error!!!"
                                End If
                            End If
                        Else
                            gb_bincut_power_list(VddBinStr2Enum(powerDomain)) = performance_mode
                        End If
                    
                        '''//Add the additional mode into the dictionary "AdditionalModeDict" for Additional Mode.
                        If additional_mode <> "" Then
                            If isParsingOutsideBinCutFlow = True Then
                                If dict_OutsideBinCut_additionalMode.Exists(additional_mode) = True Then
                                    If sheetName <> dict_OutsideBinCut_additionalMode.Item(additional_mode) Then
                                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", performance mode:" & str_mode_temp & ", it has the duplicate string about additional mode:" & additional_mode & ". Error!!!"
                                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", performance mode:" & str_mode_temp & ", it has the duplicate string about additional mode:" & additional_mode & ". Error!!!"
                                        Exit Function
                                    Else
                                        addi_mode = AdditionalModeDict.Item(additional_mode)
                                    End If
                                Else
                                    If AdditionalModeDict.Exists(additional_mode) Then
                                        addi_mode = AdditionalModeDict.Item(additional_mode)
                                    Else
                                        cntAdditionalMode = cntAdditionalMode + 1
                                        AdditionalModeDict.Add UCase(additional_mode), cntAdditionalMode
                                        addi_mode = cntAdditionalMode
                
                                        ReDim Preserve AdditionalModeName(cntAdditionalMode)
                                        AdditionalModeName(cntAdditionalMode) = additional_mode
                                        dict_OutsideBinCut_additionalMode.Add additional_mode, sheetName
                                    End If
                                End If
                            Else
                                If AdditionalModeDict.Exists(additional_mode) Then
                                    addi_mode = AdditionalModeDict.Item(additional_mode)
                                Else
                                    cntAdditionalMode = cntAdditionalMode + 1
                                    AdditionalModeDict.Add UCase(additional_mode), cntAdditionalMode
                                    addi_mode = cntAdditionalMode
            
                                    ReDim Preserve AdditionalModeName(cntAdditionalMode)
                                    AdditionalModeName(cntAdditionalMode) = additional_mode
                                End If
                            End If
                        Else
                            addi_mode = -1
                        End If
                        
                        '''//Parsing testCondition into BinCut(p_mode, bin_number) array.
                        For i = 0 To cntVddbinPin - 1
                            For passBinCut = binNumStart To binNumStop
                                col = other_voltage_start_point + i
                            
                                '''//Check if any testcondition contains "#REF!" or "#NAME?".
                                If IsError(ws_def.Cells(row, col).Value) Then
                                    testCondition = ""
                                    TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell (row:" & row & ", column:" & col & "). it has the incorrect content. Error!!!"
                                    TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & row & ", column:" & col & "). it has the incorrect content. Error!!!"
                                Else
                                    selected_powerDomain = UCase(dict_BinCutFlow_Column2Domain.Item(col))
                                    
                                    '''//Get and trim string of the testCondition from Non_Binning_Rail.
                                    testCondition = LCase(Trim(ws_def.Cells(row, col).Value))
                                    
                                    '''//Check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY)...
                                    If testCondition Like "*bin1*" Then '''Bin1
                                        bincutNum = 1
                                    ElseIf testCondition Like "*binx*" Then '''BinX
                                        bincutNum = 2
                                    ElseIf testCondition Like "*biny*" Then '''BinY
                                        bincutNum = 3
                                    Else
                                        bincutNum = 0
                                    End If
                                    
                                    '''//Check if bincutNum is greater than PassBinCut_ary(Ubound(PassBinCut_ary)).
                                    '''//PassBinCut_ary(Ubound(PassBinCut_ary)) is the highest Bin number of the BinCut voltage table(sheet "Vdd_Binning_Def").
                                    '''20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY) greater than the highest bin number.
                                    If bincutNum > PassBinCut_ary(UBound(PassBinCut_ary)) Then
                                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ",cell:(row " & row & ", column " & col & "),testCondition:" & testCondition & ", it contains the incorrect keyword about PassBin greater than the highest bin number for initVddBinCondition. Error!!!"
                                        TheExec.ErrorLogMessage "sheet:" & sheetName & ",cell:(row " & row & ", column " & col & "),testCondition:" & testCondition & ", it contains the incorrect keyword about PassBin greater than the highest bin number for initVddBinCondition. Error!!!"
                                    End If
                                    
                                    '''//Check if testCondition contains "(" but no ")".
                                    If testCondition Like "*(*" And Not (testCondition) Like "*)*" Then
                                        TheExec.Datalog.WriteComment "Please check the cell (row " & row & ", column " & col & ") of sheet " & sheetName & ". The cell contains the incorrect format. Error!!!"
                                        TheExec.ErrorLogMessage "Please check the cell (row " & row & ", column " & col & ") of sheet " & sheetName & ". The cell contains the incorrect format. Error!!!"
                                    End If
                                    
                                    '''==========================================================================================================================================='''
                                    '''//Check if testCondition contain keyword "*Evaluate*Bin*" to decide "is_BinCutJob_for_StepSearch" = True (BinCut stepSearch) while isParsingOutsideBinCutFlow = False.
                                    '''==========================================================================================================================================='''
                                    If testCondition Like LCase("*Evaluate*Bin*") And isParsingOutsideBinCutFlow = False Then
                                        strSplitted = Split(LCase(testCondition), LCase("Evaluate Bin"))
                                        
                                        '''//Check if testCondition with keyword "*Evaluate*Bin*" has the correct performance mode.
                                        '''20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
                                        If VddbinPmodeDict.Exists(UCase(Trim(strSplitted(0)))) = True And dict_IsCorePower.Exists(UCase(Trim(strSplitted(0)))) = False Then
                                            AllBinCut(VddBinStr2Enum(UCase(Trim(strSplitted(0))))).is_for_BinSearch = True
                                            is_BinCutJob_for_StepSearch = True
                                        Else
                                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ",cell:(row " & row & ", column " & col & "),testCondition:" & testCondition & ". It doesn't contain any correct performance mode in testCondition, please check sheet " & sheetName & ". Error!!!"
                                            TheExec.ErrorLogMessage "sheet:" & sheetName & ",cell:(row " & row & ", column " & col & "),testCondition:" & testCondition & ". It doesn't contain any correct performance mode in testCondition, please check sheet " & sheetName & ". Error!!!"
                                        End If
                                    End If
                                    
                                    If isParsingOutsideBinCutFlow = False Then
                                        If HVCC_flag = 0 Then
                                            If addi_mode > 0 Then
                                                BinCut(main_p_mode, passBinCut).Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                            Else
                                                BinCut(main_p_mode, passBinCut).OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                            End If
                                        ElseIf HVCC_flag = 1 Then
                                            If addi_mode > 0 Then
                                                BinCut(main_p_mode, passBinCut).HVCC_Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                            Else
                                                BinCut(main_p_mode, passBinCut).HVCC_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                            End If
                                        End If
                                    Else
                                        '''20201021: As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
                                        If addi_mode > 0 Then
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                        Else
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_HVCC_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                        End If
                                        Flag_NonbinningrailOutsideBinCut_parsed = True
                                    End If
                                End If
                            Next passBinCut
                        Next i
                    Else '''If VddbinPmodeDict.Exists(UCase(replace_p_name(0)))=false
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ",row:" & row & ",cell:" & str_mode_temp & ", it doesn't contain any correct performance mode. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & row & ",cell:" & str_mode_temp & ", it doesn't contain any correct performance mode. Error!!!"
                    End If
                End If '''If str_mainColumn_content Like "*hvcc*"
            End If '''If IsError(ws_def.Cells(row, other_voltage_start_point).Value) Then
            
            row = row + 1
        Wend
    End If '''If row_of_title > 0
    
    '''====================================================-====================================================================
    '''[Step3] Check if "cntAdditionalMode" should be "<=" with "MaxAdditionalModeCount".
    '''====================================================-====================================================================
    If cntAdditionalMode > MaxAdditionalModeCount Then
        TheExec.Datalog.WriteComment "sheet:" & sheetName & ",number of BinCut additional modes:" & cntAdditionalMode & ", it is greater than BinCut globalVariable MaxAdditionalModeCount=" & MaxAdditionalModeCount & ". Please check BinCut flow table and globalVariable MaxAdditionalModeCount. Error!!!"
        TheExec.ErrorLogMessage "sheet:" & sheetName & ",number of BinCut additional modes:" & cntAdditionalMode & ", it is greater than BinCut globalVariable MaxAdditionalModeCount=" & MaxAdditionalModeCount & ". Please check BinCut flow table and globalVariable MaxAdditionalModeCount. Error!!!"
    End If
    
    If isParsingOutsideBinCutFlow = False Then
        '''//Split pin_groups and get each powerDomain, then sort the sequence about p_mode for each powerDomain. BinCut PowerDomain consists of CorePower and OtherRail.
        If FullBinCutPowerinFlowSheet <> "" Then
            pinGroup_BinCut = Split(FullBinCutPowerinFlowSheet, ",")
            
            For i = 0 To UBound(pinGroup_BinCut)
                powerDomain = UCase(pinGroup_BinCut(i))
            
                If dict_IsCorePower.Item(powerDomain) = True Then '''CorePower
                    If FullCorePowerinFlowSheet <> "" Then
                        FullCorePowerinFlowSheet = FullCorePowerinFlowSheet & "," & powerDomain
                    Else
                        FullCorePowerinFlowSheet = powerDomain
                    End If
                    
                    dict_IsCorePowerInBinCutFlowSheet.Add powerDomain, True
                Else '''OtherRail
                    If FullOtherRailinFlowSheet <> "" Then
                        FullOtherRailinFlowSheet = FullOtherRailinFlowSheet & "," & powerDomain
                    Else
                        FullOtherRailinFlowSheet = powerDomain
                    End If
                    
                    dict_IsCorePowerInBinCutFlowSheet.Add powerDomain, False
                End If
            Next i
        Else
            TheExec.Datalog.WriteComment "FullBinCutPowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"
            TheExec.ErrorLogMessage "FullBinCutPowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"
        End If
        
        '''//Split pin_groups and get powerDomains of BinCut CorePower and OtherRail.
        '''CorePower
        If FullCorePowerinFlowSheet <> "" Then
            pinGroup_CorePower = Split(FullCorePowerinFlowSheet, ",")
        Else
            TheExec.Datalog.WriteComment "FullCorePowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"
            TheExec.ErrorLogMessage "FullCorePowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"
        End If
        
        '''OtherRail
        If FullOtherRailinFlowSheet <> "" Then
            pinGroup_OtherRail = Split(FullOtherRailinFlowSheet, ",")
        Else
            TheExec.Datalog.WriteComment "FullOtherRailinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"
            TheExec.ErrorLogMessage "FullOtherRailinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVddBinCondition"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVddBinCondition"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201223: Modified to check if any sheet contains Keyword in sheetName.
'20201222: Modified to revise the vbt function "initVddBinCondition" for multiple "Non_Binning_Rail_Outside_BinCut" sheets.
'20201222: Created to parse multiple "Non_Binning_Rail_Outside_BinCut" sheets.
Public Function parsing_OutsideBinCut_flow_table(keyword_sheetName As String)
    Dim i As Long
    Dim count_WorkSheet As Integer
    Dim outsideBincutSheetsArr() As String
    Dim idxSheet As Integer
    Dim maxSheetsNum As Integer
    Dim sheetName As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Please check sheetName keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
'''//==================================================================================================================================================================================//'''
    '''init
    idxSheet = -1

    '''//Check if the sheet exists.
    If keyword_sheetName = "Non_Binning_Rail_Outside" Then
        count_WorkSheet = Application.ActiveWorkbook.Worksheets.Count
        
        '''//Check if sheet name contains keyword_sheetName.
        For i = 1 To count_WorkSheet
            If ActiveWorkbook.Worksheets(i).Name Like "*" & keyword_sheetName & "*" Then
                idxSheet = idxSheet + 1
                ReDim Preserve outsideBincutSheetsArr(idxSheet)
                outsideBincutSheetsArr(idxSheet) = ActiveWorkbook.Worksheets(i).Name
            End If
        Next i
    ElseIf keyword_sheetName <> "" Then
        TheExec.Datalog.WriteComment keyword_sheetName & " is not the correct keyword to find Outside BinCut flow table for parsing_OutsideBinCut_flow_table. Error!!!"
        TheExec.ErrorLogMessage keyword_sheetName & " is not the correct keyword to find Outside BinCut flow table for parsing_OutsideBinCut_flow_table. Error!!!"
        Exit Function
    Else '''If keyword_sheetName is empty...
        Exit Function
    End If
    
    '''//Use sheet-loop to parse each Outside BinCut flow table.
    If idxSheet > -1 Then
        For i = 0 To idxSheet
            sheetName = outsideBincutSheetsArr(i)
            
            '''//Parsing each of sheets with Keyword in sheetName.
            initVddBinCondition sheetName
        Next i
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVddBinCondition"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVddBinCondition"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'20210806: Modified to get and print Guardband(GB) according to the BinCut testjob.
'20210706: Modified to replace is_BinCutJob_for_StepSearch with inst_info.is_BinSearch.
'20210219: Modified to check the flag "Flag_Skip_Printing_Safe_Voltage" to skip printing BV strings of BinCut Safe Voltages.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201104: Modified to replace "bincutJobName = "cp1" with inst_info.is_binsearch=True.
'20201102: Modified to update "inst_info.is_BV_Safe_Voltage_printed" and "inst_info.is_BV_Payload_Voltage_printed".
'20201029: Modified to remove the redundant arguments "str_dynamic_offset() As String" and "str_Selsrm_DSSC_Info() As String" from print_bincut_power.
'20201027: Modified to use "Public Type Instance_Info".
'20200925: Modified the branch for "indexstep_per_site".
'20200319: Modified for "Flag_PrintDcvsShadowVoltage".
'20200214: Modified to print dynamic_offset.
'20200214: Modified to print eqn information for payload voltages.
'20200211: Modified to get init voltage for DCVS shadow voltages.
'20200211: Modified to replace "FlowTestCondName" with "AdditionalModeName".
'20200206: Modified to replace "print_main_power_init" with "print_bincut_power".
'20200203: Created to merge the functions: print_main_power, print_alt_power, print_main_power_payload, print_alt_power_payload.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20191105: Modified to print offsetTestType.
'20191002: Modified to add BinCut voltageType.
'20180910: Modified to control print BinCut voltages by "remove_printing_voltage".
Public Function print_bincut_voltage(inst_info As Instance_Info, Optional passBinCut As SiteLong, Optional remove_printing_voltage As Boolean = False, Optional Flag_PrintDcvsShadowVoltage As Boolean = False, _
                                        Optional voltageType As Integer = BincutVoltageType.None, Optional DcSpecsCategoryForInitPat As String = "")
    Dim site As Variant
    Dim i As Long, j As Long
    Dim powerDomain As String
    Dim powerPin As String
    Dim strTemp As String
    Dim strPrefix As String
    Dim strOutput As String
    Dim strOutputEQN As String
    Dim voltage_PowerDomain As Double
    Dim performance_mode As String
    Dim dbl_GB_BinCutJob As Double
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''str_Selsrm_DSSC_Info     : store info about SELSRM bits comparison.
'''str_Selsrm_DSSC_Bit      : store info about SELSRM bits sequence(LSB->MSB) of each site.
'''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'''//==================================================================================================================================================================================//'''
    '''//Check the flag "Flag_Skip_Printing_Safe_Voltage" to skip printing BV strings of BinCut Safe Voltages.
    If remove_printing_voltage = True Or (Flag_Skip_Printing_Safe_Voltage = True And voltageType = BincutVoltageType.SafeVoltage) Then
        If Flag_Skip_Printing_Safe_Voltage = True And voltageType = BincutVoltageType.SafeVoltage Then
            TheExec.Datalog.WriteComment "****************separated for BinCut step****************"
        End If
        Exit Function
    Else
        '''//init
        strPrefix = ""
        
        '''//Check if Flag_PrintDcvsShadowVoltage is enabled to print DCVS shadow voltages (calculation values of BinCut payload voltages).
        If Flag_PrintDcvsShadowVoltage = True Then
            TheExec.Datalog.WriteComment "Print DCVS shadow voltages (from calculation), not DCVS real values!!!"
        End If

        '''//Check the performance mode if it has the additional mode, ex: MS003_GPU.
        '''If special_voltae_setup = True, it means that the performance mode has the additional mode.
        If inst_info.special_voltage_setup = True Then
            performance_mode = VddBinName(inst_info.p_mode) & "_" & AdditionalModeName(inst_info.addi_mode)
        Else
            performance_mode = VddBinName(inst_info.p_mode)
        End If
        
        '''//Get the prefix of voltage string.
        If voltageType = BincutVoltageType.InitialVoltage Then
            strPrefix = "Initial_Voltage_" & performance_mode
        Else
            strPrefix = "BV_" & performance_mode
        End If
        
        For Each site In TheExec.sites
            '''init
            strOutput = ""
            strOutputEQN = ""
            strTemp = ""
            
            '''//Print strings of dynamic_offset or SELSRAM_DSSC prior to BV string of BinCut payload voltages.
            If (voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None) And inst_info.offsetTestTypeIdx <> ldcbfd Then
                '''//dynamic_offset of binning p_mode.
                If inst_info.str_dynamic_offset(site) <> "" Then
                    TheExec.Datalog.WriteComment inst_info.str_dynamic_offset(site)
                End If
                
                '''//SELSRAM_DSSC
                If inst_info.str_Selsrm_DSSC_Info(site) <> "" Then
                    TheExec.Datalog.WriteComment inst_info.str_Selsrm_DSSC_Info(site)
                End If
            End If
            
            '''//In testjob "CP1", datalog always show EQN and passbin, C, M and CPGB.
            '''20190716: Modified to unify the unit for IDS. ids_current with unit mA.
            If inst_info.is_BinSearch = True And (voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None) Then
                If inst_info.step_Current(site) <> -1 Then
                    '''//Get the matched Guardband(GB) according to the BinCut testjob.
                    '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                    dbl_GB_BinCutJob = BinCut(inst_info.p_mode, DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current)).CP_GB(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(inst_info.step_Current) - 1)
                    
                    strOutputEQN = strPrefix & "," & site & ", EQN = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(inst_info.step_Current) & _
                                    ", PASSBIN = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current) & _
                                    ", C = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).c(inst_info.step_Current) & _
                                    ", M = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).M(inst_info.step_Current) & _
                                    ", GB = " & dbl_GB_BinCutJob & _
                                    ", IDS = " & inst_info.ids_current & " mA"
                Else
                    strOutputEQN = ""
                End If
            End If
            
            '''//Since tracking power is already added to pinGroup, no need to print tracking power...
            '''As per discussion with SWLINZA, ZHHUANG, and PCLIN. We decided to measure the voltage of 1st powerPin to print each power domain.
            '''We also merged the vbt code for online and offline tests. Only read the real voltage value from DCVS.
            For i = 0 To UBound(pinGroup_BinCut)
                powerDomain = pinGroup_BinCut(i)
                
                '''//Only payload voltages can use "DCVS_shadow_voltages".
                If Flag_PrintDcvsShadowVoltage = True And voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None Then
                    voltage_PowerDomain = BinCut_Payload_Voltage(VddBinStr2Enum(powerDomain)) / 1000 '''DCVS should use unit: V
                Else
                    '''//Read the real voltage value from DCVS.
                    '''ToDo: For project with UltraFlexPlus, it can directly read voltage values from DCVS by using "ValuePerSite"...
                    powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(powerDomain))
                    '''ToDo: Check if powerPin is DCVS or DCVI by checking VddbinPinDcvsType...
                    voltage_PowerDomain = TheHdw.DCVS.Pins(powerPin).Voltage.Value
                End If
                
                If strTemp <> "" Then
                    strTemp = strTemp & "," & powerDomain & "=" & Format(voltage_PowerDomain, "0.000")
                Else
                    strTemp = powerDomain & "=" & Format(voltage_PowerDomain, "0.000")
                End If
            Next i
            
            '''//Print testType and offsetType at the end of BV string.
            If voltageType <> BincutVoltageType.None Then
                If inst_info.offsetTestTypeIdx <> ldcbfd Then
                    strOutput = strPrefix & "," & site & "," & strTemp & "," & " (" & BincutVoltageTypeName(voltageType) & "," & TestTypeName(inst_info.offsetTestTypeIdx) & ")"
                Else
                    strOutput = strPrefix & "," & site & "," & strTemp & "," & " (" & BincutVoltageTypeName(voltageType) & ")"
                End If
            End If
            
            '''//Print out the string in the datalog.
            TheExec.Datalog.WriteComment strOutput
            
            '''//Update the status of "inst_info.is_BV_Safe_Voltage_printed" and "inst_info.is_BV_Payload_Voltage_printed" to control printing BV strings into the datalog.
            If voltageType = BincutVoltageType.SafeVoltage Then
                inst_info.is_BV_Safe_Voltage_printed = True
            ElseIf (voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None) Then
                inst_info.is_BV_Payload_Voltage_printed = True
            End If
            
            '''//Print info about BinCut EQN for BinSearch.
            If inst_info.is_BinSearch = True And strOutputEQN <> "" Then
                strOutputEQN = strOutputEQN & "," & strTemp
            
                '''ex:BV_VDD_SOC_MS001,1,EQN = 7, PASSBIN = 1, C = 609.375, M = 70, CPGB = 78.125, _
                '''IDS = 27.4 mA,,VDD_PCPU=0.752,VDD_ECPU=0.752,VDD_GPU=0.752,VDD_SOC=0.752,VDD_DCS_DDR=0.752,VDD_AVE=0.752,VDD_DISP=0.752, _
                '''VDD_SRAM_CPU=0.752,VDD_SRAM_ANE=0.752,VDD_SRAM_GPU=0.752,VDD_SRAM_SOC=0.752,VDD_FIXED=0.800,VDD_LOW=0.735
                TheExec.Datalog.WriteComment strOutputEQN
            End If
        Next site
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of print_bincut_voltage"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201029: Modified to remove the redundant arguments "str_dynamic_offset() As String" and "str_Selsrm_DSSC_Info() As String" from print_voltage_info_before_FuncPat.
'20201027: Modified to use "Public Type Instance_Info".
'20201026: Modified to revise the vbt code for TD pattern burst proposed by C651 Toby.
'20201008: Modified to move the vbt code of printing payload voltages for Pattern Burst(not decompose pattern set) from "print_voltage_info_before_FuncPat" to "prepare_DCVS_Output_for_RailSwitch".
'20200921: Modified to check if "Test_Type = TestType.Mbist".
'20200918: Created to print BinCut voltage before running FuncPat.
'20200113: As per discussion with Leon/Jeff/PSYAO/Minder/PCLIN, we decided to print payload voltages for pattern bursted without decomposing pattern.
Public Function print_voltage_info_before_FuncPat(inst_info As Instance_Info)
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''<For projects without Rail-Switch>
'''   Print BinCut voltages of Func Pattern for TD or Mbist test instances
'''<For projects with Rail-Switch>
'''   Print BinCut voltages of Payload Pattern(Func Pattern) for Mbist test instances
'''   Print BinCut voltages of Init Pattern in Func Patsets for TD test instances
'''//==================================================================================================================================================================================//'''
    If inst_info.test_type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
        '''//For Mbist test, it prints BinCut payload voltages before running Payload pattern. So that it can use DCVS shadow voltages.
        If inst_info.is_BV_Payload_Voltage_printed = False Then
            print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
        End If
    Else '''For TD/SA/SCAN test instances...
        '''//Use the flags "inst_info.is_BV_Safe_Voltage_printed" and "inst_info.is_BV_Payload_Voltage_printed = False" to avoid printing duplicate BV strings of safe voltages and payload voltages for TD init+pl+init+pl patset.
'        If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch
            If inst_info.is_BV_Safe_Voltage_printed = False Then
                print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.SafeVoltage
            End If
'        Else
'            If inst_info.is_BV_Payload_Voltage_printed = False Then
'                print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
'            End If
'        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of print_voltage_info_before_FuncPat"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of print_voltage_info_before_FuncPat"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210528: Modified to add inst_info to call Calculate_Harvest_Core_DSSC_Source.
'20210513: Modified to use Calculate_Harvest_Core_DSSC_Source.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201111: Modified to move "prepare_DCVS_Output_for_RailSwitch" from modudle LIB_Vdd_Binning_customer to LIB_VDD_BINNING.
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201029: Modified to remove the redundant arguments "str_dynamic_offset() As String" and "str_Selsrm_DSSC_Info() As String" from prepare_DCVS_Output_for_RailSwitch.
'20201027: Modified to use "Public Type Instance_Info".
'20201014: Modified to use "Check_PayloadPattern_with_DCVS" for offline, requested by Leon Weng.
'20201014: Modified to check if powerDomain is not "PRESERVED" or "RESERVED", requested by Leon Weng.
'20201008: Modified to replace "PrintedBVinDatalog" with "is_BV_Payload_Voltage_printed"
'20201008: Modified to move the vbt code of printing payload voltages for Pattern Burst(not decompose pattern set) from "print_voltage_info_before_FuncPat" to "prepare_DCVS_Output_for_RailSwitch".
'20201008: Modified to force DCVS to Valt and skip Check_PayloadPattern_with_DCVS for offline.
'20201008: Modified to check if "Test_Type = testType.TD".
'20201008: Modified to add the vbt code of printing BinCut payload voltages for offline simulation.
'20201007: Modified to check vbump only for online tests. requested by Leon Weng.
'20200924: Modified to add the argument "idxBlock_Selsrm_Pattern" for SELSRM DSSC signal setup.
'20200918: Created to decide syncup DCVS output and print BinCut payload voltages.
'20191009: Modified For open socket or offline simulation.
'20190626: As the discussion with TSMC PSYAO, we found vbump didn't exist in all BinCut powerDomain of BinCut patterns.
'20190121: Modified by Oscar. Signal is reusable so we remove the Recalculation and Resend.
'20181119: Resend DSSC Selsram bits for the pattern.
Public Function prepare_DCVS_Output_for_RailSwitch(inst_info As Instance_Info, str_pattern As String, Optional idxBlock_Selsrm_Pattern As Integer = -1)
    Dim i As Integer
    Dim selSramPat As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//tlDCVSVoltageOutput: This property gets or sets the output DAC used to force voltage (main or alternate).
'''It can detect DCVS output Vmain/Valt of the binning CorePower powerDomains.
'''The returned information is one of the following values:
'''     tlDCVSVoltageMain   : 1 ==> Main output voltage DAC.
'''     tlDCVSVoltageAlt    : 2 ==> Alternate output voltage DAC.
'''//==================================================================================================================================================================================//'''
    If inst_info.test_type = testType.TD Then
        '''//Notice!!! for TD tests, FuncPat pattern group consists of INIT and Payload patterns.
        '''***************************************************************************'''
        '''//TD patt_group (init+payload1+init+payload2)
        '''//Since Payload patterns in the group use the same BinCut payload voltages, it can skip SelSram bits calculation for the second and following payload patterns.
        '''***************************************************************************'''
        If idxBlock_Selsrm_Pattern > -1 Then
            '''//Resend DSSC Selsram bits for the pattern.
            If LCase(str_pattern) Like LCase(SelsramMapping(idxBlock_Selsrm_Pattern).Pattern) Then  '''for DSSC patt_group(init+pl+init+pl), 20181115
                selSramPat = str_pattern
                TheHdw.DSSC.Pins(inst_info.selsrm_DigSrc_Pin).Pattern(str_pattern).Source.Signals.DefaultSignal = inst_info.selsrm_DigSrc_SignalName
            End If
        End If
        
        '''//Calculate DSSC bit sequence of Harvest Core DSSC (FSTP and MultiFSTP), then do DSSC DigSrcWaveSetup with the pattern and DSSC bit sequence before running the pattern.
        '''Since PrePatt and FuncPat share this utility, we could use this as the common funciton for Harvest Core DSSC.
        '''Flag_Harvest_Core_DSSC_Ready = True if HarvPmodeTable and HARV_Pmode_Table.
        Call Calculate_Harvest_Core_DSSC_Source(inst_info.inst_name, VddBinName(inst_info.p_mode), str_pattern, inst_info.Harvest_Core_DigSrc_Pin, inst_info.Harvest_Core_DigSrc_SignalName, inst_info.Pattern_Pmode, inst_info.By_Mode)
        
        If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch
            '''20190626: As the discussion with TSMC PSYAO, we found vbump didn't exist in all BinCut powerDomain of BinCut patterns.
            '''//It detects that Any of SELSRAM powerDomains is switched to Valt.
            '''//If that, all BinCut powerDomains will be switched to Valt.
            For i = 0 To UBound(selsramLogicPingroup)
                '''20201014: Modified to check if powerDomain is not "PRESERVED" or "RESERVED", requested by Leon Weng.
                If UCase(selsramLogicPingroup(i)) <> "PRESERVED" And UCase(selsramLogicPingroup(i)) <> "RESERVED" Then
                    If (TheHdw.DCVS.Pins(selsramLogicPingroup(i)).Voltage.output = tlDCVSVoltageAlt) Then
                        inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                        Exit For
                    End If
                End If
            Next i
            
            '''//DCVS should be switched to Valt by Pattern with vbump prior to Payload pattern.
            '''//Check if vbump in patset before running the payload pattern, especially for project with rail-switch.
            '''20201007: Modified to check vbump only for online tests. requested by Leon Weng.
            If LCase(str_pattern) Like "*_pl*" Or Not (LCase(str_pattern) Like "*_in*") Then '''1st pattern of non-init patterns.
                '''**********************************************************************************************************'''
                '''Offline tests and OpenSocket can't detect Vmain or Valt of DCVS.
                '''So that we use the keyword "*_pl*" to detect the payload patterns and switch DCVS to Valt for TD instance.
                '''**********************************************************************************************************'''
                If Flag_VDD_Binning_Offline = True Or EnableWord_Vddbinning_OpenSocket = True Then '''offline or OpenSocket
                    select_DCVS_output_for_powerDomain tlDCVSVoltageAlt
                    inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                End If
                
                If inst_info.enable_DecomposePatt = True Then
                    Call Check_PayloadPattern_with_DCVS(inst_info.inst_name, Flag_Enable_Rail_Switch, str_pattern, inst_info.currentDcvsOutput, inst_info.enable_DecomposePatt)
                End If
            End If
            
            '''//Projects with Rail Switch might have the incomplete pin listed in the patterns, so that it detects the output status of binning power to sync up other pins.
            '''ToDo: Maybe it can skip SyncUp_DCVS_Output if "inst_info.enable_DecomposePatt = False"...
'            If Flag_SyncUp_DCVS_Output_enable Then
                Call SyncUp_DCVS_Output(inst_info.p_mode, inst_info.currentDcvsOutput, SyncUp_PowerPin_Group) '''This is to sync up logic powers and sram powers on the same DCVS output (for TD testing)
'            End If
            
            '''//Print BinCut payload voltages before running TD payload patterns of projects with Rail-Switch.
            If inst_info.is_BV_Payload_Voltage_printed = False Then
                If inst_info.previousDcvsOutput = tlDCVSVoltageMain And inst_info.currentDcvsOutput = tlDCVSVoltageAlt Then
                    '''==============================================================================================='''
                    '''[Note]: For projects with Rail Switch, it can detect the first transition Vmain-> Valt of the binning CorePower and print BinCut voltages for payload.
                    '''Especially, for those pattern set (INIT-> PL1 -> INIT -> PL2), it only needs to print BinCut payload voltages once.
                    '''No need to print BinCut payload voltages for PL2 again.
                    '''==============================================================================================='''
                    print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
                ElseIf inst_info.enable_DecomposePatt = False Then '''without decomposing pattern sets
                    '''==============================================================================================='''
                    '''//For TD/SCAN test, it prints BinCut payload voltages (values from calculation, not from DCVS) before running pattern bursted without decomposing pattern.
                    '''20200113: As per discussion with Leon/Jeff/PSYAO/Minder/PCLIN:
                    '''we decided to use shadow voltage (calculation value) for printing BinCut payload voltages for pattern bursted without decomposing pattern.
                    '''==============================================================================================='''
                    print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, True, BincutVoltageType.PayloadVoltage
                End If
            End If
            
            '''//Store current status to "PreviousDcvsOutput" for the control of printing payload voltage in the datalog.
            inst_info.previousDcvsOutput = inst_info.currentDcvsOutput
        End If
    End If
    
    '''//Check if any alarm exists.
    '''==============================================================================================='''
    '''This method forces an alarm check. It determines whether alarms are present and reports on them.
    '''This method clears alarms. For this reason, do not use it for monitoring alarms during debugging.
    '''20200106: As per discussion with SWLINZA, he suggested us to add this to check any alarm.
    '''==============================================================================================='''
    TheHdw.Alarms.Check
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of prepare_DCVS_Output_for_RailSwitch"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of prepare_DCVS_Output_for_RailSwitch"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210305: Modified to add the argument "siteResult" to the vbt function "StoreCapFailcycle".
'20210129: Modified to revise the vbt code for DevChar.
'20210125: Modified to remove "Optional voltage_Pmode_EQNbased As SiteDouble" from the arguments of the vbt function "run_patt_from_FuncPat_for_BinCut".
'20201210: Created to run pattern decomposed from FuncPat patset for BinCut online and offline.
'20201118: Modified to use "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" to get siteResult of pattern pass/fail.
Public Function run_patt_from_FuncPat_for_BinCut(inst_info As Instance_Info, indexPatt As Long, str_pattern As String, funcPatPass As SiteBoolean, _
                                                    Optional idxBlock_Selsrm_Pattern As Integer = -1, Optional CaptureSize As Long, Optional failpins As String)
    Dim siteResult As New SiteBoolean
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
'''It seemed that theExec.sites.Selected masked the failed site (not siteShutDown). But the site still ran pattern.test without updating test results.
'''Discussed this with Chihome. He saw this is ancient projects, and he suggested us to check if test results were correct.
'''We checked test results, and it seemed no error with PassBin and EQN.
'''//==================================================================================================================================================================================//'''
'''//==================================================================================================================================================================================//'''
'''//Run pattern decomposed from FuncPatt patset, and get siteResult of pattern pass/fail.
'''step1: Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
'''step2: Run the pattern decomposed from FuncPat, and get siteResult of pattern pass/fail.
'''step3: [Optional] Store Fail cycle of the pattern from Capture Memory(CMEM) for the current BinCut search step.
'''step4: Check alarmFail for pattern.
'''step5: [Optional] Save result about pattern Pass/Fail for COFInstance.
'''step6: Update pattern pass/fail to the patPass flag.
'''//==================================================================================================================================================================================//'''
    '''//step1: Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
    Call prepare_DCVS_Output_for_RailSwitch(inst_info, str_pattern, idxBlock_Selsrm_Pattern)

    '''//step2: Run the pattern decomposed from FuncPat, and get siteResult..
    If Flag_VDD_Binning_Offline = False Then '''If the test mode is Online.
        '''//Only CP1 uses CMEM.
        If inst_info.enable_CMEM_collection = True Then TheHdw.Digital.CMEM.SetCaptureConfig CaptureSize, CmemCaptFail, tlCMEMCaptureSourcePassFailData
        
        '''//Run the pattern decomposed from FuncPat.
        Call TheHdw.Patterns(str_pattern).Test(pfAlways, 0, inst_info.result_mode)
        
        '''//Get siteResult of pattern pass/fail.
        '''//Warning!!! currently "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" doesn't support "result_mode=tlResultModeModule" with PatternBurst=Yes and DecomposePatt=No.
        siteResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        
        '''//step3: [Optional] Store Fail cycle of the pattern from Capture Memory(CMEM) for the current BinCut search step.
        '''20210305: Modified to add the argument "siteResult" to the vbt function "StoreCapFailcycle".
        If inst_info.enable_CMEM_collection = True Then
            Call StoreCapFailcycle(siteResult, failpins, indexPatt, CaptureSize, inst_info.Step_CMEM_Data)
        End If
    Else '''Offline
        Call run_patt_offline_simulation(str_pattern, inst_info.result_mode, siteResult)
    End If
    
    '''//step4: Check alarmFail for pattern.
    Call check_alarmFail_for_pattern(siteResult)
    
    '''//step5: [Optional] Save result about pattern Pass/Fail for COFInstance.
    '''Use "update_patt_result_for_COFInstance" to record per pattern pass/fail and save EQN-based BinCut payload voltage of per site for "COFInstance".
    If inst_info.enable_COFInstance = True Then
        Call update_patt_result_for_COFInstance(inst_info, indexPatt, str_pattern, siteResult)
    End If

funcPatPass.Value = funcPatPass.LogicalAnd(siteResult)
    '''//step6: Update pattern pass/fail to the patPass flag.
    '''20210129: Modified to revise the vbt code for DevChar.
'    If inst_info.is_DevChar_Running = False Then
'        Call update_Pattern_result_to_PattPass(siteResult, funcPatPass)
'    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of run_patt_from_FuncPat_for_BinCut"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of run_patt_from_FuncPat_for_BinCut"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200207: Created to merge set_core_power_main and set_core_power_alt.
'20191210: Modified to use selsramPin domainGroup.
Public Function select_DCVS_output_for_powerDomain(selected_DCVS_output As Integer, Optional domainGroup As String)
    Dim split_content() As String
    Dim i As Long
    Dim PinGroup As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//tlDCVSVoltageOutput: This property gets or sets the output DAC used to force voltage (main or alternate).
'''It can detect DCVS output Vmain/Valt of the binning CorePower powerDomains
'''The returned information is one of the following values:
'''     tlDCVSVoltageMain   : 1 ==> Main output voltage DAC.
'''     tlDCVSVoltageAlt    : 2 ==> Alternate output voltage DAC.
'''//==================================================================================================================================================================================//'''
    '''//Check if the input argument "selected_DCVS_output" is correct.
    If selected_DCVS_output = tlDCVSVoltageMain Or selected_DCVS_output = tlDCVSVoltageAlt Then
        '''FullBinCutPowerinFlowSheet contains BinCut corePower and otherRail powerDomains after parsing the sheet "Non_Binning_Rail"(initVddBinCondition).
        '''20191210: Modified to use selsramPin domainGroup.
        If domainGroup = "" Then
            If selsramPin <> "" Then
                domainGroup = selsramPin
            Else
                domainGroup = FullBinCutPowerinFlowSheet
            End If
        End If
    
        split_content = Split(domainGroup, ",")
       
        '''init
        PinGroup = ""
    
        If UBound(split_content) < 0 Then
            TheExec.Datalog.WriteComment "pin_group is incorrect for set_core_power_main. Error!!!"
            TheExec.ErrorLogMessage "pin_group is incorrect for set_core_power_main. Error!!!"
        Else
            For i = 0 To UBound(split_content)
                If PinGroup = "" Then
                    PinGroup = VddbinDomain2Pin(split_content(i))
                Else
                    PinGroup = PinGroup & "," & VddbinDomain2Pin(split_content(i))
                End If
            Next i
        End If
    
        '''//Switch all BinCut powerDomains to the selected DCVS output.
        '''ToDo: Check if it needs to separate DCVS into HexVS, UVS256 and VSM groups.
        TheHdw.DCVS.Pins(PinGroup).Voltage.output = selected_DCVS_output
    Else
        TheExec.Datalog.WriteComment "The input argument of select_DCVS_output_for_powerDomain should be tlDCVSVoltageMain or tlDCVSVoltageAlt. Please check the argument selected_DCVS_output. Error!!!"
        TheExec.ErrorLogMessage "The input argument of select_DCVS_output_for_powerDomain should be tlDCVSVoltageMain or tlDCVSVoltageAlt. Please check the argument selected_DCVS_output. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of select_DCVS_output_for_powerDomain"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210113: Modified to check if cnt_DecomposedPinList>0.
'20201111: Modified to move "SyncUp_DCVS_Output" from modudle LIB_Vdd_Binning_customer to LIB_VDD_BINNING.
'20200106: Modified to remove the ErrorLogMessage.
'20200103: Modified to check the argument "powerPin" is N/C or not.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20191127: Modified for the revised InitVddBinTable.
'20181224: Modified to add "Flag_SyncUp_DCVS_Output_enable" in LIB_Vdd_Binning_GlobalVariable to control SyncUp on/off.
'20181120: Modified for GradeSearch_VT, added CurrentDcvsOutput to check Vmain/Valt for powerpin.
'20180921: SyncUp is added for switching OtherRail to Valt when detecting CorePower in Valt.
Public Function SyncUp_DCVS_Output(p_mode As Integer, currentDcvsOutput As Integer, powerGroup As String)
    Dim PinGroup As String
    Dim split_powerGroup() As String
    Dim i As Integer
    Dim j As Integer
    Dim powerDomain As String
    Dim strAry_pinSyncup() As String
    Dim cnt_DecomposedPinList As Long
    Dim domainTemp As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''This is to detect voltage source (DCVS Vmain or Valt) and sync up logic powers and sram powers on the same DCVS output (for TD tests).
'''//==================================================================================================================================================================================//'''
    '''init
    PinGroup = ""
    
    If powerGroup <> "" Then
        split_powerGroup = Split(powerGroup, ",")
        
        '''//Check if the powerGroup exists.
        For i = 0 To UBound(split_powerGroup)
            powerDomain = split_powerGroup(i)
            
            '''//Check if powerDomain belongs to BinCut CorePower or OtherRail (listed in globalVariable "FullBinCutPowerinFlowSheet").
            If UCase("*," & FullBinCutPowerinFlowSheet & ",*") Like UCase("*," & powerDomain & ",*") Then
                '''VddbinDomain2Pin
                If PinGroup <> "" Then
                    PinGroup = PinGroup & "," & VddbinDomain2Pin(powerDomain)
                Else
                    PinGroup = VddbinDomain2Pin(powerDomain)
                End If
            Else
                '''//Decompose powerDomain and check each powerPin to check the argument "powerGroup" is N/C or not.
                Call TheExec.DataManager.DecomposePinList(split_powerGroup(i), strAry_pinSyncup, cnt_DecomposedPinList)
                
                If cnt_DecomposedPinList > 0 Then
                    For j = 0 To cnt_DecomposedPinList - 1
                        If TheExec.DataManager.NumberChannelTypesForPin(strAry_pinSyncup(j)) > 0 Then
                            '''//If the powerGroup is connected to DCVS, re-assembly the pinGroup
                            If PinGroup <> "" Then
                                PinGroup = PinGroup & "," & strAry_pinSyncup(j)
                            Else
                                PinGroup = strAry_pinSyncup(j)
                            End If
                        Else
                            TheExec.Datalog.WriteComment strAry_pinSyncup(j) & " of " & powerDomain & " in " & powerGroup & " doesn't exist in PinMap or ChannelMap. Error!!!"
                            'TheExec.ErrorLogMessage strAry_pinSyncup(j) & " of " & powerDomain & " in " & powerGroup & " doesn't exist in PinMap or ChannelMap. Error!!!"
                        End If
                    Next j
                Else
                    TheExec.Datalog.WriteComment "Domain: " & powerDomain & "," & " isn't defined in PinMap for SyncUp_DCVS_Output. Error!!!"
                    TheExec.ErrorLogMessage "Domain: " & powerDomain & "," & " isn't defined in PinMap for SyncUp_DCVS_Output. Error!!!"
                End If
            End If
        Next i
    Else
        TheExec.Datalog.WriteComment "The argument powerGroup for SyncUp_DCVS_Output is empty. Error!!!"
        'TheExec.ErrorLogMessage "The argument powerGroup for SyncUp_DCVS_Output is empty. Error!!!"
    End If
    
    '''***********************************************************************************************************'''
    '''//tlDCVSVoltageOutput: This property gets or sets the output DAC used to force voltage (main or alternate).//
    '''//It can detect DCVS output Vmain/Valt of the binning CorePower powerDomains
    ''' The returned information is one of the following values:
    ''' tlDCVSVoltageMain: 1 ==> Main output voltage DAC.
    ''' tlDCVSVoltageAlt: 2 ==> Alternate output voltage DAC.
    '''***********************************************************************************************************'''
    If PinGroup <> "" Then
        If currentDcvsOutput = tlDCVSVoltageAlt Then
            TheHdw.DCVS.Pins(PinGroup).Voltage.output = tlDCVSVoltageAlt
        Else
            TheHdw.DCVS.Pins(PinGroup).Voltage.output = tlDCVSVoltageMain
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of SyncUp_DCVS_Output"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
'20201230: Modified to check powerDomain and powerPin by dictionaries of Domain2Pin and Pin2Domain.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20190625: Modified to replace the hard-code with the global variable "pinGroup_BinCut".
'20190617: Modified to use siteDouble "CorePowerStored" to save/restore voltages for BinCut powerPins.
Public Function save_core_power_vddbinning(CorePowerStored() As SiteDouble)
    Dim i As Long
    Dim powerPin As String
On Error GoTo errHandler
    '''//pinGroup_BinCut is created after initVddBinCondition (parsing "Non_Binning_Rail")
    '''It contains the pin names and sequence of corePower and otherRail.
    For i = 0 To UBound(pinGroup_BinCut)
        '''init
        powerPin = ""
    
        '''//As per discussion with SWLINZA, ZHHUANG, and PCLIN. We decided to measure the voltage of 1st powerPin of each power domain.
        '''//Check powerDomain and powerPin by dictionaries of Domain2Pin and Pin2Domain.
        If domain2pinDict.Exists(UCase(pinGroup_BinCut(i))) = True Then
            powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(pinGroup_BinCut(i)))
        ElseIf pin2domainDict.Exists(UCase(pinGroup_BinCut(i))) = True Then
            powerPin = UCase(pinGroup_BinCut(i))
        Else
            powerPin = ""
            TheExec.Datalog.WriteComment pinGroup_BinCut(i) & ", it is not BinCut powerDomain or pinGroup for save_core_power_vddbinning. Error!!!"
            TheExec.ErrorLogMessage pinGroup_BinCut(i) & ", it is not BinCut powerDomain or pinGroup for save_core_power_vddbinning. Error!!!"
        End If
        
        '''//If powerPin exists, read DCVS Vmain value.
        If powerPin <> "" Then
            For Each site In TheExec.sites
                '''20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
'                If Flag_Enable_Rail_Switch = True Then
                    CorePowerStored(i) = TheHdw.DCVS.Pins(powerPin).Voltage.Alt.Value
'                Else
'                    CorePowerStored(i) = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
'                End If
            Next site
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of save_core_power_vddbinning"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
'20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
'20200731: Modified to reset hexvsPingroup and nonhexvsPingroup for each powerDomain.
'20200210: Modified to check UltraFlex and UltraFlexPlus.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20190624: Modified to replace split_power_group with the public variable "pinGroup_BinCut".
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerPins.
Public Function restore_core_power_vddbinning(CorePowerStored() As SiteDouble)
    Dim i As Long, j As Long
    Dim split_content() As String
    Dim hexvsPingroup As String
    Dim nonhexvsPingroup As String
    Dim PinGroup As String
    Dim powerDomain As String
    Dim powerPin As String
On Error GoTo errHandler
    '''//pinGroup_BinCut is created after initVddBinCondition (parsing "Non_Binning_Rail")
    '''It contains the pin names and sequence of core_power and other rail.
    For i = 0 To UBound(pinGroup_BinCut)
        '''//init
        hexvsPingroup = ""
        nonhexvsPingroup = ""
        PinGroup = ""
        powerDomain = ""
        powerPin = ""
        
        '''//Get powerPins from powerDomain.
        powerDomain = UCase(pinGroup_BinCut(i))
        PinGroup = VddbinDomain2Pin(powerDomain)
        split_content = Split(PinGroup, ",")
        
        '''//Assembly temporary pinGroup for HexVs
        For j = 0 To UBound(split_content)
            powerPin = UCase(Trim(split_content(j)))
            
            '''20200210: Modified to check UltraFlex and UltraFlexPlus.
            If glb_TesterType = "Jaguar" Then
                If LCase(VddbinPinDcvsType(powerPin)) Like "hexvs" Then '''for HexVs
                    If hexvsPingroup = "" Then
                        hexvsPingroup = powerPin
                    Else
                        hexvsPingroup = hexvsPingroup & "," & powerPin
                    End If
                Else '''for non-HexVs
                    If nonhexvsPingroup = "" Then
                        nonhexvsPingroup = powerPin
                    Else
                        nonhexvsPingroup = nonhexvsPingroup & "," & powerPin
                    End If
                End If
            ElseIf glb_TesterType = "UltraFLEXplus" Then
                '''//Since all UltraFlexPlus DCVS instruments support siteAwareValue, we can set all powerPins in the same HexVS group.
                If hexvsPingroup = "" Then
                    hexvsPingroup = powerPin
                Else
                    hexvsPingroup = hexvsPingroup & "," & powerPin
                End If
                nonhexvsPingroup = ""
            Else
                TheExec.Datalog.WriteComment "Tester type is not UltraFlex or UltraFlexPlus. Please check this for restore_core_power_vddbinning. Error!!!"
                TheExec.ErrorLogMessage "Tester type is not UltraFlex or UltraFlexPlus. Please check this for restore_core_power_vddbinning. Error!!!"
            End If
        Next j
        
        If hexvsPingroup <> "" Then
            '''20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
            '''20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
'            If Flag_Enable_Rail_Switch = True Then
                TheHdw.DCVS.Pins(hexvsPingroup).Voltage.Alt.ValuePerSite = CorePowerStored(i) '''unit:V
'            Else
'                TheHdw.DCVS.Pins(hexvsPingroup).Voltage.Alt.ValuePerSite = CorePowerStored(i) '''unit:V
'            End If
        End If
        
        If nonhexvsPingroup = "" Then
            For Each site In TheExec.sites
                '''20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
'                If Flag_Enable_Rail_Switch = True Then
                    TheHdw.DCVS.Pins(nonhexvsPingroup).Voltage.Alt.Value = CorePowerStored(i)
'                Else
'                    TheHdw.DCVS.Pins(nonhexvsPingroup).Voltage.Main.Value = CorePowerStored(i)
'                End If
            Next site
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of restore_core_power_vddbinning"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
'20200327: Modified to use the flag "Flag_Skip_ReApplyInitVolageToDCVS" to skip "set_core_power_vddbinning_VT".
'20200210: Modified to check UltraFlex and UltraFlexPlus.
'20200210: Modified to use siteAwareValue for HexVS.
'20200210: Modified to reduce code-complexity.
'20200114: Modified to check pinGroup by "VddbinDomain2Pin".
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20191007: Modified to remove the undefined coniditions of "LV" and "HV".
'20190624: Modified to replace split_power_group with the public variable "pinGroup_BinCut".
'20190619: Modified to get voltages for BinCut Init pattern with the dedicated DC category and selectors.
'20190606: Modified to add the argument "DcSpecsCategoryForInitPat as string" for Init patterns with the new test setting DC Specs.
'20190606: Modified to use "Flag_Enable_Rail_Switch" for reading safe voltages "_VRS_GLB" or "GLB" from "Global Specs".
'20190524: Modified to use "Flag_Read_SafeVoltage_from_DCspecs" for reading safe voltages from "DC Specs".
'20190521: Modified for getting VRS from DC category of the new "DC_Specs" sheets.
'20180523: Modified for getting BinCut INIT voltages (safe voltages) from sheet "Global Specs" for projects with Rail Switch.
Public Function set_core_power_vddbinning_VT(DC_Level As String, Optional DcSpecsCategoryForInitPat As String)
    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim main_power As String
    Dim gb As Double
    Dim powerDomain As String
    Dim powerPin As String
    Dim split_content() As String
    Dim voltage_Temp As New SiteDouble
    Dim PinGroup As String
    Dim PinTEMP As String
    Dim hexvsPingroup As String
    Dim nonhexvsPingroup As String
    Dim EnableApplyInitVolageToDCVS As Boolean
    '''For instance context
    Dim DCCategory As String
    Dim DCSelector As String
    Dim ACCategory As String
    Dim ACSelector As String
    Dim TimeSetSheet As String
    Dim EdgeSetSheet As String
    Dim LevelsSheet As String
    Dim Overlay As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//"Flag_Read_SafeVoltage_from_DCspecs" for safe voltages. True: from "DC Specs"; False: from "Global Specs".
'''Discussed this with T-autogen Jeff, we decided to follow the rule of assemblying voltage names (powerDomain + "_VOP_VAR") in sheet "DC Specs".
'''"VDD_XXX_VAR" is applied to Vmain for Init patterns.
'''If DcSpecsCategoryForInitPat has the string, it uses the DcSpecsCategoryForInitPat for Init patterns.
'''If DcSpecsCategoryForInitPat is empty, it just follows the default DC Specs category.
'''//==================================================================================================================================================================================//'''
    Select Case DC_Level
        Case "NV": gb = 1
        Case Else:
            '''20191007: Modified to remove the undefined coniditions of "LV" and "HV".
            TheExec.Datalog.WriteComment "DC_Level:" & DC_Level & ", it is not defined for set_core_power_vddbinning_VT!!! Error!!!"
            TheExec.ErrorLogMessage "DC_Level:" & DC_Level & ", it is not defined for set_core_power_vddbinning_VT!!! Error!!!"
    End Select
      
    '''**********************************************************************************************************'''
    '''Note: Safe voltages for init Patt usually use the same DC category.
    If Flag_Skip_ReApplyInitVolageToDCVS = True Then
        Call TheExec.DataManager.GetInstanceContext(DCCategory, DCSelector, ACCategory, ACSelector, TimeSetSheet, EdgeSetSheet, LevelsSheet, Overlay)
        
        If DcSpecsCategoryForInitPat <> "" Then
            If LCase(DCCategory) = LCase(DcSpecsCategoryForInitPat) Then
                EnableApplyInitVolageToDCVS = False
            Else
                EnableApplyInitVolageToDCVS = True
            End If
        Else '''DcSpecsCategoryForInitPat is empty.
            '''//If initial voltages and safe voltage(init voltage) use the same DC category, it can skip "set_core_power_vddbinning_VT" after initial voltages...
            '''Note: For PTE/TTR, it can use the flag "Flag_Skip_ReApplyInitVolageToDCVS" to skip "set_core_power_vddbinning_VT".
            If IsLevelLoadedForApplyLevelsTiming = True Then
                EnableApplyInitVolageToDCVS = False
            Else
                EnableApplyInitVolageToDCVS = True
            End If
        End If
    Else
        EnableApplyInitVolageToDCVS = True
    End If
    '''**********************************************************************************************************'''
    
    If EnableApplyInitVolageToDCVS = True Then
        '''//pinGroup_BinCut is created after initVddBinCondition (parsing "Non_Binning_Rail")
        '''//It contains the pin names and sequence of CorePower and OtherRail.
        For i = 0 To UBound(pinGroup_BinCut)
            powerDomain = UCase(Trim(pinGroup_BinCut(i)))
            
            If (DC_Level = "NV") Then
                '''//init
                hexvsPingroup = ""
                nonhexvsPingroup = ""
                
                '''//Get powerPins from powerDomain
                PinGroup = VddbinDomain2Pin(powerDomain)
                split_content = Split(PinGroup, ",")
                
                '''//Assembly temporary pinGroup for HexVs
                For j = 0 To UBound(split_content)
                    PinTEMP = UCase(Trim(split_content(j)))

                    '''20200210: Modified to check UltraFlex and UltraFlexPlus.
                    If glb_TesterType = "Jaguar" Then
                        If LCase(VddbinPinDcvsType(PinTEMP)) Like "hexvs" Then '''for HexVs
                            If hexvsPingroup = "" Then
                                hexvsPingroup = PinTEMP
                            Else
                                hexvsPingroup = hexvsPingroup & "," & PinTEMP
                            End If
                        Else '''for non-HexVs
                            If nonhexvsPingroup = "" Then
                                nonhexvsPingroup = PinTEMP
                            Else
                                nonhexvsPingroup = nonhexvsPingroup & "," & PinTEMP
                            End If
                        End If
                    ElseIf glb_TesterType = "UltraFLEXplus" Then
                        '''//Since all UltraFlexPlus DCVS instruments support siteAwareValue, we can set all powerPins in the same HexVS group.
                        If hexvsPingroup = "" Then
                            hexvsPingroup = PinTEMP
                        Else
                            hexvsPingroup = hexvsPingroup & "," & PinTEMP
                        End If
                        nonhexvsPingroup = ""
                    Else
                        TheExec.Datalog.WriteComment "Tester type is not UltraFlex or UltraFlexPlus. Please check this for set_core_power_vddbinning_VT. Error!!!"
                        TheExec.ErrorLogMessage "Tester type is not UltraFlex or UltraFlexPlus. Please check this for set_core_power_vddbinning_VT. Error!!!"
                    End If
                Next j

                '''//Get 1st powerPin from powerDomain
                powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(pinGroup_BinCut(i)))
                
                If Flag_Read_SafeVoltage_from_DCspecs Then '''Get safe voltages " "VDD_***_VAR" " from "DC Specs".
                    If DcSpecsCategoryForInitPat <> "" Then '''for Mbist
                        voltage_Temp = Floor(TheExec.specs.DC.Item(powerPin & "_VAR").Categories.Item(DcSpecsCategoryForInitPat).Selectors.Item("typ").ContextValue * gb * 1000) / 1000
                    Else
                        voltage_Temp = Floor(TheExec.specs.DC.Item(powerPin & "_VAR").ContextValue * gb * 1000) / 1000
                    End If
                Else '''Get safe voltages "VDD_***_VRS_GLB" from "Global Specs".
'                    If Flag_Enable_Rail_Switch Then
                        voltage_Temp = Floor(TheExec.specs.Globals(powerPin & "_VRS" & "_GLB").ContextValue * gb * 1000) / 1000
'                    Else
'                        voltage_Temp = Floor(TheExec.specs.Globals(powerPin & "_GLB").ContextValue * gb * 1000) / 1000
'                    End If
                End If
                    
                '''20200210: Modified to use siteAwareValue for HexVS.
                '''20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
                If hexvsPingroup <> "" Then
                    TheHdw.DCVS.Pins(hexvsPingroup).Voltage.Main.ValuePerSite = voltage_Temp
                End If
                
                If nonhexvsPingroup <> "" Then
                    For Each site In TheExec.sites
                        TheHdw.DCVS.Pins(nonhexvsPingroup).Voltage.Main.Value = voltage_Temp
                    Next site
                End If
            Else
                '''20191007: Modified to remove the undefined coniditions of "LV" and "HV".
                TheExec.Datalog.WriteComment "DC_Level:" & DC_Level & ", it is not defined to get DC Levels for set_core_power_vddbinning_VT!!! Error!!!"
                TheExec.ErrorLogMessage "DC_Level:" & DC_Level & ", it is not defined to get DC Levels for set_core_power_vddbinning_VT!!! Error!!!"
            End If
        Next i
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of set_core_power_vddbinning_VT"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210906: Modified to merge the branches of the vbt function Parsing_Instance_Pmode.
'20210809: Modified to revise the vbt code to get main_performance_mode and additional_mode.
'20201230: Modified to revise the vbt code for the new naming rule of OutsideBinCut performance mode, ex: "MS004_TD_SOC_MHV".
'20200211: Modified to replace the function name "FlowTestCondStr2Enum" with "AdditionalModeStr2Enum".
'20200205: Modified to check if the input "Performance_mode" exists in the dictionary.
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20190510: Modified to merge "powerDomain = AllBinCut(p_mode).powerPin" into Parse_Performance_Mode
'20190508: Created for parsing the performance mode and getting addi_mode.
Public Function Parsing_Instance_Pmode(performance_mode As String, p_mode As Integer, addi_mode As Integer, special_voltage_setup As Boolean)
    Dim i As Integer
    Dim j As Integer
    Dim split_content() As String
    Dim strTemp As String
    Dim main_performance_mode As String
    Dim additional_mode As String
On Error GoTo errHandler
    '''init
    p_mode = 0
    addi_mode = 0
    special_voltage_setup = False
    main_performance_mode = ""
    additional_mode = ""
    
    '''//OutsideBinCut might use the new naming rules for performance modes, ex: "MS004_TD_SOC_MHV".
    split_content = Split(UCase(Trim(performance_mode)), "_")
    
    For i = 0 To UBound(split_content)
        If Trim(split_content(i)) Like UCase("m*##*") Then
            If VddbinPmodeDict.Exists(UCase(Trim(split_content(i)))) Then
                main_performance_mode = UCase(Trim(split_content(i)))
                
                '''//Check if additional_mode exists...
                If i = UBound(split_content) Then '''without additional_mode
                    special_voltage_setup = False
                Else '''with additional_mode
                    special_voltage_setup = True
                    
                    '''//Assembly the string for additional_mode.
                    additional_mode = UCase(Trim(split_content(i + 1)))
                    
                    If (i + 1) < UBound(split_content) Then
                        For j = i + 2 To UBound(split_content)
                            additional_mode = additional_mode & "_" & UCase(Trim(split_content(j)))
                        Next j
                    End If
                End If
                
                '''//Once if it gets 1st keyword with performance mode, exit loop.
                Exit For
            End If
        End If
    Next i
    
    '''//Check if it gets the performance mode.
    '''20210906: Modified to merge the branches of the vbt function Parsing_Instance_Pmode.
    If main_performance_mode <> "" Then
        '''//Check p_mode for the main_performance_mode.
        If VddbinPmodeDict.Exists(UCase(main_performance_mode)) Then
            p_mode = VddBinStr2Enum(main_performance_mode)
        Else
            TheExec.Datalog.WriteComment performance_mode & ", it wasn't the performance mode defined in BinCut voltage table for Parsing_Instance_Pmode. Error!!!"
            TheExec.ErrorLogMessage performance_mode & ", it wasn't the performance mode defined in BinCut voltage table for Parsing_Instance_Pmode. Error!!!"
            Exit Function
        End If
    Else '''If main_performance_mode = "" Then
        TheExec.Datalog.WriteComment performance_mode & ", it doesn't have the correct format of the performance mode for Parsing_Instance_Pmode. Error!!!"
        TheExec.ErrorLogMessage performance_mode & ", it doesn't have the correct format of the performance mode for Parsing_Instance_Pmode. Error!!!"
        Exit Function
    End If
    
    '''//Check addi_mode for the additional_mode.
    If special_voltage_setup = True Then
        If AdditionalModeDict.Exists(additional_mode) Then
            addi_mode = AdditionalModeStr2Enum(additional_mode)
        Else
            TheExec.Datalog.WriteComment performance_mode & ", it doesn't have the correct additional mode for Parsing_Instance_Pmode. Error!!!"
            TheExec.ErrorLogMessage performance_mode & ", it doesn't have the correct additional mode for Parsing_Instance_Pmode. Error!!!"
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_Instance_Pmode"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210730: Modified to show the error message to users if the current testJob is for BinCut search without any Efuse category "Product_Identifier", as requested by C651 Toby.
'20210707: Modified to merge the branches of checking the keyword in Efuse category for the vbt function Parsing_IDSname_from_BDF_Table.
'20210707: Modified to add the special case "product_identifier_cp1" of Efuse_BitDef_Table for the vbt function Parsing_IDSname_from_BDF_Table.
'20210707: Modified to update the string array to dict_strPmode2EfuseCategory.
'20210705: Modified to revise the vbt code to parse Efuse category with the prefix "ids_...".
'20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
'20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
'20210703: Modified to revise the vbt code of parsing Efuse_BitDef_Table.
'20210702: Modified to check column "Default or Real" in Efuse_BitDef_Table.
'20210701: Modified to revise the vbt code for BinCut search in FT.
'20210701: Modified to update AllBinCut(p_mode).listed_in_Efuse_BDF in the vbt function Parsing_IDSname_from_BDF_Table.
'20210617: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si.
'20210121: Modified to check Harvest 10 cores.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20200817: Modified to use the variable "str_alg_temp".
'20200731: Modified to get Efuse IDS name for CP1 only.
'20200731: Modified to merge MappingBincutJobName and Mapping_TPJobName_to_BincutJobName into Mapping_TestJobName_to_BincutJobName.
'20200730: Modified to check IDS_VDD_XXX_BINCHECK for Harvest powerPin.
'20200730: Modified to check if VDD_XXX__M*### is the correct BinCut performance mode.
'20200729: Modified to check if IDS name from the column "bank_config eFuse Bit Def" belongs to BinCut powerDomain by BinCut testJob.
'20200703: Modiifed to use "check_Sheet_Range".
'20200410: Modified to check "Base Voltage" in Efuse_BitDef_Table and Vdd_Binning_Def.
'20200114: Modified to check if powerDomain exists in domain2pinDict or pin2domainDict.
'20191015: Modified to use MaxSiteCount array for ids_name.
'20190710: Modified to check if "VDD_***_M*###" is listed in BinCut power_seq.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerPins.
'20190521: Modified for mapping IDS names to current BinCut testjob.
'20190514: Created for parsing the sheet "EFUSE_BitDef_Table" to get IDS names, especially for testjobs.
Public Function Parsing_IDSname_from_BDF_Table()
    Dim site As Long
    '''
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim MaxRow As Long, maxcol As Long
    Dim isSheetFound As Boolean
    '''
    Dim row As Long
    Dim col As Long
    Dim row_of_title As Long
    Dim col_Efuse_category As Long
    Dim col_stage As Long
    Dim col_alg As Long
    Dim col_defaultValue As Long
    Dim col_default_Real As Long
    Dim enableRowParsing As Boolean
    '''
    Dim k As Long
    Dim idx_powerDomain As Integer
    Dim powerDomain As String
    Dim str_category_temp As String
    Dim str_stage_temp As String
    Dim str_alg_temp As String
    Dim str_pmode_temp As String
    Dim str_default_Real As String
    Dim str_ids_name As String
    Dim str_testJob As String '''Mapping_TestJobName_to_BincutJobName
    Dim idx_testJob As Long
    Dim strAry_EfuseCategory() As String
    Dim idx_EfuseCategory As Long
    Dim str_Efuse_write_ProductIdentifier As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Caution!!!
'''1. Remember to check the column "Programming Stage" for BinCut testJob mapping.(check the function "MappingBincutJobName")
'''2. Remember to check core number of the Harvest powerPin in TestPlan and "EFUSE_BitDef_Table"!!! => ex: ids_vdd_sram_gpu_7, ids_vdd_sram_gpu_8. "7" and "8" are core numbers of Harvest powerPin.
'''3. Remember to check the format of test temperature. => ex: ids_vdd_soc_105. "105" is the test temperature.
'''4. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "EFUSE_BitDef_Table"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//Init
        dict_strPmode2EfuseCategory.RemoveAll
        dict_EfuseCategory2BinCutTestJob.RemoveAll
        
        '''//Clear the content of the globalVariable "IDS_for_BinCut" before parsing the sheet "EFUSE_BitDef_Table"
        For idx_powerDomain = 0 To UBound(pinGroup_BinCut)
            powerDomain = pinGroup_BinCut(idx_powerDomain)
            For site = 0 To MaxSiteCount - 1
                IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site) = ""
            Next site
        Next idx_powerDomain
    Else
        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", it doesn't exist in this IGXL workbook. Error!!!"
        TheExec.ErrorLogMessage "sheet:" & sheetName & ", it doesn't exist in this IGXL workbook. Error!!!"
        Exit Function
    End If '''If isSheetFound = True
    
    '''//Parse the header of the table.
    For row = 1 To MaxRow
        For col = 1 To maxcol
            '''//The header should include "bank_config eFuse Bit Def", "Programming Stage", and "Algorithm".
            If LCase(Trim(ws_def.Cells(row, col).Value)) Like LCase("bank_config eFuse Bit Def") Then
                col_Efuse_category = col
                row_of_title = row
            End If
            
            If row_of_title > 0 Then
                '''20210702: Modified to check column "Default or Real" in Efuse_BitDef_Table.
                If LCase(Trim(ws_def.Cells(row_of_title, col).Value)) Like LCase("Programming Stage") Then '''ex: "CP1", "FT1", "WLFT1".
                    col_stage = col
                ElseIf LCase(Trim(ws_def.Cells(row_of_title, col).Value)) Like LCase("Algorithm") Then '''ex: "app", "ids", "vddbin".
                    col_alg = col
                ElseIf LCase(Trim(ws_def.Cells(row_of_title, col).Value)) Like LCase("Default or Real") Then '''ex: "bincut", "Default", "Real".
                    col_default_Real = col
                ElseIf LCase(Trim(ws_def.Cells(row_of_title, col).Value)) Like LCase("Default Value") Then
                    col_defaultValue = col
                End If
            End If
        Next col
        
        '''//If row the header is found, parse each row of the table.
        If row_of_title > 0 Then
            If col_Efuse_category > 0 And col_stage > 0 And col_alg > 0 Then
                enableRowParsing = True
                Exit For
            Else
                enableRowParsing = False
                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", it doesn't have the correct header for Parsing_IDSname_from_BDF_Table. Error!!!"
                TheExec.ErrorLogMessage "sheet:" & sheetName & ", it doesn't have the correct header for Parsing_IDSname_from_BDF_Table. Error!!!"
                Exit For
            End If
        End If
    Next row

    If enableRowParsing = True Then
        For row = row_of_title + 1 To MaxRow
            '''*************************************************************************************************************************************'''
            '''//Get Efuse category, algorithm, Default/Real, and Programming Stage for each row from Efuse_BitDef_Table.
            '''*************************************************************************************************************************************'''
            '''//Efuse category.
            str_category_temp = UCase(Trim(ws_def.Cells(row, col_Efuse_category).Value))
            
            '''//Algorithm.
            str_alg_temp = UCase(Trim(ws_def.Cells(row, col_alg).Value))
            
            '''//Default or Real.
            str_default_Real = UCase(Trim(ws_def.Cells(row, col_default_Real).Value))
            
            '''//Get testJob from the column "Programming Stage".
            '''Mapping_TestJobName_to_BincutJobName
            str_stage_temp = UCase(Trim(ws_def.Cells(row, col_stage).Value))
            
            '''*****************************************************************************************************************************************'''
            '''//Check the Efuse category from the column "bank_config eFuse Bit Def" in Efuse_BitDef_Table.
            '''//If "Algorithm" = "vddbin" and Efuse category with the keyword about BinCut performance mode "VDD_XXX__M*###", the item must be Efuse product voltage.
            '''//If "Algorithm" = "ids" and Efuse category with the prefix "ids_", check if it contains the keyword about BinCut powerDomain.
            '''//If "Algorithm" = "app" and Efuse category "product_identifier", the item must be Efuse product_identifier.
            '''//If "Algorithm" = "app" and Efuse category with the keyword "power_binning", the item must be Efuse power_binning.
            '''//If "Algorithm" = "base" and Efuse category with the keyword "*base*voltage*", the item must be Efuse base voltage.
            '''*****************************************************************************************************************************************'''
            If (LCase(str_category_temp) Like "vdd*" And LCase(str_alg_temp) = "vddbin" And LCase(str_default_Real) = "bincut") _
            Or (((LCase(str_category_temp) Like "product_identifier*" Or LCase(str_category_temp) Like "power_binning*") And LCase(str_alg_temp) = "app")) Then
                '''init
                str_pmode_temp = ""
                powerDomain = ""
                               
                '''//Check if the Efuse category for Efuse product voltages with keyword “*_shadow*”, ex: "vdd_gpu_mg001_shadow".
                If LCase(str_category_temp) Like LCase("*_shadow") Then
                    str_pmode_temp = UCase(Replace(UCase(str_category_temp), UCase("_shadow"), ""))
                ElseIf LCase(str_category_temp) Like LCase("*_cp1") Then '''ex: "product_identifier_cp1".
                    str_pmode_temp = UCase(Replace(UCase(str_category_temp), UCase("_cp1"), ""))
                Else
                    str_pmode_temp = UCase(str_category_temp)
                End If
                
                '''//Check if the keyword "VDD_XXX__M*###" in Efuse category is the correct BinCut performance mode.
                If LCase(str_pmode_temp) Like "vdd*" Then
                    If VddbinPmodeDict.Exists(str_pmode_temp) = True Then
                        powerDomain = AllBinCut(VddBinStr2Enum(str_pmode_temp)).powerPin
                        
                        '''//Remove the prefix "vdd_" from str_pmode_temp.
                        If LCase(str_pmode_temp) Like LCase("vdd_*") Then '''ex: "vdd_gpu_mg001".
                            str_pmode_temp = UCase(Replace(UCase(str_pmode_temp), UCase(powerDomain & "_"), ""))
                        End If
                        
                        '''//Check if str_pmode_temp exists in gb_bincut_power_list.
                        If UCase("*," & gb_bincut_power_list(VddBinStr2Enum(powerDomain)) & ",*") Like UCase("*," & str_pmode_temp & ",*") Then
                            str_pmode_temp = VddBinName(VddBinStr2Enum(str_pmode_temp)) '''ex: "VDD_PCPU_MP001"
                        Else
                            str_pmode_temp = ""
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it doesn't contain any performance mode listed in BinCut flow table. Error!!!"
                            TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it doesn't contain any performance mode listed in BinCut flow table. Error!!!"
                        End If
                    Else
                        str_pmode_temp = ""
                    End If '''If VddbinPmodeDict.Exists(str_pmode_temp) = True
                End If
                
                '''//Store Efuse category for BinCut p_mode into the dictionary dict_strPmode2EfuseCategory.
                '''20210701: Modified to update AllBinCut(p_mode).listed_in_Efuse_BDF in the vbt function Parsing_IDSname_from_BDF_Table.
                '''20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
                '''20210707: Modified to update the string array to dict_strPmode2EfuseCategory.
                If str_pmode_temp <> "" Then
                    If dict_strPmode2EfuseCategory.Exists(str_pmode_temp) = True Then
                        '''init
                        strAry_EfuseCategory = dict_strPmode2EfuseCategory.Item(str_pmode_temp)
                        
                        '''//Check if Efuse category exists in dict_strPmode2EfuseCategory for BinCut p_mode.
                        For idx_EfuseCategory = 0 To UBound(strAry_EfuseCategory)
                            '''//Check if the Efuse category has the duplicate item in Efuse_BitDef_Table.
                            If LCase(str_category_temp) = LCase(strAry_EfuseCategory(idx_EfuseCategory)) Then
                                str_category_temp = ""
                                Exit For
                            End If
                        Next idx_EfuseCategory
                        
                        '''//Efuse category has no conflict, then store this into dict_strPmode2EfuseCategory.
                        If str_category_temp <> "" Then
                            idx_EfuseCategory = UBound(strAry_EfuseCategory) + 1
                            ReDim Preserve strAry_EfuseCategory(idx_EfuseCategory) As String
                            '''ToDo: Maybe we can do bubble sorting for the sequence of Efuse category from dict_strPmode2EfuseCategory...
                            strAry_EfuseCategory(idx_EfuseCategory) = str_category_temp
                            dict_strPmode2EfuseCategory.Item(str_pmode_temp) = strAry_EfuseCategory
                            '''//Updated the property "AllBinCut(p_mode).listed_in_Efuse_BDF" for BinCut p_mode.
                            If VddbinPmodeDict.Exists(str_pmode_temp) = True Then
                                AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF = True
                            End If
                        Else
                            str_category_temp = ""
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it has the duplicate Efuse category in sheet:" & sheetName & ". Error!!!"
                            TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it has the duplicate Efuse category in sheet:" & sheetName & ". Error!!!"
                        End If
                    Else
                        ReDim strAry_EfuseCategory(0) As String
                        strAry_EfuseCategory(0) = str_category_temp
                        dict_strPmode2EfuseCategory.Add str_pmode_temp, strAry_EfuseCategory
                        '''//Updated the property "AllBinCut(p_mode).listed_in_Efuse_BDF" for BinCut p_mode.
                        If VddbinPmodeDict.Exists(str_pmode_temp) = True Then
                            AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF = True
                        End If
                    End If
                Else
                    str_category_temp = ""
                End If '''If str_pmode_temp <> ""
                
            '''//If "Algorithm" = "ids" and Efuse category with the prefix "ids_", check if it contains the keyword about BinCut powerDomain.
            ElseIf (LCase(str_category_temp) Like "ids*" And LCase(str_alg_temp) = "ids") And LCase(str_default_Real) = "real" Then '''ex: "ids_vdd_pcpu", "ids_vdd_ecpu", "ids_vdd_cpu_sram", "ids_vdd_gpu_5", "ids_vdd_sram_gpu_5".
                '''init
                str_ids_name = ""
                powerDomain = ""
            
                '''//Check the column "Programming Stage" for BinCut testJob mapping.
                '''*****************************************************************************************************************************************'''
                '''20210617: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si.
                '''So that all BinCut testJobs use Efuse IDS values that are fused in CP1!!!
                '''ToDo: Please discuss this with project Efuse owner to see if rules about "Programming Stage" in Efuse_BitDef_Table are changed.
                '''*****************************************************************************************************************************************'''
                '''20210707: Modified to parse Efuse category with the keyword "ids_*" while the programming stage is "cp1".
                If LCase(str_stage_temp) = "cp1" Then
                    str_ids_name = UCase(Replace(LCase(str_category_temp), "ids_", ""))
                Else
                    str_ids_name = ""
                End If
                
                '''//Check if str_category_temp about IDS name from Efuse category is not empty.
                If str_ids_name <> "" Then
                    '''//Remove string about core number of the Harvest powerPin.
                    '''*****************************************************************************************************************************************'''
                    '''For Harvest powerPin, IDS_name will be defined in "VBT_LIB_DC_IDS\IDS_eFuse_Write_"
                    '''Note: Remember to check core number of the Harvest powerPin in TestPlan and "EFUSE_BitDef_Table"!!!
                    '''ex: ids_vdd_sram_gpu_7, ids_vdd_sram_gpu_8. "7" and "8" are core numbers of Harvest powerPin.
                    '''*****************************************************************************************************************************************'''
                    '''ToDo: Remember to check string about the core number of Harvest powerPin...
                    If LCase(str_ids_name) Like "*_3" Or LCase(str_ids_name) Like "*_4" Or LCase(str_ids_name) Like "*_5" _
                    Or LCase(str_ids_name) Like "*_7" Or LCase(str_ids_name) Like "*_8" Or LCase(str_ids_name) Like "*_9" Or LCase(str_ids_name) Like "*_10" Then
                        str_ids_name = UCase(Mid(str_ids_name, 1, Len(str_ids_name) - 2))
                    ElseIf LCase(str_ids_name) Like "*_bincheck" Then
                        str_ids_name = UCase(Replace(Replace(LCase(str_ids_name), "ids_", ""), "_bincheck", ""))
                    End If
                    
                    '''//Remove string of the test temperature.
                    '''*****************************************************************************************************************************************'''
                    '''Remember to check the format of test temperature, ex: ids_vdd_soc_105. "105" is the test temperature.
                    '''*****************************************************************************************************************************************'''
                    If LCase(str_ids_name) Like "*_25" Or LCase(str_ids_name) Like "*_25_*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_25", "")) '''25
                    ElseIf LCase(str_ids_name) Like "*_25c*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_25c", "")) '''25C
                    ElseIf LCase(str_ids_name) Like "*_85" Or LCase(str_ids_name) Like "*_85_*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_85", "")) '''85
                    ElseIf LCase(str_ids_name) Like "*_85c*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_85c", "")) '''85C
                    ElseIf LCase(str_ids_name) Like "*_105" Or LCase(str_ids_name) Like "*_105_*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_105", "")) '''105
                    ElseIf LCase(str_ids_name) Like "*_105c*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_105c", "")) '''105C
                    End If
                    
                    '''//Check if the string contains the keyword about powerDomain.
                    If domain2pinDict.Exists(UCase(str_ids_name)) = True Then
                        powerDomain = UCase(str_ids_name)
                    ElseIf pin2domainDict.Exists(UCase(str_ids_name)) = True Then
                        powerDomain = UCase(VddbinPin2Domain(str_ids_name))
                    Else
                        powerDomain = ""
                        '''ToDo: Maybe we can add the utility to show warning here...
                    End If
                Else
                    powerDomain = ""
                End If
                
                '''//If str_domain_temp of IDS name is not empty, check if powerDomain belongs to BinCut powerDomains.
                If powerDomain <> "" Then
                    If VddbinPmodeDict.Exists(powerDomain) = True Then
                        idx_powerDomain = VddBinStr2Enum(powerDomain)

                        '''***********************************************************************************************'''
                        '''//Check if IDS_for_BinCut(idx_PowerDomain).IDS_name of BinCut powerDomain is already defined...
                        '''***********************************************************************************************'''
                        For site = 0 To MaxSiteCount - 1
                            If IDS_for_BinCut(idx_powerDomain).ids_name(site) <> "" Then
                                '''//Check IDS_VDD_XXX_BINCHECK for Harvest powerPin.
                                If LCase(IDS_for_BinCut(idx_powerDomain).ids_name(site)) Like "*bincheck*" Then
                                    '''Do nothing, keep IDS_VDD_XXX_BINCHECK as IDS name for Harvest powerPin.
                                Else
                                    IDS_for_BinCut(idx_powerDomain).ids_name(site) = str_category_temp
                                End If
                            Else '''If IDS_for_BinCut(idx_PowerDomain).IDS_name is empty...
                                IDS_for_BinCut(idx_powerDomain).ids_name(site) = str_category_temp
                            End If
                        Next site
                    Else
                        str_category_temp = ""
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it doesn't get any correct IDS name of BinCut powerPin while Parsing_IDSname_from_BDF_Table. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it doesn't get any correct IDS name of BinCut powerPin while Parsing_IDSname_from_BDF_Table. Error!!!"
                    End If
                    
                    '''//Since str_category_temp with the keyword IDS is useful, this can be stored with "Programming stage" into dict_EfuseCategory2BinCutTestJob
                End If '''If str_domain_temp <> ""
                
            '''//If "Algorithm" = "base" and Efuse category with the keyword "*base*voltage*", check if the default value is same as "Base Voltage" in sheet Vdd_Binning_Def.
            ElseIf LCase(str_category_temp) Like "*base*voltage" And LCase(str_alg_temp) = "base" Then
                If LCase(Trim(ws_def.Cells(row, col_alg).Value)) = "base" Then
                    If ws_def.Cells(row, col_defaultValue).Value <> "" Then
                        BaseVoltageFromEfuseBDF = CDbl(ws_def.Cells(row, col_defaultValue).Value)

                        '''//Compare BaseVoltage values from from "Vdd_Binning_Def" and "EFUSE_BitDef_Table".
                        If BaseVoltageFromEfuseBDF <> VddbinningBaseVoltage Then
                            TheExec.Datalog.WriteComment "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", Base_Voltage=" & BaseVoltageFromEfuseBDF & " is inconsistent with the value in header of Vdd_Binning_Def tables. Error!!!"
                            TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", Base_Voltage=" & BaseVoltageFromEfuseBDF & " is inconsistent with the value in header of Vdd_Binning_Def tables. Error!!!"
                        End If
                    Else
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ",row:" & row & ", category:" & str_category_temp & ", Base_Voltage is undefined in Efuse_BitDef_Table. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & row & ",category:" & str_category_temp & ", Base_Voltage is undefined in Efuse_BitDef_Table. Error!!!"
                    End If
                    
                    '''//No need to store Efuse category with the keyword "*base*voltage*" into dict_EfuseCategory2BinCutTestJob.
                    str_category_temp = ""
                End If
            Else
                str_category_temp = ""
            End If '''If LCase(str_category_temp) Like...
            
            '''*****************************************************************************************************************************************'''
            '''//Store the programming stage by getBinCutJobDefinition if Efuse category is correct.
            '''*****************************************************************************************************************************************'''
            '''20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
            If str_category_temp <> "" Then
                '''//Mapping Efuse "Programming stage" to BinCut testJob.
                str_testJob = Mapping_TestJobName_to_BincutJobName(str_stage_temp)
                idx_testJob = getBinCutJobDefinition(str_testJob)
            
                If dict_EfuseCategory2BinCutTestJob.Exists(str_category_temp) = True Then
                    TheExec.Datalog.WriteComment "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it has the duplicate programming stage in sheet:" & sheetName & ". Error!!!"
                    TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & row & ",Efuse category:" & str_category_temp & ", it has the duplicate programming stage in sheet:" & sheetName & ". Error!!!"
                Else
                    dict_EfuseCategory2BinCutTestJob.Add str_category_temp, idx_testJob
                End If
            End If
        Next row
    End If '''If enableRowParsing = True Then
        
    '''//If the testJob is for BinCut search, it should have the dedicated "Product_Identifier", as commented by C651 Si and Toby, 20210727.
    '''//If the current testJob is for BinCut search without any Efuse category "Product_Identifier", show the error message to users, as requested by C651 Toby.
    '''20210730: Modified to show the error message to users if the current testJob is for BinCut search without any Efuse category "Product_Identifier", as requested by C651 Toby.
    If is_BinCutJob_for_StepSearch = True Then
        str_Efuse_write_ProductIdentifier = get_Efuse_category_by_BinCut_testJob("write", "Product_Identifier")
        
        If str_Efuse_write_ProductIdentifier = "" Then
            TheExec.Datalog.WriteComment "testJob:" & TheExec.CurrentJob & ", it is for BinCut search, but it doesn't have any Efuse category about Product_Identifier. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!"
            TheExec.ErrorLogMessage "testJob:" & TheExec.CurrentJob & ", it is for BinCut search, but it doesn't have any Efuse category about Product_Identifier. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!"
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Parsing_Efuse_BitDef_Table"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Parsing_Efuse_BitDef_Table"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
'20210707: Modified to check if ids_name (Efuse category) exists in dict_EfuseCategory2BinCutTestJob.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20200807: Modified to merge the redundant site-loop.
'20190813: Modified to use different IDS lo_limit by BinCut testjobs.
'20190630: Modified to show the error message when str_IDS_PowerDomain is empty.
'20190523: Created for the new data type "IDS_for_BinCut".
'20170810: SWLINZA modified the vbt code to Get Resolution and set it for low limit for OtherRail.
Public Function get_lo_limit_for_IDS(powerDomain As String, lo_limit As SiteDouble)
    Dim site As Variant
    Dim str_IDS_PowerDomain As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
'''2. As per discussion with TSMC SWLINZA, for powerPin group, it should use 1st powerPin to check IDS limit of powerPin group, 20210707.
'''ex: powerGroup: VDD_FIXED_GRP, and its 1st powerPin: VDD_FIXED, so that compare IDS value of VDD_FIXED with IDS_limit of VDD_FIXED_GRP. It must have Efuse category in Efuse_BitDef_Table to store IDS for VDD_FIXED.
'''3. C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs, 20210727.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        str_IDS_PowerDomain = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site)
        
        '''//If Efuse IDS name of powerDomain exists in Efuse_BitDef_Table, take Efuse IDS Resolution as IDS low limit for OtherRail.
        If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True Then
            '''For project with Efuse DSP vbt code.
            'Jeff lo_limit(site) = 1# * CFGFuse.Category(CFGIndex(str_IDS_PowerDomain)).Resoultion '''unit: mA
            lo_limit(site) = 1 * 0
        ElseIf Flag_VDD_Binning_Offline = True Or EnableWord_Vddbinning_OpenSocket = True Then '''If the tester is offline or opensocket.
            lo_limit(site) = 0
        Else
            lo_limit(site) = 0
            TheExec.Datalog.WriteComment powerDomain & " has no definition of the Efuse IDS resolution as IDS lo_limit. get_lo_limit_for_IDS has the error. Error!!!"
            TheExec.ErrorLogMessage powerDomain & " has no definition of the Efuse IDS resolution as IDS lo_limit. get_lo_limit_for_IDS has the error. Error!!!"
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of get_lo_limit_for_IDS"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191002: Created to init BincutVoltageTypeName for printing BV strings.
Public Function initBincutVoltageType()
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Please update Enum BincutVoltageType and MaxBincutVoltageType in GlobalVariable once if you want insert the new type!!!
'''//==================================================================================================================================================================================//'''
    BincutVoltageTypeName(BincutVoltageType.None) = ""
    BincutVoltageTypeName(BincutVoltageType.InitialVoltage) = "Initial Voltage"
    BincutVoltageTypeName(BincutVoltageType.SafeVoltage) = "Safe Voltage"
    BincutVoltageTypeName(BincutVoltageType.PayloadVoltage) = "Payload Voltage"
    BincutVoltageTypeName(BincutVoltageType.PostbincutBinningpower) = "Postbincut BinningPower_BinResult Voltage"
    BincutVoltageTypeName(BincutVoltageType.PostbincutAllpower) = "Payload AllPower_BinResult Voltage"
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of InitBincutVoltageType"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of InitBincutVoltageType"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191105: Created to init TestType names.
Public Function initTestTypeName()
On Error GoTo errHandler
    TestTypeName(testType.TD) = "TD"
    TestTypeName(testType.Mbist) = "Mbist"
    TestTypeName(testType.Func) = "Func"
    TestTypeName(testType.RTOS) = "RTOS"
    TestTypeName(testType.ldcbfd) = ""
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of InitTestTypeName"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of InitTestTypeName"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210113: Modified to check if cnt_DecomposedPinList>0.
'20200423: Modified to move "Dim dictPin2Dcspec As New Dictionary" into GlobalVariable.
'20191231: Modified to use TestJob of test program for DC Specs.
'20191219: Created for Domain2Pin and Pin2Domain.
Public Function initDomain2Pin(domainList As String, dictDomain2Pin As Dictionary, dictPin2Domain As Dictionary)
    Dim strAry_powerDomain() As String
    Dim i As Integer
    Dim j As Integer
    Dim strAry_pinVddbin() As String
    Dim cnt_DecomposedPinList As Long
    Dim PinGroup As String
    Dim domainTemp As String
    Dim domainTrackpower As String
    Dim PinTEMP As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//Check if powerDomain or powerPin is connected to DCVS.
'''1. BinCut powerDdomains are the pinGroup, and powerDomains include pins.
'''2. domain2pinDict, pin2domainDict are the dictionaries in GlobalVarible to store domains and pins.
'''//==================================================================================================================================================================================//'''
    '''init
    PinGroup = ""

    If domainList <> "" Then
        strAry_powerDomain = Split(domainList, ",")
        
        '''//Use TestJob of test program to find the matched DC Specs.
        ParsingDCspec UCase(TheExec.CurrentJob), dictPin2Dcspec
        
        For i = 0 To UBound(strAry_powerDomain)
            domainTemp = UCase(Trim(strAry_powerDomain(i)))
            domainTrackpower = ""
            
            If VddbinPmodeDict.Exists(domainTemp) = True Then
                PinGroup = ""
                
                '''//For main Domain
                Call TheExec.DataManager.DecomposePinList(domainTemp, strAry_pinVddbin, cnt_DecomposedPinList)
                
                If cnt_DecomposedPinList > 0 Then
                    If dictDomain2Pin.Exists(domainTemp) Then
                        '''Do nothing
                    Else
                        '''//For powerPin in main domain
                        For j = 0 To cnt_DecomposedPinList - 1
                            PinTEMP = UCase(Trim(strAry_pinVddbin(j)))
                        
                            '''//Check if DCVS is connected to the powerPin.
                            If IsDcvsConnected(PinTEMP) = True Then
                                If dictPin2Domain.Exists(PinTEMP) Then
                                    '''Do nothing
                                Else
                                    dictPin2Domain.Add PinTEMP, domainTemp
                                End If
                                
                                If PinGroup = "" Then
                                    PinGroup = PinTEMP
                                Else
                                    PinGroup = PinGroup & "," & PinTEMP
                                End If
                            End If
                        Next j
                        
                        If PinGroup = "" Then
                            TheExec.Datalog.WriteComment "Domain: " & domainTemp & " for initDomain2Pin contains no DCVS powerPin. Error!!!"
                            TheExec.ErrorLogMessage "Domain: " & domainTemp & " for initDomain2Pin contains no DCVS powerPin. Error!!!"
                        End If
                        
                        '''//Check if tracking power domain of main domain exists.
                        If AllBinCut(VddBinStr2Enum(domainTemp)).TRACKINGPOWER <> "" Then
                            domainTrackpower = AllBinCut(VddBinStr2Enum(domainTemp)).TRACKINGPOWER
                            
                            Call TheExec.DataManager.DecomposePinList(domainTrackpower, strAry_pinVddbin, cnt_DecomposedPinList)
                            
                            '''//Check if tracking power has cnt_DecomposedPinList>0.
                            If cnt_DecomposedPinList > 0 Then
                                For j = 0 To cnt_DecomposedPinList - 1
                                    PinTEMP = UCase(Trim(strAry_pinVddbin(j)))
                                    
                                    '''//Check if DCVS is connected to the powerPin.
                                    If IsDcvsConnected(PinTEMP) = True Then
                                        If dictPin2Domain.Exists(PinTEMP) Then
                                            '''Do nothing
                                        Else
                                            '''//Update dictionary of Domain2Pin
                                            dictPin2Domain.Add PinTEMP, domainTemp
                                        End If
                                        
                                        If PinGroup = "" Then
                                            TheExec.Datalog.WriteComment "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " contains no DCVS powerPin. Error!!!"
                                            TheExec.ErrorLogMessage "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " contains no DCVS powerPin. Error!!!"
                                        Else
                                            PinGroup = PinGroup & "," & PinTEMP
                                        End If
                                    End If
                                Next j
                            Else
                                TheExec.Datalog.WriteComment "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " isn't defined in PinMap for initDomain2Pin. Error!!!"
                                TheExec.ErrorLogMessage "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " isn't defined in PinMap for initDomain2Pin. Error!!!"
                            End If
                        End If
                        
                        '''//Update dictionary of Domain2Pin.
                        dictDomain2Pin.Add domainTemp, PinGroup
                        
                        If Not CheckDomainPinsDCspec(domainTemp, dictDomain2Pin, dictPin2Dcspec) Then
                            TheExec.Datalog.WriteComment "Domain: " & UCase(Trim(domainTemp)) & "," & " has different dcspec pins. Error!!!"
                            TheExec.ErrorLogMessage "Domain: " & UCase(Trim(domainTemp)) & "," & " has different dcspec pins. Error!!!"
                        End If
                    End If
                Else
                    TheExec.Datalog.WriteComment "Domain: " & UCase(Trim(domainTemp)) & "," & " isn't defined in PinMap for initDomain2Pin. Error!!!"
                    TheExec.ErrorLogMessage "Domain: " & UCase(Trim(domainTemp)) & "," & " isn't defined in PinMap for initDomain2Pin. Error!!!"
                End If
            Else
                TheExec.Datalog.WriteComment "Domain: " & UCase(Trim(domainTemp)) & "," & " doesn't enumerate in Vdd_Binning_Def. Error!!!"
                TheExec.ErrorLogMessage "Domain: " & UCase(Trim(domainTemp)) & "," & " doesn't enumerate in Vdd_Binning_Def. Error!!!"
            End If
        Next i
    Else
        TheExec.Datalog.WriteComment "Argument of domainList for initDomain2Pin should not be empty. Error!!!"
        TheExec.ErrorLogMessage "Argument of domainList for initDomain2Pin should not be empty. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initDomain2Pin"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initDomain2Pin"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created for Domain2Pin.
Public Function VddbinDomain2Pin(vddbinDomain As String) As String
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = UCase(vddbinDomain)

    If domain2pinDict.Exists(strTemp) Then
        VddbinDomain2Pin = UCase(domain2pinDict.Item(strTemp))
    Else
        VddbinDomain2Pin = "Domain_Error"
        TheExec.Datalog.WriteComment "Vddbin Domain=" & vddbinDomain & ", but it doesn't exist in VddbinDomain2Pin. Error!!!"
        'TheExec.ErrorLogMessage "Vddbin Domain=" & vddbinDomain & ", but it doesn't exist in VddbinDomain2Pin. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of VddbinDomain2Pin"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created for Pin2Domain.
Public Function VddbinPin2Domain(vddbinPin As String) As String
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = UCase(vddbinPin)

    If pin2domainDict.Exists(strTemp) Then
        VddbinPin2Domain = UCase(pin2domainDict.Item(strTemp))
    Else
        VddbinPin2Domain = "Pin_Error"
        TheExec.Datalog.WriteComment "Pin:" & vddbinPin & ", but it doesn't exist in the dictionary pin2domainDict. Please check BinCut voltage table and DC specs sheets for VddbinPin2Domain. Error!!!"
        'TheExec.ErrorLogMessage "Pin:" & vddbinPin & ", but it doesn't exist in the dictionary pin2domainDict. Please check BinCut voltage table and DC specs sheets for VddbinPin2Domain. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of VddbinDomain2Pin"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210113: Modified to check if cnt_DecomposedPinList>0.
'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created to check DCVS type for powerPin.
Public Function IsDcvsConnected(powerPin As String) As Boolean
    Dim strAry_PinName() As String
    Dim cnt_DecomposedPinList As Long
    Dim typesCount As Long
    Dim strAry_InstrumentTypes() As String
On Error GoTo errHandler
    Call TheExec.DataManager.DecomposePinList(powerPin, strAry_PinName(), cnt_DecomposedPinList)
    
    '''//If NumberPins>1, it means that powerPin might be a powerDomain(pinGroup) or an incorrect powerPin.
    If cnt_DecomposedPinList = 1 Then
        Call TheExec.DataManager.GetChannelTypes(powerPin, typesCount, strAry_InstrumentTypes())
        
        If LCase(strAry_InstrumentTypes(0)) Like "dcvs*" Then
            IsDcvsConnected = True
            If VddbinPinDcvstypeDict.Exists(UCase(powerPin)) = True Then
                '''Do nothing
            Else
                VddbinPinDcvstypeDict.Add UCase(powerPin), UCase(GetInstrument_BV(powerPin, 0))
            End If
        Else
            IsDcvsConnected = False
        End If
    Else
        IsDcvsConnected = False
        TheExec.Datalog.WriteComment "Pin:" & powerPin & ", it is not a correct powerPin for IsDcvsConnected. Error!!!"
        'TheExec.ErrorLogMessage "Pin:" & powerPin & ", it is not a correct powerPin for IsDcvsConnected. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of IsDcvsConnected"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created for Pin2Domain
Public Function VddbinPinDcvsType(vddbinPin As String) As String
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = UCase(vddbinPin)

    If VddbinPinDcvstypeDict.Exists(strTemp) Then
        VddbinPinDcvsType = UCase(VddbinPinDcvstypeDict.Item(strTemp))
    Else
        VddbinPinDcvsType = "Pin_Error"
        TheExec.Datalog.WriteComment "Pin:" & vddbinPin & ", it doesn't exist in VddbinPinDcvsType. Error!!!"
        'TheExec.ErrorLogMessage "Pin:" & vddbinPin & ", it doesn't exist in VddbinPinDcvsType. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of VddbinPinDcvsType"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created to get 1st Pin from PinGroup.
Public Function Get1stPinFromPingroup(powerDomain As String) As String
    Dim split_content() As String
On Error GoTo errHandler
    split_content = Split(powerDomain, ",")
    
    If UBound(split_content) > -1 Then
        Get1stPinFromPingroup = split_content(0)
    Else
        TheExec.Datalog.WriteComment "powerDomain: " & powerDomain & ", it is incorrect for Get1stPinFromPingroup. Error!!!"
        'TheExec.ErrorLogMessage "powerDomain: " & powerDomain & ", it is incorrect for Get1stPinFromPingroup. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of VddbinPinDcvsType"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
'20200618: Modified to move globalVariable "Public pin2dcspecstatusDict As New dictionary" to local variable of the function "CheckDomainPinsDCspec".
'20200106: Modified to remove the ErrorLogMessage.
'20191220: Created for checking group pins DCspec
Public Function CheckDomainPinsDCspec(powerDomain As String, dictDomain2Pin As Variant, DCspecTable As Variant) As Boolean
    Dim strTemp As String, DCspecTmp As String, CompareBase As String
    Dim split_content() As String
    Dim isCompResultDiff As Boolean
    Dim i As Long
    Dim AllPinExist As Boolean
    Dim pin2dcspecstatusDict As New Dictionary '''Note: "same" means group-pins has same dc spec, "diff" means group-pins has different dc spec, "none" means isn't group-pin.
On Error GoTo errHandler
    AllPinExist = True
    isCompResultDiff = False
    strTemp = dictDomain2Pin.Item(powerDomain)
    split_content = Split(strTemp, ",")
    pin2dcspecstatusDict.RemoveAll
    
    For i = 0 To UBound(split_content)
        If Not DCspecTable.Exists(split_content(i) & "_VAR") Then
            AllPinExist = False
            TheExec.Datalog.WriteComment "Pin:" & split_content(i) & "_VAR" & ", it doesn't exist in DCspec. Error!!!"
            'TheExec.ErrorLogMessage "Pin:" & split_content(i) & "_VAR" & ", it doesn't exist in DCspec. Error!!!"
        End If
    Next i
    
    If AllPinExist = True Then
        CompareBase = DCspecTable.Item(split_content(0) & "_VAR")
    
        If UBound(split_content) = 0 Then
            If Not pin2dcspecstatusDict.Exists(split_content(0)) Then
                pin2dcspecstatusDict.Add split_content(0), "none"
            End If
            CheckDomainPinsDCspec = True
        Else
            For i = 1 To UBound(split_content)
                DCspecTmp = DCspecTable.Item(split_content(i) & "_VAR")
                If Not (DCspecTmp = CompareBase) Then
                    isCompResultDiff = True
                    Exit For
                End If
            Next i
            
            If isCompResultDiff = True Then
                For i = 0 To UBound(split_content)
                    If Not pin2dcspecstatusDict.Exists(split_content(i)) Then
                        pin2dcspecstatusDict.Add split_content(i), "diff"
                    End If
                Next i
                CheckDomainPinsDCspec = False
            Else
                For i = 0 To UBound(split_content)
                    If Not pin2dcspecstatusDict.Exists(split_content(i)) Then
                        pin2dcspecstatusDict.Add split_content(i), "same"
                    End If
                Next i
                CheckDomainPinsDCspec = True
            End If
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of VddbinPinDcvsType"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of VddbinPinDcvsType"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200807: Modified to merge the branches.
'20200731: Modified to check if Spec name starts with the keyword "VDD".
'20200710: Modified to parse multiple DC_Spec sheets.
'20200703: Modiifed to use "check_Sheet_Range".
'20200508: Modified to merge "Find_Spec_Sheet" and "Find_JobList_Sheet" into "Find_Sheet".
'20191231: Modified to use TestJob of test program for DC Specs.
'20191220: Created for Parsing DCspec.
Public Function ParsingDCspec(testJob As String, DCspecTable As Variant)
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim i As Long, j As Long, k As Long
    Dim bincut_catlog_pos() As Integer
    Dim bincut_catlog_cnt As Integer
    Dim row_of_title As Integer
    Dim col_symbol As Integer
    Dim PinTEMP As String
    Dim valueTemp As String
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
    Dim strAry_DCSheetName() As String
On Error GoTo errHandler
    '''ToDo: Maybe we can check if testJob is emulated in "Enum BinCutJobDefinition"...
    If testJob <> "" Then
        '''//Find the matched DC specs sheet name for testJob from "GeneratedJobListSheet"
        Parsing_GeneratedJobList_Sheet testJob, "GeneratedJobListSheet", sheetName
        
        '''//Parse multiple DC_Spec sheets.
        strAry_DCSheetName = Split(sheetName, ",")
        
        For k = 0 To UBound(strAry_DCSheetName)
            '''*****************************************************************'''
            '''//Check if testJob has the matched DC Specs sheet.
            Set wb = Application.ActiveWorkbook
            Call check_Sheet_Range(strAry_DCSheetName(k), wb, ws_def, MaxRow, maxcol, isSheetFound)
            '''*****************************************************************'''
            If isSheetFound = True Then
                '''//Init
                '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
                bincut_catlog_cnt = 0
                row_of_title = 0
                enableRowParsing = False
                
                '''//Check the header of the table.
                '''Get the columns for the diverse coefficient.
                For i = 1 To MaxRow
                    For j = 1 To maxcol
                        If UCase(ws_def.Cells(i, j).Value) = "SYMBOL" Then
                            col_symbol = j
                            row_of_title = i
                        End If
                        
                        If row_of_title > 0 Then
                            If UCase(ws_def.Cells(row_of_title - 1, j).Value) Like "BINCUT_*" Then
                                ReDim Preserve bincut_catlog_pos(bincut_catlog_cnt)
                                bincut_catlog_pos(bincut_catlog_cnt) = j
                                bincut_catlog_cnt = bincut_catlog_cnt + 1
                            End If
                        End If
                    Next j
                    
                    If row_of_title > 0 Then
                        If bincut_catlog_cnt > 0 Then
                            enableRowParsing = True
                            Exit For
                        Else
                            enableRowParsing = False
                            If Not UCase(strAry_DCSheetName(k)) Like "*_BI" And Not UCase(strAry_DCSheetName(k)) Like "*_SC" Then
                                TheExec.Datalog.WriteComment strAry_DCSheetName(k) & " doesn't contain any DC Category name with keyword BINCUT. Error!!!"
                                TheExec.ErrorLogMessage strAry_DCSheetName(k) & " doesn't contain any DC Category name with keyword BINCUT. Error!!!"
                            End If
                            Exit For
                        End If
                    End If
                Next i
                
                '''//Start parsing the cells
                If enableRowParsing = True Then
                    For i = row_of_title + 1 To MaxRow
                        valueTemp = ""
                        PinTEMP = ws_def.Cells(i, col_symbol).Value
                        
                        '''//Check if Spec name starts with the keyword "VDD".
                        If (UCase(PinTEMP) Like "VDD*") Then
                            For j = 0 To UBound(bincut_catlog_pos)
                                If valueTemp = "" Then
                                    valueTemp = valueTemp & ws_def.Cells(i, bincut_catlog_pos(j)).Value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 1).Value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 2).Value
                                Else
                                    valueTemp = valueTemp & "," & ws_def.Cells(i, bincut_catlog_pos(j)).Value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 1).Value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 2).Value
                                End If
                            Next j
                        
                            If DCspecTable.Exists(PinTEMP) Then
                                TheExec.Datalog.WriteComment "DCspec pins repeat. Error!!!"
                                TheExec.ErrorLogMessage "DCspec pins repeat. Error!!!"
                            Else
                               DCspecTable.Add PinTEMP, valueTemp
                            End If
                        End If
                    Next i
                Else
                    If Not UCase(strAry_DCSheetName(k)) Like "*_BI" And Not UCase(strAry_DCSheetName(k)) Like "*_SC" Then
                        TheExec.Datalog.WriteComment strAry_DCSheetName(k) & " doesn't have the correct header for ParsingDCspec. Error!!!"
                        TheExec.ErrorLogMessage strAry_DCSheetName(k) & " doesn't have the correct header for ParsingDCspec. Error!!!"
                    End If
                End If
            End If '''If isSheetFound = True
        Next k '''For k = 0 To UBound(strAry_DCSheetName)
    Else
        TheExec.Datalog.WriteComment "The argument testJob of ParsingDCspec is incorrect for ParsingDCspec. Error!!!"
        TheExec.ErrorLogMessage "The argument testJob of ParsingDCspec is incorrect for ParsingDCspec. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of ParsingDCspec"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of ParsingDCspec"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
'20210809: Modified to check AllBinCut(p_mode).is_for_BinSearch to decide if it has to reset VBIN_Result for p_mode after MultiFSTP.
'20210806: Modified to remove the redundant property "IndexLevelIncDec As New SiteLong" from Public Type Instance_Step_Control.
'20210803: Modified to remove the redundant properties "IDS_current_fail As New SiteLong" and "IDS_current_Min As Double" from Public Type Instance_Step_Control.
'20210422: Modified to use Reset_VBinResult.
'20210422: Modified to remove PassBinCutByDomain and PassBinCutByPmode since the unused vbt code of GradeSearchMethod was removed.
'20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for ResetPmodePowerforBincut.
'20200807: Modified to reduce the redundant site-loop.
'20200525: Modified to remove the redundant site-loop for siteVariants.
'20200427: Modified to reset "PassBinCutByPmode".
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20191219: Created to init BinCut binning p_mode power.
Public Function ResetPmodePowerforBincut(inst_info As Instance_Info)
    Dim site As Variant
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Caution!!!
'''//If EnableWord "Vddbin_DoAll_DebugCollection" is enabled, initial performance mode result for Char. BinCut search voltage(Grade) and efuse product voltage(GradeVdd) as 0.
'''//==================================================================================================================================================================================//'''
    '''init
    inst_info.grade_found = False
    inst_info.AnySiteGradeFound = False

    For Each site In TheExec.sites
        inst_info.All_Site_Mask = inst_info.All_Site_Mask + 2 ^ site
        inst_info.IDS_ZONE_NUMBER = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).IDS_ZONE_NUMBER
    Next site
    
    '''20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
    If Flag_Vddbin_DoAll_DebugCollection = True Or EnableWord_Vddbin_PTE_Debug = True Then
        CurrentPassBinCutNum = 1 '''set the bincut number to 1.
    
        '''//Initialize bincut search voltage and efuse product voltage of the performance mode for BinCut doAll Char.
        '''20210422: Modified to use Reset_VBinResult.
        '''20210809: Modified to check AllBinCut(p_mode).is_for_BinSearch to decide if it has to reset VBIN_Result for p_mode after MultiFSTP.
        If AllBinCut(inst_info.p_mode).is_for_BinSearch = True Then
            'VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = False        'set the previous performance mode to false
            Call Reset_VBinResult(inst_info.p_mode)
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of ResetPmodePowerforBincut"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210422: Modified to use Reset_VBinResult.
'20210422: Modified to remove PassBinCutByDomain and PassBinCutByPmode since the unused vbt code of GradeSearchMethod was removed.
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20201203: Modified to use "idx_PowerDomain = VddBinStr2Enum(pinGroup_BinCut(i))".
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20200320: Modified to use the flag "Flag_Skip_ReApplyPayloadVoltageToDCVS".
'20200317: Modified for SearchByPmode.
'20200130: Modified to init DomainPassBinCutNum, BinCut_Init_Voltage, and BinCut_Payload_Voltage.
'20191230: Created for init VBIN_RESULT array.
Public Function initVbinTest()
    Dim p_mode As Integer
    Dim i As Long
    Dim idx_powerDomain As Long
On Error GoTo errHandler
    '''//Initialize VBIN_RESULT for each p_mode.
    For p_mode = 0 To MaxPerformanceModeCount
        Call Reset_VBinResult(p_mode)
    Next p_mode
    
    '''//Init BinCut_Init_Voltage, and BinCut_Payload_Voltage.
    For i = 0 To UBound(pinGroup_BinCut)
        idx_powerDomain = VddBinStr2Enum(pinGroup_BinCut(i))
        BinCut_Init_Voltage(idx_powerDomain) = 0
        BinCut_Payload_Voltage(idx_powerDomain) = 0
        Previous_Payload_Voltage(idx_powerDomain) = 0
    Next i
    
    CurrentPassBinCutNum = 1
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVbinTest"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVbinTest"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: Modified to set VBIN_RESULT(p_mode).GRADE and VBIN_RESULT(p_mode).GRADEVDD as 0.
'20210809: Modified to remove the redundant property "ALL_SITE_MIN As New SiteDouble" from Public Type VBIN_RESULT_TYPE.
'20210422: Created to reset VBin_Result(p_mode).
Public Function Reset_VBinResult(p_mode As Integer)
On Error GoTo errHandler
    VBIN_RESULT(p_mode).GRADE = 0                          '''set the lvcc result = 0
    VBIN_RESULT(p_mode).GRADEVDD = 0
    VBIN_RESULT(p_mode).tested = False                      '''set the flag "Tested" to false
    VBIN_RESULT(p_mode).FLAGFAIL = False                    '''set the fail flag to false, avoid the result always set to fail
    VBIN_RESULT(p_mode).step_in_IDS_Zone = 0                '''default is from step0
    VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone = -1
    VBIN_RESULT(p_mode).step_in_BinCut = TotalStepPerMode
    VBIN_RESULT(p_mode).passBinCut = 1
    VBIN_RESULT(p_mode).DSSC_Dec = -1
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Reset_VBinResult"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Reset_VBinResult"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201027: Modified to use "Public Type Instance_Info".
'20200130: Created to get p_mode, addi_mode, and Testtype from the test instance.
Public Function Get_Pmode_Addimode_Testtype_fromInstance(inst_info As Instance_Info)
On Error GoTo errHandler
    '''//Check if any input argument is empty.
    If inst_info.inst_name <> "" And inst_info.performance_mode <> "" Then
        '''//Init
        inst_info.p_mode = 0
        inst_info.special_voltage_setup = False
        inst_info.addi_mode = 0
        inst_info.test_type = testType.ldcbfd
        inst_info.offsetTestTypeIdx = testType.Func
        inst_info.jobIdx = BinCutJobDefinition.COND_ERROR
        
        '''//According to the instance keyword to decided "Test_Type", ex: elb, spi, rtos
        decide_test_type inst_info.test_type, inst_info.inst_name
        
        '''//According to the instance keyword to decide test type, then get dynamic_offset...
        inst_info.offsetTestTypeIdx = decide_offset_testType_byInstName(inst_info.inst_name)
        
        '''//Get BinCut test job name
        inst_info.jobIdx = getBinCutJobDefinition(bincutJobName)
        
        '=================================================================================
        ' Identify if the Performance Mode has the special test Condition (additional mode)
        '=================================================================================
        '''//Init the parameters for parsing the performance mode
        '''//Split the performance mode to get the main performance mode and the additional mode, and get powerpin from the performance mode.
        Call Parsing_Instance_Pmode(inst_info.performance_mode, inst_info.p_mode, inst_info.addi_mode, inst_info.special_voltage_setup)
        
        '''//Get powerDomain from the binning p_mode.
        inst_info.powerDomain = AllBinCut(inst_info.p_mode).powerPin
    Else
        TheExec.Datalog.WriteComment "Argument 'performance mode' of the instance:" & inst_info.inst_name & " is incorrect for Get_Pmode_Addimode_Testtype_fromInstance. Error!!!"
        TheExec.ErrorLogMessage "Argument 'performance mode' of the instance:" & inst_info.inst_name & " is incorrect for Get_Pmode_Addimode_Testtype_fromInstance. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Get_Pmode_Addimode_Testtype_fromInstance"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200320: Modified to save BinCut payload voltages of the previous instance into globalvariable "Previous_Payload_Voltage".
'20200130: Created to init BinCut_Payload_Voltage.
Public Function Init_BinCut_Voltage_Array()
    Dim i As Long
On Error GoTo errHandler
    For i = 0 To UBound(pinGroup_BinCut)
        Previous_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) = BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i)))
        'BinCut_Init_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) = 0
        BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) = 0
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Init_BinCut_Voltage_Array"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Set_PayloadVoltage_to_DCVS(Enable_Rail_Switch As Boolean, powerDomainGroup() As String, voltagePayload() As SiteDouble)

Dim Pin As Variant
'    For Each Pin In initVddBinning.hexvsPins
'        If LCase("*," & initVddBinning.SelsramBitTable.selsramPin & ",*") Like LCase("*," & Pin.Key & ",*") Then 'selsram powerPin
'            TheHdw.DCVS.Pins(Pin.Key).Voltage.Alt.ValuePerSite = voltagePayload(Pin.Value).Divide(1000)
'        Else
'            TheHdw.DCVS.Pins(Pin.Key).Voltage.ValuePerSite = voltagePayload(Pin.Value).Divide(1000)
'        End If
'    Next

    For Each site In TheExec.sites
        For Each Pin In powerDomainGroup
            If LCase("*," & selsramPin & ",*") Like LCase("*," & Pin & ",*") Then 'selsram powerPin
                TheHdw.DCVS.Pins(Pin).Voltage.Alt.Value = voltagePayload(VddBinStr2Enum(CStr(Pin))).Value(site) / 1000
            Else
                TheHdw.DCVS.Pins(Pin).Voltage.Value = voltagePayload(VddBinStr2Enum(CStr(Pin))).Value(site) / 1000
            End If
        Next
    Next site
End Function
'20210104: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
''20201215: Modified to reduce the redundant site-loop.
''20200320: Modified to use the flag "Flag_Skip_ReApplyPayloadVoltageToDCVS".
''20191210: Modified to check if powerPin belongs to selsramPin pinGroup.
'Public Function Set_PayloadVoltage_to_DCVS(Enable_Rail_Switch As Boolean, powerDomainGroup() As String, voltagePayload() As SiteDouble)
'    Dim site As Variant
'    Dim powerDomain As String
'    Dim PinGroup As String
'    Dim split_content() As String
'    Dim i As Long
'    Dim j As Long
'    Dim hexvsPingroup As String
'    Dim nonhexvsPingroup As String
'    Dim PinTEMP As String
'    Dim anySiteSelected As Boolean
'On Error GoTo errHandler
''''//==================================================================================================================================================================================//'''
''''//Note:
''''For projects with Rail-Switch, BinCut payload voltage values are applied to DCVS Valt.
''''//==================================================================================================================================================================================//'''
'    For i = 0 To UBound(powerDomainGroup)
'        '''//init
'        hexvsPingroup = ""
'        nonhexvsPingroup = ""
'        PinTEMP = ""
'        anySiteSelected = False
'
'        '''//Get powerPins from powerDomain
'        powerDomain = powerDomainGroup(i)
'        PinGroup = VddbinDomain2Pin(powerDomain)
'        split_content = Split(PinGroup, ",")
'
'        '''//Check if any site needs to update BinCut payload voltages to DCVS.
'        '''If BinCut_Payload_Voltage is same as Previous_Payload_Voltage, it could skip Re-applly payload voltages to DCVS.
'        If Flag_Skip_ReApplyPayloadVoltageToDCVS = True Then
'            If IsLevelLoadedForApplyLevelsTiming = False Then
'                For Each site In TheExec.sites
'                    If voltagePayload(VddBinStr2Enum(powerDomain))(site) <> Previous_Payload_Voltage(VddBinStr2Enum(powerDomain))(site) Then
'                        anySiteSelected = True
'                        Exit For
'                    End If
'                Next site
'            Else
'                anySiteSelected = True
'            End If
'        Else
'            anySiteSelected = True
'        End If
'
'        If anySiteSelected = True Then
'            '''//Assembly temporary pinGroup for HexVs
'            For j = 0 To UBound(split_content)
'                PinTEMP = UCase(Trim(split_content(j)))
'
'                '''20200210: Modified to check UltraFlex and UltraFlexPlus.
'                '''ToDo: Check if powerPin is DCVS or DCVI by checking VddbinPinDcvsType...
'                If glb_TesterType = "Jaguar" Then
'                    If LCase(VddbinPinDcvsType(PinTEMP)) Like "hexvs" Then '''for HexVs
'                        If hexvsPingroup = "" Then
'                            hexvsPingroup = PinTEMP
'                        Else
'                            hexvsPingroup = hexvsPingroup & "," & PinTEMP
'                        End If
'                    Else '''for non-HexVs
'                        If nonhexvsPingroup = "" Then
'                            nonhexvsPingroup = PinTEMP
'                        Else
'                            nonhexvsPingroup = nonhexvsPingroup & "," & PinTEMP
'                        End If
'                    End If
'                ElseIf glb_TesterType = "UltraFLEXplus" Then
'                    '''//Since all UltraFlexPlus DCVS instruments support siteAwareValue, we can set all powerPins in the same HexVS group.
'                    If hexvsPingroup = "" Then
'                        hexvsPingroup = PinTEMP
'                    Else
'                        hexvsPingroup = hexvsPingroup & "," & PinTEMP
'                    End If
'                    nonhexvsPingroup = ""
'                Else
'                    TheExec.Datalog.WriteComment "Tester type is not UltraFlex or UltraFlexPlus. Please check this for Set_PayloadVolage_to_DCVS. Error!!!"
'                    TheExec.ErrorLogMessage "Tester type is not UltraFlex or UltraFlexPlus. Please check this for Set_PayloadVolage_to_DCVS. Error!!!"
'                End If
'            Next j
'        End If
'
'        '''input: voltagePayload, scale & unit: mV.
'        '''DCVS applies unit:V. So that we need to convert mV into V for DCVS.
'        If hexvsPingroup <> "" Then
'            If Enable_Rail_Switch = True Then
'                '''***********************************************************************'''
'                '''//Check if powerPin belongs to selsramPin pinGroup.
'                '''selsram powerPin     : Set the payload voltage to Valt.
'                '''non-selsram powerPin : Set the payload voltage to Vmain.
'                '''***********************************************************************'''
'                '''ToDo: Maybe we can create the dictionary for selsramPin when parsing the table "SELSRM_Mapping_Table"...
'                '''20210104: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
'                If LCase("*," & selsramPin & ",*") Like LCase("*," & powerDomain & ",*") Then '''selsram powerPin
'                    TheHdw.DCVS.Pins(hexvsPingroup).Voltage.Alt.ValuePerSite = voltagePayload(VddBinStr2Enum(powerDomain)).Divide(1000)
'                Else
'                    TheHdw.DCVS.Pins(hexvsPingroup).Voltage.ValuePerSite = voltagePayload(VddBinStr2Enum(powerDomain)).Divide(1000)
'                End If
'            Else '''project without rail-switch.
'                TheHdw.DCVS.Pins(hexvsPingroup).Voltage.ValuePerSite = voltagePayload(VddBinStr2Enum(powerDomain)).Divide(1000)
'            End If
'        End If
'
'        If nonhexvsPingroup <> "" Then
'            If Enable_Rail_Switch = True Then
'                If LCase("*," & selsramPin & ",*") Like LCase("*," & powerDomain & ",*") Then '''selsram powerPin
'                    For Each site In TheExec.sites
'                        TheHdw.DCVS.Pins(nonhexvsPingroup).Voltage.Alt.Value = voltagePayload(VddBinStr2Enum(powerDomain)) / 1000
'                    Next site
'                Else '''non-selsram powerPin
'                    For Each site In TheExec.sites
'                        TheHdw.DCVS.Pins(nonhexvsPingroup).Voltage.Value = voltagePayload(VddBinStr2Enum(powerDomain)) / 1000
'                    Next site
'                End If
'            Else '''project without rail-switch.
'                For Each site In TheExec.sites
'                    TheHdw.DCVS.Pins(nonhexvsPingroup).Voltage.Value = voltagePayload(VddBinStr2Enum(powerDomain)) / 1000
'                Next site
'            End If
'        End If
'    Next i
'Exit Function
'errHandler:
'    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Set_PayloadVoltage_to_DCVS"
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'20200717: Modified to split specGrp and put voltagePayload into each spec of Overlay.
'20200618: Modified to use VddbinDomain2DcSpecGrp.
'20200615: Created for "Call Instance".
Public Function Set_PayloadVoltage_to_Overlay(Enable_Rail_Switch As Boolean, powerDomainGroup() As String, voltagePayload() As SiteDouble, OverlayName As String)
    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim powerDomain As String
    Dim specGrp As String
    Dim split_content() As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Use Overlay in test instance sheet to overwrite voltages defined in DC Specs avoid HardIP/RTOS applyLevelsTiming and ForceCondition to overwrite BinCut payload voltages.
'''2. Remember to use Remove_PayloadVoltage_from_Overlay to remove Overlay after the test.
'''3. IGXL doesn't support one input with multiple specs names, so that we have to split specGrp and setup spec by spec-loop.
'''//==================================================================================================================================================================================//'''
    With TheExec.Overlays
        If (.Contains(OverlayName) <> False) Then .Remove OverlayName
        .Add (OverlayName)
    End With
    
    For i = 0 To UBound(powerDomainGroup)
        '''//init
        powerDomain = powerDomainGroup(i)
        
        '''//Check if powerDomain belongs to powerPin or pinGroup...
        If VddbinPinDict.Exists(UCase(powerDomain)) = True Then
            specGrp = VddbinDomain2DcSpecGrp(powerDomain)

            '''//Split specGrp and put payload voltage into each spec of Overlay.
            split_content = Split(specGrp, ",")
            
            For j = 0 To UBound(split_content)
                '''//If one of them exists, create the spec item.
                If specGrp <> "" Then
                    With TheExec.Overlays(OverlayName)
                        .specs.Add (split_content(j))
                        .specs.Item(split_content(j)).Value = voltagePayload(VddBinStr2Enum(powerDomain)).Divide(1000)
                    End With
                End If
            Next j
        Else
            TheExec.Datalog.WriteComment "powerDomain:" & powerDomain & ", it is not BinCut powerPin or pinGroup. It is incorrect for Set_PayloadVoltage_to_Overlay. Error!!!"
            TheExec.ErrorLogMessage "powerDomain:" & powerDomain & ", it is not BinCut powerPin or pinGroup. It is incorrect for Set_PayloadVoltage_to_Overlay. Error!!!"
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Set_PayloadVoltage_to_Overlay"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200616: Created to remove BinCut payload voltages from Overlay for HardIP instance..
Public Function Remove_PayloadVoltage_from_Overlay(OverlayName As String)
    Dim strAry_OverlayName() As String
    Dim i As Double
On Error GoTo errHandler
    strAry_OverlayName = Split(OverlayName, ",")

    For i = 0 To UBound(strAry_OverlayName)
        With TheExec.Overlays
            If (.Contains(strAry_OverlayName(i)) <> False) Then
                .Remove strAry_OverlayName(i)
                TheExec.Datalog.WriteComment "Remove Overlay Name: " & strAry_OverlayName(i)
            End If
        End With
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Remove_PayloadVoltage_from_Overlay"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200320: Created to get instance context and assembly all information into the string.
Public Function GetInstanceContextIntoString() As String
    Dim DCCategory As String
    Dim DCSelector As String
    Dim ACCategory As String
    Dim ACSelector As String
    Dim TimeSetSheet As String
    Dim EdgeSetSheet As String
    Dim LevelsSheet As String
    Dim Overlay As String
On Error GoTo errHandler
    '''//TheExec.DataManager.GetInstanceContext(DCCategory As String, DCSelector As String, ACCategory As String, ACSelector As String, TimeSetSheet As String, EdgeSetSheet As String, LevelsSheet As String, Overlay As String, [MemberNumber As Long = -1])
    Call TheExec.DataManager.GetInstanceContext(DCCategory, DCSelector, ACCategory, ACSelector, TimeSetSheet, EdgeSetSheet, LevelsSheet, Overlay)
    
    GetInstanceContextIntoString = DCCategory & "," & DCSelector & "," & ACCategory & "," & ACSelector & "," & TimeSetSheet & "," & LevelsSheet & "," & Overlay
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of GetInstanceContextIntoString"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200424: Created to check Adjust_Max and Adjust_Min for adjust_VddBinning.
'20191023: Modified to check if "MaxPV(pmode0/pmode1)" is in the column "Comment" of "Vdd_Binning_Def" or not.
Public Function Check_Adjust_Max_Min(Adjust_Max_Enable As Boolean, Adjust_Min_Enable As Boolean, Optional Adjust_Power_Max_list As String, Optional Adjust_Power_Min_list As String)
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim group_array() As String
    Dim max_array() As String
    Dim max_value As Double
    Dim min_array() As String
    Dim min_value As Double
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Example:
'''Adjust_Power_Max_list = "P_mode1,P_mode2 + P_mode3,P_mode4" => P_mode1 and P_mode2 will fuse the max value1. P_mode3 and P_mode4 will fuse the max value2.
'''ex: MC607=500mv, MC608=550mv. If both p_modes are set as "allowequal", MC607 will be adjusted to 550mv. Then MC607 and MC608 will be 550mv.
'''//==================================================================================================================================================================================//'''
    If (Adjust_Max_Enable <> Flag_Adjust_Max_Enable) Or (Adjust_Min_Enable <> Flag_Adjust_Min_Enable) Then
        TheExec.Datalog.WriteComment "The instance arguments Adjust_Max_Enable or Adjust_Min_Enable of adjust_VddBinning might be inconsistent with column Comment of Vdd_Binning_Def_appA. Error!!!"
        TheExec.ErrorLogMessage "The instance arguments Adjust_Max_Enable or Adjust_Min_Enable of adjust_VddBinning might be inconsistent with column Comment of Vdd_Binning_Def_appA. Error!!!"
    Else
        '''//Check if "MaxPV(pmode0/pmode1)" is in the column "Comment" of "Vdd_Binning_Def" or not.
        If Adjust_Max_Enable = True Then
            If Adjust_Power_Max_list <> Adjust_Power_Max_pmode Then
                TheExec.Datalog.WriteComment "Argument Adjust_Power_Max_list: " & Adjust_Power_Max_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Max_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!"
                TheExec.ErrorLogMessage "Argument Adjust_Power_Max_list: " & Adjust_Power_Max_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Max_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!"
            End If
        End If
        
        If Adjust_Min_Enable = True Then
            If Adjust_Power_Min_list <> Adjust_Power_Min_pmode Then
                TheExec.Datalog.WriteComment "Argument Adjust_Power_Min_list: " & Adjust_Power_Min_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Min_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!"
                TheExec.ErrorLogMessage "Argument Adjust_Power_Min_list: " & Adjust_Power_Min_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Min_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!"
            End If
        End If
    End If
    
    '''//Adjust_Max_Enable is related to "Allowequal"
    If Adjust_Max_Enable = True Then
        group_array = Split(Adjust_Power_Max_list, "+")
        
        For j = 0 To UBound(group_array)
            max_array = Split(group_array(j), ",")
            
            If UBound(max_array) < 1 Then
                TheExec.Datalog.WriteComment "Adjust_Power_Max pin list of adjust_VddBinning is less than 2"
                'TheExec.ErrorLogMessage "Adjust_Power_Max pin list of adjust_VddBinning is less than 2"
            End If
            
            For Each site In TheExec.sites
                max_value = 0
                For i = 0 To UBound(max_array)
                    If AllBinCut(VddBinStr2Enum(max_array(i))).Used = True Then
                        If CDec(max_value) < CDec(VBIN_RESULT(VddBinStr2Enum(max_array(i))).GRADEVDD) Then
                             max_value = VBIN_RESULT(VddBinStr2Enum(max_array(i))).GRADEVDD
                        End If
                    Else
                        TheExec.Datalog.WriteComment "The Performance Mode " & max_array(i) & " is not used for adjust_VddBinning!!!"
                    End If
                Next i
                For i = 0 To UBound(max_array)
                    If CDec(VBIN_RESULT(VddBinStr2Enum(max_array(i))).GRADEVDD) <> CDec(max_value) Then
                        VBIN_RESULT(VddBinStr2Enum(max_array(i))).GRADEVDD = max_value
                        TheExec.Datalog.WriteComment "site:" & site & "," & max_array(i) & ",BinCut Product voltage is adjusted to Max Value: " & max_value
                    End If
                Next i
            Next site
        Next j
    End If
    
    '''//Adjust_Min_Enable is related to "Allowequal".
    '''ex: MC607=500mv, MC608=550mv. If both p_modes are set as "allowequal", MC608 will be adjusted to 500mv. Then MC607 and MC608 will be 500mv.
    If Adjust_Min_Enable = True Then
        group_array = Split(Adjust_Power_Min_list, "+")
        
        For j = 0 To UBound(group_array)
            min_array = Split(group_array(j), ",")
            
            If UBound(min_array) < 1 Then
                TheExec.ErrorLogMessage "Adjust_Power_Min pin list of adjust_VddBinning is less than 2"
            End If
            
            For Each site In TheExec.sites
                min_value = 9999
                For i = 0 To UBound(min_array)
                    If AllBinCut(VddBinStr2Enum(min_array(i))).Used = True Then
                        If CDec(min_value) > CDec(VBIN_RESULT(VddBinStr2Enum(min_array(i))).GRADEVDD) Then
                             min_value = VBIN_RESULT(VddBinStr2Enum(min_array(i))).GRADEVDD
                        End If
                    Else
                        TheExec.Datalog.WriteComment "The Performance Mode " & min_array(i) & " Doesn't Exist. Error!!!"
                        TheExec.ErrorLogMessage "The Performance Mode " & min_array(i) & " Doesn't Exist. Error!!!"
                    End If
                Next i
                For i = 0 To UBound(min_array)
                    If VBIN_RESULT(VddBinStr2Enum(min_array(i))).GRADEVDD <> min_value Then
                        VBIN_RESULT(VddBinStr2Enum(min_array(i))).GRADEVDD = min_value
                        TheExec.Datalog.WriteComment "site:" & site & "," & min_array(i) & ",BinCut Product voltage is adjusted to Min Value: " & min_value
                    End If
                Next i
            Next site
        Next j
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Check_Adjust_Max_Min"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210302: Modified to remove the If condition for DevChar.
'20210126: Modified to revise the vbt code for DevChar.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201027: Modified to use "Public Type Instance_Info".
'20200924: Modified to remove the redundant argument "IndexLevelPerSite As SiteLong" from Set_BinCut_Initial_by_ApplyLevelsTiming.
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20200425: Modified to adjust the flow for "print_bincut_power".
'20200424: Modified to use "Set_BinCut_Initial_by_ApplyLevelsTiming" to set BinCut initial voltage by ApplyLevelsTiming.
'20200324: Modified to skip ApplyLevelsTiming when current instance has the same level/timing as previous instance for project with rail-switch.
'20200320: Modified to check instance contexts of current instance and previous instance.
'20200206: Modified to replace "print_main_power_init" with "print_bincut_power".
Public Function Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info As Instance_Info)
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Set initial voltages from category "Bincut_X_X_X" in DC_Specs sheet by ApplyLevelsTiming.
'''Print the initial voltages, and applies them to DCVS Vmain and Valt by ApplyLevelsTiming (DCVS voltage source will be switched to Vmain).
'''Skip ApplyLevelsTiming when current instance has the same level/timing as previous instance for project with rail-switch.
'''//==================================================================================================================================================================================//'''
    If Flag_Skip_ReApplyPayloadVoltageToDCVS = True Then
        CurrentBinCutInstanceContext = GetInstanceContextIntoString
        
        If CurrentBinCutInstanceContext = PreviousBinCutInstanceContext And Flag_Enable_Rail_Switch = True Then
            select_DCVS_output_for_powerDomain tlDCVSVoltageMain
            inst_info.currentDcvsOutput = tlDCVSVoltageMain
            IsLevelLoadedForApplyLevelsTiming = False
        Else
            Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
            inst_info.currentDcvsOutput = tlDCVSVoltageMain
            IsLevelLoadedForApplyLevelsTiming = True
        End If
    Else
'Dim test_time As Double
'test_time = Timer
        Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
        inst_info.currentDcvsOutput = tlDCVSVoltageMain
        IsLevelLoadedForApplyLevelsTiming = True
'        TheExec.Datalog.WriteComment ("***** Test Time (VBA): ApplyLevelsTiming (s) = " & Format(Timer - test_time, "0.000000"))
    End If
    
    '''//Print initial voltages
    '''Ex: Initial_Voltage_VDD_SOC_MS003,0,VDD_PCPU=0.752,VDD_ECPU=0.752,VDD_GPU=0.752,VDD_SOC=0.752, ...
    print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages, False, BincutVoltageType.InitialVoltage
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Set_BinCut_Initial_by_ApplyLevelsTiming"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200506: Created to decide testType for string.
Public Function decide_test_type_for_string(strInput As String) As testType
    Dim strTemp As String
On Error GoTo errHandler
    strTemp = LCase(strInput)
    
    '''//Please check "Enum TestType" and "MaxTestType" in GlobalVariable.
    If strTemp Like "*td*" Then
        decide_test_type_for_string = testType.TD
    ElseIf strTemp Like "*bist*" Then
        decide_test_type_for_string = testType.Mbist
    ElseIf strTemp Like "*tmps*" Then
        decide_test_type_for_string = testType.TMPS
    ElseIf strTemp Like "*spi*" Then
        decide_test_type_for_string = testType.SPI
    ElseIf strTemp Like "*rtos*" Then
        decide_test_type_for_string = testType.RTOS
    ElseIf strTemp Like "*ldcbfd*" Then
        decide_test_type_for_string = testType.ldcbfd
    Else
        TheExec.Datalog.WriteComment "decide_test_type_for_string can't decide testType for input:" & strInput & ". Error!!!"
        TheExec.ErrorLogMessage "decide_test_type_for_string can't decide testType for input:" & strInput & ". Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of decide_test_type_for_string"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of decide_test_type_for_string"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "StepCount As Long" as "count_Step As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to move "Step_GradeFound As New SiteLong" from Public Type Instance_Step_Control to the vbt function Update_VBinResult_by_Step.
'20210901: Modified to rename "IndexFoundLevel As New SiteLong" as "Step_GradeFound As New SiteLong" for Public Type Instance_Step_Control.
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
'20210810: Modified to merge the vbt function Check_anySite_GradeFound into the vbt function Update_VBinResult_by_Step.
'20210810: Modified to use gotPassStep as flag to determine if update VBIN_Result or not.
'20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
'20210809: Modified to revise the branches of the vbt code for Linear algorithm and IDS algorithm of BinCut gradesearch.
'20210809: Modified to revise the vbt code with VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone(site).
'20210806: Modified to update VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_control.IndexFoundLevel(site)) for IDS mode.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210720: Modified to revise the vbt function Update_VBinResult_by_Step for BinCut search in FT.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210302: Modified to use "step_control.FoundLevel(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.IndexLevelPerSite(site))".
'20210226: Modified to use step_Start and step_Stop to get startVoltage and StopVoltage.
'20210125: Modified to remove "voltage_Pmode_EQNbased As SiteDouble" from the arguments of the vbt function "Update_VBinResult_by_Step".
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for Update_VBinResult_by_Step.
'20201207: Modified to add the argument "flag_All_Patt_Pass".
'20200602: Modified to remove the condition "IndexLevelPerSite(site) = 0" for "gradeAlg = GradeSearchAlgorithm.ids".
'20200317: Modified for SearchByPmode.
'20190313: Modified to add "FIRSTPASSBINCUT(p_mode)" for storing the first passbinnum of P_mode.
Public Function Update_VBinResult_by_Step(inst_info As Instance_Info)
    Dim site As Variant
    Dim gotPassStep As Boolean
    Dim step_GradeFound As New SiteLong
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//For step control:
'''count_Step: step index for DYNAMIC_VBIN_IDS_ZONE(p_mode).
'''Step_Current: step has beed tested.
'''//==================================================================================================================================================================================//'''
    '''init
    inst_info.Grade_Not_Found_Mask = 0
    inst_info.On_StopVoltage_Mask = 0
    inst_info.Grade_Found_Mask = 0

    For Each site In TheExec.sites
        '''init
        '''20210810: Modified to use gotPassStep as flag to determine if update VBIN_Result or not.
        gotPassStep = False
        
        '''//If IDS mode fails at the first step, it has to switch BinCut search algorithm from IDS mode to Linear mode immediately.
        If inst_info.gradeAlg(site) = GradeSearchAlgorithm.IDS And inst_info.All_Patt_Pass(site) = False And inst_info.count_Step = 0 Then
            '''*********************************************************************************************************************************************************************************'''
            '''//If the algorithm is IDS mode and this site is failed in first step, we will change the algorithm to linear mode and change the direction for stopvoltage.
            '''
            '''  Site0   Site1   Site2     => if the algo = IDS, means that there is lower EQ number can be tested and the IDS start EQ is must not last EQ number.
            '''  EQ2(F)  EQ2(P)  EQ2(P)    => All sites start in EQ2, and only site 0 fail in first step, the site 0 will change to linear mode.
            '''  EQ1(P)  EQ3(P)  EQ3(F)    => only site 0 run the different direction , Although site 2 is also failed, it is not failed in first step. We do not need to change the algorithm.
            '''  EQ1(P)  EQ4(P)  EQ4(F)
            '''
            '''  Grade   Grade   Grade
            '''  EQ1     EQ4     EQ2
            '''*********************************************************************************************************************************************************************************'''
            inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear
            inst_info.step_Stop(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1
        End If
        
        '''//grade found at linear algorithm.
        If (inst_info.gradeAlg = GradeSearchAlgorithm.linear And inst_info.All_Patt_Pass(site) = True) And inst_info.grade_found(site) = False _
        Or ((inst_info.gradeAlg = GradeSearchAlgorithm.IDS And inst_info.All_Patt_Pass(site) = True) And inst_info.grade_found(site) = False And inst_info.step_Current(site) = inst_info.step_Stop(site)) Then
            '''*********************************************************************************************************************************************************************************'''
            '''//When the patterns are passed in this step, there are some conditions will be identified to the grade had been found.
            ''' 1. Algorithm = IDS, And the grade was not been found in this test Instance. But the voltage of the step reach the stopvoltage (pattern never failed in this site).
            ''' 2. Algorithm = Linear, And the grade was not been found in this test Instance.
            '''
            '''  Linear                                                      IDS
            '''  site0            site1           site2                      site0   site1            site2
            '''  EQ4(F)           EQ4(F)          EQ4(P)                     EQ3(F)  EQ3(P)           EQ3(P)
            '''  EQ3(F)           EQ3(P) => found EQ3(P)                     EQ2(F)  EQ4(P) => found  EQ4(F)
            '''  EQ2(P) => found  EQ2(P)          EQ2(P)                     EQ1(P)  EQ4(P)           EQ4(F)
            '''
            '''  Grade            Grade           Grade                      Grade   Grade            Grade
            '''  EQ2              EQ3             EQ4                        EQ1     EQ4              EQ3
            '''*********************************************************************************************************************************************************************************'''
            '''//Determine the pass Step in Dynamic_IDS_zone of the p_mode.
            step_GradeFound(site) = inst_info.step_Current(site)
            gotPassStep = True
        End If
        
        '''//grade found at IDS algorithm.
        If ((inst_info.gradeAlg = GradeSearchAlgorithm.IDS And inst_info.All_Patt_Pass(site) = False) And inst_info.grade_found(site) = False) Then
            '''*********************************************************************************************************************************************************************************'''
            '''//When the patterns are failed in this step, there are some conditions will be identified to the grade had been found.
            ''' 1. Algorithm = IDS, And the grade was not been found in this test Instance.
            '''
            ''' IDS
            ''' site0   site1   site2
            ''' EQ3(F)  EQ3(P)  EQ3(P)
            ''' EQ2(F)  EQ4(P)  EQ4(F)  => found
            ''' EQ1(P)  EQ4(P)  EQ4(F)
            '''
            ''' Grade   Grade   Grade
            ''' EQ1     EQ4     EQ3
            '''*********************************************************************************************************************************************************************************'''
            '''//Determine the pass Step in Dynamic_IDS_zone of the p_mode.
            step_GradeFound(site) = inst_info.step_Current(site) + 1
            gotPassStep = True
        End If
        
        '''//If the pass step in Dynamic_IDS_zone of p_mode is found, update VBIN_Result for p_mode.
        '''20210809: Modified to revise the branches of the vbt code for Linear algorithm and IDS algorithm of BinCut gradesearch.
        '''20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
        If gotPassStep = True Then
            '''//Update PassBin, Pass step, flag"VBIN_Result(p_mode).tested", and voltage to VBIN_Result by the step in Dynamic_IDS_Zone.
            Call Set_VBinResult_by_Step(site, inst_info.p_mode, step_GradeFound(site))
            
            '''//Update the flag about grade_found.
            inst_info.grade_found(site) = True
            inst_info.AnySiteGradeFound = True
            
            '''//Update PassBin to the globalVariable CurrentPassBinCutNum.
            CurrentPassBinCutNum = VBIN_RESULT(inst_info.p_mode).passBinCut
        End If
        
        '''//Check if any site passes or fails on the current step.
        '''20210810: Modified to merge the vbt function Check_anySite_GradeFound into the vbt function Update_VBinResult_by_Step.
        '''*********************************************************************************************************************************************************************************'''
        '''//For all sites, we need to record some flags to identify if need to exit the loop when there is no reason to seach.
        '''      On_StopVoltage_Mask => record how many sites had reached the stopvoltage.
        '''      Grade_Found_Mask    => record how many sites had reached the stopvoltage.
        '''*********************************************************************************************************************************************************************************'''
        If inst_info.grade_found(site) = False Then
            inst_info.Grade_Not_Found_Mask = inst_info.Grade_Not_Found_Mask + 2 ^ site
            
            If inst_info.step_Current(site) = inst_info.step_Stop(site) Then
                inst_info.On_StopVoltage_Mask = inst_info.On_StopVoltage_Mask + 2 ^ site
            End If
        Else
            inst_info.Grade_Found_Mask = inst_info.Grade_Found_Mask + 2 ^ site
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Update_VBinResult_by_Step"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Update_VBinResult_by_Step"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210813: Created to set VBIN_Result by the step.
Public Function Set_VBinResult_by_Step(site As Variant, p_mode As Integer, lng_step_selected As Long)
On Error GoTo errHandler
    If DYNAMIC_VBIN_IDS_ZONE(p_mode).Used(site) = True Then
        '''//Update the pass step in Dynamic_IDS_zone of p_mode to VBIN_Result for p_mode.
        VBIN_RESULT(p_mode).step_in_IDS_Zone = lng_step_selected
        
        '''//Store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
        VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone = VBIN_RESULT(p_mode).step_in_IDS_Zone
        
        '''//step_in_BinCut = EQN-1.
        VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(VBIN_RESULT(p_mode).step_in_IDS_Zone(site)) - 1
        
        '''//Get the current BinCut number.
        VBIN_RESULT(p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(VBIN_RESULT(p_mode).step_in_IDS_Zone(site))
        
        '''//Update BinCut voltage(Grade) of p_mode according to the pass Step in Dynamic_IDS_zone.
        '''20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
        VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(VBIN_RESULT(p_mode).step_in_IDS_Zone(site))
        
        '''//Update Efuse product voltage(GradeVDD) of p_mode. => Efuse product voltage(GradeVDD) = BinCut voltage(Grade) + Guardband.
        VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(VBIN_RESULT(p_mode).step_in_IDS_Zone(site))
        
        VBIN_RESULT(p_mode).tested = True
    Else
        TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & ", it doesn't have any correct Dynamic_IDS_zone for Set_VBinResult_by_Step. Error!!!"
        TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(p_mode) & ", it doesn't have any correct Dynamic_IDS_zone for Set_VBinResult_by_Step. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Set_VBinResult_by_Step"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Set_VBinResult_by_Step"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
'20210909: Modified to revise the vbt code for FirstChangeBinInfo requested by C651 Si, as discussed with TSMC ZYLINI and ZQLIN.
'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210806: Modified to remove the redundant property "IndexLevelIncDec As New SiteLong" from Public Type Instance_Step_Control.
'20210629: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to determine "Print Bincut Fail Info", as suggested by Chihome.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210107: Modified to add for recording first changed binnum mode data, requested by C651 Si.
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for Decide_NextStep_for_GradeSearch.
'20200508: Created to decide Next Step for GradeSearch.
'20200317: Modified for SearchByPmode.
Public Function Decide_NextStep_for_GradeSearch(inst_info As Instance_Info)
    Dim site As Variant
    Dim sitelng_next_step As New SiteLong
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. IndexLevelPerSite: step in Dynamic_IDS_Zone has been tested.
'''2. Define the direction of each site in the IDS zone:
'''     BinCut Search Algorithm = IDS => move to small step (step3 -> step2 -> step1 -> step0 -> step0 -> step0)
'''     BinCut Search Algorithm = Linear => move to large step (step1 -> step2 -> step3 -> step3 -> step3)
'''//==================================================================================================================================================================================//'''
    '''On_StopVoltage_Mask = False
    For Each site In TheExec.sites 'decide next StepCount
        If inst_info.gradeAlg(site) = GradeSearchAlgorithm.IDS Then
            sitelng_next_step(site) = -1
        ElseIf inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear Then
            sitelng_next_step(site) = 1
        End If
        
        inst_info.step_Current(site) = inst_info.step_Current(site) + sitelng_next_step(site)
        
        '''//Check if the next step is within the correct steps of p_mode.
        '''step_inherit is the step with the lowest BinCut voltage in Dynamic_IDS_zone.
        If inst_info.step_Current(site) > DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 Then
            inst_info.step_Current(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1
        ElseIf inst_info.step_Current(site) < DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_inherit Then
            inst_info.step_Current(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_inherit
        End If
        
        If inst_info.grade_found = False And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current) <> CurrentPassBinCutNum Then
            '''********************************************************************************'''
            '''20210107: Modified to add for recording first changed binnum mode data, requested by C651 Si.
            '''20210629: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to determine "Print Bincut Fail Info", as suggested by Chihome.
            '''20210909: Modified to revise the vbt code for FirstChangeBinInfo requested by C651 Si, as discussed with TSMC ZYLINI and ZQLIN.
            If EnableWord_Vddbin_PTE_Debug = False Then
                If DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current) > CurrentPassBinCutNum Then
                    FirstChangeBinInfo.FirstChangeBinMode(site) = inst_info.p_mode
                    FirstChangeBinInfo.FirstChangeBinType(site) = inst_info.offsetTestTypeIdx
                    '''20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
                    FirstChangeBinInfo.str_Pmode_Test(site) = inst_info.inst_name
                End If
            End If
            '''********************************************************************************'''
            CurrentPassBinCutNum = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current)
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Decide_NextStep_for_GradeSearch"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Decide_NextStep_for_GradeSearch"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210830: Modified to merge the branches of the vbt function Update_PassBinCut_for_GradeNotFound.
'20210813: Modified to move "VBIN_RESULT(inst_info.p_mode).tested" from the vbt function Update_PassBinCut_for_GradeNotFound to the vbt function Set_VBinResult_by_Step.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for Update_PassBinCut_for_GradeNotFound.
'20200511: Created to update CurrentPassBinCutNum for DUT "grade_found=false".
'20200317: Modified for SearchByPmode.
Public Function Update_PassBinCut_for_GradeNotFound(inst_info As Instance_Info)
    Dim site As Variant
On Error GoTo errHandler
    For Each site In TheExec.sites
        If inst_info.grade_found(site) = False Then
            VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = inst_info.step_Current
            VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
            VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
            CurrentPassBinCutNum = VBIN_RESULT(inst_info.p_mode).passBinCut
            
            '''//If the performance mode is tested first time, set the BinCut(P_mode, CurrentPassBinCutNum).Tested = True.
'''ToDo: Maybe it needs to be modified for the failed DUT...
            VBIN_RESULT(inst_info.p_mode).tested = True
            
            '''20210830: Modified to merge the branches of the vbt function Update_PassBinCut_for_GradeNotFound.
            '''//If the grade is not found, set the result to fail.
            If VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False Then '''one of the instance fails
                VBIN_RESULT(inst_info.p_mode).FLAGFAIL = True
                VBIN_RESULT(inst_info.p_mode).GRADE = 0
                VBIN_RESULT(inst_info.p_mode).GRADEVDD = 0
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = TotalStepPerMode
            End If
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Update_PassBinCut_for_GradeNotFound"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Update_PassBinCut_for_GradeNotFound"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210416: Modified to check if DSSC DigSrc patterns of SELSRM and Harvest Core exist.
'20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
'20201029: Modified to remove the argument "result_mode As tlResultMode" and use inst_info.result_mode.
'20201029: Modified to use "Public Type Instance_Info".
'20201029: Modified to check if idxBlock_Selsrm_PrePatt = idxBlock_Selsrm_FuncPat.
'20201027: Modified to check SELSRM DSSC digsrc pattern for pattern burst requested by C651 Toby.
'20201016: Modified to adjust the sequence of arguments.
'20200918: Modified to add the argument "result_mode" for the vbt function "Check_and_Decompose_PrePatt_FuncPat".
'20200915: Modified to rename the argument "Flag_DecomposePatt_from_InstanceArg" to "str_Set_DecomposePatt"
'20200520: Created to check and decompose patsets PrePatt and FuncPat, and find SELSRAM DSSC pattern for DSSC digSrc.
Public Function Check_and_Decompose_PrePatt_FuncPat(inst_info As Instance_Info, result_mode As tlResultMode, str_Set_DecomposePatt As String, PrePatt As String, FuncPat As String)
    Dim strTemp As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''No matter DecomposePatt = "YES" or "NO", it needs to decompose the pattern set to find the DSSC digsrc pattern.
'''//==================================================================================================================================================================================//'''
    '''//Init
    inst_info.enable_DecomposePatt = True
    '''SelSRM
    inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt = ""
    inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat = ""
    inst_info.patt_SelsrmDigSrc_single = ""
    inst_info.idxBlock_Selsrm_PrePatt = -1
    inst_info.idxBlock_Selsrm_FuncPat = -1
    inst_info.idxBlock_Selsrm_singlePatt = -1
    '''str_Set_DecomposePatt comes from arguments of the instance.
    strTemp = LCase(str_Set_DecomposePatt)
    
    '''//Check if decomposing pattern is needed.
    If strTemp = "" Or strTemp = "yes" Or strTemp = "true" Then
        inst_info.enable_DecomposePatt = True
    ElseIf strTemp = "no" Or strTemp = "false" Then
        inst_info.enable_DecomposePatt = False
    Else
        inst_info.enable_DecomposePatt = True
        TheExec.Datalog.WriteComment "Argument: " & str_Set_DecomposePatt & " from this instance doesn't have the correct format to decide DecomposePat. Error!!!"
        TheExec.ErrorLogMessage "Argument: " & str_Set_DecomposePatt & " from this instance doesn't have the correct format to decide DecomposePat. Error!!!"
    End If
    
    '''//Decompose the pattern set to check if any SELSRM DSSC digsrc pattern exists in the pattern set. Requested by C651 Toby.
    '''PrePatt
    If PrePatt <> "" Then
        inst_info.PrePatt = PrePatt
        '''//Decompose PrePatt and FuncPatt to find DSSC pattern by BlockType / Pattern keyword for SELSRAM DSSC bit array according to SELSRM_Mapping_Table.
        Find_DsscPatt_fromPattSet inst_info, inst_info.PrePatt, inst_info.ary_PrePatt_decomposed, inst_info.count_PrePatt_decomposed, inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt, inst_info.idxBlock_Selsrm_PrePatt
    End If
    
    '''FuncPat
    If FuncPat <> "" Then
        inst_info.FuncPat = FuncPat
        '''//Decompose PrePatt and FuncPatt to find DSSC pattern by BlockType / Pattern keyword for SELSRAM DSSC bit array according to SELSRM_Mapping_Table.
        Find_DsscPatt_fromPattSet inst_info, inst_info.FuncPat, inst_info.ary_FuncPat_decomposed, inst_info.count_FuncPat_decomposed, inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat, inst_info.idxBlock_Selsrm_FuncPat
    End If
    
    '''//For pattern without SelSram DSSC keyword, pattern count "idxPatt_Selsrm_PrePatt" keeps as -1.
    If inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt <> "" And inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat <> "" Then
        If inst_info.idxBlock_Selsrm_PrePatt = inst_info.idxBlock_Selsrm_FuncPat Then
            inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt
            inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_PrePatt
        Else
            inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt
            inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_PrePatt
            TheExec.Datalog.WriteComment "idxBlock_Selsrm_PrePatt and idxBlock_Selsrm_FuncPat are different. Please check SelsrmPat for PrePatt and FuncPat. Error!!!"
            TheExec.ErrorLogMessage "idxBlock_Selsrm_PrePatt and idxBlock_Selsrm_FuncPat are different. Please check SelsrmPat for PrePatt and FuncPat. Error!!!"
        End If
    ElseIf inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt <> "" Then
        inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt
        inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_PrePatt
    ElseIf inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat <> "" Then
        inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat
        inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_FuncPat
    Else
        inst_info.patt_SelsrmDigSrc_single = ""
        TheExec.Datalog.WriteComment "PrePatt and FuncPat patsets don't contain any DSSC pattern."
    End If
    
    '''*********************************************************************************************************************************************'''
    '''inst_info.enable_DecomposePatt is refreshed as "False" at the beginning of each GradeSearch_XXX_VT instance.
    '''*********************************************************************************************************************************************'''
    '''//If flag of DecomposePatt = False, redim array of ary_PrePatt_decomposed and ary_FuncPat_decomposed as 0.
    If inst_info.enable_DecomposePatt = False Then '''without decomposing pattern sets
        If PrePatt <> "" And inst_info.count_PrePatt_decomposed > 0 Then
            ReDim inst_info.ary_PrePatt_decomposed(0)
            inst_info.ary_PrePatt_decomposed(0) = PrePatt
            inst_info.count_PrePatt_decomposed = 1
        End If
        
        If FuncPat <> "" And inst_info.count_FuncPat_decomposed > 0 Then
            ReDim inst_info.ary_FuncPat_decomposed(0)
            inst_info.ary_FuncPat_decomposed(0) = FuncPat
            inst_info.count_FuncPat_decomposed = 1
        End If
        
        '''//Set "result_mode = tlResultModeDomain" (return a unique pass/fail result for each module and time domain) if pattern bursted without decomposing pattern.
        '''20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
        inst_info.result_mode = tlResultModeDomain
        result_mode = tlResultModeDomain
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Check_and_Decompose_PrePatt_FuncPat"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Check_and_Decompose_PrePatt_FuncPat"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210416: Modified to check if DSSC DigSrc patterns of SELSRM and Harvest Core exist.
'20201102: Modified to add the argument "inst_info As Instance_Info".
'20201102: Modified to use "Public Type Instance_Info".
'20201027: Modified to check SELSRM DSSC digsrc pattern for pattern burst requested by C651 Toby.
'20200915: Modified to update the status of "inst_info.enable_DecomposePatt".
'20200915: Modified to rename the argument "Flag_DecomposePatt_from_InstanceArg" to "str_Set_DecomposePatt".
'20191205: Created for finding DSSC pattern in the pattern set.
Public Function Find_DsscPatt_fromPattSet(inst_info As Instance_Info, pattSet As String, ary_patt_decompose() As String, count_patt_decompose As Long, _
                                            patt_Selsrm_digsrc As String, idxBlock_Selsrm_Pattern As Integer)
    Dim i As Integer
    Dim idxBlock As Integer
On Error GoTo errHandler
    '''//Init
    patt_Selsrm_digsrc = ""
    idxBlock_Selsrm_Pattern = -1
    
    If pattSet <> "" Then
        '''//Decompose pattSet into the arrary "ary_patt_decompose".
        GetPatFromPatternSet CStr(pattSet), ary_patt_decompose, count_patt_decompose
        
        '''//If pattern count>0, use pattern-loop to find DSSC DigSrc pattern.
        If count_patt_decompose > 0 Then
            '''//Use pattern-loop.
            For i = 0 To count_patt_decompose - 1
                '''//Use pattern-loop to find DSSC DigSrc pattern of SELSRM by Pattern keyword for SELSRAM DSSC bit array according to SELSRM_Mapping_Table.
                For idxBlock = 0 To UBound(SelsramMapping)
                    If LCase(ary_patt_decompose(i)) Like LCase(SelsramMapping(idxBlock).Pattern) Then
                        patt_Selsrm_digsrc = ary_patt_decompose(i)
                        idxBlock_Selsrm_Pattern = idxBlock
                        Exit For
                    End If
                Next idxBlock
                
'''ToDo: Maybe it can use pattern-loop to find DSSC DigSrc pattern of Harvest Core here...
                '''If 1st DSSC pattern is found, exit for-loop.
                If patt_Selsrm_digsrc <> "" Then
                    Exit For
                End If
            Next i
        Else
            TheExec.Datalog.WriteComment "PatternSet: " & CStr(pattSet) & " contains no pattern. Error!!!"
            TheExec.ErrorLogMessage "PatternSet: " & CStr(pattSet) & " contains no pattern. Error!!!"
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Find_DsscPatt_fromPattSet"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201230: Modified to check if FuncPat <> "".
'20201102: Modified to add the argument "enable_DecomposePatt as boolean".
'20200520: Created to show the errorLogMessage if "burst=no" and "Decompose_Pattern=false".
Public Function Check_Pattern_NoBurst_NoDecompose(FuncPat As String, count_decompose_FuncPat As Long, enable_DecomposePatt As Boolean)
    Dim lastBurstPat As New SiteVariant
    Dim isGrp As New SiteBoolean
    Dim lastLabel As New SiteVariant
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''lastBurstPat: String or Variant. The pattern and a separated list of pattern files and groups or a pattern set name.
'''     This method returns the pattern set name to this parameter when the burst mode is enabled in the Pattern Sets sheet (Burst column is yes).
'''     When the burst mode is disabled (Burst column is no), this method returns the last executed pattern name to this parameter.
'''
'''isGrp: Boolean or Variant. Whether the last burst involved a pattern group, where:
'''     isGrp=True: The last burst involved a pattern group.
'''     isGrp=False: The last burst did not involve a pattern group.
'''//==================================================================================================================================================================================//'''
    If FuncPat <> "" And enable_DecomposePatt = False Then '''without decomposing pattern sets
        TheHdw.Digital.Patgen.ReadLastStart lastBurstPat, isGrp, lastLabel
        
        For Each site In TheExec.sites
            If count_decompose_FuncPat = 1 And lastBurstPat(site) <> FuncPat Then
                TheExec.Datalog.WriteComment "site:" & site & ", FuncPat:" & FuncPat & ", it is not pattern burst, but it is not decomposed for the instance. Error!!!"
                TheExec.ErrorLogMessage "site:" & site & ", FuncPat:" & FuncPat & ", it is not pattern burst, but it is not decomposed for the instance. Error!!!"
                Exit For
            End If
        Next site
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Check_Pattern_NoBurst_NoDecompose"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Check_Pattern_NoBurst_NoDecompose"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20201210: Modified to remove "IndexLevelPerSite As SiteLong" from the argument of the vbt function Get_PassBinNum_by_Step.
'20201030: Modified to use "Public Type Instance_Info".
'20200713: Modified to add the argument "PASSBINCUT as siteLong" for CurrentPassBinCutNum.
'20200525: Created to get PassBinNum by BinCut GradeSearch Step.
Public Function Get_PassBinNum_by_Step(inst_info As Instance_Info, PassBinCutCurrent As SiteLong, PassBinNum As SiteLong)
    Dim site As Variant
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''BinCut search (is_BinSearch = True)  : get PassBinNum from DYNAMIC_VBIN_IDS_ZONE(p_mode).PASSBINCUT(IndexLevelPerSite).
'''BinCut check  (is_BinSearch = False) : get PassBinNum from CurrentPassBinCutNum.
'''//==================================================================================================================================================================================//'''
    If inst_info.is_BinSearch = True Then '''BinCut search.
        For Each site In TheExec.sites
            PassBinNum(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current)
        Next site
    Else '''BinCut check.
        PassBinNum = PassBinCutCurrent
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Get_PassBinNum_by_Step"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200618: Created to "check powerDomain--> powerPin --> DC Spec specName".
'20200617: Modified to check if VDD_XXX_VOP_VAR or VDD_XXX_VAR of powerPin exists.
Public Function initDomain2DcSpecGrp(domainList As String, dict_Domain2DcSpecGrp As Dictionary)
    Dim i As Long
    Dim j As Long
    Dim powerDomain As String
    Dim powerPin As String
    Dim PinGroup As String
    Dim specName As String
    Dim specGrp As String
    Dim split_domainlist() As String
    Dim split_content() As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Check powerDomain--> powerPin --> DC Spec specName.
'''//==================================================================================================================================================================================//'''
    dictDomain2DcSpecGrp.RemoveAll
    split_domainlist = Split(domainList, ",")

    For i = 0 To UBound(split_domainlist)
        powerDomain = split_domainlist(i)
        PinGroup = ""
    
        '''//Check if powerDomain belongs to powerPin or pinGroup...
        If domain2pinDict.Exists(UCase(powerDomain)) = True Then
            PinGroup = VddbinDomain2Pin(powerDomain)
        ElseIf pin2domainDict.Exists(UCase(powerDomain)) = True Then
            PinGroup = powerDomain
        Else
            TheExec.Datalog.WriteComment powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"
            TheExec.ErrorLogMessage powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"
        End If
        
        If PinGroup <> "" Then
            '''//Get powerPins from powerDomain
            split_content = Split(PinGroup, ",")
            specGrp = ""
            
            For j = 0 To UBound(split_content)
                powerPin = split_content(j)
                specName = ""
                
                '''//Check if VDD_XXX_VOP_VAR or VDD_XXX_VAR exist...
                If dictPin2Dcspec.Exists(UCase(powerPin & "_VOP_VAR")) = True Then
                    specName = powerPin & "_VOP_VAR"
                ElseIf dictPin2Dcspec.Exists(UCase(powerPin & "_VAR")) = True Then
                    specName = powerPin & "_VAR"
                Else
                    specName = ""
                    TheExec.Datalog.WriteComment powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"
                    TheExec.ErrorLogMessage powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"
                End If
                
                If specName <> "" Then
                   If specGrp <> "" Then
                       specGrp = specGrp & "," & specName
                   Else
                       specGrp = specName
                   End If
                End If
            Next j
            
            '''dictDomain2DcspecGrp
            If dictDomain2DcSpecGrp.Exists(powerDomain) Then
                '''Do nothing
            Else
                '''//Update dictionary of Domain2Pin
                dictDomain2DcSpecGrp.Add powerDomain, specGrp
            End If
        End If
    Next i
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initDomain2DcSpecGrp"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initDomain2DcSpecGrp"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200618: Created for Domain2DcSpecGrp.
Public Function VddbinDomain2DcSpecGrp(vddbinDomain As String) As String
    Dim powerDomain As String
On Error GoTo errHandler
    powerDomain = UCase(vddbinDomain)

    If dictDomain2DcSpecGrp.Exists(powerDomain) Then
        VddbinDomain2DcSpecGrp = UCase(dictDomain2DcSpecGrp.Item(powerDomain))
    Else
        VddbinDomain2DcSpecGrp = "Domain_Error"
        TheExec.Datalog.WriteComment "Vddbin Domain=" & vddbinDomain & ", " & vddbinDomain & " doesn't exist in dictDomain2DcSpecGrp. Error!!!"
        'TheExec.ErrorLogMessage "Vddbin Domain=" & vddbinDomain & ", " & vddbinDomain & " doesn't exist in dictDomain2DcSpecGrp. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of VddbinDomain2DcSpecGrp"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of VddbinDomain2DcSpecGrp"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201012: Modified to check if alarmFail(site) is triggered or not.
'20200901: TER factory thought that pfAlways didn't cause "TheExec.sites(Site).LastTestResultRaw" issue...
'20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this.
'20200817: Modified to check if "TheExec.sites(site).LastTestResultRaw=tlResultFail".
'20200815: Modified to remove "BV_Pass(site)".
'20200815: Modified to use "TheExec.sites(site).LastTestResultRaw" to get testResult of Call Instance.
'20200812: Modified to use BinCut globalVariable to check HardIP pattern result for BinCut call instance.
'20200622: Created for Decide_PattPass_by_failFlag.
Public Function Decide_PattPass_by_failFlag(flagName As String, pattPass As SiteBoolean)
    Dim site As Variant
    Dim mySiteResult As Long
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check if "TheExec.sites(site).LastTestResultRaw=tlResultFail" for HardIP ELB vbt function "Meas_FreqVoltCurr_Universal_func" with "Call TheHdw.Patterns(Pat).Test(pfNever, 0)".
'''2. Check if flagState("F_BV_CALLINST") for HardIP ELB vbt function with "Call TheHdw.Patterns(Pat).Test(pfAlways, 0)".
'''
'''<LastTestResultRaw>
'''This property gets the ungated pass/fail status of the last test executed on the specified site based on the result column.
'''Read-only tlResultType.
'''tlResultFail: the test failed; tlResultNoTest: no test was available; tlResultPass: the test passed.
'''
'''<!!!Warning!!!>
'''1. "TheExec.Flow.LastFlowStepResult" has issues with "TheHdw.Patterns(Pat).test(pfAlways, 0)".
'''Please contact Teradyne factory/software team for this issue.
'''ToDo: Check if "TheExec.sites(site).LastTestResultRaw" and "TheExec.Flow.LastFlowStepResult" ready for "TheHdw.Patterns(Pat).test(pfAlways, 0)"...
'''2. For instance with pfAlways, it cau use failFlag or BV_Pass to get testResult about Pass/Fail.
'''3. Decide results of Pattern pass/fail and use-limit by failFlag of the instance and use-limit.
'''4. Remember to check if BV_Pass is used in LIB_HardIP\HardIP_WriteFuncResult.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        '''//Check if alarmFail(site) is triggered or not.
        If alarmFail(site) = True Then
            TheExec.Datalog.WriteComment "site:" & site & ", alarmFail!!!"
            pattPass(site) = False
        Else
            '''//*****************************************************************************************************************************************************************************//'''
            '''//Note:
            '''Check if "TheExec.sites(site).LastTestResultRaw=tlResultFail" for HardIP ELB vbt function "Meas_FreqVoltCurr_Universal_func" with "Call TheHdw.Patterns(Pat).Test(pfNever, 0)".
            '''Check if flagState("F_BV_CALLINST") for HardIP ELB vbt function with "Call TheHdw.Patterns(Pat).Test(pfAlways, 0)".
            '''//*****************************************************************************************************************************************************************************//'''
            '''//Get siteResult.
            'mySiteResult = TheExec.sites(site).LastTestResultRaw
            mySiteResult = TheExec.sites.Item(site).FlagState(flagName)
            
            '''//According to siteResult, update pattPass for each site.
            If TheExec.sites.Item(site).FlagState(flagName) = logicFalse Then 'And (mySiteResult = tlResultPass Or mySiteResult = tlResultNoTest) Then
                pattPass(site) = True
            ElseIf TheExec.sites.Item(site).FlagState(flagName) = logicTrue Or mySiteResult = tlResultFail Then
                '    If Flag_VDD_Binning_Offline = True Or TheExec.Flow.EnableWord("Vddbinning_OpenSocket") = True Then '''If the tester is offline or opensocket.
                '        theexec.Datalog.WriteComment "site:" & site & "," & flagName & "=" & CStr(CBool(theexec.sites.item(site).FlagState(flagName) = logicFalse)) & ",but it is forced to pass for offline or opensocket!!!"
                '        pattPass(site) = True
                '        TheExec.sites.Item(site).FlagState(flagName) = logicFalse
                '    Else
                    pattPass(site) = False
                '    End If
            End If
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Decide_PattPass_by_failFlag"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Decide_PattPass_by_failFlag"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
'20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
'20210701: Modified to reset gb_bincut_power_list().
'20210526: Modified to add "Flag_Get_column_Monotonicity_Offset" for Monotonicity_Offset check because C651 Si revised the check rules.
'20201222: Modified to reset the dictionary "dict_OutsideBinCut_additionalMode".
'20201112: Modified to reset the dictionary "dict_IsCorePowerInBinCutFlowSheet".
'20201110: Modified to reset dictionaries of powerBinning.
'20201103: Modified to set cntAdditionalMode = 0 for parsing Non_Binning_Rail.
'20201023: Modified to reset "dict_BinCutFlow_Domain2Column" and "dict_BinCutFlow_Column2Domain".
'20201021: Modified to reset "dict_IsCorePower".
'20200702: Modified to add "Flag_NonbinningrailOutsideBinCut_parsed=false".
'20200622: Created to reset BinCut globalVariable for initVddBinning.
'20191219: Modified to reset dictionaries of Domain2Pin and Pin2Domain.
Public Function Reset_BinCut_GlobalVariable_for_initVddBinning()
    Dim idx_powerDomain As Long
    Dim p_mode As Integer
    Dim addi_mode As Integer '''For the additional mode
    Dim passBinCut As Long
    Dim corePower As Long
On Error GoTo errHandler
    '''//Init variables
    Flag_PowerBinningTable_Parsed = False
    Flag_Interpolation_enable = False
    is_BinCutJob_for_StepSearch = False
    Flag_Harvest_Pmode_Table_Parsed = False
    Flag_Harvest_Mapping_Table_Parsed = False
    Flag_Harvest_Core_DSSC_Ready = False
    Flag_Get_column_Monotonicity_Offset = False
    gb_str_EfuseCategory_for_powerbinning = ""

    '''//Init power_list
    Power_List_All = ""
    
    '''//Init the flag for parsing Selsrm_Mapping_Table
    Flag_SelsrmMappingTable_Parsed = False
    
    '''//Init the flag for parsing Power_Binning_Harvest table
    Flag_Enable_PowerBinning_Harvest = False
    
    '''//Init the flag for parsing Non_Binning_Rail_Outside_BV
    Flag_NonbinningrailOutsideBinCut_parsed = False
    
    '''//Clear the dictionary of Pmode2enum, additional_mode2enum.
    VddbinPinDict.RemoveAll
    VddbinPmodeDict.RemoveAll
    AdditionalModeDict.RemoveAll
    dict_OutsideBinCut_additionalMode.RemoveAll
    
    '''//Init dictionaries of Domain2Pin and Pin2Domain.
    domain2pinDict.RemoveAll
    pin2domainDict.RemoveAll
    VddbinPinDcvstypeDict.RemoveAll
    dictPin2Dcspec.RemoveAll
    dictDomain2DcSpecGrp.RemoveAll
    dict_IsCorePower.RemoveAll
    dict_IsCorePowerInBinCutFlowSheet.RemoveAll
    
    '''//Init the pin_Groups
    dict_BinCutFlow_Domain2Column.RemoveAll
    dict_BinCutFlow_Column2Domain.RemoveAll
    FullCorePowerinFlowSheet = ""
    FullOtherRailinFlowSheet = ""
    FullBinCutPowerinFlowSheet = ""
    selsramLogicPin = ""
    selsramSramPin = ""
    selsramPin = ""
    
    '''//Define how many kinds of performance modes
    '''init the array of VddBinName
    ReDim VddBinName(0)
    ReDim AdditionalModeName(0)
    cntVddbinPin = -1
    cntVddbinPmode = -1
    cntAdditionalMode = 0
    VddBinName(0) = "None"
    AdditionalModeName(0) = "None"
    
    '''//Init the sheet dictionary for PowerBinning
    PwrBin_SheetnameDict.RemoveAll
    PwrBin_SpecIdx2SpecNameDict.RemoveAll
    dict_Binned_Mode_Column2Ratio.RemoveAll
    dict_Binned_Mode_Ratio2Column.RemoveAll
    dict_Binned_Mode_Ratio2Idx.RemoveAll
    dict_Binned_Mode_Column2Ratio.RemoveAll
    dict_Other_Mode_Ratio2Column.RemoveAll
    dict_Other_Mode_Ratio2Idx.RemoveAll
    
    '''//Clear the array gb_bincut_power_list to reset the list of all performance_modes in each powerDomain.
    '''20210701: Modified to reset gb_bincut_power_list().
    For idx_powerDomain = 0 To UBound(gb_bincut_power_list)
        gb_bincut_power_list(idx_powerDomain) = ""
    Next idx_powerDomain
    
    '''20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
    dict_strPmode2EfuseCategory.RemoveAll
    '''20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
    dict_EfuseCategory2BinCutTestJob.RemoveAll
    
    '''//Initialize the array of BV and HBV testConditions by empty string "".
    '''//The vbt function "initVddBinCondition" supported multiple "Non_Binning_Rail_Outside_BinCut" sheets.
    '''20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
    For p_mode = 0 To MaxPerformanceModeCount   '0~60
        For passBinCut = 0 To MaxPassBinCut         '0~3
            For corePower = 0 To MaxBincutPowerdomainCount
                BinCut(p_mode, passBinCut).OTHER_VOLTAGE(corePower) = ""
                BinCut(p_mode, passBinCut).HVCC_OTHER_VOLTAGE(corePower) = ""
                '''for OutsideBinCut sheet
                BinCut(p_mode, passBinCut).OutsideBinCut_OTHER_VOLTAGE(corePower) = ""
                BinCut(p_mode, passBinCut).OutsideBinCut_HVCC_OTHER_VOLTAGE(corePower) = ""
                
                For addi_mode = 0 To MaxAdditionalModeCount
                    BinCut(p_mode, passBinCut).Addtional_OTHER_VOLTAGE(corePower, addi_mode) = ""
                    BinCut(p_mode, passBinCut).HVCC_Addtional_OTHER_VOLTAGE(corePower, addi_mode) = ""
                    '''for OutsideBinCut sheet
                    BinCut(p_mode, passBinCut).OutsideBinCut_Addtional_OTHER_VOLTAGE(corePower, addi_mode) = ""
                    BinCut(p_mode, passBinCut).OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(corePower, addi_mode) = ""
                Next addi_mode
            Next corePower
        Next passBinCut
    Next p_mode
'''ToDo: Reset the siteDouble array of Previous_Payload_Voltage by powerDomain-loop.
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of Reset_BinCut_GlobalVariable_for_initVddBinning"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of Reset_BinCut_GlobalVariable_for_initVddBinning"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200703: Created to check range of row and column for the sheet.
Public Function check_Sheet_Range(sheetName As String, wb As Workbook, ws_def As Worksheet, MaxRow As Long, maxcol As Long, isSheetFound As Boolean)
On Error GoTo errHandler
    If Find_Sheet(sheetName) = True Then
        wb.Sheets(sheetName).Unprotect
        Set ws_def = wb.Sheets(sheetName)
        ws_def.Select
    
        '''//Check ranges of row and column
        MaxRow = ws_def.Cells.SpecialCells(xlCellTypeLastCell).row
        maxcol = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Column
        
        If MaxRow > 0 And maxcol > 0 Then
            isSheetFound = True
        Else
            isSheetFound = False
            MaxRow = 0
            maxcol = 0
            TheExec.Datalog.WriteComment "Content of " & sheetName & " is empty or incorrect. Error!!!"
            TheExec.ErrorLogMessage "Content of " & sheetName & " is empty or incorrect. Error!!!"
        End If
    Else
        isSheetFound = False
        MaxRow = 0
        maxcol = 0
        TheExec.Datalog.WriteComment sheetName & " doesn't exist in this workbook. Error!!!"
        TheExec.ErrorLogMessage sheetName & " doesn't exist in this workbook. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_Sheet_Range"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of check_Sheet_Range"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200709: Created to check if p_mode is tested.
Public Function isPmodeTested(p_mode As Integer) As Boolean
    Dim site As Variant
On Error GoTo errHandler
    '''init
    isPmodeTested = False
    
'''ToDo: VBIN_RESULT(p_mode).tested is siteBoolean. Maybe we can use any other method to check if p_mode is tested...
    For Each site In TheExec.sites
        If VBIN_RESULT(p_mode).tested(site) = True Then
            isPmodeTested = True
            Exit For
        End If
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of isPmodeTested"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of isPmodeTested"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201218: Modified to check if count_Pattern_decomposed > 0.
'20201208: Modified to merge the branches for inst_info.enable_PerEqnLog.
'20201030: Modified to use "Public Type Instance_Info".
'20201020: Modified to add the variables "COFInstance" and "PerEqnLog" for COFInstance.
'20201016: Created to decide the flag "Flag_Vddbin_COF_Instance".
Public Function decide_flag_for_COFInstance(inst_info As Instance_Info, count_Pattern_decomposed As Long)
On Error GoTo errHandler
    If Flag_Vddbin_COF_Instance = True And inst_info.is_BinSearch = True = True And inst_info.enable_DecomposePatt = True Then
        If Flag_IDS_Distribution_enable = True Then
            inst_info.enable_COFInstance = False
            TheExec.Datalog.WriteComment "COFInstance isn't compatible with IDS_Distribution_mode for GradeSearch_VT. Error!!!"
            TheExec.ErrorLogMessage "COFInstance isn't compatible with IDS_Distribution_mode for GradeSearch_VT. Error!!!"
        Else
            If count_Pattern_decomposed > 0 Then
                inst_info.enable_COFInstance = True
                ReDim Info_COFInstance(count_Pattern_decomposed - 1) '''use array size of the decoposed FuncPat.
            Else
                inst_info.enable_COFInstance = False
            End If
        End If
    Else
        inst_info.enable_COFInstance = False
    End If
    
    If inst_info.enable_COFInstance = True And Flag_Vddbin_COF_Instance_with_PerEqnLog = True Then
        inst_info.enable_PerEqnLog = True
    Else
        inst_info.enable_PerEqnLog = False
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of decide_flag_for_COFInstance"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of decide_flag_for_COFInstance"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "StepCount As Long" as "count_Step As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210125: Modified to remove "voltage_Binning_Pmode As SiteDouble" from the arguments of the vbt function "update_patt_result_for_COFInstance".
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for update_patt_result_for_COFInstance.
'20201016: Modfied to save EQN-based BinCut payload voltage of binning P_mode. Requested by C651 Si Li.
'20201015: Created to save result about pattern Pass/Fail for COFInstance.
Public Function update_patt_result_for_COFInstance(inst_info As Instance_Info, indexPatt As Long, Pattern As String, pattPass As SiteBoolean)
    Dim site As Variant
    Dim pattNameTemp As String
    Dim split_content() As String
On Error GoTo errHandler
    '''//Get pattern name while step_count=0.
    If inst_info.count_Step = 0 Then
        If Pattern Like "*\*" Then
            split_content = Split(UCase(Pattern), "\")
            pattNameTemp = split_content(UBound(split_content))
        Else
            pattNameTemp = UCase(Pattern)
        End If
        
        If split_content(UBound(split_content)) Like "*:*" Then
            split_content = Split(pattNameTemp, ":")
            pattNameTemp = split_content(0)
        End If
        
        If UCase(pattNameTemp) Like "*.PAT" Then
            split_content = Split(UCase(pattNameTemp), ".PAT")
            pattNameTemp = split_content(0)
        End If
        
        If inst_info.test_type = testType.Mbist Then '''Mbist instance records all patterns.
            Info_COFInstance(indexPatt).is_payload_pattern = True
            Info_COFInstance(indexPatt).Pattern = pattNameTemp
        Else '''TD/Scan instance only records patterns with keywords "*_pllp*", "*_fulp*", and "*_pl*".
            If LCase(pattNameTemp) Like "*_pllp*" Or LCase(pattNameTemp) Like "*_fulp*" Or LCase(pattNameTemp) Like "*_pl*" Then
                Info_COFInstance(indexPatt).is_payload_pattern = True
                Info_COFInstance(indexPatt).Pattern = pattNameTemp
            End If
        End If
        
        Info_COFInstance(indexPatt).grade_found = False
    End If

    If Info_COFInstance(indexPatt).is_payload_pattern = True Then
        For Each site In TheExec.sites
            If pattPass = True Then
                If Info_COFInstance(indexPatt).grade_found(site) = False Then
                    Info_COFInstance(indexPatt).grade_found(site) = True
                    Info_COFInstance(indexPatt).PassBin = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Current(site))
                    Info_COFInstance(indexPatt).EQN = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(inst_info.step_Current(site))
                    Info_COFInstance(indexPatt).Voltage = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.step_Current(site))
                End If
            Else
                If Info_COFInstance(indexPatt).grade_found(site) = True Then
                    Info_COFInstance(indexPatt).grade_found(site) = False
                    Info_COFInstance(indexPatt).PassBin = -1
                    Info_COFInstance(indexPatt).EQN = -1
                End If
            End If
        Next site
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of update_patt_result_for_COFInstance"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of update_patt_result_for_COFInstance"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201102: Modified to use "Public Type Instance_Info".
'20201016: Created to print info for COFInstance into the block "Judge_PF" in the datalog.
Public Function print_info_for_COFInstance(inst_info As Instance_Info)
    Dim site As Variant
    Dim powerDomain As String
    Dim indexPatt As Long
    Dim channel As String
On Error GoTo errHandler
    If inst_info.enable_COFInstance = True Then
        powerDomain = AllBinCut(inst_info.p_mode).powerPin
    
        For indexPatt = 0 To UBound(Info_COFInstance)
            If Info_COFInstance(indexPatt).is_payload_pattern = True Then
                For Each site In TheExec.sites
                    '''//PassBin
                    If Info_COFInstance(indexPatt).grade_found(site) = True Then
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, powerDomain, channel, 1, Info_COFInstance(indexPatt).PassBin, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_PASSBIN", scaleNone
                    Else
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestFail, parmTestLim, powerDomain, channel, 1, 0, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_PASSBIN", scaleNone
                    End If
                    TheExec.sites.Item(site).IncrementTestNumber
                    
                    '''//EQN
                    If Info_COFInstance(indexPatt).grade_found(site) = True Then
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, powerDomain, channel, 1, Info_COFInstance(indexPatt).EQN, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).Mode_Step + 1, _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_EQN", scaleNone
                    Else
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestFail, parmTestLim, powerDomain, channel, 1, 0, 1, _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_EQN", scaleNone
                    End If
                    
                    '''//BinCut voltage
'''ToDo: Maybe we can replace "_CP" with the keyword about the current testJob...
                    If Info_COFInstance(indexPatt).grade_found(site) = True Then
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestPass, parmTestLim, powerDomain, channel, 1, Info_COFInstance(indexPatt).Voltage / 1000, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_Vmax(0) / 1000, _
                                                    unitVolt, 0, unitVolt, 0, , , Info_COFInstance(indexPatt).Pattern & "_CP", scaleMilli, "%.4f"
                    Else
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.Item(site).TestNumber, logTestFail, parmTestLim, powerDomain, channel, 1, 0, 1, _
                                                    unitVolt, 0, unitVolt, 0, , , Info_COFInstance(indexPatt).Pattern & "_CP", scaleMilli, "%.4f"
                    End If
                    TheExec.sites.Item(site).IncrementTestNumber
                Next site
            End If
        Next indexPatt
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of print_info_for_COFInstance"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of print_info_for_COFInstance"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "stepcountMax As Long" as "maxStep As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210813: Modified to revise the vbt code for postBinCut or outsideBinCut instance names with keyword "*_binresult_" for the vbt initialize_inst_info and Apply_testcondition_InFlowSheet.
'20210809: Modified to check if Flag_Remove_Printing_BV_voltages = False for the vbt function initialize_inst_info.
'20210806: Modified to print the info about that the test instance is for BinCut search or check.
'20210805: Modified to update inst_info.is_BinSearch=True if testCondition for powerDomain of the binning p_mode contains the keyword "*evaluate*bin*".
'20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
'20210603: Modified to move inst_info.Pattern_Pmode and inst_info.By_Mode from initialize_inst_info to GradeSearch_XXX_VT.
'20210528: Modified to initalize inst_info.Pattern_Pmode and inst_info.By_Mode for Calculate_Harvest_Core_DSSC_Source.
'20210513: Modified to set inst_info.Harvest_Core_DigSrc_Pin and inst_info.Harvest_Core_DigSrc_SignalName.
'20210126: Modified to revise the vbt code for DevChar.
'20201217: Modified to initialize "inst_info.count_PrePatt_decomposed" and "inst_info.count_FuncPat_decomposed" as -1.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201111: Modified to initialize the siteDouble array "voltage_SelsrmBitCalc".
'20201102: Modified to add "enable_DecomposePatt" for DecomposePat.
'20201102: Modified to check if performance_mode<>"".
'20201030: Modified to move "Call Get_Pmode_Addimode_Testtype_fromInstance(inst_info)" into initialize_inst_info.
'20201029: Created to initialize inst_info for the instance.
Public Function initialize_inst_info(inst_info As Instance_Info, performance_mode As String)
    Dim site As Variant
    Dim idx_powerDomain As Integer
    Dim str_testCondition As String
    Dim split_content() As String
On Error GoTo errHandler
    inst_info.inst_name = TheExec.DataManager.instanceName

    If performance_mode <> "" Then
        inst_info.performance_mode = performance_mode
        inst_info.previousDcvsOutput = tlDCVSVoltageMain
        inst_info.currentDcvsOutput = tlDCVSVoltageMain
        inst_info.is_BV_Safe_Voltage_printed = False
        inst_info.is_BV_Payload_Voltage_printed = False
        inst_info.is_BinSearch = False                      '''True: BinSearch; False: Functional Test (Pass/Fail only).
        inst_info.enable_CMEM_collection = False
        inst_info.enable_COFInstance = False
        inst_info.enable_PerEqnLog = False
        inst_info.enable_DecomposePatt = False
        inst_info.maxStep = 0
        inst_info.ids_current = 0
        inst_info.PrePattPass = True
        inst_info.funcPatPass = True
        inst_info.sitePatPass = True
        inst_info.count_PrePatt_decomposed = -1
        inst_info.count_FuncPat_decomposed = -1
        '''//Since Step_Current=0 is 1st step in DYNAMIC_IDS_Zone, it should initialize "Step_Current = -1" prior to "finde_start_voltage".
        inst_info.step_Current = -1
        '''for DevChar.
        inst_info.is_DevChar_Running = TheExec.DevChar.Setups.IsRunning
        inst_info.Pattern_Pmode = ""
        inst_info.By_Mode = ""
        
        If inst_info.is_DevChar_Running = True Then
            inst_info.DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
            
            '''get_DevChar_Precondition
            If TheExec.DevChar.Results(inst_info.DevChar_Setup).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(inst_info.DevChar_Setup).startTime Like "0001/1/1*" Then
                inst_info.get_DevChar_Precondition = False
            Else
                inst_info.get_DevChar_Precondition = True
            End If
        Else
            inst_info.DevChar_Setup = ""
            inst_info.get_DevChar_Precondition = False
        End If
        
        '''20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
        For idx_powerDomain = 0 To MaxBincutPowerdomainCount
            inst_info.voltage_SelsrmBitCalc(idx_powerDomain) = 0
            inst_info.sram_Vth(idx_powerDomain) = 0
        Next idx_powerDomain
    
        For Each site In TheExec.sites
            inst_info.str_dynamic_offset(site) = ""
            inst_info.str_Selsrm_DSSC_Bit(site) = ""
            inst_info.str_Selsrm_DSSC_Info(site) = ""
        Next site
        
        '''//Get p_mode, addi_mode, testtype, and offsettestype from test instance and its argument.
        Call Get_Pmode_Addimode_Testtype_fromInstance(inst_info)
        
        '''//Check if testCondition of the binning powerDomain contains the keyword "*evaluate*bin*".
        '''20210805: Modified to update inst_info.is_BinSearch=True if testCondition for powerDomain of the binning p_mode contains the keyword "*evaluate*bin*".
        For Each site In TheExec.sites.Active
            str_testCondition = LCase(Trim(Get_BinCut_TestCondition(inst_info, VddBinStr2Enum(inst_info.powerDomain), CurrentPassBinCutNum(site))))
                        
            '''20210813: Modified to revise the vbt code for postBinCut or outsideBinCut instance names with keyword "*_binresult_" for the vbt initialize_inst_info and Apply_testcondition_InFlowSheet.
            '''*************************************************************************************'''
            '''//Keyword replacement of BinCut test condition of p_mode.
            '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
            '''Since PostBinCut_Voltage_Set_VT support BV and HBV tests, it should replace keyword of testCondition from BV with "bin result".
            '''//For the special case, ex: "900mV (MS003)", do not replace the keyword of testCondtion with with "*Bin*Result*".
            '''20181009: As the request from KTCHAN, he defined that postbincut instances must have the keyword "*_binresult_*".
            '''20210126: Modified to revise the vbt code for DevChar.
            '''20210302: Modified to optimize the keyword replacement to "M*### Bin Result".
            If is_BinCutJob_for_StepSearch = True Then '''only for postBinCut or outsideBincut in CP1.
                If (LCase(inst_info.inst_name) Like "*_binresult_*" And Not (str_testCondition) Like "*bin*result*" And Not (str_testCondition) Like "*#mv*") Or inst_info.is_DevChar_Running = True Then
                    split_content = Split(str_testCondition, " ")
                    
                    '''//Check if any correct keyword of performance_mode exists...
                    If VddbinPmodeDict.Exists(UCase(Trim(split_content(0)))) Then
                        '''//p_mode of non_binning CorePower.
                        If VBIN_RESULT(VddBinStr2Enum(UCase(split_content(0)))).tested = True Then
                            str_testCondition = split_content(0) & " " & "bin result"
                        End If
                    End If
                End If
            End If
            '''*************************************************************************************'''
            
            If str_testCondition <> "" Then
                If str_testCondition Like "*evaluate*bin*" Then
                    inst_info.is_BinSearch = True
                Else
                    inst_info.is_BinSearch = False
                End If
                Exit For
            Else
                inst_info.is_BinSearch = False
                TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                TheExec.Datalog.WriteComment "Instance: " & inst_info.inst_name & "," & inst_info.performance_mode & ",powerDomain:" & inst_info.powerDomain & ", testCondition is incorrect to determine if the instance is for BinCut search or check. Please check argument of the instance. Error!!!"
                TheExec.ErrorLogMessage "Instance: " & inst_info.inst_name & "," & inst_info.performance_mode & ",powerDomain:" & inst_info.powerDomain & ", testCondition is incorrect to determine if the instance is for BinCut search or check. Please check argument of the instance. Error!!!"
            End If
        Next site
        
        '''//Check if the test instance is for BinCut search or check.
        '''20210806: Modified to print the info about that the test instance is for BinCut search or check.
        '''20210809: Modified to check if Flag_Remove_Printing_BV_voltages = False for the vbt function initialize_inst_info.
        If Flag_Remove_Printing_BV_voltages = False Then
            If inst_info.is_BinSearch = True Then
                TheExec.Datalog.WriteComment "instance:" & inst_info.inst_name & "," & inst_info.performance_mode & ", the instance is for BinCut search"
            Else
                TheExec.Datalog.WriteComment "instance:" & inst_info.inst_name & "," & inst_info.performance_mode & ", the instance is for BinCut check"
            End If
        End If
    Else
        TheExec.Datalog.WriteComment "Instance: " & inst_info.inst_name & " doesn't have the correct performance_mode. Please check argument of the instance. Error!!!"
        TheExec.ErrorLogMessage "Instance: " & inst_info.inst_name & " doesn't have the correct performance_mode. Please check argument of the instance. Error!!!"
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initialize_inst_info"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initialize_inst_info"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210805: Modified to remove the redundant vbt function initialize_step_control since it initialized step_control.All_Site_Mask = 0 in the vbt function decide_binSearch_and_start_voltage.
'20201211: Created to initialize control flags from "inst_info" and "step_control" at the beginning of each step in step-loop.
Public Function initialize_control_flag_for_step_loop(inst_info As Instance_Info)
On Error GoTo errHandler
    inst_info.PrePattPass = True                    'initail the flag for init pattern
    inst_info.funcPatPass = True
    inst_info.sitePatPass = True
    inst_info.is_BV_Safe_Voltage_printed = False
    inst_info.is_BV_Payload_Voltage_printed = False
    inst_info.Grade_Not_Found_Mask = 0           'grade not found flag for all site
    inst_info.On_StopVoltage_Mask = 0            'already on stop voltage flag for all site
    inst_info.All_Patt_Pass = True               'initialize the flag for all sites.
    inst_info.AllSiteFailPatt = 0
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initialize_control_flag_for_step_loop"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to check inst_info.HarvestBinningFlag for HarvestBinning.
'20201218: Modified to remove "enable_CMEM_collection As Boolean" from the arguments of the vbt function "check_flag_to_enable_CMEM_collection".
'20201218: Modified to add "count_FuncPat_decomposed As Long" to the arguments of the vbt function "check_flag_to_enable_CMEM_collection".
'20201218: Modified to move "resize_CMEM_Data_by_pattern_number" from "decide_bincut_feature_for_stepsearch" to "check_flag_to_enable_CMEM_collection".
'20201217: Created to decide if BinCut features are OK to be enabled for BinCut stepSearch.
Public Function decide_bincut_feature_for_stepsearch(inst_info As Instance_Info, count_FuncPat_decomposed As Long, Optional CaptureSize As Long, Optional failpins As String)
On Error GoTo errHandler
    '''//Check the flag "Flag_Enable_CMEM_Collection" to enable CMEM collection if tester is online.
    '''If inst_info.enable_CMEM_Collection = True, check and decide CaptureSize, failpins, and PrintSize for CMEM.
    Call check_flag_to_enable_CMEM_collection(inst_info, Flag_Enable_CMEM_Collection, count_FuncPat_decomposed, CaptureSize, failpins)

    '''//Decide if it's OK to enable COFInstance. If that, redim array size to store payload patterns pass/fail.
    Call decide_flag_for_COFInstance(inst_info, count_FuncPat_decomposed)
    
    '''//Checkscript uses this info to check if BinCut new features are activated or not.
    '''20210901: Modified to check inst_info.HarvestBinningFlag for HarvestBinning.
    TheExec.Datalog.WriteComment "Instance_Condition" & ", COFInstance:" & CStr(inst_info.enable_COFInstance) & ", PerEqnLog:" & CStr(inst_info.enable_PerEqnLog) & _
                                    ", Enable_CMEM_Collection:" & CStr(inst_info.enable_CMEM_collection) & ", HarvestBinningFlag:" & inst_info.HarvestBinningFlag
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initialize_control_flag_for_step_loop"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210730: Modified to show the error message to users if p_mode is for BinCut search without any Efuse category, as requested by C651 Toby.
'20210701: Created to update AllBinCut(p_mode).used after parsing BinCut flow table and Efuse_BitDef_Table.
Public Function update_bincut_pmode_list()
    Dim idx_powerDomain As Long
    Dim powerDomain As String
    Dim split_content() As String
    Dim str_pmodeGroup_list As String
    Dim idx_pmode As Long
    Dim str_pmode_temp As String
    Dim str_Efuse_write_pmode As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''allbincut(p_mode).used is decided after parsing BinCut flow table and Efuse_BitDef_Table, and it can check if p_mode can be tested and fused for BinCut...
'''//==================================================================================================================================================================================//'''
    For idx_powerDomain = 0 To UBound(pinGroup_BinCut)
        '''//Get the BinCut powerDomain.
        powerDomain = pinGroup_BinCut(idx_powerDomain)
        
        If gb_bincut_power_list(VddBinStr2Enum(powerDomain)) <> "" Then
            '''init
            str_pmodeGroup_list = ""
        
            '''//Get array of performance modes for powerDomain.
            split_content = Split(gb_bincut_power_list(VddBinStr2Enum(powerDomain)), ",")
            
            '''//Check AllBinCut(p_mode).listed_in_Efuse_BDF and update AllBinCut(p_mode).used for all performance modes from BinCut flow table.
            '''This step can make sure that the performances exists in BinCut flow table and Efuse_BitDef_Table definitely.
            For idx_pmode = 0 To UBound(split_content)
                str_pmode_temp = split_content(idx_pmode)
                
                '''//Check if the performance mode is listed in Efuse_BitDef_Table.
                '''ToDo: Maybe we can add the option here to skip checking AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF...
                If AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF = True Then
                    '''//If that, Update AllBinCut(p_mode).used for the performance mode.
                    AllBinCut(VddBinStr2Enum(str_pmode_temp)).Used = True
                    
                    '''//Add the performance mode to the new pmode_list.
                    If str_pmodeGroup_list <> "" Then
                        str_pmodeGroup_list = str_pmodeGroup_list & "," & str_pmode_temp
                    Else
                        str_pmodeGroup_list = str_pmode_temp
                    End If
                End If
                
                '''//Check if p_mode for BinCut search has the dedicated Efuse category in the current testJob.
                '''20210730: Modified to show the error message to users if p_mode is for BinCut search without any Efuse category, as requested by C651 Toby.
                If AllBinCut(VddBinStr2Enum(str_pmode_temp)).is_for_BinSearch = True Then
                    str_Efuse_write_pmode = get_Efuse_category_by_BinCut_testJob("write", VddBinName(VddBinStr2Enum(str_pmode_temp)))
                    
                    If str_Efuse_write_pmode = "" Then
                        TheExec.Datalog.WriteComment str_pmode_temp & ", it is for BinCut search, but it doesn't have any Efuse category about the performance mode. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!"
                        TheExec.ErrorLogMessage str_pmode_temp & ", it is for BinCut search, but it doesn't have any Efuse category about the performance mode. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!"
                    End If
                End If
            Next idx_pmode
            
            '''//Update the new pmodeGroup_list to gb_bincut_power_list for the BinCut powerDomain.
            gb_bincut_power_list(VddBinStr2Enum(powerDomain)) = str_pmodeGroup_list
            
            '''*******************************************************************************************************************'''
            '''//Sort the Performance mode by MAX_ID to define the inherit sequence for different PowerDomain.
            '''//Enable the Performance mode by the Flow Table.
            '''*******************************************************************************************************************'''
            If gb_bincut_power_list(VddBinStr2Enum(powerDomain)) <> "" Then
                sort_power_seqence gb_bincut_power_list(VddBinStr2Enum(powerDomain)), BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
            Else
                BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq = Split(gb_bincut_power_list(VddBinStr2Enum(powerDomain)), ",")
            End If
            
            '''*******************************************************************************************************************'''
            '''Check the performance_mode, and determine its previous performance_mode from the power_seq for voltage inheritance.
            '''*******************************************************************************************************************'''
            InitVddBinInherit BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
        End If
    Next idx_powerDomain
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of update_bincut_pmode_list"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of update_bincut_pmode_list"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210831: Modified to remove the unused vbt code.
'20210720: Modified to revise the branches.
'20210706: Created to get Efuse category by BinCut testJob.
Public Function get_Efuse_category_by_BinCut_testJob(str_selector As String, str_keyword_EfuseCategory As String) As String
    Dim i As Long
    Dim idx_CurrentBinCutJob As Long
    Dim strAry_temp_EfuseCategory() As String
    Dim str_Efuse_temp As String
    Dim gotCorrectSelector As Boolean
On Error GoTo errHandler
    '''init
    str_Efuse_temp = ""
    
    '''//Check if str_selector is "read" or "write".
    If LCase(str_selector) = "read" Or LCase(str_selector) = "write" Then
        '''//Get BinCutJob definition for the current BinCut testJob.
        idx_CurrentBinCutJob = getBinCutJobDefinition(bincutJobName)
    
        '''//Get Efuse category for str_keyword_EfuseCategory.
        If dict_strPmode2EfuseCategory.Exists(UCase(str_keyword_EfuseCategory)) = True Then
            strAry_temp_EfuseCategory = dict_strPmode2EfuseCategory.Item(UCase(str_keyword_EfuseCategory))
            
            '''//Check the array of Efuse category related to str_keyword_EfuseCategory.
            For i = 0 To UBound(strAry_temp_EfuseCategory)
                '''//Check if any Efuse category is fused prior to the current BinCut testJob.
                If dict_EfuseCategory2BinCutTestJob.Exists(strAry_temp_EfuseCategory(i)) = True Then
                    If LCase(str_selector) = "read" Then
                        If idx_CurrentBinCutJob > dict_EfuseCategory2BinCutTestJob.Item(strAry_temp_EfuseCategory(i)) Then
                            If str_Efuse_temp <> "" Then
                                If dict_EfuseCategory2BinCutTestJob.Item(strAry_temp_EfuseCategory(i)) > dict_EfuseCategory2BinCutTestJob.Item(str_Efuse_temp) Then
                                    str_Efuse_temp = strAry_temp_EfuseCategory(i)
                                End If
                            Else '''If str_selected_EfuseCategory is empty...
                                str_Efuse_temp = strAry_temp_EfuseCategory(i)
                            End If
                        End If
                    ElseIf LCase(str_selector) = "write" Then
                        If idx_CurrentBinCutJob = dict_EfuseCategory2BinCutTestJob.Item(strAry_temp_EfuseCategory(i)) Then
                            str_Efuse_temp = strAry_temp_EfuseCategory(i)
                        End If
                    End If
                Else
                    str_Efuse_temp = ""
                    'TheExec.Datalog.WriteComment "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", it doesn't have any correct programming stage to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
                    'TheExec.ErrorLogMessage "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", it doesn't have any correct programming stage to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
                End If
            Next i
        Else
            str_Efuse_temp = ""
            TheExec.Datalog.WriteComment "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", it isn't a correct keyword to find the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
            TheExec.ErrorLogMessage "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", , it isn't a correct keyword to find the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
        End If
    Else
        str_Efuse_temp = ""
        TheExec.Datalog.WriteComment "str_selector:" & str_selector & ", it isn't 'read' or 'write' to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
        TheExec.ErrorLogMessage "str_selector:" & str_selector & ", it isn't 'read' or 'write' to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
    End If
    
    '''//Output the string.
    get_Efuse_category_by_BinCut_testJob = str_Efuse_temp
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of get_Efuse_category_by_BinCut_testJob"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of get_Efuse_category_by_BinCut_testJob"
    If AbortTest Then Exit Function Else Resume Next
End Function
