Attribute VB_Name = "LIB_EFUSE_Type_Syntax"
Option Explicit

Public Type EFuseCategoryParamResultSyntax
    'BitStrL() As String ''''(LSB...MSB), using dynamic array
    'BitStrM() As String ''''(MSB...LSB), using dynamic array
    BitStrL As New SiteVariant ''''20180726 update
    BitStrM As New SiteVariant ''''20180726 update
    
    Decimal As New SiteVariant
    BitSummation As New SiteLong
    
    Value As New SiteVariant
    'ValStr() As String  ''''using dynamic array
    ValStr As New SiteVariant ''''20180726 update
    
    ''''<NOTICE> 20160907 must be with '0x' to stand for the Hex-String
    ''''Then it's easy to identify as a meaningful HEX Value
    HexStr As New SiteVariant
    
    ''''20180625 Add
    BitArrWave As New DSPWave ''''[Definition] its .Element[0] is LSBbit value
End Type


''''-------------------------------------------------------
''''20151006, update for the New eFuse ChkList format
''''-------------------------------------------------------
Public Type EFuseCategoryParamSyntax ''''was Private
    index As Long
    Name As String
    ''''-----------------------------------------------------------------------
    SeqStart As Long    ''''change it to LSBBit, keep it to compatible existed
    SeqEnd As Long      ''''change it to MSBBit, keep it to compatible existed
    ''''-----------------------------------------------------------------------
    LSBbit As Long      ''''ECID:was SeqEnd,   CFG/UID/UDR/SEN: was SeqStart
    MSBbit As Long      ''''ECID:was SeqStart, CFG/UID/UDR/SEN: was SeqEnd
    ''''-----------------------------------------------------------------------
    BitWidth As Long
    
    ''''UDR: USI LSB-Bit Cycle   USI MSB-Bit Cycle   USO LSB-Bit Cycle   USO MSB-Bit Cycle
    USILSBCycle As Long
    USIMSBCycle As Long
    USOLSBCycle As Long
    USOMSBCycle As Long
    
    Stage As String          ''''parameter:: Programming Stage
    LoLMT As Variant
    HiLMT As Variant
    Resoultion As Double     ''''parameter:: IDS Resolution
    algorithm As String
    comment As String        ''''parameter:: Comment or Description
    Default_Real As String   ''''parameter:: Use Default or Real Value
    DefaultValue As Variant
    MSBFirst As String       ''''parameter:: At present, ECID is 'Y', UID/CFG/UDR/SEN is 'N'

    ''''20150625 New for HardIP pattern test result
    PatTestPass_Flag As New SiteBoolean
    
    ''''20151228 reserved for the limit application
    LoLMT_R As Variant
    HiLMT_R As Variant
    
    Read As EFuseCategoryParamResultSyntax
    Write As EFuseCategoryParamResultSyntax

    DefValBitArr() As Long      ''''Bit[0] is LSB Bit
    BitIndexWave As New DSPWave ''''index of each BitArrWave element
End Type


Public Type EFuseCategorySyntax
    Category() As EFuseCategoryParamSyntax ''''using dynamic array
End Type

''''----------------------------------------------------------------
'''' Type structure for the Config Table Sheet (Config_Table_appX)
''''

Private Type CondCategoryParamSyntax
    Name As String
    ''''-------------------------------------------------------------
    MSBbit As Long
    LSBbit As Long
    BitWidth As Long
    Stage As String
    comment As String
    ''''-------------------------------------------------------------
    HexStr As String
    Decimal As Variant  ''''is possible to over/equal 32bits
    BitStrM As String   ''''[MSB......LSB]
    BitVal() As Long    ''''using dynamic array BitVal[0] is LSB
End Type

Private Type CondCate32bitParamSyntax
    Name As String
    ''''-------------------------------------------------------------
    MSBbit As Long
    LSBbit As Long
    BitWidth As Long
    Stage As String
    comment As String
    ''''-------------------------------------------------------------
    HexStr As String
    Decimal As Variant  ''''is possible to over/equal 32bits
    BitStrM As String   ''''[MSB......LSB]
    BitVal() As Long    ''''using dynamic array BitVal[0] is LSB
    ''''-------------------------------------------------------------
    ''''New to get bits for the programming by Stage
    ''''-------------------------------------------------------------
    HexStr_byStage As String
    Decimal_byStage As Variant ''''is possible to over/equal 32bits
    BitStrM_byStage As String  ''''[MSB......LSB]
    BitVal_byStage() As Long   ''''using dynamic array BitVal[0] is LSB
    ''''-------------------------------------------------------------
End Type

Private Type ConfigCategoryParamSyntax
    row As Long ''''Cell Row
    col As Long ''''Cell Column
    pkgName As String
    FuseName As String
    BitStr() As String ''''using dynamic array
    BitVal() As Long   ''''using dynamic array
    BitStrM As String
    ''''-----------------------------------------------
    ''''New to get bits for the programming by Stage
    ''''-----------------------------------------------
    BitStr_byStage() As String ''''using dynamic array
    BitVal_byStage() As Long   ''''using dynamic array
    BitStrM_byStage As String
    ''''-----------------------------------------------
    ''''New for CFG_Condition_Table
    ''''-----------------------------------------------
    condition() As CondCategoryParamSyntax ''''20170630
    ''''-----------------------------------------------
    Cate32bit() As CondCate32bitParamSyntax ''''using dynamic array, 20170630, Category per 32bits
    ''''-----------------------------------------------
End Type

Public Type ConfigTableSyntax
    Category() As ConfigCategoryParamSyntax ''''using dynamic array
End Type
''''----------------------------------------------------------------
