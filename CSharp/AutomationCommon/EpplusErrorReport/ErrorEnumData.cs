namespace AutomationCommon.EpplusErrorReport
{
    public enum ErrorLevel
    {
        Warning = 1,
        Error = 2
    }

    public enum PmicErrorType
    {
        MissingPinName,
        MissingPattern,
        MissingRegister,
        FormatError,
    };

    public enum HardIpErrorType
    {
        //common
        Existential,
        FormatError,
        MissingNeededSheets,
        MissingHeader,
        UnrecognisedHeader,
        MissingHardIpSheet,
        //pin
        MissingHardIpDcPin,
        MissingPinName,
        MissingPatInfoPinInPinMap,
        //PatInfo
        DuplicatePatternInPatInfo,
        MissingSendBitOrSendBitStr,
        WrongSendInformation,
        WrongMeasSequence,
        //Scgh
        DuplicatePayloadInScgh,
        //Pattern
        MissingPatternInPatList,
        MissingPatternInScgh,
        MissingPatternInTestPlan,
        MissingPatternInPatInfo,
        MisPatternForMeasurement,
        PatternExistMeasurement,
        WrongPatternName,
        //Force condition
        WrongForceCondition,
        //Register
        WrongRegisterAssignment,
        DupBitName,
        DictionaryKey,
        //Misc Info
        MissingParameter,
        ManualItems,
        RepeatSubBlock,
        MissingSubBlock,
        //Meas pin
        WrongTotalMeasCount,
        WrongMeasCountInPatInfo,
        WrongMeasPinInPatInfo,
        WrongMeasType,
        WrongMeasContent,
        WrongMeasC,
        MissingPinSeq,
        PartialCorePower,
        WrongDiffPinLevel,
        //Limit
        WrongLimit,
        WrongLimitValue,
        OppositeLimit,
        MisMatchLimitUnit,
        MissingLimitUnit,
        MissingLimit,
        //Bin
        MissingBinNum,
        //GPIO
        MisMatchGpioPinGroup,
        //Others
        ConflictTName,
        MisVbtModule,
        MisCalculationParaDefine,
        IgxlVersion,
        DcDefault
    };

    public enum EFuseErrorType
    {
        Existential,
        FormatError,
        Business
    };

    public enum BasicErrorType
    {
        Existential,
        FormatError,
        FormatWarning,
        Warning,
        Business,
        MissingPinName
    };

    public enum PreActionErrorType
    {
        Existential,
        FormatError,
        Business,
        DuplicateVbtModule,
        PinGroupNotMatch
    };

    public enum PostActionErrorType
    {
        Existential,
        FormatError,
        Business
    };

    public enum MbistErrorType
    {
        ReferenceFileError,
        Existential,
        FormatError,
        Business
    };

    public enum ScanErrorType
    {
        Existential,
        FormatError,
        Business,
        NonUsedApplication
    }

    public enum EvsErrorType
    {
        Existential,
        FormatError,
        Business
    }

    public enum RtosErrorType
    {
        Existential,
        FormatError,
        Business
    }

    public enum MainFlowErrorType
    {
        Existential,
        FormatError,
        Business,
        FlowSheet,
        InstanceSheet
    }

    public enum RelayErrorType
    {
        Existential,
        FormatError,
        Business
    };

    public enum NWireErrorType
    {
        Existential,
        FormatError,
        Business
    };

    public enum BinCutErrorType
    {
        Existential,
        FormatError,
        Business,
        IdsMax,
        CpMax,
        CpMin,
        Inherit,
        EfuseBitDef,
        CpGb,
        Cp2Gb,
        FtRoomGb,
        FtHotGb,
        AllowEqual 
    };

    public enum DuplicateInstance
    {
        Duplicate
    };

    public enum UnusedDcCategory
    {
        UnusedDcCategory
    };

    public enum ZeroVoltageDcCategory
    {
        ZeroVoltageDcCategory
    };
}
