namespace PmicAutogen.Inputs.ScghFile.ProChar.Base
{
    public enum NopType
    {
        NonUsage, // Column[Usage] = 0
        BlankInit, // Can not find Init for some payload
        WrongTimeSet, // Can not find TimeSet
        NoUse // Don't use in PatternList.csv
    }
}