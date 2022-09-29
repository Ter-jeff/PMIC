namespace PmicAutogen.Inputs.Setting.BinNumber
{
    public class BinNumberRuleCondition
    {
        private string _block;
        private string _condition;

        public BinNumberRuleCondition(EnumBinNumberBlock block, string condition)
        {
            Block = block.ToString();
            Condition = condition;
        }

        public string Block
        {
            set { _block = value; }
            get { return _block.Replace("_", "").Replace(" ", ""); }
        }

        public string Condition
        {
            set { _condition = value; }
            get { return _condition.Replace("_", "").Replace(" ", ""); }
        }
    }
}