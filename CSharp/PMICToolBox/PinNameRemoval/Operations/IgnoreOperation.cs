using System.Text.RegularExpressions;

namespace PinNameRemoval.Operations
{
    class IgnoreOperation : Operation
    {
        public static new Regex RegexScript = new Regex(@"^\s*//", RegexOptions.IgnoreCase);
        public IgnoreOperation(string line) : base(line) { }
    }
}
