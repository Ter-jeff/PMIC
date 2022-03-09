using AutomationCommon.DataStructure;

namespace PmicAutogen.InputPackages
{
    public class Response
    {
        private static InputPackageBase.WriteMessage _writeMessage;

        public static void Report(string message, MessageLevel messageLevel, int percentage)
        {
            _writeMessage(message, messageLevel, percentage);
        }

        public static void Initialize(InputPackageBase.WriteMessage writeMessage)
        {
            _writeMessage = writeMessage;
        }
    }
}