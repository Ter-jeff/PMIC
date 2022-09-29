using CommonLib.WriteMessage;
using System;

namespace PmicAutogen.InputPackages
{
    public class Response
    {
        private static InputPackageBase.WriteMessage _writeMessage;
        public static Progress<ProgressStatus> Progress { get; set; }

        public static void Report(string message, MessageLevel messageLevel, int percentage)
        {
            _writeMessage(message, messageLevel, percentage);
        }

        public static void Initialize(InputPackageBase.WriteMessage writeMessage)
        {
            _writeMessage = writeMessage;
        }

        public static void Report(string v)
        {
            //throw new NotImplementedException();
        }
    }
}