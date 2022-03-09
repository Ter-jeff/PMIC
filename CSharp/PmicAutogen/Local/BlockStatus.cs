using System.Collections.Generic;

namespace PmicAutogen.Local
{
    public class Status
    {
        public bool Down;
        public bool Enable;
    }

    public static class BlockStatus
    {
        public const string Basic = "Basic";
        public const string Scan = "Scan";
        public const string Mbist = "Mbist";
        public const string Otp = "OTP";
        public const string Vbt = "VBT";
        private static readonly Dictionary<string, Status> AutomationBlockStatus = new Dictionary<string, Status>();

        static BlockStatus()
        {
            SetAllDefault();
        }

        public static void SetAllDefault()
        {
            Create();
            SetDownDefault();
            SetEnableDefault();
        }

        public static void Create()
        {
            AutomationBlockStatus.Clear();

            AutomationBlockStatus.Add(Scan, new Status());
            AutomationBlockStatus.Add(Mbist, new Status());
            AutomationBlockStatus.Add(Basic, new Status());
            AutomationBlockStatus.Add(Vbt, new Status());
            AutomationBlockStatus.Add(Otp, new Status());
        }

        public static void SetEnableDefault()
        {
            foreach (var automationStatus in AutomationBlockStatus.Values) automationStatus.Enable = false;
        }

        public static void SetDownDefault()
        {
            foreach (var automationStatus in AutomationBlockStatus.Values)
                automationStatus.Down = false;
        }

        public static Status GetAutomationBlockStatus(string blockName)
        {
            if (AutomationBlockStatus.ContainsKey(blockName))
                return AutomationBlockStatus[blockName];
            return null;
        }

        public static void UpdateAutomationBlockStatus(PmicMainForm pmicMainForm)
        {
            GetAutomationBlockStatus(Basic).Down = pmicMainForm.button_Basic.Checked;
            GetAutomationBlockStatus(Scan).Down = pmicMainForm.button_Scan.Checked;
            GetAutomationBlockStatus(Mbist).Down = pmicMainForm.button_Mbist.Checked;
            GetAutomationBlockStatus(Otp).Down = pmicMainForm.button_Otp.Checked;
            GetAutomationBlockStatus(Vbt).Down = pmicMainForm.button_VBT.Checked;

            GetAutomationBlockStatus(Basic).Enable = pmicMainForm.button_Basic.Enabled;
            GetAutomationBlockStatus(Scan).Enable = pmicMainForm.button_Scan.Enabled;
            GetAutomationBlockStatus(Mbist).Enable = pmicMainForm.button_Mbist.Enabled;
            GetAutomationBlockStatus(Otp).Enable = pmicMainForm.button_Otp.Enabled;
            GetAutomationBlockStatus(Vbt).Enable = pmicMainForm.button_VBT.Enabled;
        }
    }
}