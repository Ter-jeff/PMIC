using MyWpf.Properties;
using System.Windows.Input;

namespace MyWpf.Controls.Ribbon.UI
{
    public static class WindowCommands
    {
        public static RoutedUICommand Help =
           new RoutedUICommand(Resources.HelpCommandTooltip, Resources.HelpCommandName, typeof(Ribbon));

        public static RoutedUICommand Minimize =
            new RoutedUICommand(Resources.MinimizeCommandTooltip, Resources.MinimizeCommandName, typeof(Ribbon));

        public static RoutedUICommand Maximize =
            new RoutedUICommand(Resources.MaximizeCommandTooltip, Resources.MaximizeCommandName, typeof(Ribbon));

        public static RoutedUICommand RestoreDown =
            new RoutedUICommand(Resources.RestoreDownCommandTooltip, Resources.RestoreDownCommandName, typeof(Ribbon));
    }
}