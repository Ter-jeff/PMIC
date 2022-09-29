using MyWpf.Properties;
using System.Windows.Input;

namespace MyWpf.Controls.Ribbon.UI
{
    public static class RibbonCommands
    {
        public static RoutedUICommand OpenAppMenu =
            new RoutedUICommand(Resources.OpenAppMenuCommandTooltip, Resources.OpenAppMenuCommandName, typeof(Ribbon));

        public static RoutedUICommand CloseAppMenu =
            new RoutedUICommand(Resources.CloseAppMenuCommandTooltip, Resources.CloseAppMenuCommandName, typeof(Ribbon));

        public static RoutedUICommand BlendInRibbon =
            new RoutedUICommand(Resources.BlendInRibbonCommandTooltip, Resources.BlendInRibbonCommandName, typeof(Ribbon));

        public static RoutedUICommand OpenRibbonOptions =
            new RoutedUICommand(Resources.OpenRibbonOptionsCommandTooltip, Resources.OpenRibbonOptionsCommandName, typeof(Ribbon));

        public static RoutedUICommand AddQuickAccess =
            new RoutedUICommand(Resources.AddQuickAccessCommandTooltip, Resources.AddQuickAccessCommandName, typeof(Ribbon));

        public static RoutedUICommand RemoveQuickAccess =
            new RoutedUICommand(Resources.RemoveQuickAccessCommandTooltip, Resources.RemoveQuickAccessCommandName, typeof(Ribbon));

    }
}