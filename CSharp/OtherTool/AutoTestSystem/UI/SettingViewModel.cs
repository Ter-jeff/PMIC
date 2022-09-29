using AutoTestSystem.UI.Enable;
using AutoTestSystem.UI.Site;
using System.Collections.ObjectModel;

namespace AutoTestSystem.UI
{
    public class SettingViewModel
    {
        public ObservableCollection<EnableRow> EnableRows = new ObservableCollection<EnableRow>();

        public ObservableCollection<SiteRow> SiteRows = new ObservableCollection<SiteRow>();
    }
}