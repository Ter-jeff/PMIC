using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AutoTestSystem.UI.Site
{
    public class SiteRow : INotifyPropertyChanged
    {
        private string _site;
        public string Site
        {
            get { return _site; }
            set
            {
                if (_site != value)
                {
                    _site = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _select;
        public bool Select
        {
            get { return _select; }
            set
            {
                if (_select != value)
                {
                    _select = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyname = null)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }
    }
}