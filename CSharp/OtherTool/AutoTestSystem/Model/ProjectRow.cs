using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AutoTestSystem.Model
{
    public class ProjectRow
    {
        private string _path;
        private string _project;

        public string Project
        {
            get { return _project; }
            set
            {
                if (_project != value)
                {
                    _project = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Path
        {
            get { return _path; }
            set
            {
                if (_path != value)
                {
                    _path = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyname = null)
        {
            if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
        }
    }
}