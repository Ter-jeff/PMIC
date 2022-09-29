using AutoTestSystem.Model;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AutoTestSystem.ViewModel
{
    public class ViewModelMain : INotifyPropertyChanged
    {
        private ObservableCollection<PathRow> _pathRows;

        private ObservableCollection<ProjectRow> _projectRows;

        public ViewModelMain()
        {
            ProjectRows = new ObservableCollection<ProjectRow>();
            PathRows = new ObservableCollection<PathRow>();
        }

        public ProjectRow SelectProject { get; set; }

        public ObservableCollection<ProjectRow> ProjectRows
        {
            get { return _projectRows; }
            set
            {
                if (_projectRows != value)
                {
                    _projectRows = value;
                    OnPropertyChanged();
                }
            }
        }

        public PathRow SelectPathRow { get; set; }

        public ObservableCollection<PathRow> PathRows
        {
            get { return _pathRows; }
            set
            {
                if (_pathRows != value)
                {
                    _pathRows = value;
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