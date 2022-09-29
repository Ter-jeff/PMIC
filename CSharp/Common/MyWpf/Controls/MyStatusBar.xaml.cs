using System.Windows;
using System.Windows.Controls;

namespace MyWpf.Controls
{
    /// <summary>
    /// Interaction logic for MyStatusBar.xaml
    /// </summary>
    public partial class MyStatusBar : UserControl
    {
        public MyStatusBar()
        {
            InitializeComponent();
        }

        public int ProgressBarValue
        {
            get { return (int)GetValue(ProgressBarValueProperty); }
            set { SetValue(ProgressBarValueProperty, value); }
        }

        public static readonly DependencyProperty ProgressBarValueProperty =
              DependencyProperty.Register("ProgressBarValue", typeof(int), typeof(MyStatusBar), new PropertyMetadata(null));

        public string StatusText
        {
            get { return (string)GetValue(StatusTextProperty); }
            set { SetValue(StatusTextProperty, value); }
        }

        public static readonly DependencyProperty StatusTextProperty =
              DependencyProperty.Register("StatusLabelText", typeof(string), typeof(MyStatusBar), new PropertyMetadata(null));

        public string ProcessTimeText
        {
            get { return (string)GetValue(ProcessTimeTextProperty); }
            set { SetValue(ProcessTimeTextProperty, value); }
        }

        public static readonly DependencyProperty ProcessTimeTextProperty =
              DependencyProperty.Register("ProcessTimeText", typeof(string), typeof(MyStatusBar), new PropertyMetadata(null));

    }
}
