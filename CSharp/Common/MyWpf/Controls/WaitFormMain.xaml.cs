using System.Windows;
using System.Windows.Input;

namespace MyWpf.Controls
{
    /// <summary>
    /// Interaction logic for WaitFormMain.xaml
    /// </summary>
    public partial class WaitFormMain : Window
    {
        public WaitFormMain()
        {
            InitializeComponent();
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            DragMove();
        }
    }
}
