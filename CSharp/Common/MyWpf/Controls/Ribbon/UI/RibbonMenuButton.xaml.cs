using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
using System.Windows.Media;

namespace MyWpf.Controls.Ribbon.UI
{
    [ContentProperty("MenuItems")]
    public class RibbonMenuButton : ToggleButton
    {
        static RibbonMenuButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(RibbonMenuButton),
               new FrameworkPropertyMetadata(typeof(RibbonMenuButton)));
        }

        public ImageSource ImageSource
        {
            get { return (ImageSource)GetValue(ImageSourceProperty); }
            set { SetValue(ImageSourceProperty, value); }
        }

        public static readonly DependencyProperty ImageSourceProperty =
           DependencyProperty.Register("ImageSource", typeof(ImageSource), typeof(RibbonMenuButton), new PropertyMetadata(null));

        public string Header
        {
            get { return (string)GetValue(HeaderProperty); }
            set { SetValue(HeaderProperty, value); }
        }

        public static readonly DependencyProperty HeaderProperty =
              DependencyProperty.Register("Header", typeof(string), typeof(RibbonMenuButton), new PropertyMetadata(null));

        public ObservableCollection<DependencyObject> MenuItems
        {
            get { return _menuItems; }
            set { _menuItems = value; }
        }

        private ObservableCollection<DependencyObject> _menuItems = new ObservableCollection<DependencyObject>();

        protected override void OnClick()
        {
            ContextMenu.IsOpen = true;
        }

        public override void OnApplyTemplate()
        {
            ContextMenu = new ContextMenu();
            foreach (var menuItem in MenuItems)
            {
                if (menuItem is MenuItem)
                {
                    var item = (MenuItem)menuItem;
                    item.StaysOpenOnClick = true;
                    ContextMenu.Items.Add(item);
                    ContextMenu.PlacementTarget = this;
                    ContextMenu.Placement = PlacementMode.Bottom;
                }
            }
        }
    }
}
