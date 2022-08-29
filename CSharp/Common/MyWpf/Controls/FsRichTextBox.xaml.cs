using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace MyWpf.Controls
{
    /// <summary>
    ///     Interaction logic for FsRichTextBox.xaml
    /// </summary>
    public partial class FsRichTextBox
    {
        #region Dependency Property Declarations

        // Document property
        public static readonly DependencyProperty DocumentProperty =
            DependencyProperty.Register("Document", typeof(FlowDocument),
                typeof(FsRichTextBox), new PropertyMetadata(OnDocumentChanged));

        #endregion

        #region Constructor

        /// <summary>
        ///     Default constructor.
        /// </summary>
        public FsRichTextBox()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties

        /// <summary>
        ///     The WPF FlowDocument contained in the control.
        /// </summary>
        public FlowDocument Document
        {
            get { return (FlowDocument)GetValue(DocumentProperty); }
            set { SetValue(DocumentProperty, value); }
        }

        #endregion

        #region PropertyChanged Callback Methods

        /// <summary>
        ///     Called when the Document property is changed
        /// </summary>
        private static void OnDocumentChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            /* For unknown reasons, this method gets called twice when the 
             * Document property is set. Until we figure out why, we initialize
             * the flag to 2 and decrement it each time through this method. */

            // Initialize
            var thisControl = (FsRichTextBox) d;

            // Exit if this update was internally generated
            if (thisControl._internalUpdatePending > 0)
            {
                // Decrement flags and exit
                thisControl._internalUpdatePending--;
                return;
            }

            // Set Document property on RichTextBox
            thisControl.TextBox.Document = e.NewValue == null ? new FlowDocument() : (FlowDocument) e.NewValue;

            // Reset flag
            thisControl._textHasChanged = false;
        }

        #endregion

        #region Event Handlers

        /// <summary>
        ///     Invoked when the user changes text in this user control.
        /// </summary>
        private void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            // Set the TextChanged flag
            _textHasChanged = true;
            TextBox.ScrollToEnd();
        }

        #endregion

        #region Fields

        // Static member variables
        private int _internalUpdatePending;
        private bool _textHasChanged;

        #endregion
    }
}