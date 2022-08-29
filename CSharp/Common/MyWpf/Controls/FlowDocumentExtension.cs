using System;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;
using Color = System.Drawing.Color;

namespace MyWpf.Controls
{
    public static class FlowDocumentExtension
    {
        public static void AppendText(this FlowDocument flowDocument, string text, Color color)
        {
            if (flowDocument.Dispatcher.CheckAccess())
                UpdateControl(flowDocument, text, color);
            else
                flowDocument.Dispatcher.Invoke(DispatcherPriority.Normal,
                    new Action<FlowDocument, string, Color>(UpdateControl), flowDocument, text, color);
        }

        private static void UpdateControl(FlowDocument flowDocument, string text, Color color)
        {
            var r = new Run(text);
            var brush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B));
            r.Foreground = brush;
            var paragraph = new Paragraph(r);
            flowDocument.Blocks.Add(paragraph);
        }
    }
}