using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PrintTest;

public partial class MainWindow
{
    private readonly PrintDocument _document;

    public MainWindow()
    {
        InitializeComponent();

        _document = new PrintDocument();
        _document.PrinterSettings.PrinterName = "Microsoft Print to PDF";
        _document.PrinterSettings.PrintToFile = true;
        _document.PrintPage += PrintPage;
    }

    private void PrintPage(object sender, PrintPageEventArgs e)
    {
        var source = Canvas.FrameworkElementToBitmapSource();
        var encoder = new BmpBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(source));
        using var ms = new MemoryStream();
        encoder.Save(ms);
        var image = Image.FromStream(ms);

        e.Graphics?.DrawImage(image, new System.Drawing.Point(0, 0));
    }

    private void OnPrintClick(object sender, RoutedEventArgs e)
    {
        _document.PrinterSettings.PrintFileName = $"pdf_{DateTime.Now:yy-MM-dd_HH-mm-ss_fff}.pdf";

        _document.Print();
    }
}

public static class FrameworkElementExtensions
{
    public static BitmapSource FrameworkElementToBitmapSource(this FrameworkElement element)
    {
        element.UpdateLayout();
        var width = element.ActualWidth;
        var height = element.ActualHeight;
        var dv = new DrawingVisual();
        using (var dc = dv.RenderOpen())
        {
            dc.DrawRectangle(new BitmapCacheBrush(element), null, new Rect(0, 0, width, height));
        }
        var rtb = new RenderTargetBitmap((int)width, (int)height, 96d, 96d, PixelFormats.Pbgra32);
        rtb.Render(dv);
        return rtb;
    }
}
