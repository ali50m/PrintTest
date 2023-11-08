using System.Windows.Controls;

namespace PrintTest;

public partial class MainWindow
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void OnPrintClick(object sender, System.Windows.RoutedEventArgs e)
    {
        var printDialog = new PrintDialog();

        if (printDialog.ShowDialog() == true)
        {
            printDialog.PrintVisual(Canvas, "A Simple Drawing");
        }
    }
}
