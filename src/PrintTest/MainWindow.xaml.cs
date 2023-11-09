using System;
using System.Windows;
using Microsoft.Win32;

namespace PrintTest;

public partial class MainWindow
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void OnPrintClick(object sender, RoutedEventArgs e)
    {
        var fileDialog = new SaveFileDialog
        {
            Filter = "PDF|*.pdf",
            FileName = $"pdf_{DateTime.Now:yy-MM-dd_HH-mm-ss_fff}.pdf"
        };

        if (fileDialog.ShowDialog() is false)
            return;

        XpsPrintHelper.PrintVisual(Canvas, fileDialog.FileName);
    }
}
