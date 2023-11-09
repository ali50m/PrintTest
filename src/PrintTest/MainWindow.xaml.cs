using System.ComponentModel;
using System.IO;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO.Packaging;
using System.Linq;
using System.Printing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Xps.Packaging;
using Microsoft.Win32;

namespace PrintTest;

public partial class MainWindow
{
    private const string MicrosoftPrintToPdf = "Microsoft Print to PDF";

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

        if (fileDialog.ShowDialog(this) is false)
            return;

        var filePath = fileDialog.FileName;

        using var svr = new LocalPrintServer();
        using var queue = svr.GetPrintQueues()
            .FirstOrDefault(queue => queue.Name == MicrosoftPrintToPdf);
        if (queue is null)
            return;

        var ticket = queue.DefaultPrintTicket;

        using var streamXps = new MemoryStream();
        using (var pack = Package.Open(streamXps, FileMode.CreateNew))
        {
            using var doc = new XpsDocument(pack, CompressionOption.SuperFast);
            var writer = XpsDocument.CreateXpsDocumentWriter(doc);
            writer.Write(Canvas, ticket); //PrintVisual
        }
        streamXps.Position = 0;

        //https://social.msdn.microsoft.com/Forums/vstudio/ja-JP/bbc1f202-b73a-4da7-839b-6945c020e9db/#answers
        XpsPrintHelper.Print(streamXps, queue.Name, "Print to PDF Job", false, filePath);
    }
}

[SuppressMessage(
    "CodeQuality",
    "IDE0079:Remove unnecessary suppression",
    Justification = "<Pending>"
)]
public static class XpsPrintHelper
{
    public static void Print(
        Stream stream,
        string printerName,
        string jobName,
        bool isWait,
        string outputFileName
    )
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));
        if (printerName is null)
            throw new ArgumentNullException(nameof(printerName));

        var completionEvent = CreateEvent(IntPtr.Zero, true, false, null);
        if (completionEvent == IntPtr.Zero)
            throw new Win32Exception();

        try
        {
            StartJob(
                printerName,
                jobName,
                completionEvent,
                out var job,
                out var jobStream,
                outputFileName
            );

            CopyJob(stream, job, jobStream);

            if (isWait)
            {
                WaitForJob(completionEvent);
                CheckJobStatus(job);
            }
        }
        finally
        {
            if (completionEvent != IntPtr.Zero)
                CloseHandle(completionEvent);
        }
    }

    private static void StartJob(
        string printerName,
        string jobName,
        IntPtr completionEvent,
        out IXpsPrintJob job,
        out IXpsPrintJobStream jobStream,
        string outputFileName
    )
    {
        var result = StartXpsPrintJob(
            printerName,
            jobName,
            outputFileName,
            IntPtr.Zero,
            completionEvent,
            null,
            0,
            out job,
            out jobStream,
            IntPtr.Zero
        );
        if (result != 0)
            throw new Win32Exception(result);
    }

    private static void CopyJob(Stream stream, IXpsPrintJob job, IXpsPrintJobStream jobStream)
    {
        try
        {
            var buff = new byte[4096];
            while (true)
            {
                var read = (uint)stream.Read(buff, 0, buff.Length);
                if (read == 0)
                    break;

                jobStream.Write(buff, read, out var written);

                if (read != written)
                    throw new Exception("Failed to copy data to the print job stream.");
            }

            // Indicate that the entire document has been copied.
            jobStream.Close();
        }
        catch (Exception)
        {
            // Cancel the job if we had any trouble submitting it.
            job.Cancel();
            throw;
        }
    }

    private static void WaitForJob(IntPtr completionEvent)
    {
        const int infinite = -1;
        switch (WaitForSingleObject(completionEvent, infinite))
        {
            case WaitResult.WaitObject0:
                // Expected result, do nothing.
                break;
            case WaitResult.WaitFailed:
                throw new Win32Exception();
            default:
                throw new Exception("Unexpected result when waiting for the print job.");
        }
    }

    private static void CheckJobStatus(IXpsPrintJob job)
    {
        job.GetJobStatus(out var jobStatus);
        switch (jobStatus.Completion)
        {
            case XpsJobCompletion.XpsJobCompleted:
                // Expected result, do nothing.
                break;
            case XpsJobCompletion.XpsJobFailed:
                throw new Win32Exception(jobStatus.JobStatus);
            default:
                throw new Exception("Unexpected print job status.");
        }
    }

    [DllImport("XpsPrint.dll", EntryPoint = "StartXpsPrintJob")]
    private static extern int StartXpsPrintJob(
        [MarshalAs(UnmanagedType.LPWStr)] string printerName,
        [MarshalAs(UnmanagedType.LPWStr)] string jobName,
        [MarshalAs(UnmanagedType.LPWStr)] string outputFileName, //こいつ
        IntPtr progressEvent, // HANDLE
        IntPtr completionEvent, // HANDLE
        [MarshalAs(UnmanagedType.LPArray)] byte[]? printablePagesOn,
        uint printablePagesOnCount,
        out IXpsPrintJob xpsPrintJob,
        out IXpsPrintJobStream documentStream,
        IntPtr printTicketStream
    ); // This is actually "out IXpsPrintJobStream", but we don't use it and just want to pass null, hence IntPtr.

    [DllImport("Kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateEvent(
        IntPtr lpEventAttributes,
        bool bManualReset,
        bool bInitialState,
        string? lpName
    );

    [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
    private static extern WaitResult WaitForSingleObject(IntPtr handle, int milliseconds);

    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);

    [Guid("0C733A30-2A1C-11CE-ADE5-00AA0044773D")] // This is IID of ISequentialSteam.
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IXpsPrintJobStream
    {
        // ISequentialStream methods.
        void Read([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbRead);
        void Write([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbWritten);

        // IXpsPrintJobStream methods.
        void Close();
    }

    [Guid("5ab89b06-8194-425f-ab3b-d7a96e350161")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IXpsPrintJob
    {
        void Cancel();
        void GetJobStatus(out XpsJobStatus jobStatus);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct XpsJobStatus
    {
        public uint JobId;
        public int CurrentDocument;
        public int CurrentPage;
        public int CurrentPageTotal;
        public XpsJobCompletion Completion;
        public int JobStatus; // UInt32
    };

    [SuppressMessage("ReSharper", "UnusedMember.Local")]
    private enum XpsJobCompletion
    {
        XpsJobInProgress = 0,
        XpsJobCompleted = 1,
        XpsJobCancelled = 2,
        XpsJobFailed = 3
    }

    [SuppressMessage("ReSharper", "UnusedMember.Local")]
    private enum WaitResult
    {
        WaitObject0 = 0,
        WaitAbandoned = 0x80,
        WaitTimeout = 0x102,
        WaitFailed = -1 // 0xFFFFFFFF
    }
}
