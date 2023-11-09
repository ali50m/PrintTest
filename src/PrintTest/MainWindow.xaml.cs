using System.ComponentModel;
using System.IO;
using System;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;

namespace PrintTest;

public partial class MainWindow
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void OnPrintClick(object sender, System.Windows.RoutedEventArgs e)
    {
        var fileDialog = new Microsoft.Win32.SaveFileDialog();
        fileDialog.Filter = "PDF|*.pdf";
        fileDialog.FileName = $"pdf_{DateTime.Now:yy-MM-dd_HH-mm-ss_fff}.pdf";

        if (fileDialog.ShowDialog(this) != true)
        {
            return;
        }

        var filePath = fileDialog.FileName;
        var desctiption = "A Simple Drawing";

        var svr = new System.Printing.LocalPrintServer();
        var queue = svr.GetPrintQueues().FirstOrDefault(_ => _.Name == "Microsoft Print to PDF");
        if (queue == null)
        {
            return;
        }

        var ticket = queue.DefaultPrintTicket;

        var streamXPS = new MemoryStream();
        using (var pack = Package.Open(streamXPS, FileMode.CreateNew))
        {
            using var doc = new XpsDocument(pack, CompressionOption.SuperFast);
            var writer = XpsDocument.CreateXpsDocumentWriter(doc);
            writer.Write(Canvas, ticket); //PrintVisual
        }
        streamXPS.Position = 0;

        //https://social.msdn.microsoft.com/Forums/vstudio/ja-JP/bbc1f202-b73a-4da7-839b-6945c020e9db/#answers
        XpsPrintHelper.Print(streamXPS, queue.Name, desctiption, false, filePath);
    }
}

public class XpsPrintHelper
{
    public static void Print(
        Stream stream,
        string printerName,
        string jobName,
        bool isWait,
        string outputFileName
    ) //ファイル名
    {
        if (stream == null)
            throw new ArgumentNullException("stream");
        if (printerName == null)
            throw new ArgumentNullException("printerName");

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
            ); //ファイル名

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
    ) //ファイル名
    {
        var result = StartXpsPrintJob(
            printerName,
            jobName,
            outputFileName //ファイル名
            ,
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
        const int INFINITE = -1;
        switch (WaitForSingleObject(completionEvent, INFINITE))
        {
            case WAIT_RESULT.WAIT_OBJECT_0:
                // Expected result, do nothing.
                break;
            case WAIT_RESULT.WAIT_FAILED:
                throw new Win32Exception();
            default:
                throw new Exception("Unexpected result when waiting for the print job.");
        }
    }

    private static void CheckJobStatus(IXpsPrintJob job)
    {
        job.GetJobStatus(out var jobStatus);
        switch (jobStatus.completion)
        {
            case XPS_JOB_COMPLETION.XPS_JOB_COMPLETED:
                // Expected result, do nothing.
                break;
            case XPS_JOB_COMPLETION.XPS_JOB_FAILED:
                throw new Win32Exception(jobStatus.jobStatus);
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
        [MarshalAs(UnmanagedType.LPArray)] byte[] printablePagesOn,
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
        string lpName
    );

    [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
    private static extern WAIT_RESULT WaitForSingleObject(IntPtr handle, int milliseconds);

    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);

    [Guid("0C733A30-2A1C-11CE-ADE5-00AA0044773D")] // This is IID of ISequenatialSteam.
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IXpsPrintJobStream
    {
        // ISequentualStream methods.
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
        void GetJobStatus(out XPS_JOB_STATUS jobStatus);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct XPS_JOB_STATUS
    {
        public UInt32 jobId;
        public Int32 currentDocument;
        public Int32 currentPage;
        public Int32 currentPageTotal;
        public XPS_JOB_COMPLETION completion;
        public Int32 jobStatus; // UInt32
    };

    private enum XPS_JOB_COMPLETION
    {
        XPS_JOB_IN_PROGRESS = 0,
        XPS_JOB_COMPLETED = 1,
        XPS_JOB_CANCELLED = 2,
        XPS_JOB_FAILED = 3
    }

    private enum WAIT_RESULT
    {
        WAIT_OBJECT_0 = 0,
        WAIT_ABANDONED = 0x80,
        WAIT_TIMEOUT = 0x102,
        WAIT_FAILED = -1 // 0xFFFFFFFF
    }
}
