using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OnenoteCapabilities
{
    public class CrashDumpWriter
    {
        private static class MINIDUMP_TYPE
        {
            public const int MiniDumpNormal = 0x00000000;
            public const int MiniDumpWithDataSegs = 0x00000001;
            public const int MiniDumpWithFullMemory = 0x00000002;
            public const int MiniDumpWithHandleData = 0x00000004;
            public const int MiniDumpFilterMemory = 0x00000008;
            public const int MiniDumpScanMemory = 0x00000010;
            public const int MiniDumpWithUnloadedModules = 0x00000020;
            public const int MiniDumpWithIndirectlyReferencedMemory = 0x00000040;
            public const int MiniDumpFilterModulePaths = 0x00000080;
            public const int MiniDumpWithProcessThreadData = 0x00000100;
            public const int MiniDumpWithPrivateReadWriteMemory = 0x00000200;
            public const int MiniDumpWithoutOptionalData = 0x00000400;
            public const int MiniDumpWithFullMemoryInfo = 0x00000800;
            public const int MiniDumpWithThreadInfo = 0x00001000;
            public const int MiniDumpWithCodeSegs = 0x00002000;
        }

        [DllImport("dbghelp.dll")]
        private static extern bool MiniDumpWriteDump(IntPtr hProcess,
            Int32 ProcessId,
            IntPtr hFile,
            int DumpType,
            IntPtr ExceptionParam,
            IntPtr UserStreamParam,
            IntPtr CallackParam);
        public static bool WriteFullDump(string fileName)
        {
            using(FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                using(System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess())
                {
                    return MiniDumpWriteDump(process.Handle,
                        process.Id,
                        fs.SafeFileHandle.DangerousGetHandle(),
                        MINIDUMP_TYPE.MiniDumpWithFullMemory,
                        IntPtr.Zero,
                        IntPtr.Zero,
                        IntPtr.Zero);
                }
            }
        }

        // NOTE: Probably in wrong class - but can move when it makes more sense.
        // NOTE: Add exception if installed twice.
        public static void InstallReportAndCreateCrashDumpUnhandledExceptionHandler()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(ReportAndCreateCrashDumpExceptionHandler);
        }

        private static void ReportAndCreateCrashDumpExceptionHandler(object sender, UnhandledExceptionEventArgs e)
        {
            var result = MessageBox.Show(e.ExceptionObject.ToString(), "Unhandled Exception - Write a crash dump?",
                MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                var tempFileName = Path.GetTempFileName() + "_onom.dmp";
                CrashDumpWriter.WriteFullDump(tempFileName);
                MessageBox.Show(tempFileName, "Crash Dump Location");
            }
            else
            {
                MessageBox.Show("Crash dump creation skipped");
            }
        }
    }
}