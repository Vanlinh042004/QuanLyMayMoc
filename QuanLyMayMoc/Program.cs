using QuanLyMayMoc.Common.Helper;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyMayMoc
{
    static class Program
    {
        public static string Root = AppDomain.CurrentDomain.BaseDirectory;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            CheckduplicateApp();
            var culture = CultureInfo.GetCultureInfo("vi-VN");
            CultureInfo.DefaultThreadCurrentCulture = culture;

            //Culture for UI in any thread
            CultureInfo.DefaultThreadCurrentUICulture = culture;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            initBaseFrom();
            Application.Run(new QuanLyMayMoc360());
        }
        private static void initBaseFrom()
        {
            BaseFrom.m_tempPath = AppDomain.CurrentDomain.BaseDirectory + @"\Temp";
            BaseFrom.m_UnsavedPath = AppDomain.CurrentDomain.BaseDirectory + @"\Unsaved";
            BaseFrom.m_TempFilePath = AppDomain.CurrentDomain.BaseDirectory + @"\TempFile";

            if (Directory.Exists(BaseFrom.m_TempFilePath))
            {
                MyFunction.DirectoryDelete(BaseFrom.m_TempFilePath);
            }
            Directory.CreateDirectory(BaseFrom.m_TempFilePath);
            //BaseFrom.m_templatePath = BaseFrom.m_path + @"\Template";

            if (!Directory.Exists(BaseFrom.m_tempPath))
                Directory.CreateDirectory(BaseFrom.m_tempPath);

            if (!Directory.Exists(BaseFrom.m_UnsavedPath))
                Directory.CreateDirectory(BaseFrom.m_UnsavedPath);


#if DEBUG
            BaseFrom.m_path = $@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\.."; //Chạy file excel
            BaseFrom.m_resourceChatPath = $@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\ChatBox\Resource\";
#else
            BaseFrom.m_path = $@"{AppDomain.CurrentDomain.BaseDirectory}";
                                    BaseFrom.m_resourceChatPath = $@"{BaseFrom.m_path}\Resource";

#endif
            AppDomain.CurrentDomain.SetData("DataDirectory", Path.Combine(BaseFrom.m_path, "Database"));


            //string dbString2 = $"INSERT INTO test (so) VALUES (' @So ')";
            //DataProvider.InstanceTBT.ExecuteNonQuery(dbString, new object[] { a });

            //fcn_init();
        }
        private static void CheckduplicateApp()
        {
            Process crProc = Process.GetCurrentProcess();
            Process[] procs = Process.GetProcesses();

            foreach (Process p in procs)
            {

                try
                {

                    if (p.Id != crProc.Id && (p.MainModule.FileName == crProc.MainModule.FileName))
                    {
                        MessageShower.ShowError("Phần mềm đã khởi chạy trước đó!");
                        Environment.Exit(0);
                    }

                }
                catch { continue; }
            }
        }
        private static void InitWin32Bit()
        {
            if (!File.Exists(Root + @"SDX.dll"))
                File.Copy(Root + @"\x86\SDX.dll", Root + @"SDX.dll");
            if (!File.Exists(Root + @"SDX.lib"))
                File.Copy(Root + @"\x86\SDX.lib", Root + @"SDX.lib");
            if (!File.Exists(Root + @"SecureDongle_Control.dll"))
                File.Copy(Root + @"\x86\SecureDongle_Control.dll", Root + @"SecureDongle_Control.dll");
        }
        private static void InitWin64Bit()
        {
            if (!File.Exists(Root + @"SDX.dll"))
                File.Copy(Root + @"\x64\SDX.dll", Root + @"SDX.dll");
            if (!File.Exists(Root + @"SDX.lib"))
                File.Copy(Root + @"\x64\SDX.lib", Root + @"SDX.lib");
            if (!File.Exists(Root + @"SecureDongle_Control.dll"))
                File.Copy(Root + @"\x64\SecureDongle_Control.dll", Root + @"SecureDongle_Control.dll");
        }
        private static void CheckOSWindow()
        {
            //If win11os
            if (Environment.OSVersion.Version.Build > 20000)
            {
                InitWin32Bit();
            }
            else
            {
                //Win 64bit
                if (IntPtr.Size == 8)
                {
                    InitWin64Bit();
                }
                else
                {
                    InitWin32Bit();
                }
            }
        }
    }
}
