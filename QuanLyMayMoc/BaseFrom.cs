using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMayMoc
{
    public static class BaseFrom
    {
        public static string m_path, m_resourceChatPath; //Thư mục làm việc - thư mục chứa mẫu WORD/EXCEL
        public static string m_tempPath { get; set; }
        public static string m_templatePath
        {
            get
            {
                return $@"{m_path}\Template";
            }
        }
        public static string m_TempFilePath { get; set; }
        private static string _crTempDATH { get; set; }
        public static string m_UnsavedPath { get; set; }
        public static string m_FullTempathDA
        {
            get => Path.Combine(m_tempPath, _crTempDATH);
        }
        public static bool THDAChanged { get; set; } = false;
        public static string m_crTempDATH
        {
            get => _crTempDATH;
            set
            {
                _crTempDATH = value;
                THDAChanged = false;
                //SharedControls.watcher.Changed += new FileSystemEventHandler(OnDirChanged);
                //SharedControls.watcher.Created += new FileSystemEventHandler(OnDirChanged);
                //SharedControls.watcher.Deleted += new FileSystemEventHandler(OnDirChanged);
                //SharedControls.watcher.Renamed += new RenamedEventHandler(OnDirChanged);
            }
        }//Thư mục tạm - Tên dự án tổng hợp
    }
}
