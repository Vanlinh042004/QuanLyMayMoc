using DevExpress.Data.Mask.Internal;
using QuanLyMayMoc.Common.Enums;
using QuanLyMayMoc.Common.Helper;
using StackExchange.Profiling.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMayMoc.Model
{
    public class FileSQLiteViewModel 
    {
        public string Code { get; set; } = "";
        public string CodeParent { get; set; }
        public string FileDinhKem { get; set; } = "";
        public string Checksum { get; set; } = "";
        public DateTime? Ngay { get; set; }
        public int State { get; set; } = 0;
        public string Link { get; set; }
        public int? SortId { get; set; }
        public bool ModifiedFromServer { get; set; } = false;
        public bool? Modified { get; set; } = false;
        public DateTime? LastChange { get; set; }
        public DateTime? CreatedOn { get; set; }
        public string ParentCode
        {
            get
            {
                return ParentCodeCustom ?? (((Link is null) ? nameof(FileLinkEnum.FileVatLy) : nameof(FileLinkEnum.FileLink)) + $"_{StateString}");
            }
        }
        public bool Checked { get; set; } = false;
        
        public string StateString
        {
            get
            {
                return ((FileStateEnum)State).GetEnumName();
            }
        }

        public string ParentCodeCustom { get; set; }

    }
}
