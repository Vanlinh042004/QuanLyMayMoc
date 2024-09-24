using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMayMoc.Model
{
    public class ModelGiaLoiLinhKien
    {
        public string Code { get; set; }
        public string MaHieu { get; set; }
        public string Ten { get; set; }
        public string CodeMay { get; set; }
        public string CodeLinhKien { get; set; }
        public string TenMayCha { get; set; }
        public long Gia { get; set; }
    }
}
