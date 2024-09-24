using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars.Alerter;
using DevExpress.XtraEditors;

namespace QuanLyMayMoc
{
    public static class SharedControls
    {
        public static Form Form { get; set; }
        public static AlertControl alertControl { get; set; } = null;
        public static SaveFileDialog saveFileDialog;
        public static SearchLookUpEdit slke_ThongTinDuAn;
    }
}
