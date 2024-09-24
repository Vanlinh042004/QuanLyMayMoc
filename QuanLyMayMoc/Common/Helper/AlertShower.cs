using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMayMoc.Common.Helper
{
    public class AlertShower
    {
        public static void ShowInfo(string message, string caption = null)
        {
            if (SharedControls.alertControl is null)
                return;
            caption = caption ?? "Thông báo";

            SharedControls.alertControl.Show(SharedControls.Form, caption, message);
        }
    }
}
