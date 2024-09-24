using DevExpress.Spreadsheet;
using DevExpress.XtraPdfViewer;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Menu;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraSpreadsheet.Menu;
using QuanLyMayMoc.Common.Constant;
using QuanLyMayMoc.Common.Helper;
using QuanLyMayMoc.Model;
using StackExchange.Profiling.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyMayMoc
{
    public static class MyFunction
    {
        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();

            // If the destination directory doesn't exist, create it.       
            Directory.CreateDirectory(destDirName);

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string tempPath = Path.Combine(destDirName, file.Name);
                file.CopyTo(tempPath, false);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string tempPath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, tempPath, copySubDirs);
                }
            }
        }
        public static int xemTruocFileCoBan(FileViewModel file, Control parent = null, string extension = null)
        {
            if (parent != null)
            {
                foreach (Control ctrl in parent.Controls)
                {
                    ctrl.Dispose();
                }
            }

            if (extension is null)
                extension = Path.GetExtension(file.FileNameInDb);


            if (CommonConstants.GetFileDocExt().Contains(extension))
            {
                xemTruocWord(file, parent);
            }
            else if (CommonConstants.GetFileExcelExt().Contains(extension) || CommonConstants.GetFileCsvExt().Contains(extension))
            {
                xemTruocEXECL(file, parent);

            }
            else if (CommonConstants.GetFilePdfExt().Contains(extension))
            {
                xemTruocPDF(file, parent);
            }
            else if (CommonConstants.GetFileImgExt().Contains(extension))
            {
                xemTruocHINHANH(file, parent);

            }
            else
            {
                MessageShower.ShowInformation("Không hỗ trợ xem trước file này! Bạn có thể tải file về để xem trên máy tính!");
                return 1;
            }

            return 0;
        }
        public static void xemTruocWord(FileViewModel file, Control parent = null)
        {
            RichEditControl word;
            XtraForm_wordPreview form = null;
            if (parent is null)
            {
                form = new XtraForm_wordPreview();
                word = form.GetRec();
            }
            else
            {
                word = new RichEditControl();
                parent.Controls.Add(word);
            }
            word.Unit = DevExpress.Office.DocumentUnit.Centimeter;

            word.RemoveShortcutKey(Keys.S, Keys.Control);
            word.ModifiedChanged += (s, e) =>
            {
                if (!word.Modified)
                    word.Parent.Text = file.FileNameInDb;
                else
                    word.Parent.Text = "*" + file.FileNameInDb;
            };

            if (file.Table.HasValue())
            {
                word.PopupMenuShowing += (s, e) =>
                {
                    RichEditMenuItem myItem = new RichEditMenuItem("Lưu đè mẫu hiện tại",
                        new EventHandler((s1, e1) =>
                        {
                            //using (FileStream stream = new FileStream("D:\\Doc.xlsx", FileMode.Create, FileAccess.ReadWrite))
                            //{
                            if (!word.Modified)
                            {
                                MessageShower.ShowInformation("File chưa có sự thay đổi");
                                return;
                            }
                            else
                            {


                                word.Document.SaveDocument(file.FilePath, DevExpress.XtraRichEdit.DocumentFormat.OpenXml);

                                var crCRC = FileHelper.CalculateFileMD5(file.FilePath);
                                string dbString = $"UPDATE {file.Table} SET Checksum = '{crCRC}' WHERE Code = '{file.Code}'";
                                DataProvider.InstanceTHDA.ExecuteNonQuery(dbString);
                                file.Checksum = crCRC;
                                MessageShower.ShowInformation("Đã cập nhật");

                                return;
                            }
                            //}
                        }));

                    e.Menu.Items.Add(myItem);
                };
            }

            word.Dock = DockStyle.Fill;

            try
            {
                FileHelper.fcn_wordStreamDocument(word, file.FilePath);
            }
            catch
            {
                MessageShower.ShowInformation("Lỗi tải file");
            }

            //if (parent == null)
            //{


            //    using (var fm_xt = new Form())
            //    {
            //        fm_xt.WindowState = FormWindowState.Maximized;
            //        fm_xt.Controls.Add(word);
            //        fm_xt.Text = file.FileNameInDb;
            //        fm_xt.ShowDialog();
            //    }
            //}
            //else
            //{
            //    parent.Controls.Add(word);
            //    parent.Text = file.FileNameInDb;
            //}
            word.Parent.Text = file.FileNameInDb;
            if (form != null)
                form.ShowDialog();

        }
        public static void xemTruocEXECL(FileViewModel file, Control parent = null)
        {
            SpreadsheetControl spsheet;// = new SpreadsheetControl();
            XtraForm_ExcelPreview form = null;
            if (parent is null)
            {
                form = new XtraForm_ExcelPreview();
                spsheet = form.GetSpsheet();
            }
            else
            {
                spsheet = new SpreadsheetControl();
                parent.Controls.Add(spsheet);
            }

            spsheet.RemoveShortcutKey(Keys.S, Keys.Control);
            spsheet.ModifiedChanged += (s, e) =>
            {
                if (!spsheet.Modified)
                    spsheet.Parent.Text = file.FileNameInDb;
                else
                    spsheet.Parent.Text = "*" + file.FileNameInDb;
            };
            spsheet.Unit = DevExpress.Office.DocumentUnit.Centimeter;
            if (file.Table.HasValue())
            {
                spsheet.PopupMenuShowing += (s, e) =>
                {
                    SpreadsheetMenuItem myItem = new SpreadsheetMenuItem("Lưu đè mẫu hiện tại",
                        new EventHandler((s1, e1) =>
                        {
                            //using (FileStream stream = new FileStream("D:\\Doc.xlsx", FileMode.Create, FileAccess.ReadWrite))
                            //{
                            if (!spsheet.Modified)
                            {
                                MessageShower.ShowInformation("File chưa có sự thay đổi");
                                return;
                            }
                            else
                            {
                                spsheet.Document.SaveDocument(file.FilePath, DevExpress.Spreadsheet.DocumentFormat.OpenXml);


                                var crCRC = FileHelper.CalculateFileMD5(file.FilePath);
                                string dbString = $"UPDATE {file.Table} SET Checksum = '{crCRC}' WHERE Code = '{file.Code}'";
                                DataProvider.InstanceTHDA.ExecuteNonQuery(dbString);
                                file.Checksum = crCRC;
                                MessageShower.ShowInformation("Đã cập nhật");
                                return;
                            }
                            //}
                        }));

                    e.Menu.Items.Add(myItem);
                };
            }

            spsheet.Dock = DockStyle.Fill;

            try
            {
                FileHelper.fcn_spSheetStreamDocument(spsheet, file.FilePath);
            }
            catch
            {
                MessageShower.ShowInformation("Lỗi tải file");
            }

            spsheet.Parent.Text = file.FileNameInDb;
            if (form != null)
                form.ShowDialog();


        }
        //}
        //else
        //{
        //    var fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
        //    ctrl.LoadDocument(fs);
        //    fs.Close();
        //    fs.Dispose();
        //    ctrl.Dock = DockStyle.Fill;
        //    ctrl.BringToFront();
        //}
        public static Dictionary<string, string> fcn_getDicOfColumn(CellRange range, int row2getName = 0)
        {
            Worksheet ws = range.Worksheet;
            IWorkbook wb = ws.Workbook;
            Dictionary<string, string> dic = new Dictionary<string, string>();
            //Cập nhật Dic
            //CellRange range = wb.Range[rangeName];
            for (int i = range.LeftColumnIndex; i <= range.RightColumnIndex; i++)
            {
                if (!string.IsNullOrEmpty(ws.Rows[row2getName][i].Value.ToString()))
                    dic.Add(ws.Rows[row2getName][i].Value.ToString(), ws.Columns[i].Heading);
            }
            return dic;
        }

        public static void xemTruocPDF(FileViewModel file, Control parent = null)
        {
            PdfViewer pdf = new PdfViewer();

            pdf.Dock = DockStyle.Fill;

            if (parent == null)
            {

                try
                {
                    pdf.LoadDocument(file.FilePath);
                }
                catch
                {
                    MessageShower.ShowInformation("Lỗi tải file");
                }
                using (Form fm_xt = new Form())
                {
                    fm_xt.WindowState = FormWindowState.Maximized;
                    fm_xt.Text = file.FileNameInDb;
                    fm_xt.Controls.Add(pdf);
                    fm_xt.ShowDialog();

                }
            }
            else
            {
                try
                {
                    FileHelper.fcn_PdfStreamDoc(pdf, file.FilePath);
                }
                catch
                {
                    MessageShower.ShowInformation("Lỗi tải file");
                }
                parent.Controls.Add(pdf);
                parent.Text = file.FileNameInDb;
            }

        }
        public static void xemTruocHINHANH(FileViewModel file, Control parent = null)
        {

            PictureBox hinh = new PictureBox();
            hinh.SizeMode = PictureBoxSizeMode.StretchImage;
            hinh.Dock = DockStyle.Fill;
            try
            {

                FileHelper.fcn_ImageStreamDoc(hinh, file.FilePath);
            }
#pragma warning disable CS0168 // The variable 'ex' is declared but never used
            catch (Exception ex)
#pragma warning restore CS0168 // The variable 'ex' is declared but never used
            {
                MessageShower.ShowInformation("Lỗi tải file");
            }

            if (parent == null)
            {
                using (Form fm_xt = new Form())
                {
                    fm_xt.WindowState = FormWindowState.Maximized;
                    fm_xt.Text = file.FileNameInDb;
                    fm_xt.Controls.Add(hinh);
                    fm_xt.ShowDialog();
                }
            }
            else
            {
                parent.Controls.Add(hinh);
                parent.Text = file.FileNameInDb;
            }


        }

        public static void DirectoryDelete(string target_dir)
        {
            if (!Directory.Exists(target_dir))
                return;

            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DirectoryDelete(dir);
            }

            Directory.Delete(target_dir, false);
        }

        public static int GetIntFromString(string str)
        {
            string numStr = "";
            foreach (char c in str)
            {
                if (char.IsDigit(c))
                    numStr += c;
            }

            if (numStr == "")
                return -1;
            return int.Parse(numStr);
        }

        public static int GetRowCountFromStringRange(string str)
        {
            string[] lsStr = str.Split(':');
            if (lsStr.Length < 2)
                return -1;
            int TopRow = GetIntFromString(lsStr[0]);
            int BotRow = GetIntFromString(lsStr[1]);

            if (TopRow < 0 || BotRow < 0)
                return -1;
            return (BotRow - TopRow + 1);
        }
        public static string fcn_RemoveAccents(string text)
        {
            string[] arr1 = new string[] { "á", "à", "ả", "ã", "ạ", "â", "ấ", "ầ", "ẩ", "ẫ", "ậ", "ă", "ắ", "ằ", "ẳ", "ẵ", "ặ",
                                            "đ",
                                            "é","è","ẻ","ẽ","ẹ","ê","ế","ề","ể","ễ","ệ",
                                            "í","ì","ỉ","ĩ","ị",
                                            "ó","ò","ỏ","õ","ọ","ô","ố","ồ","ổ","ỗ","ộ","ơ","ớ","ờ","ở","ỡ","ợ",
                                            "ú","ù","ủ","ũ","ụ","ư","ứ","ừ","ử","ữ","ự",
                                            "ý","ỳ","ỷ","ỹ","ỵ",};
            string[] arr2 = new string[] { "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a",
                                            "d",
                                            "e","e","e","e","e","e","e","e","e","e","e",
                                            "i","i","i","i","i",
                                            "o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o",
                                            "u","u","u","u","u","u","u","u","u","u","u",
                                            "y","y","y","y","y",};
            for (int i = 0; i < arr1.Length; i++)
            {
                text = text.Replace(arr1[i], arr2[i]);
                text = text.Replace(arr1[i].ToUpper(), arr2[i].ToUpper());
            }
            return text;
        }
        public static string fcn_Array2listQueryCondition(IEnumerable<string> array)
        {
            if (array is null)
                return "";
            var newArr = array.Where(x => x != null).Select(x => $"'{x.Replace("'", "''")}'");
            return string.Join(", ", newArr);
        }
    }
}
