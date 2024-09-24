using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QuanLyMayMoc.Controls.NhanVien;
using System.IO;
using QuanLyMayMoc.Common.Helper;
using MSETTING = QuanLyMayMoc.Properties.Settings;
using PhanMemQuanLyThiCong.Common.Helper;
using QuanLyMayMoc.Common.Constant;
using DevExpress.Compression;
using DevExpress.Spreadsheet;
using QuanLyMayMoc.Model;

namespace QuanLyMayMoc
{
    public partial class QuanLyMayMoc360 : DevExpress.XtraEditors.XtraForm
    {
        public QuanLyMayMoc360()
        {
            InitializeComponent();
        }

        private void QuanLyMayMoc360_Load(object sender, EventArgs e)
        {
            MainLoading();
        }
        private void  MainLoading()
        {
            SharedControls.Form = this;
            SharedControls.saveFileDialog = saveFileDialog;
            SharedControls.slke_ThongTinDuAn = slke_ThongTinDuAn;
            CheckingFileUnsaved();
            fcn_init();
            this.Enabled = true;
        }
        private void fcn_initFile()
        {
            WaitFormHelper.ShowWaitForm("Đang khởi tạo file");
            FileHelper.fcn_spSheetStreamDocument(spred_DichVuTheoThang, $@"{BaseFrom.m_templatePath}\FileExcel\Template_DichVuThang.xlsx"); // Tên lưới excel trong bảng (tự đặt ví dụ spsheet_Thongtinchinh) và tên file excel đọc vào (ví dụ:1.aThongTinCoBan.xlsx). 
            FileHelper.fcn_spSheetStreamDocument(spread_GiaLoiLinhKien, $@"{BaseFrom.m_templatePath}\FileExcel\Template_GiaLoiLinhKien.xlsx"); // Tên lưới excel trong bảng (tự đặt ví dụ spsheet_Thongtinchinh) và tên file excel đọc vào (ví dụ:1.aThongTinCoBan.xlsx). 
            WaitFormHelper.CloseWaitForm();
        }
        private void fcn_init()
        {
            fcn_initFile();
            if (openFormUnsavedFile(true) != DialogResult.OK)
                fcn_initDA_CT();
        }
        private DialogResult openFormUnsavedFile(bool isFstTime = false)
        {
            XtraForm_NotYetSaveProject form = new XtraForm_NotYetSaveProject(isFstTime);
            form.OpenBackup = new XtraForm_NotYetSaveProject.DE_OPENBACKUP(OpenBackup);
            return form.ShowDialog();
        }
        private void OpenBackup(string file)
        {
            if (BaseFrom.THDAChanged)
            {
                var dr = MessageShower.ShowYesNoQuestion("Dự án hiện tại có thay đổi, bạn có muốn lưu lại trước khi mở file không?");
                if (dr == DialogResult.Yes)
                {
                    //goto BeginSave;
                }
                else if (dr == DialogResult.No)
                {
                    BaseFrom.THDAChanged = false;
                    //goto BeginCreate;
                }
                else
                {
                    return;
                }
            }
            //BeginSave:
            //fcn_saveDA(false);

            //BeginCreate:
            /*            if (!TaoMoiTongDuAn())
                        {
                            MessageShower.ShowWarning("Đã hủy thao tác tải về!");
                        }*/

            if (!fcn_openDAwithPath(file))
            {
                MessageShower.ShowError("Không thể mở file!");
                TaoMoiTongDuAn();
                return;
            }

            File.Delete(file);
            MSETTING.Default.PathHienTai = "";
            MSETTING.Default.Save();
        }
        private void CheckingFileUnsaved()
        {
            var dirs = Directory.GetDirectories(BaseFrom.m_tempPath);
            WaitFormHelper.ShowWaitForm("Đang sao lưu file chưa lưu");
            foreach (string dir in dirs)
            {
                try
                {
                    if (!Directory.Exists(BaseFrom.m_UnsavedPath))
                        Directory.CreateDirectory(BaseFrom.m_UnsavedPath);

                    //string lastChange = Directory.GetLastWriteTime(dir).ToString(MyConstant.CONST_DATE_FORMAT_SQLITE_WithTime);
                    string fileName = Path.GetFileName(dir);

                    string newFileGoc = Path.Combine(BaseFrom.m_UnsavedPath,
                                                    $"{fileName}_Backup{{0}}.qlmm");
                    string newFile = string.Format(newFileGoc, "");
                    string newTemp = Path.Combine(BaseFrom.m_tempPath, Path.GetFileNameWithoutExtension(newFile));
                    int count = 0;
                    while (File.Exists(newFile) || File.Exists(newTemp))
                    {
                        newFile = string.Format(newFileGoc, (count++).ToString());
                        newTemp = Path.Combine(BaseFrom.m_tempPath, Path.GetFileNameWithoutExtension(newFile));
                    }
                    Directory.Move(dir, newTemp);
                    TongHopHelper.fcn_SaveTHDAToQLTC(newFile);
                    //MyFunction.DeleteDirectory(dir);
                    MyFunction.DirectoryDelete(newTemp);
                }
                catch (Exception ex)
                {
                    AlertShower.ShowInfo($"{ex.Message}__Inner: {ex.InnerException?.Message}", "Lỗi check Unsaved File");
                    continue;
                }
            }
            WaitFormHelper.CloseWaitForm();
        }

        private void ToolStripTaoMoi_Click(object sender, EventArgs e)
        {
            TaoMoiTongDuAn();
        }
        private void fcn_AddThongTinKemDuAnMoi(string tenDuAn)
        {
            string codeDA = Guid.NewGuid().ToString();
            MSETTING.Default.DuAnHienTai = codeDA;
            MSETTING.Default.Save();
            MSETTING.Default.DuAnHienTai = codeDA;
            string ngayBD = DateTime.Now.ToString(MyConstant.CONST_DATE_FORMAT_SQLITE);
            string ngaykt = DateTime.Now.AddDays(30).ToString(MyConstant.CONST_DATE_FORMAT_SQLITE);
            string codeCT = Guid.NewGuid().ToString();
            string codeCTKHVT = Guid.NewGuid().ToString();
            string codeTTH = Guid.NewGuid().ToString();
            string SoHopDong = $"01-{DateTime.Now.Date.ToString("MM/dd/yyyy")}-HĐTTH";
            string db_string = $"INSERT INTO {MyConstant.Tbl_ThongTinDuAn} (\"Code\",\"TenDuAn\") " +
                $"VALUES ('{codeDA}','{tenDuAn}')";
            DataProvider.InstanceTHDA.ExecuteNonQuery(db_string);

        }
        private bool TaoDuAnMoi()
        {
            //if (!BaseFrom.IsFullAccess && !BaseFrom.allPermission.HaveInitProjectPermission
            //    && !BaseFrom.allPermission.HaveCreateProjectPermission)
            //{
            //    MessageShower.ShowError("Bạn không có quyền khởi tạo dự án!");
            //    return false;
            //}
            WaitFormHelper.ShowWaitForm("Đang tạo dự án mới");
            IWorkbook workbook = spred_DichVuTheoThang.Document;

            MSETTING.Default.Save();

            int STT_TempDA = 0;
            bool isExist = false;//Kiểm tra số thứ tự công trình tạm
            string db_string = $"Select \"TenDuAn\" from {MyConstant.Tbl_ThongTinDuAn}";
            DataTable dt = DataProvider.InstanceTHDA.ExecuteQuery(db_string);
            do
            {
                isExist = false;
                STT_TempDA++;
                if (dt.AsEnumerable().Where(x => x.Field<string>("TenDuAn") == $"Dự án {STT_TempDA}").Count() > 0)
                    isExist = true;

            } while (isExist);

            string tenDuAn = $"Dự án {STT_TempDA}";

            fcn_AddThongTinKemDuAnMoi(tenDuAn);
            TongHopHelper.fcn_updateCbbThongTinDuAnCongTrinh();

            //if (xtraTabPage_ThongTinDuAn.PageVisible)
            //{
            //    xtraTabControl_TabMain.SelectedTabPage = xtraTabPage_ThongTinDuAn;
            //    if (xtraTab_ThongTin.PageVisible)
            //        xtraTabControl_ThongTinDA.SelectedTabPage = xtraTab_ThongTin;
            //}
            WaitFormHelper.CloseWaitForm();
            return true;


        }
        private bool TaoMoiTongDuAn()
        {
            if (!TongHopHelper.CheckSaveDA())
                return false;

            MSETTING.Default.DuAnHienTai = string.Empty;
            //MSETTING.Default.PathHienTai = string.Empty;
            MSETTING.Default.Save();

            try
            {
                MyFunction.DirectoryDelete($@"{BaseFrom.m_tempPath}\{BaseFrom.m_crTempDATH}");
            }
            catch (Exception ex)
            {
                MessageShower.ShowInformation("Không thể xóa thư mục" + ex.ToString());
            }
            MSETTING.Default.PathHienTai = "";
            MSETTING.Default.Save();
            fcn_initDA_CT();
            return true;
        }
        private void fcn_initDA_CT()
        {
            if (MSETTING.Default.PathHienTai == "" || !File.Exists(MSETTING.Default.PathHienTai))
            {
                MSETTING.Default.PathHienTai = "";
                MSETTING.Default.Save();

                int STT_TempDA = 0;
                bool isExist = false;//Kiểm tra số thứ tự công trình tạm
                do
                {
                    isExist = Directory.Exists($@"{BaseFrom.m_tempPath}\TongHopDuAn_{++STT_TempDA}");
                } while (isExist);
                //m_crFileDA = "";
                string crTempDATH = $"TongHopDuAn_{STT_TempDA}";


                Directory.CreateDirectory($@"{BaseFrom.m_tempPath}\{crTempDATH}");
                MyFunction.DirectoryCopy($@"{BaseFrom.m_templatePath}\DuAnMau", $@"{BaseFrom.m_tempPath}\{crTempDATH}", true);
                DataProvider.InstanceTHDA.changePath($@"{BaseFrom.m_tempPath}\{crTempDATH}\{MyConstant.CONST_DbFromPathDA}");
                BaseFrom.m_crTempDATH = crTempDATH;
                this.Text = $"{crTempDATH}-Phần mềm quản lý máy móc";                ;
                //fcn_AddThongTinKemDuAnMoi("Dự án 1");
                de_ThangTinhLuong.DateTime = DateTime.Now;

                TongHopHelper.fcn_updateCbbThongTinDuAnCongTrinh();
                //tabGIAODIENCHINHQLTC.SelectedIndex = 2;
            }
            else
            {
                fcn_openDAwithPath(MSETTING.Default.PathHienTai);
            }

        }
        private bool fcn_openDAwithPath(string pathFile)
        {
            try
            {
                MSETTING.Default.PathHienTai = pathFile;
                MSETTING.Default.Save();
                //this.Text = Path.GetFileName(pathFile);
                string fileName = Path.GetFileNameWithoutExtension(pathFile);
                string fullPath = $@"{BaseFrom.m_tempPath}\{fileName}";


                MyFunction.DirectoryDelete(fullPath);
                //slke_ThongTinDuAn.Properties.DataSource = null;
                //cbb_DauViecLon.DataSource = null;
                //cbb_DauViecNho.DataSource = null;
                //cbb_MenuCongTrinhThucHien.Items.Clear();
                //m_crFileDA = MSETTING.Default.PathHienTai = mopenFileDialog.FileName;
                //MSETTING.Default.Save();
                using (ZipArchive archive = ZipArchive.Read(pathFile))
                {
                    //DevExpress.Compression.EncryptionType encryptionType = DevExpress.Compression.EncryptionType.PkZip;
                    //archive.EncryptionType = encryptionType;
                    //archive.Password = m_pw;
                    //if (!Directory.Exists($@"{BaseFrom.m_tempPath}\{fileName}"))
                    //    Directory.CreateDirectory($@"{BaseFrom.m_tempPath}\{fileName}");
                    foreach (ZipItem item in archive)
                    {

                        item.Extract(fullPath);
                    }
                    BaseFrom.m_crTempDATH = fileName;
                    //fcn_getPrvDA();
                    //fcn_getPrvDA();
                    //fcn_loadAllSpSheet();
                    //fcn_updateThongTinCongTrinh();
                    //fcn_updateThongTinHangMuc();

                }
                DataProvider.InstanceTHDA.changePath($@"{BaseFrom.m_tempPath}\{BaseFrom.m_crTempDATH}\{MyConstant.CONST_DbFromPathDA}");
                //var GiaoViecCha = MyFunction.Fcn_CalKLKHNew(TypeKLHN.GiaoViecCha);
                //var GiaoViecCon = MyFunction.Fcn_CalKLKHNew(TypeKLHN.GiaoViecCon);
                //var CongTac = MyFunction.Fcn_CalKLKHNew(TypeKLHN.CongTac);
                //var HaoPhiVatTu = MyFunction.Fcn_CalKLKHNew(TypeKLHN.HaoPhiVatTu);
                //var VatLieu = MyFunction.Fcn_CalKLKHNew(TypeKLHN.VatLieu);



                //DataProvider.InstanceTHDA.ExecuteNonQuery($"UPDATE {MyConstant.TBL_THONGTINTHDA} SET \"Tentonghopduan\"='{BaseFrom.m_crTempDATH}'");
                TongHopHelper.fcn_updateCbbThongTinDuAnCongTrinh();
                this.Text = $"{BaseFrom.m_crTempDATH}-Phần mềm quản lý máy móc";
                return true;
            }
            catch (Exception ex)
            {
                string err = $"{ex.Message}_Inner: {ex.InnerException?.Message}";
                AlertShower.ShowInfo(err, "Lỗi mở dự án!");
                MSETTING.Default.PathHienTai = "";
                MSETTING.Default.Save();
                fcn_initDA_CT();
                return false;

            }
        }
        private void ToolStripMoFile_Click(object sender, EventArgs e)
        {
            if (!TongHopHelper.CheckSaveDA(false))
                return;

            fcn_openDA();
        }
        private bool fcn_openDA()
        {
            openFileDialog.DefaultExt = "qlmm";
            openFileDialog.Filter = "TBT QLMM (*.qlmm)|*.qlmm";
            openFileDialog.AddExtension = false;
            DialogResult rs = openFileDialog.ShowDialog();

            if (rs == DialogResult.Cancel)
                return false;

            try
            {
                MyFunction.DirectoryDelete($@"{BaseFrom.m_tempPath}\{BaseFrom.m_crTempDATH}");
            }
            catch (Exception ex)
            {
                MessageShower.ShowInformation("Không thể xóa thư mục" + ex.ToString());
            }
            MSETTING.Default.PathHienTai = "";
            MSETTING.Default.Save();

            fcn_openDAwithPath(openFileDialog.FileName);
            return true;
        }
        private void lưuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TongHopHelper.fcn_saveDA(false);
        }

        private void lưuMớiToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cácDựÁnChưaĐượcLưuLạiToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void kiểmTraWaitformToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void đóngToànBộMànHìnhChờToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void sửaLỗiThứTựHạngMụcToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btn_TaoDuAnMoi_Click(object sender, EventArgs e)
        {
            TaoDuAnMoi();
        }

        private void xTabMain_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if(e.Page== TageTenNhanVien)
            {
                danhSachNhanVien.Fcn_Update();
            }
            else if(e.Page== TagDichVuTheoThang)
            {
                Fcn_LoadDataDichVuTheoThang();
            }
        }
        private void Fcn_LoadDataLinhKien()
        {
            string dbString = $"SELECT GM.*,MAY.Ten as TenMayCha " +
        $" FROM {MyConstant.Tbl_MTC_GiaLinhKien} GM" +
        $" LEFT JOIN {MyConstant.Tbl_MTC_TenLinhKien} MAY ON MAY.Code=GM.CodeLinhKien ";
            List<ModelGiaLoiLinhKien> InforMay = DataProvider.InstanceTHDA.ExecuteQueryModel<ModelGiaLoiLinhKien>(dbString);
            IWorkbook wb = spread_GiaLoiLinhKien.Document;
            Worksheet ws = wb.Worksheets[1];
            Worksheet SheetMau = wb.Worksheets["SheetMau"];
            CellRange MauMay = SheetMau.Range["DataMauLinhKien"];
            CellRange DataLinhKien = ws.Range["DataLinhKien"];
            wb.BeginUpdate();
            if (DataLinhKien.RowCount > 5)
            {
                ws.Rows.Remove(DataLinhKien.TopRowIndex + 1, DataLinhKien.RowCount - 4);
            }
            spread_GiaLoiLinhKien.CellValueChanged -= spread_GiaLoiLinhKien_CellValueChanged;
            DataLinhKien.ClearContents();
            spread_GiaLoiLinhKien.CellValueChanged += spread_GiaLoiLinhKien_CellValueChanged;

            int RowBegin = DataLinhKien.TopRowIndex;
            var grMay = InforMay.GroupBy(x => new { x.CodeLinhKien, x.TenMayCha });
            ws.Rows.Insert(RowBegin + 1, InforMay.Count+ grMay.Count(), RowFormatMode.FormatAsNext);
            foreach(var itemmay in grMay)
            {
                Row Crow = ws.Rows[RowBegin++];
                Crow["A"].CopyFrom(MauMay,PasteSpecial.All);
                Crow["A"].SetValueFromText(itemmay.Key.TenMayCha);
                Crow["E"].SetValueFromText(itemmay.Key.CodeLinhKien);
                int STTLinhKien = 1;
                foreach (var itemLK in itemmay)
                {
                    Crow = ws.Rows[RowBegin++];
                    Crow["A"].SetValue(STTLinhKien++);
                    Crow["B"].SetValueFromText(itemLK.MaHieu);
                    Crow["C"].SetValueFromText(itemLK.Ten);
                    Crow["D"].SetValue(itemLK.Gia);
                    Crow["E"].SetValueFromText(itemLK.Code);

                }
            }
            ws.Rows.Remove(ws.Range["DataLinhKien"].BottomRowIndex - 4, 4);
            wb.EndUpdate();
        }   
        private void Fcn_LoadDataGiaLoiLinhKien()
        {
            string dbString = $"SELECT GM.*,MAY.Ten as TenMayCha " +
        $" FROM {MyConstant.Tbl_MTC_ChiTietGiaMay} GM" +
        $" LEFT JOIN {MyConstant.Tbl_MTC_TenMay} MAY ON MAY.Code=GM.CodeMay ";
            List<ModelGiaLoiLinhKien> InforMay = DataProvider.InstanceTHDA.ExecuteQueryModel<ModelGiaLoiLinhKien>(dbString);
            IWorkbook wb = spread_GiaLoiLinhKien.Document;
            Worksheet ws = wb.Worksheets["GIÁ LÕI"];
            Worksheet SheetMau = wb.Worksheets["SheetMau"];
            CellRange MauMay = SheetMau.Range["DataMauGiaLoi"];
            CellRange DataGiaLoi = ws.Range["DataGiaLoi"];
            wb.BeginUpdate();
            if (DataGiaLoi.RowCount > 5)
            {
                ws.Rows.Remove(DataGiaLoi.TopRowIndex + 1, DataGiaLoi.RowCount - 4);
            }
            spread_GiaLoiLinhKien.CellValueChanged -= spread_GiaLoiLinhKien_CellValueChanged;
            DataGiaLoi.ClearContents();
            spread_GiaLoiLinhKien.CellValueChanged += spread_GiaLoiLinhKien_CellValueChanged;
            int RowBegin = DataGiaLoi.TopRowIndex;
            var grMay = InforMay.GroupBy(x => new { x.CodeMay, x.TenMayCha });
            ws.Rows.Insert(RowBegin + 1, InforMay.Count+ grMay.Count(), RowFormatMode.FormatAsNext);
            foreach(var itemmay in grMay)
            {
                Row Crow = ws.Rows[RowBegin++];
                Crow["A"].CopyFrom(MauMay,PasteSpecial.All);
                Crow["A"].SetValueFromText(itemmay.Key.TenMayCha);
                Crow["E"].SetValueFromText(itemmay.Key.CodeMay);
                int STTLinhKien = 1;
                foreach (var itemLK in itemmay)
                {
                    Crow = ws.Rows[RowBegin++];
                    Crow["A"].SetValue(STTLinhKien++);
                    Crow["B"].SetValueFromText(itemLK.MaHieu);
                    Crow["C"].SetValueFromText(itemLK.Ten);
                    Crow["D"].SetValue(itemLK.Gia);
                    Crow["E"].SetValueFromText(itemLK.Code);

                }
            }
            ws.Rows.Remove(ws.Range["DataGiaLoi"].BottomRowIndex - 4, 3);
            wb.EndUpdate();
        }
        private void slke_ThongTinDuAn_EditValueChanged(object sender, EventArgs e)
        {
            xTabMain.SelectedTabPage = null;
        }

        private void spread_GiaLoiLinhKien_CellValueChanged(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellEventArgs e)
        {
            IWorkbook wb = spread_GiaLoiLinhKien.Document;
            Worksheet ws = e.Worksheet;
            var dic = MyFunction.fcn_getDicOfColumn(ws.GetUsedRange());
            string Code = ws.Rows[e.RowIndex][dic["Code"]].Value.ToString();
            string CodeMay = ws.Rows[e.RowIndex][dic["CodeMay"]].Value.ToString();
            string dbString = string.Empty;
            string col = ws.Rows[0][e.ColumnIndex].Value.TextValue;
            if (!string.IsNullOrEmpty(Code))
            {
                dbString = $"UPDATE {MyConstant.Tbl_MTC_ChiTietGiaMay} SET {col} = '{e.Value}' WHERE \"Code\" = '{Code}'";
                DataProvider.InstanceTHDA.ExecuteNonQuery(dbString);
            }
            else if (!string.IsNullOrEmpty(CodeMay))
            {
                dbString = $"UPDATE {MyConstant.Tbl_MTC_TenMay} SET \"Ten\" = '{e.Value}' WHERE \"Code\" = '{CodeMay}'";
                DataProvider.InstanceTHDA.ExecuteNonQuery(dbString);
            }
        }

        private void spread_GiaLoiLinhKien_CellBeginEdit(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellCancelEventArgs e)
        {
            IWorkbook wb = spread_GiaLoiLinhKien.Document;
            Worksheet ws = e.Worksheet;
            if(ws.Name == "GIÁ LÕI")
            {
                CellRange DataGiaLoi = ws.Range["DataGiaLoi"];
                if (!DataGiaLoi.Contains(e.Cell))
                    e.Cancel = true;
                var dic = MyFunction.fcn_getDicOfColumn(ws.GetUsedRange());
                string Code = ws.Rows[e.RowIndex][dic["Code"]].Value.ToString();
                string CodeMay = ws.Rows[e.RowIndex][dic["CodeMay"]].Value.ToString();
                if (string.IsNullOrEmpty(Code) && string.IsNullOrEmpty(CodeMay))
                    e.Cancel = true;
            }
        }

        private void spread_GiaLoiLinhKien_ActiveSheetChanged(object sender, ActiveSheetChangedEventArgs e)
        {
            if (e.NewActiveSheetName == "GIÁ LÕI")
                Fcn_LoadDataGiaLoiLinhKien();
            else  if (e.NewActiveSheetName == "LINH KIỆN")
                Fcn_LoadDataLinhKien();


        }
        private int Fcn_CreateEmty(DataTable dt,int STT,DateTime Date)
        {
            DataRow Row = dt.NewRow();
            Row["STT"] = STT++;
            Row["TypeRow"] = MyConstant.TYPEROW_CVCha;
            Row["Ngay"] =Row["Date"] = Date.ToString(MyConstant.CONST_DATE_FORMAT_SPSHEET);
            dt.Rows.Add(Row);
            return STT;
        }
        private void Fcn_LoadDataDichVuTheoThang()
        {
            FileHelper.fcn_spSheetStreamDocument(spred_DichVuTheoThang, $@"{BaseFrom.m_templatePath}\FileExcel\Template_DichVuThang.xlsx"); // Tên lưới excel trong bảng (tự đặt ví dụ spsheet_Thongtinchinh) và tên file excel đọc vào (ví dụ:1.aThongTinCoBan.xlsx). 
            DateTime Month = de_ThangTinhLuong.DateTime;
            var firstDayOfMonth = new DateTime(Month.Year, Month.Month, 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            
            string dbString = $"SELECT DVT.*,COALESCE(MAY.Ten, LK.Ten) AS Model,COALESCE(MAY.Code, LK.Code) AS CodeModel," +
                $"COALESCE(GIAMAY.MaHieu, GIALK.MaHieu) AS MaLinhKien,COALESCE(GIAMAY.Ten, GIALK.Ten) AS TenLinhKien,LKTHEOTHANG.CodeGiaMay,LKTHEOTHANG.CodeGiaLinhKien " +
       $" FROM {MyConstant.Tbl_DichVuThang} DVT" +
       $" LEFT JOIN {MyConstant.Tbl_MTC_TenMay} MAY ON MAY.Code=DVT.CodeLoiMay " +
       $" LEFT JOIN {MyConstant.Tbl_MTC_TenLinhKien} LK ON LK.Code=DVT.CodeLinhkien " +
       $" LEFT JOIN {MyConstant.Tbl_LinhKienDiKemTheoThang} LKTHEOTHANG ON LKTHEOTHANG.CodeDichVu=DVT.Code " +
       $" LEFT JOIN {MyConstant.Tbl_MTC_ChiTietGiaMay} GIAMAY ON LKTHEOTHANG.CodeGiaMay=GIAMAY.Code " +
       $" LEFT JOIN {MyConstant.Tbl_MTC_GiaLinhKien} GIALK ON LKTHEOTHANG.CodeGiaLinhKien=GIALK.Code " +
       $"WHERE DVT.CodeDuAn='{slke_ThongTinDuAn.EditValue}' " +
       $"AND DVT.Ngay>='{firstDayOfMonth.ToString(MyConstant.CONST_DATE_FORMAT_SQLITE)}' AND DVT.Ngay<='{lastDayOfMonth.ToString(MyConstant.CONST_DATE_FORMAT_SQLITE)}' ";
            List<DichVuThang> DVThang = DataProvider.InstanceTHDA.ExecuteQueryModel<DichVuThang>(dbString);
            List<DateTime> lstDate = DVThang.Select(x => x.Ngay).Distinct().ToList();
            for(DateTime i = firstDayOfMonth; i <= lastDayOfMonth; i=i.AddDays(1))
            {
                if (!lstDate.Contains(i))
                    DVThang.Add(new DichVuThang
                    {
                        Ngay=i
                    });
            }
            IWorkbook workbook = spred_DichVuTheoThang.Document;
            Worksheet ws = workbook.Worksheets[0];
            CellRange definedName = ws.Range["Data"];
            Dictionary<string, string> dic = MyFunction.fcn_getDicOfColumn(definedName);
            DataTable dtData = new DataTable();
            foreach (var item in dic)
            {
                dtData.Columns.Add(item.Key, typeof(object));
            }
            var crRowInd = definedName.TopRowIndex;
            int rowTongInd = crRowInd;
            int STT = 1;
            var grDate = DVThang.GroupBy(x => x.Ngay).OrderBy(x => x.Key);
            foreach(var Date in grDate)
            {
                var grDV = Date.Where(x=>x.Code!=null).GroupBy(x => x.Code);
                foreach(var DV in grDV)
                {
                    DataRow RowCha = dtData.NewRow();


                }
                STT=Fcn_CreateEmty(dtData, STT, Date.Key);
                STT=Fcn_CreateEmty(dtData, STT, Date.Key);
            }
            crRowInd = rowTongInd - 1 + dtData.Rows.Count;
            int numRow = dtData.Rows.Count;
            workbook.BeginUpdate();
            ws.Rows.Insert(rowTongInd + 1, numRow, RowFormatMode.FormatAsNext);
            ws.Import(dtData, false, rowTongInd, 0);
            definedName = ws.Range["Data"];
            SpreadsheetHelper.ReplaceAllFormulaAfterImport(definedName);
            ws.Rows.Remove(definedName.BottomRowIndex - 2, 3);
            SpreadsheetHelper.FormatRowsInRange(definedName, dic[MyConstant.COL_TypeRow], dic["Ngay"]);
            workbook.EndUpdate();



        }

        private void spred_DichVuTheoThang_CellValueChanged(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellEventArgs e)
        {
            Worksheet ws = e.Worksheet;
            string newVal = e.Value.ToString();
            string colHeading = ws.Columns[e.ColumnIndex].Heading;
            string colInDb = ws.Rows[0][e.ColumnIndex].Value.ToString();
            CellRange range = ws.Range["Data"];
            Dictionary<string, string> NAME_COL = MyFunction.fcn_getDicOfColumn(range);
            Row crRow = ws.Rows[e.RowIndex];
            string code = crRow[NAME_COL[MyConstant.COL_CodeCT]].Value.ToString();
            string dbString = string.Empty;
            if (string.IsNullOrEmpty(code))
            {
                string Code = Guid.NewGuid().ToString();
                string LoaiModel = crRow[NAME_COL["LoaiModel"]].Value.ToString();
                if (string.IsNullOrEmpty(LoaiModel))
                {
                    crRow[NAME_COL["LoaiModel"]].SetValueFromText("LÕI");
                    LoaiModel = "LÕI";
                }
                DateTime Date =DateTime.Parse(crRow[NAME_COL["Ngay"]].Value.ToString());
                dbString = $"INSERT INTO {MyConstant.Tbl_DichVuThang} (\"Code\", {colInDb}, \"Ngay\", \"LoaiModel\") VALUES " +
                      $"('{Code}', '{e.Value}','{Date.ToString(MyConstant.CONST_DATE_FORMAT_SQLITE)}','{LoaiModel}')";
                DataProvider.InstanceTHDA.ExecuteNonQuery(dbString);
                crRow[NAME_COL["Code"]].SetValue(Code);
            }
            else
            {
                dbString = $"UPDATE {MyConstant.Tbl_DichVuThang} SET {colInDb}='{e.Value}' WHERE \"Code\" = '{crRow[NAME_COL["Code"]].Value}'";
                DataProvider.InstanceTHDA.ExecuteNonQuery(dbString);
            }




        }

        private void spred_DichVuTheoThang_CellBeginEdit(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellCancelEventArgs e)
        {
            IWorkbook wb = e.Worksheet.Workbook;
            Worksheet ws = e.Worksheet;
            string colHeading = ws.Columns[e.Cell.ColumnIndex].Heading;
            if (e.Cell.ColumnIndex == 0|| e.Cell.ColumnIndex == 1|| !ws.Range["Data"].Contains(e.Cell)) //Ô STT và ô ngày
            {
                e.Cancel = true;
                return;
            }
        }
    }
}