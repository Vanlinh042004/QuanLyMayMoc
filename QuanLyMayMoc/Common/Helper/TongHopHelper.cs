using DevExpress.Compression;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using StackExchange.Profiling.Internal;
using MSETTING = QuanLyMayMoc.Properties.Settings;
using QuanLyMayMoc.Common.Constant;
using QuanLyMayMoc.Common.ViewModel;

namespace QuanLyMayMoc.Common.Helper
{
    public static class TongHopHelper
    {
        public static bool fcn_SaveTHDAToQLTC(string newPath)
        {
            try
            {
                string m_crFileDA = newPath;
                string filename = Path.GetFileNameWithoutExtension(m_crFileDA);
                string tempDirInTemp = $@"{BaseFrom.m_tempPath}\~${filename}";
                string tempFileBeforeSave = $@"{Path.GetDirectoryName(m_crFileDA)}/~${Path.GetFileName(m_crFileDA)}";

                if (Directory.Exists(tempDirInTemp))
                    MyFunction.DirectoryDelete(tempDirInTemp);

                Directory.CreateDirectory(tempDirInTemp);

                MyFunction.DirectoryCopy($@"{BaseFrom.m_tempPath}\{filename}", tempDirInTemp, true);

                if (File.Exists(tempFileBeforeSave))
                    File.Delete(tempFileBeforeSave);

                //Directory.CreateDirectory(tempFileBeforeSave);

                DevExpress.Compression.EncryptionType encryptionType = DevExpress.Compression.EncryptionType.PkZip;


                using (ZipArchive archive = new ZipArchive())
                {

                    //archive.Password = m_pw;
                    archive.EncryptionType = encryptionType;
                    archive.AddDirectory(tempDirInTemp, " /");
                    archive.Save(tempFileBeforeSave);
                }

                if (File.Exists(m_crFileDA))
                    File.Delete(m_crFileDA);

                File.Move(tempFileBeforeSave, m_crFileDA);
                MyFunction.DirectoryDelete(tempDirInTemp);
                File.Delete(tempFileBeforeSave);
            }
            catch (Exception ex)
            {
                MessageShower.ShowError($"Lỗi lưu dự án: {ex.Message}__{ex.InnerException?.Message}");
                return false;
            }
            return true;
        }
        public static bool fcn_saveDA(bool isSaveAs, string ThongBaoSaveAs = null, bool isShowErrorDialog = true)
        {
            try
            {
                string m_crFileDA = MSETTING.Default.PathHienTai;
                string filename = Path.GetFileNameWithoutExtension(m_crFileDA);

                //if (newPath != null)
                //{
                //    m_crFileDA = newPath;
                //    filename = Path.GetFileNameWithoutExtension(m_crFileDA);
                //    goto SavingBegin;
                //}

                if (!File.Exists(m_crFileDA) || isSaveAs)
                {
                    if (ThongBaoSaveAs.HasValue())
                        MessageShower.ShowInformation(ThongBaoSaveAs);

                    var saveFileDialog = SharedControls.saveFileDialog;

                    saveFileDialog.DefaultExt = "qlmm";
                    saveFileDialog.Filter = "Quan Ly May Moc (*.qlmm)|*.qlmm";
                    saveFileDialog.AddExtension = false;

                    if (saveFileDialog.ShowDialog() == DialogResult.Cancel)
                        return false;
                    m_crFileDA = saveFileDialog.FileName;
                    filename = Path.GetFileNameWithoutExtension(m_crFileDA);
                    Directory.Move($@"{BaseFrom.m_tempPath}\{BaseFrom.m_crTempDATH}", $@"{BaseFrom.m_tempPath}\{filename}");
                    BaseFrom.m_crTempDATH = filename;
                    DataProvider.InstanceTHDA.changePath($@"{BaseFrom.m_tempPath}\{filename}\{MyConstant.CONST_DbFromPathDA}");
                }

                if (fcn_SaveTHDAToQLTC(m_crFileDA))
                {
                    MSETTING.Default.PathHienTai = m_crFileDA;
                    MSETTING.Default.Save();
                    AlertShower.ShowInfo("Đã lưu dự án");
                    BaseFrom.THDAChanged = false;
                    return true;
                }
                else return false;

            }
            catch (Exception ex)
            {
                string err = $"Lỗi lưu dự án: {ex.Message}__{ex.InnerException?.Message}";
                if (isShowErrorDialog)
                    MessageShower.ShowError(err);

                AlertShower.ShowInfo(err);
                return false;
            }
            //return true;
        }
        public static bool CheckSaveDA(bool deleteFolder = true)
        {
            if (BaseFrom.m_crTempDATH.HasValue())
            {
                if (BaseFrom.THDAChanged)
                {
                    DialogResult rs = MessageShower.ShowYesNoCancelQuestion("Lưu dự án", "Dự án đang làm có thay đổi, bạn có muốn lưu không?");

                    if (rs == DialogResult.Cancel)
                        return false;

                    if (rs == DialogResult.Yes)
                    {
                        if (fcn_saveDA(false))
                        {
                            if (deleteFolder)
                            {
                                MyFunction.DirectoryDelete(BaseFrom.m_FullTempathDA);
                                DataProvider.InstanceTHDA.changePath("");
                            }
                            return true;
                        }
                        else return false;
                    }
                    else
                    {
                        if (deleteFolder)
                        {
                            MyFunction.DirectoryDelete(BaseFrom.m_FullTempathDA);
                            DataProvider.InstanceTHDA.changePath("");
                        }
                        return true;
                    }
                }
                else if (deleteFolder)
                    MyFunction.DirectoryDelete(BaseFrom.m_FullTempathDA);
            }


            return true;
        }
        public static void fcn_updateCbbThongTinDuAnCongTrinh()
        {
            var slke_ThongTinDuAn = SharedControls.slke_ThongTinDuAn;
            //slke_ThongTinDuAn.SelectedIndexChanged -= slke_ThongTinDuAn_SelectedIndexChanged;
            //slke_ThongTinDuAn.Properties.DataSource = null;

            //slke_ThongTinDuAn.EditValue = null;

            string db_string = $"SELECT * from {MyConstant.Tbl_ThongTinDuAn} ";
            var DAs = DataProvider.InstanceTHDA.ExecuteQueryModel<Tbl_ThongTinDuAnViewModel>(db_string);

            slke_ThongTinDuAn.Properties.DataSource = DAs;//new BindingSource(DicCbbDA, null);

            var crDA = DAs.FirstOrDefault(x => x.Code == MSETTING.Default.DuAnHienTai);
            //string crCodeInSlke = slke_ThongTinDuAn.EditValue as string;
            slke_ThongTinDuAn.EditValue = crDA?.Code;
            //if ()
            return;
        }

    }
}
