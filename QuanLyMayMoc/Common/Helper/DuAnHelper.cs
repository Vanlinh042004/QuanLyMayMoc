using QuanLyMayMoc.Common.Constant;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMayMoc.Common.Helper
{
    public static class DuAnHelper
    {
        public static bool DeleteDataRows(string fstTbl, IEnumerable<string> ids = null, bool AddToDeletedTable = true)
        {
            try
            {
                var allTbls = MyConstant.LS_CONST_TYPE_SERVER_TBL;
                int indTbl = Array.IndexOf(allTbls, fstTbl);
                if (indTbl < 0)
                {
                    MessageShower.ShowError("Bảng dữ liệu không tồn tại");
                    return false;
                }

                Dictionary<string, List<string>> dicTblCodes = new Dictionary<string, List<string>>();//{Table, Code}
                                                                                                      //SQLiteAllDataViewModel AllData = new SQLiteAllDataViewModel();
                Dictionary<string, DataTable> dicTblFile_dtDeleted = new Dictionary<string, DataTable>();
                string dbString = "";
                for (int i = indTbl; i < allTbls.Length; i++)
                {
                    string tbl = allTbls[i];
                    string colCode = "Code";



                    dbString = $"SELECT * FROM {tbl}";
                    List<string> whereCondition = new List<string>();


                    if (tbl != fstTbl)
                    {
                        Dictionary<string, string> lsFk = MyConstant.dicFks[tbl];


                        string[] fks = lsFk.Where(x => dicTblCodes.ContainsKey(x.Value) && dicTblCodes[x.Value].Any())
                            .Select(x => $"{x.Key} IN ({MyFunction.fcn_Array2listQueryCondition(dicTblCodes[x.Value])})")
                            .ToArray();

                        if (fks.Any())
                            whereCondition.Add($"({string.Join(" OR ", fks)})");
                        else if (!lsFk.Any(x => dicTblCodes.ContainsKey(x.Value)) || lsFk.Any(x => dicTblCodes.ContainsKey(x.Value) && !dicTblCodes[x.Value].Any()))
                            continue;

                    }
                    else
                    {
                        if (ids != null)
                            whereCondition.Add($"{colCode} IN ({MyFunction.fcn_Array2listQueryCondition(ids)})");//IN ({SQLRawHelper.fcn_Array2listQueryCondition(allPermission.AllProject.ToArray())})";
                    }



                    if (whereCondition.Any())
                    {
                        dbString += $" WHERE {string.Join(" AND ", whereCondition)}";
                    }
                    else if (tbl != fstTbl)
                        continue;


                    DataTable dt = DataProvider.InstanceTHDA.ExecuteQuery(dbString);

                    if (MyConstant.tblsFileDinhKem.ContainsKey(tbl))
                    {
                        dicTblFile_dtDeleted.Add(tbl, dt);
                    }

                    if (!dicTblCodes.ContainsKey(tbl))
                        dicTblCodes[tbl] = new List<string>() { };

                    var codes = dt.AsEnumerable().Select(x => x[colCode].ToString());
                    dicTblCodes[tbl].AddRange(codes);

                    //Setnull các code nhóm;
                }



                for (int i = allTbls.Length - 1; i >= indTbl; i--)
                {
                    string tbl = allTbls[i];
                    Type ModelSQL = Type.GetType($"VChatCore.Model.SyncSqlite.{tbl}");

                    string nameDBSet = $"{tbl}s";
                    string colCode = "Code";
                    if (!dicTblCodes.ContainsKey(tbl) || !dicTblCodes[tbl].Any())
                        continue;

                    string con = MyFunction.fcn_Array2listQueryCondition(dicTblCodes[tbl].ToArray());
                    dbString = $"DELETE FROM {tbl} WHERE {colCode} IN ({con})";

                    int num = DataProvider.InstanceTHDA.ExecuteNonQuery(dbString, AddToDeletedDataTable: AddToDeletedTable);
                }
                foreach (var item in MyConstant.tblsFileDinhKem)
                {
                    string tblFile = item.Key;
                    string tblParentFile = item.Value;
                    if (!dicTblCodes.ContainsKey(tblParentFile))
                        continue;

                    foreach (string code in dicTblCodes[tblParentFile])
                    {
                        if (string.IsNullOrEmpty(code))
                            continue;
                        string dir = Path.Combine(BaseFrom.m_FullTempathDA,
                            string.Format(CusFilePath.SQLiteFile, tblFile, code));

                        if (Directory.Exists(dir))
                            Directory.Delete(dir, true);
                    }

                    if (!dicTblCodes.ContainsKey(tblFile))
                        continue;

                    foreach (string code in dicTblCodes[tblFile])
                    {
                        if (!dicTblFile_dtDeleted.ContainsKey(tblFile) || string.IsNullOrEmpty(code))
                            continue;

                        DataTable dtFile = dicTblFile_dtDeleted[tblFile];
                        DataRow dr = dtFile.AsEnumerable().Single(x => x["Code"].ToString() == code);
                        string dir = Path.Combine(BaseFrom.m_FullTempathDA, string.Format(CusFilePath.SQLiteFile, tblFile, dr["CodeParent"].ToString()));

                        if (!Directory.Exists(dir))
                            continue;

                        string fullFilePath = Path.Combine(dir, code);

                        if (File.Exists(fullFilePath))
                        {
                            File.Delete(fullFilePath);
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                MessageShower.ShowError($"Lỗi xóa data: {ex.Message}__{ex.InnerException?.Message}");
                return false;
            }
            //MessageShower.ShowInformation("Xóa dữ liệu thành công");
            return true;

        }
    }
}
