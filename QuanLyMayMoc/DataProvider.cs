using System.Data.SQLite;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Dapper;

namespace QuanLyMayMoc
{
    public class DataProvider
    {
        private static DataProvider instanceTHDA;
        public string m_conString = string.Empty;
        public string m_NameDb = string.Empty;
        public bool isEncrypted = true;
        const string TBTkey = "TBT@@1!360";
        //string TBTlic = "13453-563-7092387-29456";

        static string sqlKEY = $"PRAGMA key='{TBTkey}';";

        string sqlLIC = "PRAGMA lic='13453-563-7092387-29456';";
        string patternParam = @"(@\w+)(,|\)|\s|;)";

        public static DataProvider InstanceTHDA
        {
            get { if (instanceTHDA == null) instanceTHDA = new DataProvider(); return DataProvider.instanceTHDA; }
            private set { DataProvider.instanceTHDA = value; }
        }
        public DataProvider()
        {
            m_conString = $"Data Source={m_NameDb};Version=3;foreign keys=true";
            //m_con.ConnectionString = m_conString;

            if (!File.Exists(m_NameDb))
            {
                Debug.WriteLine("Database not exist!");
            }
            else
            {
                Debug.WriteLine("Database existed!");
            }
        }
        public int ExecuteNonQuery(string query, bool withEncrypt = true, object[] parameter = null, bool AddToDeletedDataTable = true)
        {
            query = query.Trim() + " ";
            string pattern = $@"('-*\d+)(,+)(\d+')";
            //Match m = Regex.Match(query, pattern, RegexOptions.IgnoreCase); 

            query = Regex.Replace(query, pattern, "$1.$3");

            int data = 0;
            Debug.WriteLine("DB Non query: " + query);


            DataTable dtDeleted = null;
            if (query.StartsWith("DELETE FROM") && AddToDeletedDataTable && m_conString == instanceTHDA.m_conString)
            {
                string patternTbl = @"(\s+)(Tbl_([a-z]|[A-Z]|_)+)(\s+)";
                Match mTbl = Regex.Match(query, patternTbl, RegexOptions.IgnoreCase);


                if (mTbl != Match.Empty)
                {
                    string tbl = mTbl.Groups[2].Value;

                    string colCode = "Code";


                    //DataTable dt = ExecuteQuery(query.Replace("DELETE", $"SELECT {colCode}"));
                    //var codesDeleted = dt.AsEnumerable().Select(x => (string)x[0]);

                    //if (codesDeleted.Any())
                    //{
                    //    string dbString = $"SELECT * FROM {MyConstant.TBL_DeletedRecored} WHERE Code IN ({MyFunction.fcn_Array2listQueryCondition(codesDeleted)}) ";
                    //    dtDeleted = DataProvider.instanceTHDA.ExecuteQuery(dbString);

                    //    foreach (string codeDeleted in codesDeleted)
                    //    {
                    //        DataRow dr = dtDeleted.AsEnumerable().SingleOrDefault(x => (string)x["Code"] == codeDeleted);
                    //        if (dr is null)
                    //        {
                    //            dr = dtDeleted.NewRow();
                    //            dtDeleted.Rows.Add(dr);
                    //            dr["Code"] = codeDeleted;
                    //            dr["TableName"] = tbl;
                    //            dr["CodeDuAn"] = SharedControls.slke_ThongTinDuAn.EditValue as string;
                    //        }

                    //    }
                    //    UpdateDataTableFromSqliteSource(dtDeleted, Server.Tbl_DeletedRecored);
                    //}

                }
            }

            using (SQLiteConnection connection = new SQLiteConnection(m_conString))
            {
                connection.Open();

                if (isEncrypted && withEncrypt)
                {
                    SQLiteCommand command = new SQLiteCommand(sqlKEY, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();

                    command = new SQLiteCommand(sqlLIC, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();
                }

                using (SQLiteTransaction transaction = connection.BeginTransaction())
                {

                    using (SQLiteCommand command = new SQLiteCommand(query, connection))
                    {
                        if (parameter != null)
                        {
                            Match result = Regex.Match(query, patternParam, RegexOptions.IgnoreCase);
                            List<string> listPara = new List<string>();

                            if (result != Match.Empty)
                            {
                                do
                                {
                                    //Console.WriteLine(result.ToString());
                                    listPara.Add(result.Groups[1].Value);
                                    result = result.NextMatch(); // Chuyển qua kết quả trùng khớp kế tiếp
                                }
                                while (result != Match.Empty); // Kiểm tra xem đã hết kết quả trùng khớp chưa
                            }
                            int i = 0;
                            foreach (string item in listPara)
                            {
                                if (item.Contains("@"))
                                {
                                    command.Parameters.AddWithValue(item, parameter[i]);
                                    i++;
                                }
                            }
                        }

                        data = command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
                connection.Close();

            }
            //if (dtDeleted != null)
            //{
            //    UpdateDataTableFromSqliteSource(dtDeleted, MyConstant.TBL_DeletedRecored);
            //}
            return data;
        }
        public DataTable ExecuteQuery(string query, bool withEncrypt = true, object[] parameter = null)
        {
            Debug.WriteLine("DB query: " + query);
            DataTable data = new DataTable();
            query = query.Trim() + " ";
            using (SQLiteConnection connection = new SQLiteConnection(m_conString))
            {
                connection.Open();

                if (isEncrypted && withEncrypt)
                {
                    SQLiteCommand command = new SQLiteCommand(sqlKEY, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();

                    command = new SQLiteCommand(sqlLIC, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();

                }

                using (SQLiteCommand command = new SQLiteCommand(query, connection))
                {
                    if (parameter != null)
                    {


                        Match result = Regex.Match(query, patternParam, RegexOptions.IgnoreCase);
                        List<string> listPara = new List<string>();

                        if (result != Match.Empty)
                        {
                            do
                            {
                                //Console.WriteLine(result.ToString());
                                listPara.Add(result.Groups[1].Value);
                                result = result.NextMatch(); // Chuyển qua kết quả trùng khớp kế tiếp
                            }
                            while (result != Match.Empty); // Kiểm tra xem đã hết kết quả trùng khớp chưa
                        }
                        int i = 0;
                        foreach (string item in listPara)
                        {
                            if (item.Contains("@"))
                            {
                                //var type = parameter[i].GetType();
                                ////Type t = (type.GenericTypeArguments.Any()) ? type.GenericTypeArguments[0] : type;

                                //if (type != null && (type == typeof(DateTime) || type == typeof(DateTime?)))
                                //{
                                //    parameter[i] = ((DateTime)parameter[i]).ToString(MyConstant.CONST_DATE_FORMAT_SQLITE);
                                //}
                                command.Parameters.AddWithValue(item, parameter[i]);
                                i++;
                            }
                        }
                    }


                    SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);

                    adapter.Fill(data);
                    //command.Dispose();
                    //connection.Dispose();
                }
                connection.Close();
            }

            return data;
        }

        public List<T> ExecuteQueryModel<T>(string query, bool withEncrypt = true, object[] parameter = null)
        {
            Debug.WriteLine("DB query: " + query);
            List<T> data = new List<T>();

            using (SQLiteConnection connection = new SQLiteConnection(m_conString))
            {

                connection.Open();

                if (isEncrypted && withEncrypt)
                {
                    SQLiteCommand command = new SQLiteCommand(sqlKEY, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();

                    command = new SQLiteCommand(sqlLIC, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();

                }

                var result = connection.Query<T>(query, null);
                if (result.Any())
                    data = result.ToList();
                else
                    data = new List<T>();

                connection.Close();


            }

            return data;
            //return dt.fcn_DataTable2List<T>();
        }
        public int UpdateDataTableFromSqliteSource(DataRow[] drs, string tbl, bool isReplace = false, bool withEncrypt = true, string queryStringCondition = "")
        {

            if (isReplace)
                ExecuteNonQuery($"DELETE FROM \"{tbl}\" {queryStringCondition}");
            string queryString = $"SELECT * FROM \"{tbl}\" {queryStringCondition}";

            int numOfRow = 1;



            using (SQLiteConnection connection = new SQLiteConnection(m_conString))
            {
                connection.Open();

                if (isEncrypted && withEncrypt)
                {
                    SQLiteCommand command = new SQLiteCommand(sqlKEY, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();

                    command = new SQLiteCommand(sqlLIC, connection);
                    command.ExecuteNonQuery();
                    command.Dispose();

                }
                var eachQr = 10000;
                var num = drs.Length / eachQr;
                for (int i = 0; i <= num; i++)
                {
                    var drss = drs.Skip(i * eachQr).Take(eachQr).ToArray();

                    using (var transaction = connection.BeginTransaction())
                    {
                        using (var sqliteAdapter = new SQLiteDataAdapter(queryString, connection))
                        {
                            var cmdBuilder = new SQLiteCommandBuilder(sqliteAdapter);
                            cmdBuilder.ConflictOption = ConflictOption.OverwriteChanges;
                            numOfRow = sqliteAdapter.Update(drss);
                            cmdBuilder.Dispose();
                        }

                        transaction.Commit();
                    }
                }

                connection.Close();
            }

            return numOfRow;
        }
        public void changePath(string path, bool Encrypt = true)
        {
            m_NameDb = path;
            m_conString = $"Data Source={path};Version=3;foreign keys=true";
            //m_con.ConnectionString = m_conString;

            isEncrypted = Encrypt;
            //if (Encrypt)
            //{
            //    SQLiteConnection db = new SQLiteConnection(m_conString);
            //    //SQLiteConnection db = new SQLiteConnection("Data Source=DatabaseTBT.sqlite;Version=3;");

            //    db.Open();

            //    SQLiteCommand command = new SQLiteCommand(sqlLIC, db);
            //    command.ExecuteNonQuery();


            //    command = new SQLiteCommand(sqlREKEY, db);
            //    command.ExecuteNonQuery();
            //    db.Close();
            //    db.Dispose();
            //}    

            if (!File.Exists(m_NameDb))
            {
                Debug.WriteLine("Database not exist!");
            }
            else
            {
                Debug.WriteLine("Database existed!");
            }
        }
        public int UpdateDataTableFromSqliteSource(DataTable dt, string tbl, bool isReplace = false, bool withEncrypt = true, string queryStringCondition = "")
        {

            return UpdateDataTableFromSqliteSource(dt.AsEnumerable().ToArray(), tbl, isReplace, withEncrypt, queryStringCondition);

            return 0;
        }
    }
}
