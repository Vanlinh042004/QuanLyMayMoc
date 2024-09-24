using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyMayMoc.Common.Constant
{
    public class MyConstant
    {
        public const string Tbl_ThongTinDuAn = "Tbl_ThongTinDuAn";
        public const string Tbl_DichVuThang = "Tbl_DichVuThang";
        public const string Tbl_LinhKienDiKemTheoThang = "Tbl_LinhKienDiKemTheoThang";
        public const string Tbl_ChamCong_BangNhanVien = "Tbl_ChamCong_BangNhanVien";
        public const string Tbl_ChamCong_BangPhongBan = "Tbl_ChamCong_BangPhongBan";
        public const string Tbl_ChamCong_ViTriNhanVien = "Tbl_ChamCong_ViTriNhanVien";
        public const string Tbl_BangNhanVien_FileDinhKem = "Tbl_BangNhanVien_FileDinhKem";
        public const string Tbl_MTC_TenMay = "Tbl_MTC_TenMay";
        public const string Tbl_MTC_TenLinhKien = "Tbl_MTC_TenLinhKien";
        public const string Tbl_MTC_ChiTietGiaMay = "Tbl_MTC_ChiTietGiaMay";
        public const string Tbl_MTC_GiaLinhKien = "Tbl_MTC_GiaLinhKien";
        public const string CONST_DATE_FORMAT_SQLITE = "yyyy-MM-dd";
        public const string CONST_DATE_FORMAT_SPSHEET = "dd/MM/yyyy";
        public const string CONST_DbFromPathDA = "db_QuanLyMayMoc.sqlite3";
        public const string PrefixFormula = "TBTFormula";
        public const string PrefixDate = "TBTDate";
        public const string PrefixMerge = "TBTMerge";
        public const string DefaultProvince = "HoChiMinh";
        public const string COL_TypeRow = "TypeRow";
        public const string COL_CodeCT = "Code";
        public static string[] LS_CONST_TYPE_SERVER_TBL =
            {Tbl_ThongTinDuAn,
            Tbl_ChamCong_BangNhanVien,
            Tbl_ChamCong_ViTriNhanVien,
            Tbl_ChamCong_BangPhongBan,
            Tbl_BangNhanVien_FileDinhKem

        };
        public const int MaxFileSize = 157286400; //150MB
        public const int MaxRequestSize = 209715200; //200MB
        public const int MaxSoNgayThucHien = 3650;
        public static Dictionary<string, string> tblsFileDinhKem = new Dictionary<string, string>()
        {
            { Tbl_BangNhanVien_FileDinhKem, Tbl_ChamCong_BangNhanVien}
        };
        public static Dictionary<string, Dictionary<string, string>> dicFks = new Dictionary<string, Dictionary<string, string>>()
        {
//AllDicForeign
#region SQLITE
			{Tbl_BangNhanVien_FileDinhKem, new Dictionary<string, string>()
            {
                {"CodeParent","Tbl_ChamCong_BangNhanVien"}
            } },

            {Tbl_ChamCong_BangNhanVien, new Dictionary<string, string>()
            {
                {"PhongBan","Tbl_ChamCong_BangPhongBan"},
                {"ChucVu","Tbl_ChamCong_ViTriNhanVien"}
            } },

            {Tbl_ChamCong_BangPhongBan, new Dictionary<string, string>()
            {

            } },
                   {Tbl_ChamCong_ViTriNhanVien, new Dictionary<string, string>()
            {

            } },

#endregion


        };
        public static CultureInfo culture = (CultureInfo)CultureInfo.GetCultureInfo("vi-VN").Clone();

        public const string TYPEROW_NoType = ""; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_CongTrinh = "CTR"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_ChiPhi = "CP"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_LamTron = "TongLamTron"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_HangMuc = "HM"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới

        public const string TYPEROW_MuiThiCong = "MuiThiCong"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_PhanTuyen = "PhanTuyen"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_HTPhanTuyen = "HTPhanTuyen"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_Nhom = "Nhom"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_NhomDienGiai = "NhomCon"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_CVCha = "CVCha"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_CVCON = "CVCon"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_CVCHIA = "CVCHIA"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        //public const string TYPEROW_THEMMOI = "Add"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới
        public const string TYPEROW_CVTong = "Tong"; //Hàng tổng chứa con là kế hoạch và phát sinh
        public const string TYPEROW_SUMTT = "SUMTT"; //Hàng tổng BÁO CÁO TRƯỜNG SƠN
        public const string TYPEROW_GOP = "GOP"; //Hàng chưa CV Con, CV Cha hay hàng nhập mới

    }
}
