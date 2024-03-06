using EDR_Report.Interfaces;
using EDR_Report.Models.CPM;
using EDR_Report.Models.ERP;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace EDR_Report.Controllers
{
    public partial class ReportController : BaseController
    {
        private readonly ILogger<ReportController> _logger;
        /// <summary>
        /// 套表存放位置
        /// </summary>
        public string BasePath { get; set; } = string.Empty;
        /// <summary>
        /// 暫存檔存放位置
        /// </summary>
        public string TempFolder { get; set; } = string.Empty;
        private readonly IRptXlsHelper _rptHelper;
        private readonly I1CB004DbContext _dbcontext;
        public ReportController(IRptXlsHelper rptHelper, I1CB004DbContext dbcontext, ILogger<ReportController> logger)
        {
            BasePath = configuration.GetValue<string>("ReportTemplatePath");
            TempFolder = configuration.GetValue<string>("TempFolder");
            if (!Directory.Exists(TempFolder))
            {
                Directory.CreateDirectory(TempFolder);
            }
            _rptHelper = rptHelper;
            _dbcontext = dbcontext;
            _logger = logger;
        }
        [HttpGet, AllowAnonymous, Route("edrrpt/{act}/{state}/{cd?}"), EnableCors("allow_link")]
        public IActionResult EDRRpt(string? vEmpno, string? VSSessID, int? proID, string? act, string? state, string? cd)
        {
            if (string.IsNullOrEmpty(vEmpno) || string.IsNullOrEmpty(VSSessID)) return NotFound("SessionID遺失");
            if (string.IsNullOrEmpty(cd) || cd.Length != 8) return NotFound("參數錯誤");
            if (string.IsNullOrEmpty(state)) state = "2";
            var db = new DBFunc();
            var cli = db.query<CPM_SESSIONSModel>("cpm", "SELECT * FROM CPM_SESSIONS WHERE SESSION_ID = :VSSessID AND EMPNO = :vEmpno", new
            {
                VSSessID,
                vEmpno
            });
            if (cli.Count() == 0) return NotFound("尚未登入");
            var uli = db.query<VS_USERSModel>("erp", "SELECT * FROM VS_USERS WHERE USER_NAME = :vEmpno", new
            {
                vEmpno
            });
            if (uli.Count() == 0) return NotFound("工號錯誤");
            var u = uli.First();
            UserInfo.EMPNO = u.USER_NAME;
            UserInfo.UserID = u.USER_ID;
            UserInfo.SessionID = VSSessID;
            UserInfo.ProjectID = proID;
            UserInfo.UserData = u;

            return (act ?? "").ToLower() switch
            {
                // 如果有加入新的報表，記得要在這邊加上要呼叫的 Method
                "1cb001" => RPT1CB001(state, cd),
                "1cb004_1" => RPT1CB004_1(cd),
                "1cb004_2" => RPT1CB004_2(cd),
                ////////////
                // 編輯區4//
                ////////////
                "1ca803" => RPTLAYOUT(state, cd), //4780(1CA803)鳥嘴潭工務所 烏溪鳥嘴潭人工湖工程...
                "1cb102" => RPTLAYOUT(state, cd), //5541(1CB102)大埔工務所 曾文水庫放水渠道及擴...
                "1bb101" => RPTLAYOUT(state, cd), //5520(1CB102)大林蒲工務所 森霸二期
                "1cb109" => RPTLAYOUT(state, cd), //5780(1CB109)華桂聯合工務所 鳥嘴潭淨水場新建統包工程(續) 
                _ => NotFound("錯誤的報表")
            };
        }
        /// <summary>
        /// 插入一個Row，後面的資料往下移
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="count">插入幾個Row</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static IRow InsertRow(ISheet sheet, int rowIndex, int count = 1)
        {
            if (rowIndex < 0) throw new Exception("rowIndex 錯誤！");
            if (count < 1) throw new Exception("count 錯誤！");
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, count);
            return sheet.CreateRow(rowIndex);
        }
        /// <summary>
        /// 寫入欄位（由系統判斷並建立）
        /// <para>可合併欄位</para>
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        /// <param name="rowSpan">合併幾個 Row</param>
        /// <param name="colSpan">合併幾個 Cell</param>
        /// <returns></returns>
        public static ICell SetCell(ISheet sheet, int rowIndex, int colIndex, object? data, ICellStyle? style = null, int rowSpan = 0, int colSpan = 0)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = SetCell(row, colIndex, data, style);
            if (rowSpan > 0 && colSpan > 0)
            {
                var rm = rowIndex + rowSpan - 1;
                var cm = colIndex + colSpan - 1;
                for (var r = rowIndex; r <= rm; r++)
                {
                    for (var c = colIndex; c <= cm; c++)
                    {
                        if (r == rowIndex && c == colIndex) continue;
                        SetCell(sheet, r, c, "", style);
                    }
                }
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rm, colIndex, cm));
            }
            return cell;
        }
        /// <summary>
        /// 寫入欄位（由系統判斷並建立）
        /// </summary>
        /// <param name="row"></param>
        /// <param name="colIndex"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        /// <returns></returns>
        public static ICell SetCell(IRow row, int colIndex, object? data, ICellStyle? style) =>
            SetCell(row.GetCell(colIndex) ?? row.CreateCell(colIndex), data, style);

        /// <summary>
        /// NPOI共用Method: 用來建立空的Excel欄位
        /// </summary>
        /// <param name="c"></param>
        /// <param name="style"></param>
        public static ICell SetEmpty(ICell c, ICellStyle? style = null) => SetCell(c, string.Empty, style);

        /// <summary>
        /// NPOI共用Method: 自動判斷要寫入的內容
        /// </summary>
        /// <param name="c"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        public static ICell SetCell(ICell c, object? data, ICellStyle? style = null)
        {
            if (data is null) return SetEmpty(c, style);
            else if (data is decimal m) return SetCell(c, m, style);
            else if (data is string s) return SetCell(c, s, style);
            else if (data is int i) return SetCell(c, i, style);
            else if (data is double D) return SetCell(c, D, style);
            else return SetCell(c, data.ToString(), style);
        }

        /// <summary>
        /// NPOI共用Method: 用來在Excel欄位寫入字串
        /// </summary>
        /// <param name="c"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        public static ICell SetCell(ICell c, string? data, ICellStyle? style = null)
        {
            c.SetCellValue(data ?? "");
            if (style != null) c.CellStyle = style;
            return c;
        }
        /// <summary>
        /// NPOI共用Method: 用來在Excel欄位寫入浮點數
        /// </summary>
        /// <param name="c"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        public static ICell SetCell(ICell c, decimal? data, ICellStyle? style = null)
        {
            c.SetCellValue(data.HasValue ? decimal.ToDouble(data.Value) : 0);
            if (style != null) c.CellStyle = style;
            return c;
        }
        /// <summary>
        /// NPOI共用Method: 用來在Excel欄位寫入浮點數
        /// </summary>
        /// <param name="c"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        public static ICell SetCell(ICell c, double? data, ICellStyle? style = null)
        {
            c.SetCellValue(data ?? 0);
            if (style != null) c.CellStyle = style;
            return c;
        }
        /// <summary>
        /// NPOI共用Method: 用來在Excel欄位寫入數值
        /// </summary>
        /// <param name="c"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        public static ICell SetCell(ICell c, int? data, ICellStyle? style = null)
        {
            c.SetCellValue(data.HasValue ? data.Value : 0);
            if (style != null) c.CellStyle = style;
            return c;
        }
        /// <summary>
        /// NPOI共用Method: 用來在Excel欄位寫入日期時間(自訂格式)
        /// </summary>
        /// <param name="c"></param>
        /// <param name="data"></param>
        /// <param name="format"></param>
        /// <param name="style"></param>
        public static ICell SetCell(ICell c, DateTime? data, string format, ICellStyle? style = null)
        {
            c.SetCellValue(data.HasValue ? data.Value.ToString(format) : "");
            if (style != null) c.CellStyle = style;
            return c;
        }
        /// <summary>
        /// 取得指定欄位
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        public static ICell GetCell(ISheet sheet, int rowIndex, int colIndex)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            return row.GetCell(colIndex) ?? row.CreateCell(colIndex);
        }
        /// <summary>
        /// 讀取欄位字串
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static string? GetCellString(ISheet sheet, int rowIndex, int colIndex, string? defVal = "") => ReadCellString(GetCell(sheet, rowIndex, colIndex), defVal);
        /// <summary>
        /// 讀取欄位數值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static decimal? GetCellDecimal(ISheet sheet, int rowIndex, int colIndex, decimal? defVal = null) => ReadCellDecimal(GetCell(sheet, rowIndex, colIndex), defVal);
        /// <summary>
        /// 讀取欄位數值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static int? GetCellInt(ISheet sheet, int rowIndex, int colIndex, int? defVal = null) => ReadCellInt(GetCell(sheet, rowIndex, colIndex), defVal);
        /// <summary>
        /// 讀取欄位數值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static double? GetCellDouble(ISheet sheet, int rowIndex, int colIndex, double? defVal = null) => ReadCellDouble(GetCell(sheet, rowIndex, colIndex), defVal);
        /// <summary>
        /// 讀取欄位日期資料
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static DateTime? GetCellDate(ISheet sheet, int rowIndex, int colIndex, DateTime? defVal = null) => ReadCellDate(GetCell(sheet, rowIndex, colIndex), defVal);
        /// <summary>
        /// 讀取欄位布林資料
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static bool? GetCellBool(ISheet sheet, int rowIndex, int colIndex, bool? defVal = null) => ReadCellBool(GetCell(sheet, rowIndex, colIndex), defVal);
        /// <summary>
        /// 讀取欄位字串
        /// </summary>
        /// <param name="c"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static string? ReadCellString(ICell? c, string? defVal = "")
        {
            if (c == null) return defVal;
            return c.StringCellValue;
        }
        /// <summary>
        /// 讀取欄位數值
        /// </summary>
        /// <param name="c"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static decimal? ReadCellDecimal(ICell? c, decimal? defVal = null)
        {
            if (c == null) return defVal;
            return Convert.ToDecimal(c.NumericCellValue);
        }
        /// <summary>
        /// 讀取欄位數值
        /// </summary>
        /// <param name="c"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static int? ReadCellInt(ICell? c, int? defVal = null)
        {
            if (c == null) return defVal;
            return Convert.ToInt32(c.NumericCellValue);
        }
        /// <summary>
        /// 讀取欄位數值
        /// </summary>
        /// <param name="c"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static double? ReadCellDouble(ICell? c, double? defVal = null)
        {
            if (c == null) return defVal;
            return c.NumericCellValue;
        }
        /// <summary>
        /// 讀取欄位日期資料
        /// </summary>
        /// <param name="c"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static DateTime? ReadCellDate(ICell? c, DateTime? defVal = null)
        {
            if (c == null) return defVal;
            return c.DateCellValue;
        }
        /// <summary>
        /// 讀取欄位布林資料
        /// </summary>
        /// <param name="c"></param>
        /// <param name="defVal"></param>
        /// <returns></returns>
        public static bool? ReadCellBool(ICell? c, bool? defVal = null)
        {
            if (c == null) return defVal;
            return c.BooleanCellValue;
        }

        /// <summary>
        /// 計算Excel欄寬(for ISheet.SetColumnWidth)
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        public int CalcCellWidth(double n) => Convert.ToInt32((n + .71) * 256);
        /// <summary>
        /// 計算Excel欄寬(for ISheet.SetColumnWidth)
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        public int CalcCellWidth(int n) => Convert.ToInt32((n + .71) * 256);
        /// <summary>
        /// 計算Excel欄寬(for ISheet.SetColumnWidth)
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        public int CalcCellWidth(float n) => Convert.ToInt32((n + .71) * 256);
        /// <summary>
        /// 計算Excel欄寬(for ISheet.SetColumnWidth)
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        public int CalcCellWidth(decimal n) => Convert.ToInt32((n + .71m) * 256);
        /// <summary>
        /// 取得對應的MIME
        /// </summary>
        public static class FileMime
        {
            /// <summary>Word (.doc)</summary>
            public static readonly string Word2003 = "application/msword";
            /// <summary>Word (.docx)</summary>
            public static readonly string Word = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            /// <summary>Excel (.xls)</summary>
            public static readonly string Excel2003 = "application/vnd.ms-excel";
            /// <summary>Excel (.xlsx)</summary>
            public static readonly string Excel = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            /// <summary>PDF (.pdf)</summary>
            public static readonly string Pdf = "application/pdf";
        }
        /// <summary>
        /// 字型名稱
        /// </summary>
        public static class FontName
        {
            /// <summary>
            /// 新細明體
            /// </summary>
            public static readonly string MingLiU = "新細明體";
            /// <summary>
            /// 標楷體
            /// </summary>
            public static readonly string BiauKai = "標楷體";
            /// <summary>
            /// 微軟正黑體
            /// </summary>
            public static readonly string MSJhengHei = "微軟正黑體";
        }
    }
}
