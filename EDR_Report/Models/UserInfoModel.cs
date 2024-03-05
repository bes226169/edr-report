using EDR_Report.Models.CPM;
using EDR_Report.Models.ERP;

namespace EDR_Report
{
    public class UserInfoModel
    {
        /// <summary>
        /// 是否登入
        /// </summary>
        public bool IsLogin { get; set; } = false;
        /// <summary>
        /// Session ID
        /// </summary>
        public string? SessionID { get; set; }
        /// <summary>
        /// User ID
        /// </summary>
        public int? UserID { get; set; }
        /// <summary>
        /// 工號
        /// </summary>
        public string? EMPNO { get; set; }
        /// <summary>
        /// 部門代碼
        /// </summary>
        public string? Dept {  get; set; }
        /// <summary>
        /// Email Address
        /// </summary>
        public string MailAddress
        {
            get => string.IsNullOrEmpty(EMPNO) ? string.Empty : $"{EMPNO}@bes.com.tw";
        }
        /// <summary>
        /// VS_USERS
        /// </summary>
        public VS_USERSModel UserData { get; set; } = new();
        /// <summary>
        /// 部門資訊
        /// </summary>
        public DeptInfoModel DeptInfo { get; set; } = new();
        /// <summary>
        /// 工務所代碼
        /// </summary>
        public string? SDIV { get; set; } = "ZZZ";
        /// <summary>
        /// 專案ID
        /// </summary>
        public int? ProjectID { get; set; }
        /// <summary>
        /// 專案名稱
        /// </summary>
        public string? ProjectName { get; set; }
        /// <summary>
        /// 權限ID
        /// </summary>
        public int? RespID { get; set; }
        /// <summary>
        /// 權限名稱
        /// </summary>
        public string? RespName { get; set; }
        public void ChangeProject(string _ProjectName, int _ProjectID, int _RespID, string _RespName)
        {
            ProjectName = _ProjectName;
            ProjectID = _ProjectID;
            RespID = _RespID;
            RespName = _RespName;
        }
        /// <summary>
        /// 從 CPM_AUTH 取得該使用者在目前專案的權限設定
        /// Key: PRI_NAME (大寫)
        /// Value: S_VALUES
        /// </summary>
        public Dictionary<string, dynamic> Auth { get; set; } = new();
        /// <summary>
        /// 檢查是否有指定的權限
        /// <para>若輸入多個，只要有一個符合就回傳 true</para>
        /// </summary>
        /// <param name="a"></param>
        /// <returns></returns>
        public bool HasAuth(params string[] a) =>
            a != null && a.Length != 0 && a.Where(x => Auth.Where(x => new string[] { "T", "Y" }.Contains((((string?)x.Value) ?? "").ToUpper())).ToDictionary(x => x.Key, x => x.Value).ContainsKey(x.ToUpper())).Count() > 0;
    }
}
