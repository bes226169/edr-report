using EDR_Report.Models.ERP;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;

namespace EDR_Report
{
    public partial class BaseController : Controller
    {
        public BaseController() 
        {
            configuration = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            DebugMode = configuration.GetValue<bool?>("DebugMode") ?? false;
        }

        bool CheckLogin()
        {
            if (User.Identity == null || !User.Identity.IsAuthenticated) return false;
            var db = new DBFunc();
            var sid = HttpContext.Session.GetString("vsSessionID");
            if (string.IsNullOrEmpty(sid)) return false;
            UserInfo.SessionID = User.Claims.First(x => x.Type == ClaimTypes.Sid)?.Value;
            UserInfo.EMPNO = User.Identity.Name;
            Console.WriteLine($"Empno: {UserInfo.EMPNO}, SessionID: {UserInfo.SessionID}");
            if (string.IsNullOrEmpty(UserInfo.SessionID) || string.IsNullOrEmpty(UserInfo.EMPNO)) return false;
            UserInfo.UserID = HttpContext.Session.GetInt32("vUserID");
            if (!UserInfo.UserID.HasValue) return false;
            UserInfo.ProjectID = HttpContext.Session.GetInt32("vProjectID");
            UserInfo.ProjectName = HttpContext.Session.GetString("vProjectName");
            UserInfo.RespID = HttpContext.Session.GetInt32("vRespID");
            UserInfo.RespName = HttpContext.Session.GetString("vRespName");
            UserInfo.SDIV = HttpContext.Session.GetString("sdiv");
            if (db.query<dynamic>("erp", "SELECT SESSION_ID, USER_ID FROM VS_SESSIONS WHERE SESSION_ID = :SessionID OR USER_ID = :UserID", new
            {
                UserInfo.SessionID,
                UserInfo.UserID
            }).Count() == 0) return false;
            if (db.query<dynamic>("cpm", "SELECT SESSION_ID FROM CPM_SESSIONS WHERE SESSION_ID = :SessionID AND SESSION_NAME = 'vEmpno' AND SESSION_VALUES = :EMPNO", new
            {
                UserInfo.SessionID,
                UserInfo.EMPNO
            }).Count() == 0) return false;
            var uli = db.query<VS_USERSModel>("erp", "SELECT * FROM VS_USERS WHERE USER_NAME = :EMPNO", new
            {
                UserInfo.EMPNO
            });
            if (uli.Count() == 0) return false;
            UserInfo.UserData = uli.First();
            var dli = GetUserDeptInfo();
            if (dli.Count() == 0) return false;
            UserInfo.DeptInfo = dli.First();
            UserInfo.Dept = UserInfo.DeptInfo.VNDREMP == "N" ? UserInfo.DeptInfo.DEPT : UserInfo.DeptInfo.DEPT_TO_99;
            UserInfo.Auth = GetCPMAuth();
            return UserInfo.IsLogin = true;
        }

        /// <summary>
        /// 寫入登入狀態
        /// </summary>
        /// <returns></returns>
        public async Task WriteSignIn()
        {
            await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(new ClaimsIdentity(new List<Claim>()
            {
                new(ClaimTypes.Name, UserInfo.EMPNO!),
                new(ClaimTypes.Sid, UserInfo.SessionID!),
            }, CookieAuthenticationDefaults.AuthenticationScheme)), new AuthenticationProperties { });
        }

        Dictionary<string, dynamic> GetCPMAuth() => 
            new DBFunc()
            .query<dynamic>("cpm", "SELECT DISTINCT PRI_NAME, S_VALUES FROM CPM_AUTH WHERE EMPNO = :EMPNO AND SESSION_ID = :SessionID", new
            {
                UserInfo.EMPNO,
                UserInfo.SessionID
            })
            .ToDictionary(
                keySelector: x => ((string)x.PRI_NAME).ToUpper(),
                elementSelector: x => x.S_VALUES
                );
    }
}
