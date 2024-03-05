using EDR_Report.Models.CPM;
using EDR_Report.Models.ERP;
using Lib;
using Lib.BesSSO;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Runtime.Versioning;

namespace EDR_Report.Controllers
{
    [AllowAnonymous]
    public class LoginController : BaseController
    {
        readonly ILogger<LoginController> _logger;
        readonly BesSSOLib BesSSOSet;

        public LoginController(ILogger<LoginController> logger)
        {
            _logger = logger;
            BesSSOSet = new(configuration.GetSection("BesSSO").Get<BesSSOModel>());
        }

        [HttpGet]
        public IActionResult Index(string? ReturnUrl)
        {
            if (!string.IsNullOrEmpty(ReturnUrl)) HttpContext.Session.SetString("jumpTo", ReturnUrl);
            if (DebugMode)
            {
                HttpContext.Session.SetString("ts", DateTime.Now.ToString("yyyyMMddHHmmss")); //防止Session 遺失
                UserInfo.SessionID = $"00{HttpContext.Session.Id.Replace("-", "")}";
                HttpContext.Session.SetString("vsSessionID", UserInfo.SessionID);
                HttpContext.Session.SetString("VSSessID", UserInfo.SessionID);
                return View("Login");
            }
            ViewBag.ErrorMsg = "SessionID遺失，請重新登入！";
            return View("ErrorMessage");
        }

        [HttpPost]
        public IActionResult test() => Ok(new
        {
            state = true,
            msg = "this is a book"
        });

        /// <summary>
        /// 取得圖片驗證碼
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [SupportedOSPlatform("windows")]
        public IActionResult GetCaptchaImage()
        {
            var captcha = Captcha.GenerateCaptchaImage(100, 50, Captcha.GenerateCaptchaCode());
            HttpContext.Session.SetString("CaptchaCode", captcha.CaptchaCode);
            return File(captcha.CaptchaByteData, "image/png");
        }

        /// <summary>
        /// 從外部連線登入
        /// </summary>
        /// <param name="vEmpno"></param>
        /// <param name="VSSessID"></param>
        /// <param name="proID"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> SSO(string? vEmpno, string? VSSessID, int? proID, string? url)
        {
            if (string.IsNullOrEmpty(vEmpno) || string.IsNullOrEmpty(VSSessID))
            {
                Response.StatusCode = 400;
                ViewBag.ErrorMsg = "SessionID遺失，請重新登入！";
                return View("ErrorMessage");
            }
            else if (string.IsNullOrEmpty(url))
            {
                Response.StatusCode = 400;
                ViewBag.ErrorMsg = "導頁錯誤！";
                return View("ErrorMessage");
            }
            else
            {
                var cli = new DBFunc().query<CPM_SESSIONSModel>("cpm", "SELECT * FROM CPM_SESSIONS WHERE SESSION_ID = :VSSessID AND EMPNO = :vEmpno", new
                {
                    VSSessID,
                    vEmpno
                }).ToList();
                if (cli.Count == 0)
                {
                    Response.StatusCode = 400;
                    ViewBag.ErrorMsg = "您已登出，請重新登入！";
                    return View("ErrorMessage");
                }
                cli.ForEach(x =>
                {
                    if (string.IsNullOrEmpty(x.SESSION_NAME)) return;
                    HttpContext.Session.SetString(x.SESSION_NAME, x.SESSION_VALUES ?? "");
                });
                Response.StatusCode = 200;
                UserInfo.EMPNO = vEmpno;
                UserInfo.SessionID = VSSessID;
                if (User.Identity == null || (User.Identity != null && User.Identity.Name != vEmpno))
                {
                    if (User.Identity != null) await HttpContext.SignOutAsync();
                    await WriteSignIn();
                }
                return Redirect(url);
            }
        }

        /// <summary>
        /// 連線到 Bes SSO 登入
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult ERP_SSO() =>
            Redirect(BesSSOSet.AuthUrl($"{BaseUrl.TrimEnd('/')}/login/erp_ssoreturn"));

        /// <summary>
        /// 帳號密碼登入
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="cacode"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoginCheck(string? username, string? password, string? cacode)
        {
            var now = DateTime.Now;
            var captchaCode = HttpContext.Session.GetString("CaptchaCode");
            var logincount = HttpContext.Session.GetInt32("logincount") ?? 0;
            UserInfo.SessionID = HttpContext.Session.GetString("vsSessionID");
            if (string.IsNullOrEmpty(captchaCode))
            {
                ViewBag.ErrorMsg = "SessionID遺失，請聯絡管理員！";
                return View("Login");
            }
            else if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                ViewBag.ErrorMsg = "You have some form errors. Please check below.";
                return View("Login");
            }
            else if (string.IsNullOrEmpty(cacode) || cacode.ToLower() != captchaCode.ToLower())
            {
                ViewBag.ErrorMsg = "驗證碼錯誤！";
                return View("Login");
            }
            var db = new DBFunc();
            var uli = db.query<VS_USERSModel>("erp", @"SELECT USER_ID, USER_NAME, PORTALPAS, STOPFLAG, STOPTIME, BES_CHKPAS_FUN(PORTALPAS, 1) PORTALPAS_DEC 
FROM VS_USERS 
WHERE USER_NAME = :username", new
            {
                username
            });
            if (uli.Count() == 0)
            {
                logincount++;
                HttpContext.Session.SetInt32("logincount", logincount);
                if (logincount > 3)
                {
                    var loginLockMinutes = configuration.GetValue<double?>("LoginLockMinutes");
                    if (!loginLockMinutes.HasValue || loginLockMinutes < 0) loginLockMinutes = 5;
                    HttpContext.Session.SetString("locktime", DateTime.Now.AddMinutes(loginLockMinutes.Value).ToString("yyyy/MM/dd HH:mm:ss"));
                }
                else
                {
                    ViewBag.ErrorMsg = "帳號或密碼輸入錯誤，請注意若錯誤超過3次，此帳號將暫停使用！";
                }
                return View("Login");
            }
            else
            {
                var u = uli.First();
                UserInfo.UserData = u;
                UserInfo.UserID = u.USER_ID;
                UserInfo.EMPNO = u.USER_NAME;
                if (u.STOPFLAG == "Y") // 鎖定
                {
                    var hn0 = configuration.GetSection("ContactHN0").Get<string[]>();
                    var msg = $"注意！此使用者帳號 {u.CH_NAME}" + (u.STOPTIME.HasValue ? $"，因於 {u.STOPTIME:yyyy/MM/dd} 連續密碼輸入錯誤，該帳號將暫停使用" : "尚未啟用") +
                            "，欲恢復使用權請聯絡總公司資訊處。";
                    if (hn0 != null && hn0.Length > 0)
                    {
                        foreach (var h in hn0) msg += $"<br />{h}";
                    }
                    ViewBag.ErrorMsg = msg;
                    return View("Login");
                }
                else if (u.PORTALPAS_DEC != password) // 密碼錯誤
                {
                    if (u.LOGINTIME > 3)
                    {
                        u.LOGINTIME = 0;
                        u.STOPFLAG = "Y";
                        u.STOPTIME = now;
                    }
                    ViewBag.ErrorMsg = "帳號或密碼輸入錯誤，請注意若錯誤超過3次，此帳號將暫停使用！";
                    return View("Login");
                }
                u.STOPFLAG = "N";
                HttpContext.Session.SetString("Passed", "true");
                var m = await WriteLoginInfo();
                if (!string.IsNullOrEmpty(m))
                {
                    ViewBag.ErrorMsg = m;
                    return View("Login");
                }
                return RedirectToAction("index", "home");
            }
        }

        /// <summary>
        /// Bes SSO Response
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> ERP_SSOReturn(ResponseModel data)
        {
            var jumpTo = HttpContext.Session.GetString("JumpTo");
            var sid = HttpContext.Session.GetString("vsSessionID");
            if (string.IsNullOrEmpty(sid))
            {
                ViewBag.ErrorMsg = "Session 遺失！";
                return View("Login");
            }
            UserInfo.SessionID = sid;
            if (data == null || !data.IsSuccess)
            {
                ViewBag.ErrorMsg = data == null ? "接收參數時發生錯誤！" : data.error_description;
                return View("Login");
            }
            var profileRes = await BesSSOSet.GetProfile(data);
            if (profileRes == null || !profileRes.IsSuccess)
            {
                ViewBag.ErrorMsg = profileRes == null ? "接收參數時發生錯誤！" : profileRes.error_description;
                return View("Login");
            }
            else if (profileRes.profile == null || string.IsNullOrEmpty(profileRes.profile.email))
            {
                ViewBag.ErrorMsg = "SSO異常！";
                return View("Login");
            }
            var db = new DBFunc();
            var ma = profileRes.profile.email.Split('@');
            if (ma[1] != "bes.com.tw")
            {
                ViewBag.ErrorMsg = "您的email不是由中華工程發出，請登出Google，再重新登入一次!";
                return View("Login");
            }
            UserInfo.EMPNO = ma[0];
            var uli = db.query<VS_USERSModel>("erp", "SELECT * FROM VS_USERS WHERE USER_NAME = :EMPNO", new
            {
                UserInfo.EMPNO
            });
            if (uli.Count() == 0)
            {
                ViewBag.ErrorMsg = "無法取得使用者資料，請聯絡管理員！";
                return View("Login");
            }
            var user = uli.First();
            UserInfo.UserID = user.USER_ID;
            HttpContext.Session.SetString("Passed", "false");
            var m = await WriteLoginInfo();
            if (!string.IsNullOrEmpty(m))
            {
                ViewBag.ErrorMsg = m;
                return View("Login");
            }
            if (string.IsNullOrEmpty(jumpTo))
            {
                return RedirectToAction("index", "home");
            }
            else
            {
                return Redirect(jumpTo);
            }
        }

        /// <summary>
        /// 寫入登入資訊及初始化
        /// </summary>
        /// <returns></returns>
        async Task<string> WriteLoginInfo()
        {
            var db = new DBFunc();
            var now = DateTime.Now;
            db.exec("erp", "DELETE FROM VS_SESSIONS WHERE SESSION_ID = :SessionID OR USER_ID = :UserID", new
            {
                UserInfo.SessionID,
                UserInfo.UserID
            });
            db.Insert("erp", "VS_SESSIONS", new
            {
                SESSION_ID = UserInfo.SessionID,
                USER_ID = UserInfo.UserID,
                ORGANIZATION_ID = 26,
                RESPONSIBILITY_ID = 107,
                PROJECT_ID = 59,
                LAST_UPDATE_DATE = now,
                LAST_UPDATED_BY = UserInfo.UserID,
                CREATION_DATE = now,
                CREATED_BY = UserInfo.UserID,
                REMOTE_ADDR = HttpContext.Connection.RemoteIpAddress?.ToString(),
                REMOTE_HOST = HttpContext.Connection.RemoteIpAddress?.ToString()
            });
            var dli = GetUserDeptInfo();
            if (dli.Count() == 0)
            {
                return "您的部門資料遺失，請聯絡管理員！";
            }
            UserInfo.DeptInfo = dli.First();
            ChangeProject();

            try
            {
                // 寫入 使用者每日上網狀況表
                db.Insert("erp", "VS_SESSIONS_HISTORY", new
                {
                    SESSION_ID = UserInfo.SessionID,
                    USER_ID = UserInfo.UserID,
                    ORGANIZATION_ID = 26,
                    LAST_UPDATE_DATE = now,
                    LAST_UPDATED_BY = UserInfo.UserID,
                    CREATION_DATE = now,
                    CREATED_BY = UserInfo.UserID,
                    REMOTE_ADDR = HttpContext.Connection.RemoteIpAddress?.ToString(),
                    REMOTE_HOST = HttpContext.Connection.RemoteIpAddress?.ToString(),
                    NAME = UserInfo.UserData.CH_NAME,
                    UserInfo.DeptInfo.DEPT,
                    UserInfo.DeptInfo.JOB_CH,
                    SVR_NAME = HttpContext.Request.Host.Host
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "寫入 使用者每日上網狀況表 失敗");
            }
            await WriteSignIn();
            return string.Empty;
        }
    }
}
