using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace EDR_Report.Controllers
{
    [AllowAnonymous]
    public class LogoutController : BaseController
    {
        private readonly ILogger<LogoutController> _logger;

        public LogoutController(ILogger<LogoutController> logger)
        {
            _logger = logger;
        }
        public async Task<IActionResult> Index()
        {
            var db = new DBFunc();
            if (UserInfo.IsLogin)
            {
                db.Delete("cpm", "CPM_AUTH", new
                {
                    UserInfo.EMPNO
                });
            }
            var SESSION_ID = HttpContext.Session.GetString("vsSessionID");
            if (!string.IsNullOrEmpty(SESSION_ID))
            {
                try
                {
                    db.Update("erp", "VS_SESSIONS_HISTORY", new
                    {
                        LOGOUT_TIME = DateTime.Now
                    }, new
                    {
                        SESSION_ID
                    });
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"更新使用者每日上網狀況表時發生錯誤！");
                }
            }
            await HttpContext.SignOutAsync();
            HttpContext.Session.Clear();
            UserInfo = new();
            return RedirectToAction("index", "login");
        }
    }
}
