using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Newtonsoft.Json;

namespace EDR_Report
{
    public partial class BaseController
    {
        public override void OnActionExecuting(ActionExecutingContext context)
        {
            CheckLogin();
            ViewBag.HostName = configuration.GetValue<string?>("ServerHostName");
            ViewBag.BaseUrl = BaseUrl = $"https://{HttpContext.Request.Host.Value}{HttpContext.Request.PathBase}";
            ViewBag.UserInfo = UserInfo;
            ViewBag.DebugMode = DebugMode;
            if ((context.ActionDescriptor.RouteValues["controller"] ?? "").ToLower() == "report" && (context.ActionDescriptor.RouteValues["action"] ?? "").ToLower() == "edrrpt")
            {
                base.OnActionExecuting(context);
            }
            else if ((context.ActionDescriptor.RouteValues["controller"] ?? "").ToLower() == "login")
            {
                if (UserInfo.IsLogin) context.Result = RedirectToAction("index", "home");
            }
            else if (!UserInfo.IsLogin && (context.ActionDescriptor.RouteValues["controller"] ?? "").ToLower() != "logout")
            {
                context.Result = RedirectToAction("index", "logout");
            }
            base.OnActionExecuting(context);
        }
        /// <summary>
        /// 複寫輸出內容，預設MIME Type: text/html
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        public override ContentResult Content(string content)
        {
            var c = @"<!DOCTYPE html>
<html lang=""zh-tw"">
<head>
    <meta charset=""utf-8"" />
    <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />
    <title>營建管理系統</title>
</head>
<body>{0}</body>
</html>";
            return base.Content(string.Format(c, content), "text/html");
        }
        /// <summary>
        /// 複寫 OkResult，自動將內容轉換成 JSON 字串
        /// JsonResult 會發生大小寫的問題
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public override OkObjectResult Ok([ActionResultObjectValue] object? value) =>
            value is string or int or decimal or bool or DateTime or long or double or float or char or short or byte or ushort or sbyte or uint or ulong or null ?
            base.Ok(value) :
            base.Ok(JsonConvert.SerializeObject(value));
    }
}
