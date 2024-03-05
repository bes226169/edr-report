using EDR_Report.Models.ERP;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace EDR_Report.Controllers
{
    [Authorize]
    public class HomeController : BaseController
    {
        /// <summary>
        /// Block的圖片
        /// </summary>
        Dictionary<int, string> block_pic_name = new()
        {
            { 0, "chart.png" },
            { 1, "chartline.png" },
            { 2, "documents.png" },
            { 3, "excel.png" },
            { 4, "news.png" },
            { 5, "search.png" },
            { 6, "tips.png" },
            { 7, "wallet.png" },
            { 8, "chart.png" },
            { 9, "chartline.png" },
            { 10, "documents.png" },
            { 11, "excel.png" },
            { 12, "news.png" },
            { 13, "search.png" }
        };
        readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index() => View();

        [HttpPost]
        public IActionResult GetRptList() => Ok(new
        {
            data = new DBFunc().query<BES_PROJ_REPORT_INDEXModel>("erp", "SELECT * FROM BES_PROJ_REPORT_INDEX WHERE SYS_NAME = 'EDR' AND REPORT_DOC = 'edr' ORDER BY PROJECT_ID, SEQUENCE_NO")
        });

        [HttpPost]
        public IActionResult ChangeProj(int? pid, string? url)
        {
            if (!pid.HasValue) return Ok(new
            {
                state = false,
                msg = "參數傳遞時發生錯誤，請重新登入！"
            });
            if (!string.IsNullOrEmpty(url)) HttpContext.Session.SetString("jumpTo", url);
            var li = new DBFunc().query<BES_USER_RES_VModel>("erp", "SELECT * FROM VSUSER.BES_USER_RES_V WHERE EMPNO = :EMPNO AND PID = :pid", new
            {
                UserInfo.EMPNO,
                pid
            });
            if (li.Count() > 0)
            {
                var p = li.First();
                UserInfo.SDIV = p.DIV;
                ChangeProject(true, p.PNAME!, p.PID!.Value, p.RID!.Value, p.RNAME!);
            }
            else
            {
                return Ok(new
                {
                    state = false,
                    msg = "您沒有權限！"
                });
            }
            return Ok(new
            {
                state = true
            });
        }

        /// <summary>
        /// View: 載入管報Page
        /// </summary>
        /// <param name="pageNumber"></param>
        /// <param name="GROUPKIND2"></param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult VSList(int pageNumber, string GROUPKIND2)
        {
            DBFunc db = new();
            int pageSize = 16;
            var list = new List<dynamic>();
            if (!string.IsNullOrEmpty(GROUPKIND2))
            {
                // 取得pro_id
                var pro_id = db.query<dynamic>("CPM", "SELECT DISTINCT C.PRO_ID FROM CPM_AUTH A " +
                    "INNER JOIN CPM_PRIVILEGES_LIST B ON UPPER(B.PRI_NAME) = UPPER(A.PRI_NAME) " +
                    "INNER JOIN CPM_PRIVILEGES_PROGRAMME C ON C.PRI_ID = B.PRI_ID " +
                    $"WHERE A.EMPNO = '{UserInfo.EMPNO}'").Select(x => ((int)x.PRO_ID).ToString()).ToArray();
                if (pro_id.Length > 0)
                {
                    var s = (pageNumber - 1) * pageSize;

                    // 取得管報項目
                    list = db.query<dynamic>("cpm", $"SELECT A.*, CASE WHEN EXISTS(SELECT B.PRO_ID FROM CPM_FAVOURITE B WHERE B.PRO_ID = A.PRO_ID AND B.EMPNO = '{UserInfo.EMPNO}' ) THEN 'Y' ELSE 'N' END IS_FAV " +
                        $"FROM CPM_PROGRAMME_LIST A WHERE A.GROUPKIND2 = :GROUPKIND2 AND A.FLAG = '1' AND PRO_ID IN ({string.Join(",", pro_id)}) " +
                        $"ORDER BY A.PRO_NAME OFFSET {s} ROWS FETCH NEXT {pageSize} ROWS ONLY", new
                        {
                            GROUPKIND2
                        }).ToList();
                }
            }
            ViewBag.List = list;
            ViewBag.TotalRow = list.Count;
            ViewBag.BlockImg = block_pic_name;
            return View("VSList");
        }

        [HttpPost]
        public IActionResult GetProjectList()
        {
            var db = new DBFunc();
            var deptLi = db.query<dynamic>("erp", @"SELECT DISTINCT 
    DECODE(
        (
			SELECT DIV
			FROM (SELECT DEPT DIV, DEPT_CH DIV_CH FROM BES_DEPT_DATA WHERE DEPT_KND = 'K' OR DEPT_KND = 'E') A
			WHERE DIV = M.DIV
        ), 
        NULL, 'ZZZ', 
        (
			SELECT DIV
			FROM (SELECT DEPT DIV, DEPT_CH DIV_CH FROM BES_DEPT_DATA WHERE DEPT_KND = 'K' OR DEPT_KND = 'E') A
			WHERE DIV = M.DIV
        )
    ) DIV,
	DECODE(
        (
			SELECT DIV_CH
			FROM (SELECT DEPT DIV, DEPT_CH DIV_CH FROM BES_DEPT_DATA WHERE DEPT_KND = 'K' OR DEPT_KND = 'E') A
			WHERE DIV = M.DIV
        ), 
        NULL, '(行政)' || (SELECT DEPT_CH FROM BEERP.E104_PSNMST_ALL_NOW_V W WHERE W.EMPNO = U.USER_NAME), 
        (
			SELECT DIV || '：' || DIV_CH DIV_CH
			FROM (SELECT DEPT DIV, DEPT_CH DIV_CH FROM BES_DEPT_DATA WHERE DEPT_KND = 'K' OR DEPT_KND = 'E') A
			WHERE DIV = M.DIV
        )
    ) DIV_CH
FROM 
    VS_USER_RESPONSIBILITIES UR, 
    VS_USERS U, 
    (
		SELECT 
            RE.RESPONSIBILITY_ID RID, 
            RE.RESPONSIBILITY_NAME RNAME, 
            RP.PROJECT_ID PID, 
            P.PROJECT_NAME PNAME, 
            (SELECT DIV FROM ORDMST_REP WHERE ORD_NO = P.BES_ORD_NO) DIV, 
            P.PROJECT_SNAPSHOT PJPG
		FROM VS_RESPONSIBILITIES RE, VS_RESPONSIBILITY_PROJECTS RP, VS_PROJECTS P
		WHERE RE.ORGANIZATION_ID = 26 AND RP.RESPONSIBILITY_ID(+) = RE.RESPONSIBILITY_ID AND P.PROJECT_ID(+) = RP.PROJECT_ID
    ) M
WHERE U.USER_ID = UR.USER_ID AND UR.RESPONSIBILITY_ID = M.RID AND U.USER_ID = :UserID
ORDER BY 1 DESC", new
            {
                UserInfo.UserID
            });
            var projLi = db.query<dynamic>("erp", "SELECT DISTINCT PNAME ,PID, DIV FROM VSUSER.BES_USER_RES_V WHERE EMPNO = :EMPNO ORDER BY PNAME DESC", new
            {
                UserInfo.EMPNO
            });

            return Ok(new
            {
                deptLi,
                projLi
            });
        }
            
    }
}
