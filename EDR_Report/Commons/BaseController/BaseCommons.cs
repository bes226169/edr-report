using EDR_Report.Models.CPM;
using EDR_Report.Models.ERP;
using Microsoft.AspNetCore.Mvc;
using System.Data;

namespace EDR_Report
{
    public partial class BaseController : Controller
    {
        public UserInfoModel UserInfo { get; set; } = new();
        public bool DebugMode { get; set; } = false;
        public IConfiguration configuration { get; set; }
        public string BaseUrl { get; set; } = string.Empty;

        /// <summary>
        /// 取得使用者部門、職務等基本資訊
        /// </summary>
        /// <returns></returns>
        public List<DeptInfoModel> GetUserDeptInfo() => new DBFunc().query<DeptInfoModel>("erp", @"SELECT M.DEPT, M.DEPT_KND, M.JOB_CH, (SELECT DEPT FROM BESRSP99_ST WHERE DEPT99 = M.DEPT) DEPT_TO_99, 
DECODE(SIGN((SELECT COUNT(*) FROM BESRSP99_ST WHERE DEPT99 = M.DEPT )), 1, 'Y', 'N') VNDREMP, 
NVL((SELECT VNDR_TYP FROM BESRSP_ST WHERE DEPT = M.DEPT), '') VNDR_TYPE, 
NVL((SELECT DEPT_CH FROM BESRSP_ST WHERE DEPT = M.DEPT), '') DEPT_CH,
NVL((SELECT DEPTID FROM VSUSER.BPM_PSNMST_TRN_V W WHERE EMPNO = M.EMPNO), '') BPM_DEPT 
FROM BEEMP_ALL_V M WHERE M.EMPNO = :EMPNO", new
        {
            UserInfo.EMPNO
        }).ToList();

        public void ChangeProject(bool upd = false, string pro_name = "資訊查詢作業", int pro_id = 59, int resp_id = 107, string resp_name = "查詢資料")
        {
            var db = new DBFunc();
            if (upd)
            {
                db.Update("erp", "VS_SESSIONS", new
                {
                    PROJECT_ID = pro_id,
                    ORGANIZATION_ID = 26,
                    RESPONSIBILITY_ID = resp_id,
                }, new
                {
                    SESSION_ID = UserInfo.SessionID,
                    USER_ID = UserInfo.UserID
                });
            }
            UserInfo.ChangeProject(pro_name, pro_id, resp_id, resp_name);
            SetPermission();
            SetSessions();
        }

        void SetPermission()
        {
            var now = DateTime.Now;
            object insert_data(string PRI_NAME, string? S_VALUES) => new
            {
                SESSION_ID = UserInfo.SessionID,
                UserInfo.EMPNO,
                UPDATE_DATE = now,
                PRI_NAME,
                S_VALUES
            };
            var db = new DBFunc();
            var rstr = GetPriName().Where(x => !string.IsNullOrEmpty(x)).ToList();
            var role = db.query<dynamic>("cpm", "SELECT * FROM CPM_PRI_LIST_V WHERE EMPNO = :EMPNO", new { UserInfo.EMPNO }).ToList();
            var pri_list = db.query<dynamic>("cpm", "SELECT UPPER(pri_name) PRI_NAME, SESSIONVALUES FROM CPM_PRIVILEGES_LIST").ToList();
            var urllist = db.query<dynamic>("cpm", "SELECT 'Bpm0Link' SESSION_NAME, BPMPATH SESSION_VALUES FROM jde.jde_info").ToList();
            db.Delete("cpm", "CPM_AUTH", new { UserInfo.EMPNO });
            rstr.ForEach(pn =>
            {
                var pri = pri_list.Where(x => x.PRI_NAME == pn.ToUpper());
                var sval = pri.Count() > 0 ? (string?)pri.First().SESSIONVALUES : string.Empty;
                db.Insert("cpm", "CPM_AUTH", insert_data(pn, sval));
            });
            role.ForEach(r =>
            {
                var pn = (string?)r.PRI_NAME;
                if (string.IsNullOrEmpty(pn)) return;
                db.Insert("cpm", "CPM_AUTH", insert_data(pn, (string?)r.SESSIONVALUES));
            });
            urllist.ForEach(r =>
            {
                var pn = (string?)r.SESSION_NAME;
                if (string.IsNullOrEmpty(pn)) return;
                db.Insert("cpm", "CPM_AUTH", insert_data(pn, (string?)r.SESSION_VALUES));
            });
        }

        /// <summary>
        /// 取得登入人員權限名稱
        /// </summary>
        /// <returns></returns>
        List<string> GetPriName()
        {
            var db = new DBFunc();
            var dataTB = DBFunc.ToDataTable(db.query<dynamic>("erp", $"SELECT A.* FROM BES_RESPONSIBILITIES A " +
                $"INNER JOIN VSUSER.BES_USER_RES_V B ON A.RESPONSIBILITY_ID IN B.RID " +
                $"WHERE B.EMPNO = '{UserInfo.EMPNO}' AND B.PID = {UserInfo.ProjectID} AND B.DIV = '{UserInfo.SDIV}'"));
            var itemTB = DBFunc.ToDataTable(db.query<dynamic>("erp", "SELECT * FROM BES_RESPONSIBILITIES_ITEM").ToList()); //權限對照列表
            var rstr = new List<string>();

            foreach (DataRow dr in dataTB.Rows)
            {
                foreach (DataRow ir in itemTB.Rows)
                {
                    int seq = int.Parse(ir["SEQUENCE_NO"].ToString()!); //200以下沒有權限名稱所以必須自給
                    if (ir["CREATE_PRIV_ID"].ToString() != "0")
                    {
                        if (seq <= 201)
                            rstr.Add(proc__Responsibilitiesstr(ir["CREATE_PRIV_ID"].ToString()!, "_c", dr, seq));
                        else
                            rstr.Add(proc__Responsibilitiesstr(ir["CREATE_PRIV_ID"].ToString()!, "_c", dr, 800, ir["PRV_NAME"].ToString()!));
                    }
                    if (ir["DELETE_PRIV_ID"].ToString() != "0")
                    {
                        if (seq <= 201)
                            rstr.Add(proc__Responsibilitiesstr(ir["DELETE_PRIV_ID"].ToString()!, "_d", dr, seq));
                        else
                            rstr.Add(proc__Responsibilitiesstr(ir["DELETE_PRIV_ID"].ToString()!, "_d", dr, 800, ir["PRV_NAME"].ToString()!));
                    }
                    if (ir["VIEW_PRIV_ID"].ToString() != "0")
                    {
                        if (seq <= 201)
                            rstr.Add(proc__Responsibilitiesstr(ir["VIEW_PRIV_ID"].ToString()!, "_v", dr, seq));
                        else
                            rstr.Add(proc__Responsibilitiesstr(ir["VIEW_PRIV_ID"].ToString()!, "_v", dr, 800, ir["PRV_NAME"].ToString()!));
                    }
                    if (ir["UPDATE_PRIV_ID"].ToString() != "0")
                    {
                        if (seq <= 201)
                            rstr.Add(proc__Responsibilitiesstr(ir["UPDATE_PRIV_ID"].ToString()!, "_u", dr, seq));
                        else
                            rstr.Add(proc__Responsibilitiesstr(ir["UPDATE_PRIV_ID"].ToString()!, "_u", dr, 800, ir["PRV_NAME"].ToString()!));
                    }
                    if (ir["APPROVE_PRIV_ID"].ToString() != "0")
                    {
                        if (seq <= 201)
                            rstr.Add(proc__Responsibilitiesstr(ir["APPROVE_PRIV_ID"].ToString()!, "_a", dr, seq));
                        else
                            rstr.Add(proc__Responsibilitiesstr(ir["APPROVE_PRIV_ID"].ToString()!, "_a", dr, 800, ir["PRV_NAME"].ToString()!));
                    }
                }
            }
            return rstr;
        }
        /// <summary>
        /// 組權限名稱
        /// </summary>
        /// <param name="PRIVILEGE_COUNT"></param>
        /// <param name="type"></param>
        /// <param name="dr"></param>
        /// <param name="SEQUENCE_NO"></param>
        /// <param name="PRV_NAME"></param>
        /// <returns></returns>
        string proc__Responsibilitiesstr(string PRIVILEGE_COUNT, string type, DataRow dr, int SEQUENCE_NO = 0, string PRV_NAME = "")
        {
            if (dr["PRIVILEGE" + PRIVILEGE_COUNT].ToString() != "T") return string.Empty;
            return SEQUENCE_NO switch
            {
                10 => "prvProgressUpdate_" + type,
                20 => "prvPCO" + type,
                30 => "prvCO" + type,
                40 => "prvRFI" + type,
                50 => "prvContact" + type,
                60 => "prvDrawing" + type,
                70 => "prvSchedule" + type,
                80 => "prvResourceUsage" + type,
                81 => "prvResourceCost" + type,
                90 => "prvCriticalPath" + type,
                100 => "prvDelay" + type,
                110 => "prvQuatityTakeoff" + type,
                120 => "prvContractRpt" + type,
                130 => "prvDailyRpt" + type,
                140 => "prvSpec" + type,
                150 => "prvFile" + type,
                160 => "prvNote" + type,
                170 => "prvConflict" + type,
                180 => "prvProd" + type,
                190 => "prvMS" + type,
                200 => "prvOrgAdmin" + type,
                201 => "prvPlanner" + type,
                _ => PRV_NAME + type,
            };
        }

        /// <summary>
        /// 設定Session
        /// </summary>
        void SetSessions()
        {
            var now = DateTime.Now;
            var db = new DBFunc();
            var perdb = db.query<dynamic>("erp", @"SELECT m.empno, m.name, m.dept, m.dept_knd, m.job_ch, ( 
SELECT dept FROM besrsp99_st WHERE dept99 = m.dept ) dept_to_99, decode(sign(( SELECT count(*) FROM besrsp99_st WHERE dept99 = m.dept )), 1, 'Y', 'N') VndrEMP, 
nvl(( SELECT vndr_typ FROM besrsp_st WHERE dept = m.dept ), '') vndr_type, 
nvl(( SELECT deptid FROM vsuser.bpm_psnmst_trn_v w WHERE empno = m.empno ), '') bpm_dept FROM beemp_all_v m 
WHERE m.empno = ( SELECT user_name FROM vs_users WHERE user_id = :UserID)", new
            {
                UserInfo.UserID
            });
            if (perdb.Count() > 0)
            {
                var per = perdb.First();
                var vbpm = ((string?)per.VBPM_DEPT) ?? "";
                HttpContext.Session.SetString("vbpm_dept", vbpm);
                if ((string?)per.EMPNO == "900127")
                {
                    var jobch = (string?)per.JOB_CH ?? "";
                    HttpContext.Session.SetString("VndrEMP", "Y");
                    HttpContext.Session.SetString("vJobch", jobch);
                }
            }
            var vpridiv = db.query<dynamic>("erp", "SELECT DIV, DIV_CH FROM VSUSER.ORDMST_V WHERE PROJECT_ID = :ProjectID", new
            {
                UserInfo.ProjectID
            });
            if (vpridiv.Count() > 0)
            {
                HttpContext.Session.SetString("prjDiv", (string?)vpridiv.First().DIV ?? "");
                HttpContext.Session.SetString("prjDiv_ch", (string?)vpridiv.First().DIV_CH ?? "");
            }
            else
            {
                HttpContext.Session.SetString("prjDiv", "");
                HttpContext.Session.SetString("prjDiv_ch", "");
            }
            HttpContext.Session.SetString("vsSessionID", UserInfo.SessionID ?? "");
            HttpContext.Session.SetString("VSSessID", UserInfo.SessionID ?? "");
            HttpContext.Response.Cookies.Append("ASP.NET_SessionId", UserInfo.SessionID ?? "");
            HttpContext.Session.SetString("l_type", "zh-tw");
            HttpContext.Session.SetString("vCharset", "utf-8");
            HttpContext.Session.SetInt32("vUserID", UserInfo.UserID ?? 0);
            HttpContext.Session.SetString("vEmpno", UserInfo.EMPNO ?? "");
            HttpContext.Session.SetString("vChName", UserInfo.UserData.CH_NAME ?? "");
            HttpContext.Session.SetString("BESSSO_EMPNO", UserInfo.EMPNO ?? "");
            HttpContext.Session.SetString("BESSSO_GMAIL", UserInfo.MailAddress);
            HttpContext.Session.SetString("BESSSO_DOMAIN", UserInfo.MailAddress.Split('@')[1]);
            HttpContext.Session.SetString("vDept_ch", UserInfo.DeptInfo.DEPT_CH ?? "");
            HttpContext.Session.SetString("vDept_knd", UserInfo.DeptInfo.DEPT_KND ?? "");
            HttpContext.Session.SetString("vbpm_dept", UserInfo.DeptInfo.BPM_DEPT ?? "");
            HttpContext.Session.SetString("need_sub", "N");
            if (UserInfo.DeptInfo.VNDREMP == "N")
            {
                HttpContext.Session.SetString("vDept", UserInfo.DeptInfo.DEPT ?? "");
                UserInfo.Dept = UserInfo.DeptInfo.DEPT ?? "";
            }
            else
            {
                HttpContext.Session.SetString("vDept", UserInfo.DeptInfo.DEPT_TO_99 ?? "");
                UserInfo.Dept = UserInfo.DeptInfo.DEPT_TO_99 ?? "";
            }
            HttpContext.Session.SetString("svr_name", HttpContext.Request.Host.Host);
            HttpContext.Session.SetString("svr_path", "01");
            HttpContext.Session.SetString("svr_name_n", "01");
            HttpContext.Session.SetString("sdiv", UserInfo.SDIV ?? "ZZZ");
            HttpContext.Session.SetInt32("pProjectID", UserInfo.ProjectID ?? 59);
            HttpContext.Session.SetInt32("vProjectID", UserInfo.ProjectID ?? 59);
            HttpContext.Session.SetString("vProjectName", UserInfo.ProjectName ?? "");
            HttpContext.Session.SetInt32("vRespID", UserInfo.RespID ?? 107);
            HttpContext.Session.SetString("vRespName", UserInfo.RespName ?? "");

            db.exec("cpm", "DELETE CPM_SESSIONS WHERE EMPNO = :EMPNO OR SESSION_ID = :SessionID", new
            {
                UserInfo.EMPNO,
                UserInfo.SessionID
            });
            foreach (var s in HttpContext.Session.Keys)
            {
                db.Insert("cpm", "CPM_SESSIONS", new
                {
                    SESSION_ID = UserInfo.SessionID,
                    SESSION_NAME = s,
                    SESSION_VALUES = HttpContext.Session.GetString(s),
                    LOGINDATETIME = now,
                    UserInfo.EMPNO,
                    USER_ID = UserInfo.UserID
                });
            }
        }
    }
}
