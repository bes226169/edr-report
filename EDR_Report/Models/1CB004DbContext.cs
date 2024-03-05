using EDR_Report.Commons;
using EDR_Report.Controllers;
using EDR_Report.Interfaces;

namespace EDR_Report.Models
{
    public class _1CB004DbContext : I1CB004DbContext
    {
        DBFunc _db = new DBFunc();
        private readonly ILogger<ReportController> _logger;
        public _1CB004DbContext(ILogger<ReportController> logger)
        {
            _logger = logger;
        }


        /// <summary>
        /// 工程記事
        /// </summary>
        /// <param name="project_id"></param>
        /// <param name="report_date"></param>
        /// <returns></returns>
        public IEnumerable<dynamic> QueryDailyWorkNotes(int project_id, string report_date)
        {
            _logger.Log(LogLevel.Information, "Entered QueryProjDailyNotes");
            IEnumerable<dynamic> data = new List<dynamic>();

            try
            {
                var query =
                    "SELECT PROJECT_ID,                                               " +
                    "       CALENDAR_DATE,                                            " +
                    "       WEATHER_AM,                                               " +
                    "       WEATHER_PM,                                               " +
                    "       WORK_HOUR,                                                  " +
                    "       WORKDAY,                                                    " +
                    "       NOTES,                                                    " +
                    "       NOTES_A,                                                 " +
                    "       nvl(NOTES_B, '無') NOTES_B,                                " +
                    "       nvl(NOTES_C, '無') NOTES_C,                                 " +
                    "       NOTES_D,                                                " +
                    "       NOTES_E,                                                    " +
                    "       NOTES_F,                                                     " +
                    "       EXP_PERCENT,                                             " +
                    "       ACT_PERCENT,                                                 " +
                    "       ACT_SUM,                                                   " +
                    "       NOCAL_DAY                                                    " +
                    "       EXTEND_DAY                                                    " +
                    "FROM   vsuser.BES_EDR_NOTES                     " +
                    $"WHERE  PROJECT_ID = {project_id}               " +
                    $"       AND CALENDAR_DATE = To_date('{report_date}', 'yyyy/mm/dd')   ";
                data = _db.query<dynamic>("edr", query).ToList();
                _logger.Log(LogLevel.Information, "Existing QueryProjDailyNotes");
                return data;

            }
            catch (Exception ex)
            {
                string message = string.Format("{0}\n{1}\n{2}\n{3}", ex.Message, ex.InnerException);
                _logger.Log(LogLevel.Error, message);
                return data;
            }
        }


        /// <summary>
        /// 所有施工項目
        ///  case 1: 標題內有[全部有施作項目] 則印至目前有施作項目(施工項目)
        ///  case 2: 標題內有[特定施工項目]SPECIAL_ITEM='Y'						
        /// </summary>
        /// <param name="project_id"></param>
        /// <returns></returns>
        public IEnumerable<dynamic> QueryWorkItems(int project_id, WorkItemsEnum workitem)
        {
            _logger.Log(LogLevel.Information, "Entered QueryWorkItems");
            IEnumerable<dynamic> data = new List<dynamic>();

            try
            {
                var query =
                    "SELECT a.TMP_CODE                up_own_code,               " +
                    "       a.OWN_CODE,                                          " +
                    "       a.OWNITEM_NO,                                        " +
                    "       a.NAME,                                              " +
                    "       a.OWN_CONTROL_ITEM,                                  " +
                    "       a.UP_WORKITEM_ID,                                    " +
                    "       a.UNIT,                                              " +
                    "       Nvl(a.QUANTITY, 0)        quantity,                  " +
                    "       Nvl(a.OWN_UNIT_PRICE, 0)  own_unit_price,            " +
                    "       Nvl(NOW_EDR_QUANTITY, 0)  now_quantity,              " +
                    "       Nvl(PRV_EDR_QUANTITY, 0)  prv_quantity,              " +
                    "       Nvl(SUM_EDR_QUANTITY, 0)  sum_quantity,              " +
                    "       Nvl(QTY_DECIMAL_PLACE, 0) qty_decimal_place,         " +
                    "       Nvl(BUILD_SUB, ' ')       build_sub                  " +
                    "FROM   VSUSER.BES_EDR_TEMP_QUANTITY a                              " +
                    $"WHERE  ROWNUM <=10 AND a.PROJECT_ID = {project_id}                         " +
                    "       AND a.CREATED_BY = 25452                             " + //TODO
                                                                                     //"       AND a.CREATED_BY = 19537                             " +
                    "       :FILTER                                              " +
                    "ORDER  BY 1                                                 ";
                string filter = string.Empty;

                switch (workitem)
                {
                    case WorkItemsEnum.施工項目:
                    default:
                        filter = "";
                        break;
                    case WorkItemsEnum.特定施工項目:
                        filter = $" AND SPECIAL_ITEM = 'Y'             ";
                        break;
                }
                query = query.Replace(":FILTER", filter);

                data = _db.query<dynamic>("edr", query).ToList();
                _logger.Log(LogLevel.Information, "Existing QueryWorkItems");
                return data;

            }
            catch (Exception ex)
            {
                string message = string.Format("{0}\n{1}\n{2}\n{3}", ex.Message, ex.InnerException);
                _logger.Log(LogLevel.Error, message);
                return data;
            }

        }

        /// <summary>
        /// --第三部表身開始施工日報2023head1body.html	
        /// 人工:2
        /// 材料:4	
        /// 機具:3機具+5器材
        /// </summary>
        /// <param name="project_id"></param>
        /// <param name="report_date"></param>
        /// <returns></returns>
        public IEnumerable<dynamic> QueryResourceItems(int project_id, string report_date, ResourceClassEnum resClass)
        {
            _logger.Log(LogLevel.Information, "Entered QueryResourceItems");
            IEnumerable<dynamic> data = new List<dynamic>();

            try
            {
                var query = string.Empty;
                switch (resClass)
                {
                    case ResourceClassEnum.機具器材: //3機具+5器材 or 任何複數種類
                        query =
                            "SELECT s.OWN_RES_CODE      pbg_code,                                        " +
                            "       s.NAME,                                                              " +
                            "       Nvl(e.TODAY_QTY, 0) today_qty,                                       " +
                            "       Nvl(s.SUM_QTY, 0)   sum_qty,                                         " +
                            "       Nvl(e.NUM_SET, 0)   q_num,                                           " +
                            "       Nvl(s.S_NUM, 0)     s_num                                            " +
                            "FROM   (SELECT OWN_RES_CODE,                                                " +
                            "               SUM(Nvl(QUANTITY, 0)) today_qty,                             " +
                            "               SUM(Nvl(NUM_SET, 0))  num_set                                " +
                            "        FROM   vsuser.BES_EDR_RESQTY_V                                      " +
                            "        WHERE  PROJECT_ID = 4780                                            " +
                            $"               AND DATA_DATE = To_date('{report_date}', 'yyyy/mm/dd')      " +
                            "               AND QUANTITY >= 0                                            " +
                            $"               AND ( RESOURCE_CLASS = '{(int)ResourceClassEnum.機具}'            " +
                            $"                      OR RESOURCE_CLASS = '{(int)ResourceClassEnum.器材}' )      " +
                            "        GROUP  BY OWN_RES_CODE) e,                                          " +
                            "       (SELECT OWN_RES_CODE,                                                " +
                            "               NAME,                                                        " +
                            "               SUM(QUANTITY)        sum_qty,                                " +
                            "               SUM(Nvl(NUM_SET, 0)) s_num                                   " +
                            "        FROM   vsuser.BES_EDR_RESQTY_V                                      " +
                            "        WHERE  PROJECT_ID = 4780                                            " +
                            $"               AND DATA_DATE <= To_date('{report_date}', 'yyyy/mm/dd')     " +
                            "               AND QUANTITY >= 0                                            " +
                            $"               AND ( RESOURCE_CLASS =  '{(int)ResourceClassEnum.機具}'           " +
                            $"                      OR RESOURCE_CLASS = '{(int)ResourceClassEnum.器材}' )      " +
                            "        GROUP  BY OWN_RES_CODE,                                             " +
                            "                  NAME) s                                                   " +
                            "WHERE  e.OWN_RES_CODE = s.OWN_RES_CODE                                      ";

                        break;
                    default: //人工:2  材料:4	 等單一種類
                        query =
                            "SELECT s.OWN_RES_CODE      pbg_code,                                      " +
                            "       s.NAME,                                                            " +
                            "       UNIT,                                                              " +
                            "       Nvl(e.TODAY_QTY, 0) today_qty,                                     " +
                            "       Nvl(s.SUM_QTY, 0)   sum_qty                                        " +
                            "FROM   (SELECT OWN_RES_CODE,                                              " +
                            "               SUM(Nvl(QUANTITY, 0)) today_qty                            " +
                            "        FROM   vsuser.BES_EDR_RESQTY_V                                    " +
                            "        WHERE  PROJECT_ID = 4780                                          " +
                            //TODO $"               AND DATA_DATE = To_date('{report_date}', 'yyyy/mm/dd')    " +
                            "               AND QUANTITY >= 0                                          " +
                            $"               AND RESOURCE_CLASS = '{(int)resClass}'            " +
                            "        GROUP  BY OWN_RES_CODE) e,                                        " +
                            "       (SELECT OWN_RES_CODE,                                              " +
                            "               NAME,                                                      " +
                            "               UNIT,                                                      " +
                            "               Nvl(SUM(QUANTITY), 0) sum_qty                              " +
                            "        FROM   vsuser.BES_EDR_RESQTY_V                                    " +
                            "        WHERE  PROJECT_ID = 4780                                          " +
                            //TODO $"               AND DATA_DATE <= To_date('{report_date}', 'yyyy/mm/dd')   " +
                            "               AND QUANTITY >= 0                                          " +
                            $"               AND RESOURCE_CLASS = '{(int)resClass}'            " +
                            "        GROUP  BY OWN_RES_CODE,                                           " +
                            "                  NAME,                                                   " +
                            "                  UNIT) s                                                 " +
                            "WHERE  e.OWN_RES_CODE = s.OWN_RES_CODE                                    ";
                        break;
                }
                data = _db.query<dynamic>("edr", query).ToList();
                _logger.Log(LogLevel.Information, "Existing QueryResourceItems");
                return data;

            }
            catch (Exception ex)
            {
                string message = string.Format("{0}\n{1}\n{2}\n{3}", ex.Message, ex.InnerException);
                _logger.Log(LogLevel.Error, message);
                return data;
            }
        }

        /// <summary>
        /// 工程進度查詢
        /// </summary>
        /// <param name="project_id"></param>
        /// <param name="report_date"></param>
        /// <returns></returns>
        public IEnumerable<dynamic> QueryWorkProgress(int project_id, string report_date)
        {
            _logger.Log(LogLevel.Information, "Entered QueryProjProgress");
            IEnumerable<dynamic> data = new List<dynamic>();

            try
            {
                var query =
                    "SELECT O.ORD_NO,                                                                                                                             " +
                    "       Nvl(O.ORD_CH, '')                                                ord_ch,                                                              " +
                    "       DIV_CH,                                                                                                                               " +
                    "       DIV,                                                                                                                                  " +
                    "       Nvl(A.OWN_CONTRACT_NO, ' ')                                      own_no,                                                              " +
                    "       O.ABDATE,                                                                                                                             " +
                    "       O.BEDATE,                                                                                                                             " +
                    "       O.OWN_CH  owner_name,                                                                                                                 " +
                    "       O.SUPV,                                                                                                                               " +
                    "       O.BBDATE,                                                                                                                             " +
                    "       Nvl(O.CUR_NT_ADD, 0) cur_nt_add,                                                                                                      " +
                    "       O.EDESC,                                                                                                                              " +
                    "       Nvl(o.CUR_NT, 0)  cur_nt,                                                                                                             " +
                    "       Nvl(SPREAD_DAY, 0) Spread_day,                                                                                                        " +
                    "       A.LOCATION,                                                                                                                           " +
                    $"       Nvl(( To_date(To_char(A.PROJ_SPREAD_DATE, 'yyyy/mm/dd'), 'yyyy/mm/dd')- To_date('{report_date}', 'yyyy/mm/dd') ), 0)  v_day,          " +
                    "       Nvl(ORIGINAL_DAY, 0)                                             t_day,                                                               " +
                    $"       ( ( To_date('{report_date}', 'yyyy/mm/dd') - To_date(To_char(A.START_DATE, 'yyyy/mm/dd'), 'yyyy/mm/dd') ) + 1 ) s_day,                " +
                    "       Nvl(EDR_ACTQTY_DECIMAL_PLACE, 0)                                                                                                      " +
                    "       edr_actqty_decimal_place                                                                                                              " +
                    "FROM   VSUSER.VS_PROJECTS A,                                                                                                                 " +
                    "       VSUSER.ORDMST_REP O                                                                                                                   " +
                    $"WHERE  A.PROJECT_ID = {project_id}                                                                                                           " +
                    "       AND A.BES_ORD_NO = O.ORD_NO                                                                                                           ";
                //"SELECT o.ORD_NO,                                 " +
                //"       Nvl(o.ORD_CH, '')            ord_ch,      " +
                //"       DIV_CH,                                                         " +
                //"       DIV,                                                           " +
                //"       Nvl(a.OWN_CONTRACT_NO, ' ')                                      own_no,      " +
                //"       o.ABDATE,                        " +
                //"       o.BEDATE,                         " +
                //"       o.OWN_CH   owner_name,                                                        " +
                //"       o.SUPV,                                                                       " +
                //"       o.BBDATE,                                                                     " +
                //"       Nvl(o.CUR_NT_ADD, 0)  cur_nt_add,                                             " +
                //"       Nvl(o.CUR_NT, 0)  cur_nt,                                                     " +
                //"       o.EDESC,                                                                      " +
                //"       Nvl(SPREAD_DAY, 0)                                                            " +
                //"       Spread_day,                                                 " +
                //"       a.LOCATION,                                                                   " +
                //$"       Nvl((To_date(To_char(a.PROJ_SPREAD_DATE,'yyyy/mm/dd'),'yyyy/mm/dd')-To_date('{report_date}','yyyy/mm/dd') ), 0)  v_day,   " +
                //"       Nvl(ORIGINAL_DAY, 0)   t_day," +
                //$"       Nvl(To_date('{report_date}','yyyy/mm/dd')-To_date(a.START_DATE,'yyyy/mm/dd')+ 1 ,0) s_day,       " +
                //"       Nvl(EDR_ACTQTY_DECIMAL_PLACE, 0)                                              " +
                //"       edr_actqty_decimal_place                                                      " +
                //"FROM   vsuser.VS_PROJECTS a,                                                         " +
                //"       vsuser.ORDMST_REP o                                                           " +
                //$"WHERE  a.PROJECT_ID = {project_id}                                                  " +
                //"       AND a.BES_ORD_NO = o.ORD_NO                                                   ";

                data = _db.query<dynamic>("edr", query).ToList();
                _logger.Log(LogLevel.Information, "Existing QueryProjProgress");
                return data;

            }
            catch (Exception ex)
            {
                string message = string.Format("{0}\n{1}\n{2}\n{3}", ex.Message, ex.InnerException);
                _logger.Log(LogLevel.Error, message);
                return data;
            }
        }
    }
}
