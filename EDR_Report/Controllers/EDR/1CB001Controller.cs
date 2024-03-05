using Dapper;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace EDR_Report.Controllers
{
    public partial class ReportController
    {
        /// <summary>
        /// B0A 新展工務所
        /// <para>Project ID: 5320</para>
        /// <para>工令: 1CB001</para>
        /// </summary>
        /// <param name="state"></param>
        /// <param name="cd"></param>
        /// <returns></returns>
        IActionResult RPT1CB001(string? state, string cd)
        {
            // 這支沒有使用套表，所以不會有套表的檔案路徑
            var yy = int.Parse(cd[..4]);
            var mm = int.Parse(cd[4..6]);
            var dd = int.Parse(cd[6..8]);
            var calendar_date = new DateTime(yy, mm, dd); // 日報日期
            var userid = UserInfo.UserID; // User ID
            var projectid = UserInfo.ProjectID; // Project ID

            string format_chdate(string? dt)
            {
                if (string.IsNullOrEmpty(dt)) return "";
                try
                {
                    return $"{dt[..3]}年{dt[3..5]}月{dt[5..7]}日";
                }
                catch
                {
                    return "";
                }
            }
            var db = new DBFunc();

            #region 產生施作項目
            var dp = new DynamicParameters();
            dp.Add("myproject_id", projectid);
            dp.Add("myuser_id", userid);
            dp.Add("MYDATE", calendar_date);
            dp.Add("MYSTATE", state); // 1:全部, 2:本日, 3: 至本日累計
            db.sp("edr", "VSUSER.BES_EDR_REP1", dp);
            #endregion

            #region 取資料
            var prjLi = db.query<ProjectInfoModel>("edr", @"SELECT 
    O.ORD_NO, -- 工令
    NVL(O.ORD_CH, '') ORD_CH, --工程名稱
    DIV_CH,--工務所名稱
    DIV, --工務所代碼
    NVL(A.OWN_CONTRACT_NO, ' ') OWN_NO,
    O.ABDATE, --開工日期 (1080812--> 要轉換 108年08月12日)
    O.BEDATE, --完工日期(1121225--> 要轉換112年12月25日)
    O.OWN_CH OWNER_NAME,
    O.SUPV,
    O.BBDATE, 
    NVL(O.CUR_NT_ADD, 0) CUR_NT_ADD,
    O.EDESC,
    NVL(SPREAD_DAY, 0) SPREAD_DAY, --  展延天數
    A.LOCATION,
    NVL(
        (
            TO_DATE(TO_CHAR(A.PROJ_SPREAD_DATE, 'YYYY/MM/DD'), 'YYYY/MM/DD') - 
            TO_DATE(:cd1, 'YYYY/MM/DD')
        ),
        0
    ) V_DAY,  --剩餘工期
    NVL(ORIGINAL_DAY, 0) T_DAY, --核定工期
    (
        (
            TO_DATE(:cd2, 'YYYY/MM/DD') -
            TO_DATE(TO_CHAR(A.START_DATE, 'YYYY/MM/DD'), 'YYYY/MM/DD')
        ) + 1
    ) S_DAY, --累計工期
    NVL(EDR_ACTQTY_DECIMAL_PLACE, 0) EDR_ACTQTY_DECIMAL_PLACE
FROM VSUSER.VS_PROJECTS A, VSUSER.ORDMST_REP O
WHERE
    A.PROJECT_ID = :projectid AND 
    A.BES_ORD_NO = O.ORD_NO", new
            {
                cd1 = cd,
                cd2 = cd,
                projectid
            });
            var noteLi = db.query<EDRNotesModel>("edr", @"SELECT 
    PROJECT_ID, 
    CALENDAR_DATE,  -- 日報日期
    WEATHER_AM,     -- 天氣(上)
    WEATHER_PM,     -- 天氣(下)
    WORK_HOUR,      --  工作時數
    WORKDAY,        -- 工作天(Y/N)
    NOTES,          --工地記事 (其他事項)
    NOTES_A,        --工地記事 (六、施工取樣試驗紀錄：)
    NOTES_B,        --工地記事 (七、通知協力廠商辦理事項：)
    NOTES_C,        --工地記事 (八、重要事項記錄：)
    NOTES_D,        --工地記事 (九、主要工作項目：)
    NOTES_E,        --工地記事
    NOTES_F,        --工地記事
    EXP_PERCENT,    --預定進度(%) (至本日累計預定進度)
    ACT_PERCENT,    --本日實際進度
    ACT_SUM,        --實際進度(%) (至本日累計實際進度)
    NOCAL_DAY,      --免計工期(天)
    EXTEND_DAY      --展延工期(天)
FROM VSUSER.BES_EDR_NOTES
WHERE 
    PROJECT_ID = :projectid AND 
    CALENDAR_DATE = TO_DATE(:cd, 'YYYY/MM/DD')", new
            {
                projectid,
                cd
            });
            var etqLi = db.query<EDRTempQuantityModel>("edr", @"SELECT 
    A.TMP_CODE UP_OWN_CODE,
    A.OWN_CODE,
    A.OWNITEM_NO,
    A.NAME,
    A.OWN_CONTROL_ITEM,
    A.UP_WORKITEM_ID,
    A.UNIT,
    A.SPECIAL_ITEM, 
    NVL(A.QUANTITY, 0) QUANTITY,
    NVL(A.OWN_UNIT_PRICE, 0) OWN_UNIT_PRICE,
    NVL(NOW_EDR_QUANTITY, 0) NOW_QUANTITY,
    NVL(PRV_EDR_QUANTITY, 0) PRV_QUANTITY,
    NVL(SUM_EDR_QUANTITY, 0) SUM_QUANTITY,
    NVL(QTY_DECIMAL_PLACE, 0) QTY_DECIMAL_PLACE,	
    NVL(BUILD_SUB, ' ') BUILD_SUB
FROM VSUSER.BES_EDR_TEMP_QUANTITY A				
WHERE A.PROJECT_ID = :projectid AND A.CREATED_BY = :userid", new
            {
                projectid,
                userid
            });
            var eqrLi = db.query<EDRListCountModel>("edr", @"    
    
SELECT 
    S.OWN_RES_CODE PBG_CODE,
    E.RESOURCE_CLASS,
    S.NAME,
    UNIT,
    NVL(E.TODAY_QTY, 0) TODAY_QTY,
    NVL(S.SUM_QTY, 0) SUM_QTY
FROM 
    (
        SELECT 
            OWN_RES_CODE, 
            RESOURCE_CLASS,
            SUM(NVL(QUANTITY, 0)) TODAY_QTY	
        FROM VSUSER.BES_EDR_RESQTY_V	
        WHERE 
            PROJECT_ID = :projectid1 AND 
            DATA_DATE = TO_DATE(:cd1, 'YYYY/MM/DD')	AND 
            QUANTITY >= 0
        GROUP BY OWN_RES_CODE, RESOURCE_CLASS
     ) E,
    (
        SELECT 
            OWN_RES_CODE, 
            RESOURCE_CLASS,
            NAME, 
            UNIT, 
            NVL(SUM(QUANTITY), 0) SUM_QTY
        FROM VSUSER.BES_EDR_RESQTY_V
        WHERE 
            PROJECT_ID = :projectid2 AND 
            DATA_DATE <= TO_DATE(:cd2, 'YYYY/MM/DD') AND 
            QUANTITY >= 0
        GROUP BY OWN_RES_CODE, RESOURCE_CLASS, NAME, UNIT
    ) S
WHERE 
    E.OWN_RES_CODE = S.OWN_RES_CODE AND 
    E.RESOURCE_CLASS = S.RESOURCE_CLASS", new
            {
                projectid1 = projectid,
                cd1 = cd,
                projectid2 = projectid,
                cd2 = cd,
            });
            #endregion

            if (prjLi.Count() == 0 || noteLi.Count() == 0) return NotFound("找不到資料");
            var prj = prjLi.First();
            var note = noteLi.First();

            IWorkbook wb = new XSSFWorkbook();
            var ws = wb.CreateSheet("公共工程施工日誌"); // 建立Sheet

            // 欄位寬度
            var col_width = new int[]
            {
                CalcCellWidth(13), // A
                CalcCellWidth(8.5), // B
                CalcCellWidth(8.5), // C
                CalcCellWidth(8.5), // D
                CalcCellWidth(8.5), // E
                CalcCellWidth(8.5), // F
                CalcCellWidth(8.5), // G
                CalcCellWidth(8.5), // H
                CalcCellWidth(8.5), // I
                CalcCellWidth(13), // J
                CalcCellWidth(8.5), // K
                CalcCellWidth(8.5), // L
                CalcCellWidth(8.5), // M
                CalcCellWidth(8.5), // N
                CalcCellWidth(8.5), // O
                CalcCellWidth(8.5), // P
                CalcCellWidth(8.5), // Q
                CalcCellWidth(8.5), // R
                CalcCellWidth(8.5), // S
            };

            #region Font
            var h1_font = wb.CreateFont();
            h1_font.FontName = "標楷體";
            h1_font.FontHeightInPoints = 18;
            h1_font.IsBold = true;

            var n_font = wb.CreateFont();
            n_font.FontName = "標楷體";
            n_font.FontHeightInPoints = 12;

            var s_font = wb.CreateFont();
            s_font.FontName = "標楷體";
            s_font.FontHeightInPoints = 10;

            var ss_font = wb.CreateFont();
            ss_font.FontName = "標楷體";
            ss_font.FontHeightInPoints = 8;
            #endregion

            #region Style
            var h1_style = wb.CreateCellStyle();
            h1_style.SetFont(h1_font);
            h1_style.Alignment = HorizontalAlignment.Center;

            var h2_style = wb.CreateCellStyle();
            h2_style.SetFont(n_font);

            var h3_style = wb.CreateCellStyle();
            h3_style.SetFont(n_font);
            h3_style.Alignment = HorizontalAlignment.Center;

            var l_style = wb.CreateCellStyle();
            l_style.SetFont(n_font);
            l_style.Alignment = HorizontalAlignment.Left;
            l_style.VerticalAlignment = VerticalAlignment.Center;
            l_style.WrapText = true;
            l_style.BorderBottom = l_style.BorderTop = l_style.BorderLeft = l_style.BorderRight = BorderStyle.Thin;

            var tl_style = wb.CreateCellStyle();
            tl_style.SetFont(n_font);
            tl_style.Alignment = HorizontalAlignment.Left;
            tl_style.VerticalAlignment = VerticalAlignment.Top;
            tl_style.WrapText = true;
            tl_style.BorderBottom = tl_style.BorderTop = tl_style.BorderLeft = tl_style.BorderRight = BorderStyle.Thin;

            var sl_style = wb.CreateCellStyle();
            sl_style.SetFont(s_font);
            sl_style.Alignment = HorizontalAlignment.Left;
            sl_style.VerticalAlignment = VerticalAlignment.Center;
            sl_style.WrapText = true;
            sl_style.BorderBottom = sl_style.BorderTop = sl_style.BorderLeft = sl_style.BorderRight = BorderStyle.Thin;

            var ssl_style = wb.CreateCellStyle();
            ssl_style.SetFont(ss_font);
            ssl_style.Alignment = HorizontalAlignment.Left;
            ssl_style.VerticalAlignment = VerticalAlignment.Center;
            ssl_style.WrapText = true;
            ssl_style.BorderBottom = ssl_style.BorderTop = ssl_style.BorderLeft = ssl_style.BorderRight = BorderStyle.Thin;

            var c_style = wb.CreateCellStyle();
            c_style.SetFont(n_font);
            c_style.Alignment = HorizontalAlignment.Center;
            c_style.VerticalAlignment = VerticalAlignment.Center;
            c_style.WrapText = true;
            c_style.BorderBottom = c_style.BorderTop = c_style.BorderLeft = c_style.BorderRight = BorderStyle.Thin;

            var sc_style = wb.CreateCellStyle();
            sc_style.SetFont(s_font);
            sc_style.Alignment = HorizontalAlignment.Center;
            sc_style.VerticalAlignment = VerticalAlignment.Center;
            sc_style.WrapText = true;
            sc_style.BorderBottom = sc_style.BorderTop = sc_style.BorderLeft = sc_style.BorderRight = BorderStyle.Thin;

            var r_style = wb.CreateCellStyle();
            r_style.SetFont(n_font);
            r_style.Alignment = HorizontalAlignment.Right;
            r_style.VerticalAlignment = VerticalAlignment.Center;
            r_style.WrapText = true;
            r_style.BorderBottom = r_style.BorderTop = r_style.BorderLeft = r_style.BorderRight = BorderStyle.Thin;

            var sr_style = wb.CreateCellStyle();
            sr_style.SetFont(s_font);
            sr_style.Alignment = HorizontalAlignment.Right;
            sr_style.VerticalAlignment = VerticalAlignment.Center;
            sr_style.WrapText = true;
            sr_style.BorderBottom = sr_style.BorderTop = sr_style.BorderLeft = sr_style.BorderRight = BorderStyle.Thin;
            #endregion

            var ri = 0;
            SetCell(ws, ri, 0, "公共工程施工日誌", h1_style, 1, 19);

            ri++;
            SetCell(ws, ri, 0, "表報編號：", h2_style);
            SetCell(ws, ri, 1, $"1100220-1{prj.S_DAY:00000000}", h3_style, 1, 5);

            ri++;
            SetCell(ws, ri, 0, "本日天氣：", h2_style);
            SetCell(ws, ri, 1, $"上午：　{note.WEATHER_AM}", h2_style, 1, 3);
            SetCell(ws, ri, 4, $"下午：　{note.WEATHER_PM}", h2_style, 1, 3);
            SetCell(ws, ri, 11, "填表日期：", h2_style, 1, 2);
            SetCell(ws, ri, 13, note.CALENDAR_DATE?.ToString("yyyy年MM月dd日"), h3_style, 1, 3);
            SetCell(ws, ri, 16, note.CALENDAR_DATE?.ToString("dddd"), h3_style, 1, 3);

            ri++;
            SetCell(ws, ri, 0, "工程名稱", c_style, 1, 2);
            SetCell(ws, ri, 2, prj.ORD_CH, l_style, 1, 7);
            SetCell(ws, ri, 9, "承攬廠商名稱", c_style, 1, 2);
            SetCell(ws, ri, 11, "中華工程股份有限公司", l_style, 1, 8);

            ri++;
            SetCell(ws, ri, 0, "核定工程", c_style, 1, 2);
            SetCell(ws, ri, 2, $"{prj.T_DAY}日曆天", l_style, 1, 3);
            SetCell(ws, ri, 5, "累計工期", c_style, 1, 2);
            SetCell(ws, ri, 7, $"{prj.S_DAY}天", l_style, 1, 2);
            SetCell(ws, ri, 9, "剩餘工期", c_style, 1, 2);
            SetCell(ws, ri, 11, $"{prj.V_DAY}天", l_style, 1, 3);
            SetCell(ws, ri, 14, "工期展延天數", c_style, 1, 2);
            SetCell(ws, ri, 16, $"{prj.SPREAD_DAY}天", l_style, 1, 3);

            ri++;
            SetCell(ws, ri, 0, "開工日期", c_style, 1, 3);
            SetCell(ws, ri, 3, format_chdate(prj.ABDATE), c_style, 1, 6);
            SetCell(ws, ri, 9, "完工日期", c_style, 1, 3);
            SetCell(ws, ri, 12, format_chdate(prj.BEDATE), c_style, 1, 7);

            ri++;
            SetCell(ws, ri, 0, "預定進度(%)", c_style, 1, 3);
            SetCell(ws, ri, 3, $"{note.EXP_PERCENT}%", c_style, 1, 6);
            SetCell(ws, ri, 9, "實際進度(%)", c_style, 1, 3);
            SetCell(ws, ri, 12, $"{note.ACT_SUM}%", c_style, 1, 7);

            ri++;
            SetCell(ws, ri, 0, "一、依施工計畫書執行按圖施工概況（含約定之重要施工項目及完成數量等）：", l_style, 1, 19);

            ri++;
            SetCell(ws, ri, 0, "施工項目", c_style, 1, 11);
            SetCell(ws, ri, 11, "單位", c_style);
            SetCell(ws, ri, 12, "契約數量", c_style, 1, 2);
            SetCell(ws, ri, 14, "本日完成數量", c_style, 1, 2);
            SetCell(ws, ri, 16, "累計完成數量", c_style, 1, 2);
            SetCell(ws, ri, 18, "備註", c_style);

            var etqN = etqLi.Where(x => x.SPECIAL_ITEM != "Y");
            if (etqN.Count() > 0)
            {
                foreach (var etq in etqN)
                {
                    ri++;
                    SetCell(ws, ri, 0, etq.NAME, sl_style, 1, 11);
                    SetCell(ws, ri, 11, etq.UNIT, c_style);
                    SetCell(ws, ri, 12, etq.QUANTITY?.ToString("#,#0"), r_style, 1, 2);
                    SetCell(ws, ri, 14, etq.NOW_QUANTITY?.ToString("#,#0.00"), r_style, 1, 2);
                    SetCell(ws, ri, 16, etq.SUM_QUANTITY?.ToString("#,#0.00"), r_style, 1, 2);
                    SetCell(ws, ri, 18, "", sl_style);
                }
            } 
            else
            {
                ri++;
                SetCell(ws, ri, 0, "-", sl_style, 1, 11);
                SetCell(ws, ri, 11, "-", c_style);
                SetCell(ws, ri, 12, "-", r_style, 1, 2);
                SetCell(ws, ri, 14, "-", r_style, 1, 2);
                SetCell(ws, ri, 16, "-", r_style, 1, 2);
                SetCell(ws, ri, 18, "-", sl_style);
            }

            ri++;
            SetCell(ws, ri, 0, "營造業專業工程特定施工項目", l_style, 1, 19);

            var etqSp = etqLi.Where(x => x.SPECIAL_ITEM == "Y");
            if (etqSp.Count() > 0)
            {
                foreach (var etq in etqSp)
                {
                    ri++;
                    SetCell(ws, ri, 0, etq.NAME, sl_style, 1, 11);
                    SetCell(ws, ri, 11, etq.UNIT, c_style);
                    SetCell(ws, ri, 12, etq.QUANTITY?.ToString("#,#0"), r_style, 1, 2);
                    SetCell(ws, ri, 14, etq.NOW_QUANTITY?.ToString("#,#0.00"), r_style, 1, 2);
                    SetCell(ws, ri, 16, etq.SUM_QUANTITY?.ToString("#,#0.00"), r_style, 1, 2);
                    SetCell(ws, ri, 18, "", sl_style);
                }
            }
            else
            {
                ri++;
                SetCell(ws, ri, 0, "-", sl_style, 1, 11);
                SetCell(ws, ri, 11, "-", c_style);
                SetCell(ws, ri, 12, "-", r_style, 1, 2);
                SetCell(ws, ri, 14, "-", r_style, 1, 2);
                SetCell(ws, ri, 16, "-", r_style, 1, 2);
                SetCell(ws, ri, 18, "-", sl_style);
            }

            ri++;
            SetCell(ws, ri, 0, "二、工地材料管理概況（含約定之重要材料使用狀況及數量等）：", l_style, 1, 19);

            ri++;
            SetCell(ws, ri, 0, "材料名稱", sc_style, 1, 2);
            SetCell(ws, ri, 2, "單位", sc_style);
            SetCell(ws, ri, 3, "契約數量", sc_style, 1, 2);
            SetCell(ws, ri, 5, "本日使用數量", sc_style, 1, 2);
            SetCell(ws, ri, 7, "累計使用數量", sc_style, 1, 2);
            SetCell(ws, ri, 8, "備註", sc_style);
            SetCell(ws, ri, 9, "材料名稱", sc_style, 1, 2);
            SetCell(ws, ri, 11, "單位", sc_style);
            SetCell(ws, ri, 12, "契約數量", sc_style, 1, 2);
            SetCell(ws, ri, 14, "本日使用數量", sc_style, 1, 2);
            SetCell(ws, ri, 16, "累計使用數量", sc_style, 1, 2);
            SetCell(ws, ri, 18, "備註", sc_style);

            var res_4 = eqrLi.Where(x => x.RESOURCE_CLASS == 4).ToList();
            for (var i = 0; i < res_4.Count; i += 2)
            {
                var h = 1;
                ri++;
                var d = res_4[i];
                if (h < (d.NAME ?? "").Length) h = (d.NAME ?? "").Length;
                SetCell(ws, ri, 0, d.NAME, sc_style, 1, 2);
                SetCell(ws, ri, 2, d.UNIT, sc_style);
                SetCell(ws, ri, 3, "", sr_style, 1, 2);
                SetCell(ws, ri, 5, d.TODAY_QTY, sr_style, 1, 2);
                SetCell(ws, ri, 7, d.SUM_QTY, sr_style, 1, 2);
                SetCell(ws, ri, 8, "", sc_style);

                if (i + 1 < res_4.Count)
                {
                    d = res_4[i + 1];
                    if (h < (d.NAME ?? "").Length) h = (d.NAME ?? "").Length;
                    SetCell(ws, ri, 9, d.NAME, sc_style, 1, 2);
                    SetCell(ws, ri, 11, d.UNIT, sc_style);
                    SetCell(ws, ri, 12, "", sr_style, 1, 2);
                    SetCell(ws, ri, 14, d.TODAY_QTY, sr_style, 1, 2);
                    SetCell(ws, ri, 16, d.SUM_QTY, sr_style, 1, 2);
                    SetCell(ws, ri, 18, "", sc_style);
                }
                else
                {
                    SetCell(ws, ri, 9, "", sc_style, 1, 2);
                    SetCell(ws, ri, 11, "", sc_style);
                    SetCell(ws, ri, 12, "", sr_style, 1, 2);
                    SetCell(ws, ri, 14, "", sr_style, 1, 2);
                    SetCell(ws, ri, 16, "", sr_style, 1, 2);
                    SetCell(ws, ri, 18, "", sc_style);
                }
                var r = ws.GetRow(ri);
                r.HeightInPoints = (h / 11 + 1) * 16.25f;
            }

            ri++;
            SetCell(ws, ri, 0, "工地人員及機具管理（含約定支出工人數及機具使用情形及數量）：", l_style, 1, 19);

            ri++;
            SetCell(ws, ri, 0, "工別", c_style, 1, 4);
            SetCell(ws, ri, 4, "本日人數", c_style, 1, 2);
            SetCell(ws, ri, 6, "累計人數", c_style, 1, 3);
            SetCell(ws, ri, 9, "機具名稱", c_style, 1, 4);
            SetCell(ws, ri, 13, "本日使用數量", c_style, 1, 2);
            SetCell(ws, ri, 15, "累計使用數量", c_style, 1, 4);

            var max = 0;
            var res_2 = eqrLi.Where(x => x.RESOURCE_CLASS == 2).ToList();
            var res_35 = eqrLi.Where(x => x.RESOURCE_CLASS == 3 || x.RESOURCE_CLASS == 5).ToList();
            if (res_2.Count > res_35.Count) max = res_2.Count;
            else if (res_35.Count > res_2.Count) max = res_35.Count;
            else max = res_2.Count;

            if (max > 0)
            {
                for (var i = 0; i < max; i++)
                {
                    ri++;
                    if (i < res_2.Count)
                    {
                        var d = res_2[i];
                        SetCell(ws, ri, 0, d.NAME, c_style, 1, 4);
                        SetCell(ws, ri, 4, d.TODAY_QTY, c_style, 1, 2);
                        SetCell(ws, ri, 6, d.SUM_QTY, c_style, 1, 3);
                    }
                    else
                    {
                        SetCell(ws, ri, 0, "", c_style, 1, 4);
                        SetCell(ws, ri, 4, "", c_style, 1, 2);
                        SetCell(ws, ri, 6, "", c_style, 1, 3);
                    }
                    if (i < res_35.Count)
                    {
                        var d = res_35[i];
                        SetCell(ws, ri, 9, d.NAME, c_style, 1, 4);
                        SetCell(ws, ri, 13, d.TODAY_QTY, c_style, 1, 2);
                        SetCell(ws, ri, 15, d.SUM_QTY, c_style, 1, 4);
                    }
                    else
                    {
                        SetCell(ws, ri, 9, "", c_style, 1, 4);
                        SetCell(ws, ri, 13, "", c_style, 1, 2);
                        SetCell(ws, ri, 15, "", c_style, 1, 4);
                    }
                }
            }
            else
            {
                ri++;
                SetCell(ws, ri, 0, "-", c_style, 1, 4);
                SetCell(ws, ri, 4, "-", c_style, 1, 2);
                SetCell(ws, ri, 6, "-", c_style, 1, 3);
                SetCell(ws, ri, 9, "-", c_style, 1, 4);
                SetCell(ws, ri, 13, "-", c_style, 1, 2);
                SetCell(ws, ri, 15, "-", c_style, 1, 4);
            }

            ri++;
            var c = SetCell(ws, ri, 0, "四、本日施工項目是否需依「營造業專業工程特定施工項目應置之技術士總類、比率或人數標準表」規定應設置技術士\r\n" +
                "　　之專業工程：□有 □無 （此項如勾選＂有＂，則應填寫後附「公共工程施工日誌之技術士簽章表」）", l_style, 1, 19);
            c.Row.HeightInPoints = 33.25f;

            ri++;
            SetCell(ws, ri, 0, "五、工地職業安全衛生事項之督導、公共環境與安全之維護及其他工地行政事務：", l_style, 1, 19);

            ri++;
            var str5 = "（一）施工前檢查事項：\r\n" +
                "１、實施勤前教育（含工地預防災變及危害告知）：□有 □無\r\n" +
                "２、確認新進勞工是否提報勞工保險（或其他商業保險）資料及安全衛生教育訓練紀錄：□有 □無 □無新進勞工\r\n" +
                "３、檢查勞工個人防護具：□有 □無\r\n" +
                $"（二）其他事項：自主檢查工地實際進用移工與核准名冊相符：□有 □無 □無外籍移工\r\n{note.NOTES}";
            c = SetCell(ws, ri, 0, str5, l_style, 1, 19);
            c.Row.HeightInPoints = str5.Split('\n').Length * 16.25f;

            ri++;
            var str6 = $"六、施工取樣試驗紀錄：\r\n{note.NOTES_A}";
            c = SetCell(ws, ri, 0, str6, l_style, 1, 19);
            c.Row.HeightInPoints = str6.Split('\n').Length * 16.25f;

            ri++;
            var str7 = $"七、通知協力廠商辦理事項：\r\n{note.NOTES_B}";
            c = SetCell(ws, ri, 0, str7, l_style, 1, 19);
            c.Row.HeightInPoints = str7.Split('\n').Length * 16.25f;

            ri++;
            SetCell(ws, ri, 0, "八、其他重要事項紀錄：", l_style, 1, 19);

            ri++;
            c = SetCell(ws, ri, 0, note.NOTES_C, sl_style, 1, 19);
            c.Row.HeightInPoints = (note.NOTES_C ?? "").Split('\n').Length * 14.5f;

            ri++;
            var strOth = $"【其他重要紀錄】\r\n{note.NOTES_D}";
            c = SetCell(ws, ri, 0, strOth, tl_style, 1, 19);
            c.Row.HeightInPoints = strOth.Split('\n').Length * 15.75f;

            ri++;
            c = SetCell(ws, ri, 0, "簽章：【工地主任】（註3）：", tl_style, 1, 19);
            c.Row.HeightInPoints = 135.75f;

            ri++;
            c = SetCell(ws, ri, 0, "註：\r\n" +
                "1. 依營造業法第32條第1項第2款規定，工地主任應按日填報施工誌。\r\n" +
                "2. 本施工日誌格式僅供參考，惟原則應包含上開欄位，各機關亦得依工程性質及契約約定事項自行增訂之。\r\n" +
                "3. 本工程依營造業法第30條規定須置工地主任者，由工地主任簽章；依上開規定免置工地主任者，則由營造業法第32條第2項所定之人員簽章。廠商非屬營造業者，由工地負責人簽章。\r\n" +
                "4. 契約工期如有修正，應填修正後之契約工期，含展延工期及不計工期天數；如有依契約變更設計，預定進度及實際進度應填變更設計後計算之進度。\r\n" +
                "5. 上開重要事項記錄包含（１）主辦機關及監造單位指示（２）工地遇緊急異常狀況之通報處理情形（３）本日是否由專任工程人員督察按圖施工、解決施工技術問題等。\r\n" +
                "6. 上開施工前檢查事項所列工作應由職業安全衛生管理辦法第3條規定所置職業安全衛生人員於每日施工前辦理（檢查紀錄參考範例如附工地職業安全衛生施工前檢查表），工地主任負責督導及確認事項完成後於施工日誌填載。\r\n" +
                "7. 公共工程屬建築物者，請依內政部最新訂頒之「建築物施工日誌」填寫。\r\n" +
                "8. 本表單係公共工程委員會頒訂格式，使用時請再確認使用最新版本。", ssl_style, 1, 19);
            c.Row.HeightInPoints = 103.5f;

            for (var w = 0; w < col_width.Length; w++)
            {
                ws.SetColumnWidth(w, col_width[w]);
            }

            using var ms = new MemoryStream();
            wb.Write(ms);
            return File(ms.ToArray(), FileMime.Excel);
        }

        class ProjectInfoModel
        {
            public string? ORD_NO { get; set; }
            public string? ORD_CH { get; set; }
            public string? DIV_CH { get; set; }
            public string? DIV { get; set; }
            public string? OWN_NO { get; set; }
            public string? ABDATE { get; set; }
            public string? BEDATE { get; set; }
            public string? OWNER_NAME { get; set; }
            public string? SUPV { get; set; }
            public string? BBDATE { get; set; }
            public decimal? CUR_NT_ADD { get; set; }
            public string? EDESC { get; set; }
            public decimal? SPREAD_DAY { get; set; }
            public string? LOCATION { get; set; }
            public decimal? V_DAY { get; set; }
            public decimal? T_DAY { get; set; }
            public decimal? S_DAY { get; set; }
            public decimal? EDR_ACTQTY_DECIMAL_PLACE { get; set; }
        }

        class EDRNotesModel
        {
            public decimal? PROJECT_ID { get; set; }
            public DateTime? CALENDAR_DATE { get; set; }
            public string? WEATHER_AM { get; set; }
            public string? WEATHER_PM { get; set; }
            public decimal? WORK_HOUR { get; set; }
            public decimal? WORKDAY { get; set; }
            public string? NOTES { get; set; }
            public string? NOTES_A { get; set; }
            public string? NOTES_B { get; set; }
            public string? NOTES_C { get; set; }
            public string? NOTES_D { get; set; }
            public string? NOTES_E { get; set; }
            public string? NOTES_F { get; set; }
            public decimal? EXP_PERCENT { get; set; }
            public decimal? ACT_PERCENT { get; set; }
            public decimal? ACT_SUM { get; set; }
            public decimal? NOCAL_DAY { get; set; }
            public decimal? EXTEND_DAY { get; set; }
        }

        class EDRTempQuantityModel
        {
            public string? UP_OWN_CODE { get; set; }
            public string? OWN_CODE { get; set; }
            public string? OWNITEM_NO { get; set; }
            public string? NAME { get; set; }
            public string? OWN_CONTROL_ITEM { get; set; }
            public decimal? UP_WORKITEM_ID { get; set; }
            public string? UNIT { get; set; }
            public string? SPECIAL_ITEM { get; set; }
            public decimal? QUANTITY { get; set; }
            public decimal? OWN_UNIT_PRICE { get; set; }
            public decimal? NOW_QUANTITY { get; set; }
            public decimal? PRV_QUANTITY { get; set; }
            public decimal? SUM_QUANTITY { get; set; }
            public decimal? QTY_DECIMAL_PLACE { get; set; }
            public string? BUILD_SUB { get; set; }
        }

        class EDRListCountModel
        {
            public string? PBG_CODE { get; set; }
            public decimal? RESOURCE_CLASS { get; set; }
            public string? NAME { get; set; }
            public string? UNIT { get; set; }
            public decimal? TODAY_QTY { get; set; }
            public decimal? SUM_QTY { get; set; }
        }
    }
}
