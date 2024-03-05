using Dapper;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Text.RegularExpressions;

namespace EDR_Report.Controllers
{
    public partial class ReportController
    {
        /// <summary>
        /// B0C 欣台工務所
        /// <para>Project ID: 5400</para>
        /// <para>工令: 1CB004</para>
        /// <para>業主：中油</para>
        /// </summary>
        /// <param name="cd"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="Exception"></exception>
        IActionResult RPT1CB004_2(string? cd)
        {
            // 這邊組合套表的檔案路徑，所有的套表請都放在 Report 這個資料夾裡面
            string filePath = $"{BasePath.TrimEnd('\\').TrimEnd('/')}\\rpt1cb004_2.xlsx";
            if (!System.IO.File.Exists(filePath)) filePath = "report/rpt1cb004_2.xlsx";
            if (!System.IO.File.Exists(filePath)) return NotFound("找不到套表檔");

            // 從資料庫取得工令資料
            var db = new DBFunc();

            // 報表參數
            //var projectName = "烏溪鳥嘴潭人工湖工程計畫-湖區工程";
            var projectId = UserInfo.ProjectID!.Value;
            var myuserId = UserInfo.UserID!.Value;
            var yy = int.Parse(cd[..4]);
            var mm = int.Parse(cd[4..6]);
            var dd = int.Parse(cd[6..8]);
            var calendarDate = new DateTime(yy, mm, dd);

            //var projectName = "通霄電廠第二期更新改建計畫";
            //var calendarDate = new DateTime(2023, 12, 28);


            // 拉取SQL資料
            static Dictionary<string, object[]> ListToDictionary<T>(List<T> list)
            {
                DataTable dataTable = new DataTable();
                typeof(T)
                    .GetProperties()
                    .ToList()
                    .ForEach(property => dataTable.Columns.Add(property.Name));

                list.ForEach(info =>
                {
                    var row = dataTable.NewRow();
                    typeof(T)
                        .GetProperties()
                        .ToList()
                        .ForEach(property => row[property.Name] = property.GetValue(info));
                    dataTable.Rows.Add(row);
                });

                Dictionary<string, object[]> dict = new Dictionary<string, object[]>();

                foreach (DataColumn column in dataTable.Columns)
                {
                    string columnName = column.ColumnName;
                    object[] columnData = new object[dataTable.Rows.Count];
                    for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                    {
                        columnData[rowIndex] = dataTable.Rows[rowIndex][columnName];
                    }
                    dict[columnName.ToUpper()] = columnData;
                }
                return dict;
            }

            Dictionary<string, object[]> GetProjInfo(int projectId, DateTime calendarDate)
            {
                //var projectNameMod = "%" + projectName + "%";
                var calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
                var list = db.query<B0C_PROJECT_INFO>("edr", @"
                SELECT
                      PROJ.PROJECT_NAME_SHORT AS PROJECT_NAME
                    , PROJ.PROJECT_ID
                    , PROJ.ORGANIZATION_ID 
                    , PROJ.BES_ORD_NO
                    , TO_CHAR(PROJ.START_DATE, 'yyy""年""mm""月""dd""日""(fmDay)', 'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS START_DATE
                    , TO_CHAR(PROJ.END_DATE, 'yyy""年""mm""月""dd""日""(fmDay)', 'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS END_DATE
                    , PROJ.LAST_UPDATE_DATE
                    , TO_CHAR(NOTE.CALENDAR_DATE, 'yyy""年""mm""月""dd""日""(fmDay)', 'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS CALENDAR_DATE
                    , NOTE.WEATHER_AM     -- 天氣(上)
                    , NOTE.WEATHER_PM     -- 天氣(下)
                    , NOTE.WORK_HOUR      --  工作時數
                    , NOTE.WORKDAY        -- 工作天(Y/N)
                    , TO_CHAR(NOTE.EXP_PERCENT * 1) || '%'  AS EXP_PERCENT --預定進度(%) (至本日累計預定進度)
                    , NOTE.ACT_PERCENT    --本日實際進度
                    , TO_CHAR(NOTE.ACT_SUM * 1) || '%'  AS ACT_SUM --實際進度(%) (至本日累計實際進度)
                    , NOTE.NOCAL_DAY      --免計工期(天)
                    , NOTE.EXTEND_DAY     --展延工期(天)
                    , NVL(PROJ.SPREAD_DAY, 0) AS SPREAD_DAY --  展延天數
                    , NVL(PROJ.PROJ_SPREAD_DATE - TO_DATE(:calendarDateStr, 'yyyy/mm/dd'), 0) AS V_DAY  --剩餘工期
                    , NVL(PROJ.ORIGINAL_DAY, 0) AS T_DAY --核定工期
                    , TO_DATE(:calendarDateStr, 'yyyy/mm/dd') - PROJ.START_DATE + 1 AS S_DAY --累計工期
                    , NVL(NOTE.M50, 'none') AS M50
                    , NVL(NOTE.M51, 'none') AS M51
                    , NVL(NOTE.M52, 'none') AS M52
                    , NVL(NOTE.M53, 'none') AS M53
                    , NVL(NOTE.M54, 'none') AS M54
                FROM VSUSER.VS_PROJECTS PROJ
                INNER JOIN VSUSER.BES_EDR_NOTES NOTE
                ON PROJ.PROJECT_ID = NOTE.PROJECT_ID
                --WHERE PROJ.PROJECT_NAME LIKE :projectNameMod
                WHERE PROJ.PROJECT_ID = :projectId
                AND NOTE.CALENDAR_DATE = TO_DATE(:calendarDateStr, 'yyyy/mm/dd') --and CALENDAR_DATE = 列印日期
                ", new
                {
                    //projectNameMod = projectNameMod,
                    projectId = projectId,
                    calendarDateStr = calendarDateStr
                }).ToList();

                return ListToDictionary(list);

            }

            Dictionary<string, object[]> GetProjConsOverview(int projectId, DateTime calendarDate, int myuserId, int myState, bool specific = false)
            {
                //var projectNameMod = "%" + projectName + "%";
                var calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
                //string query1 = @"SELECT PROJECT_ID FROM VSUSER.VS_PROJECTS WHERE PROJECT_NAME LIKE :projectNameMod";
                //var projInfoList = db.query<PROJECT_INFO>("edr", query1, new
                //{
                //    projectNameMod,
                //}).ToList();
                //var projectId = projInfoList.First().PROJECT_ID;
                var query2 = @"vsuser.bes_edr_rep1";
                var dp = new DynamicParameters();
                dp.Add("myproject_id", projectId);
                dp.Add("myuser_id", myuserId);
                dp.Add("MYDATE", calendarDate);
                dp.Add("MYSTATE", myState);
                db.sp("edr", query2, dp);

                string query3_;
                if (specific)
                {
                    query3_ = @"AND SPECIAL_ITEM = 'Y'";
                }
                else
                {
                    query3_ = "";
                }

                string query3 = $@"
                SELECT 
                      OWNITEM_NO
                    , NAME
                    , NVL(UNIT, ' ') UNIT
                    , NVL(QUANTITY, 0) QUANTITY
                    , NVL(NOW_EDR_QUANTITY, 0) NOW_EDR_QUANTITY
                    , NVL(SUM_EDR_QUANTITY, 0) SUM_EDR_QUANTITY
                FROM VSUSER.BES_EDR_TEMP_QUANTITY
                WHERE PROJECT_ID = {projectId}
                AND CREATED_BY = {myuserId}
                {query3_}
            ";
                var list = db.query<B0C_PROJECT_CONSTRUCTION_OVERVIEW>("edr", query3).ToList();

                return ListToDictionary(list);
            }

            Dictionary<string, object[]> GetProjMaterial(int projectId, DateTime calendarDate)
            {
                //var projectNameMod = "%" + projectName + "%";
                var calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
                //string query1 = @"SELECT PROJECT_ID FROM VSUSER.VS_PROJECTS WHERE PROJECT_NAME LIKE :projectNameMod";
                //var projInfoList = db.query<PROJECT_INFO>("edr", query1, new
                //{
                //    projectNameMod,
                //}).ToList();
                //var projectId = projInfoList.First().PROJECT_ID;

                var list = db.query<B0C_PROJECT_MATERIAL>("edr", @"
                SELECT 
                      own_res_code AS pbg_code
                    , ROW_NUMBER() OVER (ORDER BY MAX(name)) AS row_number
                    , MAX(name) AS name
                    , MAX(unit) AS unit
                    , SUM(CASE WHEN data_date = TO_DATE(:calendarDateStr, 'yyyy/MM/dd') THEN NVL(quantity, 0) ELSE 0 END) AS today_qty
                    , SUM(CASE WHEN data_date <= TO_DATE(:calendarDateStr, 'yyyy/MM/dd') THEN NVL(quantity, 0) ELSE 0 END) AS sum_qty
                FROM vsuser.bes_edr_resqty_v
                WHERE project_id = :projectId
                    AND quantity >= 0
                    AND resource_class = 4
                GROUP BY own_res_code
                ORDER BY name
                ", new
                {
                    projectId = projectId,
                    calendarDateStr = calendarDateStr
                }).ToList();

                return ListToDictionary(list);

            }

            Dictionary<string, object[]> GetProjPeopleOrMachine(int projectId, DateTime calendarDate, string resource)
            {
                //var projectNameMod = "%" + projectName + "%";
                var calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
                //string query1 = @"SELECT PROJECT_ID FROM VSUSER.VS_PROJECTS WHERE PROJECT_NAME LIKE :projectNameMod";
                //var projInfoList = db.query<PROJECT_INFO>("edr", query1, new
                //{
                //    projectNameMod,
                //}).ToList();
                //var projectId = projInfoList.First().PROJECT_ID;
                var resourceClass = "";
                if (resource == "people")
                {
                    resourceClass = "('2')";
                }
                else if (resource == "machine")
                {
                    resourceClass = "('3', '5')";
                }
                var list = db.query<B0C_PROJECT_PEOPLE>("edr", $@"
                SELECT 
                    s.own_res_code AS pbg_code,
                    s.name,
                    NVL(e.today_qty, 0) AS today_qty,
                    NVL(s.sum_qty, 0) AS sum_qty
                FROM (
                    SELECT 
                        own_res_code, 
                        SUM(NVL(quantity, 0)) AS today_qty
                    FROM vsuser.bes_edr_resqty_v
                    WHERE 
                        project_id = :projectId
                        AND data_date = TO_DATE(:calendarDateStr, 'yyyy/mm/dd')
                        AND quantity >= 0
                        AND resource_class IN {resourceClass}
                    GROUP BY own_res_code
                ) e
                FULL OUTER JOIN (
                    SELECT 
                        own_res_code, 
                        name, 
                        NVL(SUM(quantity), 0) AS sum_qty
                    FROM vsuser.bes_edr_resqty_v
                    WHERE 
                        project_id = :projectId
                        AND data_date <= TO_DATE(:calendarDateStr, 'yyyy/mm/dd')
                        AND quantity >= 0
                        AND resource_class IN {resourceClass}
                    GROUP BY own_res_code, name
                ) s ON e.own_res_code = s.own_res_code
                ", new
                {
                    projectId = projectId,
                    calendarDateStr = calendarDateStr
                }).ToList();

                return ListToDictionary(list);

            }

            Dictionary<string, object[]> GetProjNote(int projectId, DateTime calendarDate)
            {
                //var projectNameMod = "%" + projectName + "%";
                var calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
                //string query1 = @"SELECT PROJECT_ID FROM VSUSER.VS_PROJECTS WHERE PROJECT_NAME LIKE :projectNameMod";
                //var projInfoList = db.query<PROJECT_INFO>("edr", query1, new
                //{
                //    projectNameMod,
                //}).ToList();
                //var projectId = projInfoList.First().PROJECT_ID;

                var list = db.query<B0C_PROJECT_NOTE>("edr", $@"
                SELECT 
                     NVL(notes_a, ' ') AS NOTES_A        --工地記事 (六、施工取樣試驗紀錄：)
                   , NVL(notes_b, ' ') AS NOTES_B       --工地記事 (七、通知協力廠商辦理事項：)
                   , NVL(notes_c, ' ') AS NOTES_C        --工地記事 (八、重要事項記錄：)
                   , NVL(notes, ' ') AS NOTES 
                FROM vsuser.BES_EDR_NOTES
                WHERE project_id = :projectId
                AND CALENDAR_DATE = to_date(:calendarDateStr, 'yyyy/MM/dd') --and CALENDAR_DATE = 列印日期
                ", new
                {
                    projectId = projectId,
                    calendarDateStr = calendarDateStr
                }).ToList();

                return ListToDictionary(list);

            }

            // 準備資料
            var projInfo = GetProjInfo(projectId, calendarDate);
            var projConsOverview = GetProjConsOverview(projectId, calendarDate, myuserId, 2);
            var projConsOverviewSpec = GetProjConsOverview(projectId, calendarDate, myuserId, 2, true);
            var projMaterial = GetProjMaterial(projectId, calendarDate);
            var projPeople = GetProjPeopleOrMachine(projectId, calendarDate, "people");
            var projMachine = GetProjPeopleOrMachine(projectId, calendarDate, "machine");
            var projNote = GetProjNote(projectId, calendarDate);

            // 待確認資料來源
            projInfo["CONSTRUCTOR"] = new object[] { "中華工程股份有限公司" };
            projInfo["REPORT_SERIAL"] = new object[] { "" };

            //放入對應資料
            string chk_true = "▓有 □無";
            string chk_false = "□有 ▓無";
            string chk_none = "□有 □無";

            string ch4_chk_value = projInfo["M50"][0].ToString() ?? "none";
            string ch4_chk_string = "    ";
            switch (ch4_chk_value)
            {
                case "Y":
                    ch4_chk_string = ch4_chk_string + chk_true;
                    break;
                case "N":
                    ch4_chk_string = ch4_chk_string + chk_false;
                    break;
                default:
                    ch4_chk_string = ch4_chk_string + chk_none;
                    break;
            }
            object[] ch4_chk = { ch4_chk_string };

            string ch5_chk1_value = projInfo["M51"][0].ToString() ?? "none";
            string ch5_chk1_string = "";
            switch (ch5_chk1_value)
            {
                case "Y":
                    ch5_chk1_string = chk_true;
                    break;
                case "N":
                    ch5_chk1_string = chk_false;
                    break;
                default:
                    ch5_chk1_string = chk_none;
                    break;
            }
            ch5_chk1_string = "    1.實施勤前教育(含工地預防災變及危害告知)：" + ch5_chk1_string;
            object[] ch5_chk1 = { ch5_chk1_string };

            string ch5_chk2_value = projInfo["M52"][0].ToString() ?? "none";
            string ch5_chk2_string = "    2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：";
            switch (ch5_chk2_value)
            {
                case "1":
                    ch5_chk2_string = ch5_chk2_string + chk_true + " □無新進勞工";
                    break;
                case "2":
                    ch5_chk2_string = ch5_chk2_string + chk_false + " □無新進勞工";
                    break;
                case "3":
                    ch5_chk2_string = ch5_chk2_string + chk_none + " ▓無新進勞工";
                    break;
                default:
                    ch5_chk2_string = ch5_chk2_string + chk_none + " □無新進勞工";
                    break;
            }
            object[] ch5_chk2 = { ch5_chk2_string };

            string ch5_chk3_value = projInfo["M53"][0].ToString() ?? "none";
            string ch5_chk3_string = "";
            switch (ch5_chk3_value)
            {
                case "Y":
                    ch5_chk3_string = chk_true;
                    break;
                case "N":
                    ch5_chk3_string = chk_false;
                    break;
                default:
                    ch5_chk3_string = chk_none;
                    break;
            }
            ch5_chk3_string = "    3.檢查勞工個人防護具：" + ch5_chk3_string;
            object[] ch5_chk3 = { ch5_chk3_string };

            string ch5_chk4_value = projInfo["M54"][0].ToString() ?? "none";
            string ch5_chk4_string = "    自主檢查工地實際進用移工與核准名冊相符：";
            switch (ch5_chk4_value)
            {
                case "1":
                    ch5_chk4_string = ch5_chk4_string + chk_true + " □無外籍移工";
                    break;
                case "2":
                    ch5_chk4_string = ch5_chk4_string + chk_false + " □無外籍移工";
                    break;
                case "3":
                    ch5_chk4_string = ch5_chk4_string + chk_none + " ▓無外籍移工";
                    break;
                default:
                    ch5_chk4_string = ch5_chk4_string + chk_none + " □無外籍移工";
                    break;
            }
            object[] ch5_chk4 = { ch5_chk4_string };

            // 建立資料字典檔
            Dictionary<string, object[]> variables = new Dictionary<string, object[]>
            {
                { "$report_serial$", projInfo["REPORT_SERIAL"]},
                { "$weather_am$", projInfo["WEATHER_AM"] },
                { "$weather_pm$", projInfo["WEATHER_PM"] },
                { "$calendar_date$", projInfo["CALENDAR_DATE"] },
                { "$project_name$", projInfo["PROJECT_NAME"] },
                { "$contractor$", projInfo["CONSTRUCTOR"]  },
                { "$t_day$", projInfo["T_DAY"] },
                { "$s_day$", projInfo["S_DAY"] },
                { "$v_day$", projInfo["V_DAY"] },
                { "$spread_day$", projInfo["SPREAD_DAY"] },
                { "$state_date$", projInfo["START_DATE"] },
                { "$end_date$", projInfo["END_DATE"] },
                { "$exp_percent$", projInfo["EXP_PERCENT"] },
                { "$act_sum$", projInfo["ACT_SUM"] },

                { "$ch1_col1$", projConsOverview["OWNITEM_NO"]},
                { "$ch1_col2$", projConsOverview["NAME"]},
                { "$ch1_col3$", projConsOverview["UNIT"]},
                { "$ch1_col4$", projConsOverview["QUANTITY"]},
                { "$ch1_col5$", projConsOverview["NOW_EDR_QUANTITY"]},
                { "$ch1_col6$", projConsOverview["SUM_EDR_QUANTITY"]},
                //{ "$ch1_col7$", projConsOverview["REMARK"]},
                { "$ch1ex_col1$", projConsOverviewSpec["OWNITEM_NO"]},
                { "$ch1ex_col2$", projConsOverviewSpec["NAME"]},
                { "$ch1ex_col3$", projConsOverviewSpec["UNIT"]},
                { "$ch1ex_col4$", projConsOverviewSpec["QUANTITY"]},
                { "$ch1ex_col5$", projConsOverviewSpec["NOW_EDR_QUANTITY"]},
                { "$ch1ex_col6$", projConsOverviewSpec["SUM_EDR_QUANTITY"]},
                //{ "$ch1ex_col7$", projConsOverviewSpec["REMARK"]},
                { "$ch2_col1$", projMaterial["ROW_NUMBER"]},
                { "$ch2_col2$", projMaterial["NAME"]},
                { "$ch2_col3$", projMaterial["UNIT"]},
                //{ "$ch2_col4$", projMaterial["QUANTITY"]},
                { "$ch2_col5$", projMaterial["TODAY_QTY"]},
                { "$ch2_col6$", projMaterial["SUM_QTY"]},
                //{ "$ch2_col7$", projMaterial["REMARK"]},
                { "$ch3_col1$", projPeople["NAME"]},
                { "$ch3_col2$", projPeople["TODAY_QTY"]},
                { "$ch3_col3$", projPeople["SUM_QTY"]},
                { "$ch3_col4$", projMachine["NAME"]},
                { "$ch3_col5$", projMachine["TODAY_QTY"]},
                { "$ch3_col6$", projMachine["SUM_QTY"]},
                { "$ch4_chk$", ch4_chk },
                { "$ch5_chk1$", ch5_chk1 },
                { "$ch5_chk2$", ch5_chk2 },
                { "$ch5_chk3$", ch5_chk3 },
                { "$ch5_chk4$", ch5_chk4 },
                { "$ch5$", projNote["NOTES"]},
                { "$ch6$", projNote["NOTES_A"]},
                { "$ch7$", projNote["NOTES_B"] },
                { "$ch8$", projNote["NOTES_C"] }
            };

            IWorkbook wb;
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fileStream);
                ISheet ws = wb.GetSheetAt(0);

                void WriteCellValueByPlaceholder(ISheet sheet, string placeholder, object cellValue)
                {
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                        {
                            ICell cell = row.GetCell(colIndex);
                            if (cell != null && cell.CellType == CellType.String && cell.StringCellValue == placeholder)
                            {
                                if (cellValue is int intValue)
                                {
                                    cell.SetCellValue(intValue);
                                }
                                else if (cellValue is short shortValue)
                                {
                                    cell.SetCellValue(shortValue);
                                }
                                else if (cellValue is long longValue)
                                {
                                    cell.SetCellValue(longValue);
                                }
                                else if (cellValue is float floatValue)
                                {
                                    cell.SetCellValue(floatValue);
                                }
                                else if (cellValue is double doubleValue)
                                {
                                    cell.SetCellValue(doubleValue);
                                }
                                else if (cellValue is decimal decimalValue)
                                {
                                    cell.SetCellValue(Convert.ToDouble(decimalValue));
                                }
                                else if (cellValue is string stringValue)
                                {
                                    cell.SetCellValue(stringValue);
                                }
                            }
                        }
                    }
                }

                void WriteCellValueByIndex(ISheet sheet, int rowIndex, int colIndex, object cellValue)
                {
                    rowIndex--;
                    colIndex--;
                    IRow row = sheet.GetRow(rowIndex);
                    ICell cell = row.GetCell(colIndex);
                    if (cellValue is int intValue)
                    {
                        cell.SetCellValue(intValue);
                    }
                    else if (cellValue is short shortValue)
                    {
                        cell.SetCellValue(shortValue);
                    }
                    else if (cellValue is long longValue)
                    {
                        cell.SetCellValue(longValue);
                    }
                    else if (cellValue is float floatValue)
                    {
                        cell.SetCellValue(floatValue);
                    }
                    else if (cellValue is double doubleValue)
                    {
                        cell.SetCellValue(doubleValue);
                    }
                    else if (cellValue is decimal decimalValue)
                    {
                        cell.SetCellValue(Convert.ToDouble(decimalValue));
                    }
                    else if (cellValue is string stringValue)
                    {
                        cell.SetCellValue(stringValue);
                    }
                }

                int[,] GetCellIndexByPlaceholder(ISheet sheet, string placeholder)
                {
                    List<int[]> indexesList = new List<int[]>();
                    for (int rowIndex = 0; rowIndex < sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                        {
                            ICell cell = row.GetCell(colIndex);
                            if (cell != null && cell.CellType == CellType.String && cell.StringCellValue == placeholder)
                            {
                                int[] indexes = new int[] { rowIndex + 1, colIndex + 1 };
                                indexesList.Add(indexes);
                            }
                        }
                    }
                    int[,] indexesArray = new int[indexesList.Count, 2];
                    for (int i = 0; i < indexesList.Count; i++)
                    {
                        indexesArray[i, 0] = indexesList[i][0];
                        indexesArray[i, 1] = indexesList[i][1];
                    }
                    return indexesArray;
                }

                void RemoveRangeRows(ISheet sheet, int rowIndexStart, int rowIndexEnd)
                {
                    rowIndexStart--;
                    rowIndexEnd--;
                    int rowCount = rowIndexEnd - rowIndexStart + 1;
                    for (int i = 0; i < rowCount; i++)
                    {
                        IRow row = ws.GetRow(rowIndexStart);
                        if (row != null)
                        {
                            sheet.RemoveRow(row);
                            sheet.ShiftRows(rowIndexStart + 1, sheet.LastRowNum, -1);
                        }
                    }
                }

                void CopyRowDown(ISheet sheet, int rowIndexStart, int repeatTimes)
                {
                    rowIndexStart--; // 將 rowIndex 轉換為以 0 為基底的索引

                    for (int i = 0; i < repeatTimes; i++)
                    {
                        IRow sourceRow = sheet.GetRow(rowIndexStart);
                        if (sourceRow != null)
                        {
                            int lastRowNum = sheet.LastRowNum;
                            sheet.ShiftRows(rowIndexStart + 1, lastRowNum, 1, true, false);
                            IRow newRow = sheet.CreateRow(rowIndexStart + 1);
                            newRow.Height = sourceRow.Height;

                            // 复制合并单元格信息
                            for (int m = 0; m < sheet.NumMergedRegions; m++)
                            {
                                var mergedRegion = sheet.GetMergedRegion(m);
                                if (mergedRegion.FirstRow == rowIndexStart && mergedRegion.LastRow == rowIndexStart)
                                {
                                    // 合并单元格在源行，需要在新行中创建合并单元格
                                    int newFirstRow = rowIndexStart + 1;
                                    int newLastRow = rowIndexStart + 1 + (mergedRegion.LastRow - mergedRegion.FirstRow);
                                    int newFirstCol = mergedRegion.FirstColumn;
                                    int newLastCol = mergedRegion.LastColumn;
                                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(newFirstRow, newLastRow, newFirstCol, newLastCol));
                                }
                            }

                            for (int j = 0; j < sourceRow.LastCellNum; j++)
                            {
                                ICell sourceCell = sourceRow.GetCell(j);
                                if (sourceCell != null)
                                {
                                    ICell newCell = newRow.CreateCell(j);
                                    newCell.CellStyle = sourceCell.CellStyle;
                                    switch (sourceCell.CellType)
                                    {
                                        case CellType.Blank:
                                            newCell.SetCellType(CellType.Blank);
                                            break;
                                        case CellType.Boolean:
                                            newCell.SetCellValue(sourceCell.BooleanCellValue);
                                            break;
                                        case CellType.Numeric:
                                            newCell.SetCellValue(sourceCell.NumericCellValue);
                                            break;
                                        case CellType.String:
                                            newCell.SetCellValue(sourceCell.StringCellValue);
                                            break;
                                    }
                                }
                            }

                            rowIndexStart++; // 更新 rowIndex 以适应下一次的复制
                        }
                    }
                }

                void MergeCells(ISheet sheet, int startRow, int startColumn, int endRow, int endColumn)
                {
                    if (sheet == null)
                    {
                        throw new ArgumentNullException(nameof(sheet), "Sheet cannot be null");
                    }

                    if (startRow < 0 || startColumn < 0 || endRow < 0 || endColumn < 0)
                    {
                        throw new ArgumentException("Row and column indexes cannot be negative");
                    }
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRow - 1, endRow - 1, startColumn - 1, endColumn - 1));
                }

                string[] FindPlaceholderAppearingTimes(ISheet sheet, int times)
                {
                    Dictionary<string, int> textCounts = new Dictionary<string, int>();
                    if (sheet == null)
                    {
                        throw new ArgumentNullException(nameof(sheet), "Sheet cannot be null");
                    }
                    for (int rowIndex = 0; rowIndex < sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            foreach (ICell cell in row.Cells)
                            {
                                if (cell.CellType == CellType.String)
                                {
                                    string cellValue = cell.StringCellValue;
                                    Regex regex = new Regex(@"\$\w+\$");
                                    MatchCollection matches = regex.Matches(cellValue);
                                    foreach (System.Text.RegularExpressions.Match match in matches)
                                    {
                                        string foundText = match.Value;
                                        if (textCounts.ContainsKey(foundText))
                                        {
                                            textCounts[foundText]++;
                                        }
                                        else
                                        {
                                            textCounts[foundText] = 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (times == 1)
                    {
                        string[] result = textCounts.Where(pair => pair.Value == 1).Select(pair => pair.Key).ToArray();
                        return result;
                    }
                    else if (times == 2)
                    {
                        string[] result = textCounts.Where(pair => pair.Value > 1).Select(pair => pair.Key).ToArray();
                        return result;
                    }
                    else
                    {
                        throw new Exception("times 的值只能是 1 (單次)或 2 (多次)");
                    }
                }

                void ReplaceAllSinglePlaceholder(ISheet sheet)
                {
                    string[] uniquePlaceholder = FindPlaceholderAppearingTimes(sheet, 1);
                    foreach (string placeholder in uniquePlaceholder)
                    {
                        if (variables.ContainsKey(placeholder))
                        {
                            if (variables[placeholder].Length == 1)
                            {
                                WriteCellValueByPlaceholder(sheet, placeholder, variables[placeholder][0]);
                            }
                            else if (variables[placeholder].Length == 0)
                            {
                                WriteCellValueByPlaceholder(sheet, placeholder, "");
                            }
                        }
                    }
                }

                void ReplaceMultiplePlaceholder(ISheet sheet, string placeholder, object[] data)
                {
                    int[,] indexesArray = GetCellIndexByPlaceholder(sheet, placeholder);
                    for (int i = 0; i < data.Length; i++)
                    {
                        int rowIndex = indexesArray[i, 0];
                        int colIndex = indexesArray[i, 1];
                        if (data[i] is int intValue)
                        {
                            WriteCellValueByIndex(sheet, rowIndex, colIndex, intValue);
                        }
                        else if (data[i] is long longValue)
                        {
                            WriteCellValueByIndex(sheet, rowIndex, colIndex, longValue);
                        }
                        else if (data[i] is double doubleValue)
                        {
                            WriteCellValueByIndex(sheet, rowIndex, colIndex, doubleValue);
                        }
                        else if (data[i] is decimal decimalValue)
                        {
                            WriteCellValueByIndex(sheet, rowIndex, colIndex, Convert.ToDouble(decimalValue));
                        }
                        else if (data[i] is short shortValue)
                        {
                            WriteCellValueByIndex(sheet, rowIndex, colIndex, shortValue);
                        }
                        else if (data[i] is float floatValue)
                        {
                            WriteCellValueByIndex(sheet, rowIndex, colIndex, floatValue);
                        }
                        else if (data[i] is string stringValue)
                        {
                            WriteCellValueByIndex(sheet, rowIndex, colIndex, stringValue);
                        }
                    }
                }

                void ReplaceAllMultipleVariable(ISheet sheet)
                {
                    string[] uniqueVariables = FindPlaceholderAppearingTimes(sheet, 2);
                    foreach (string placeholder in uniqueVariables)
                    {
                        if (variables.ContainsKey(placeholder))
                        {
                            ReplaceMultiplePlaceholder(sheet, placeholder, (object[])variables[placeholder]);
                        }
                    }
                }

                void AdjustRowNums(ISheet sheet, string placeholder1, string? placeholder2 = null)
                {
                    string placeholder;
                    if (!string.IsNullOrEmpty(placeholder2))
                    {
                        if (variables[placeholder1].GetLength(0) >= variables[placeholder2].GetLength(0))
                        {
                            placeholder = placeholder1;
                        }
                        else
                        {
                            placeholder = placeholder2;
                        }
                    }
                    else
                    {
                        placeholder = placeholder1;
                    }
                    int[,] indexesArray = GetCellIndexByPlaceholder(sheet, placeholder);
                    object[] data = (object[])variables[placeholder];
                    int rowNumDiff = data.GetLength(0) - indexesArray.GetLength(0);
                    if (rowNumDiff >= 0)
                    {
                        RemoveRangeRows(sheet, indexesArray[0, 0] + 1, indexesArray[indexesArray.GetLength(0) - 1, 0]);
                        CopyRowDown(sheet, indexesArray[0, 0], indexesArray[indexesArray.GetLength(0) - 1, 0] - indexesArray[0, 0] + Math.Abs(rowNumDiff));
                    }
                    else if (rowNumDiff < 0)
                    {
                        RemoveRangeRows(sheet, indexesArray[0, 0] + 1, indexesArray[indexesArray.GetLength(0) - 1, 0]);
                        CopyRowDown(sheet, indexesArray[0, 0], indexesArray[indexesArray.GetLength(0) - 1, 0] - indexesArray[0, 0] - Math.Abs(rowNumDiff));
                    }
                }

                void ImportDataIntoPlaceholder(ISheet sheet)
                {
                    AdjustProjNameRowHeight(sheet);
                    ReplaceAllSinglePlaceholder(sheet);
                    ReplaceAllMultipleVariable(sheet);
                    ReplacePlaceholdersWithEmpty(sheet);
                    //AutoSizeRows(sheet);
                }

                void HorizontalMergeCols(ISheet sheet, string placeholder, int[,] mergeColRangeArray)
                {
                    int[,] Index = GetCellIndexByPlaceholder(sheet, placeholder);
                    foreach (int i in Enumerable.Range(1, Index.GetLength(0) - 1))
                    {
                        foreach (int j in Enumerable.Range(0, mergeColRangeArray.GetLength(0)))
                        {
                            MergeCells(sheet, Index[i, 0], mergeColRangeArray[j, 0], Index[i, 0], mergeColRangeArray[j, 1]);
                        }
                    }
                }

                void VerticalMergeRows(ISheet sheet, string placeholder, int[] mergeCol)
                {
                    int[,] Index = GetCellIndexByPlaceholder(sheet, placeholder);
                    foreach (int i in mergeCol)
                    {
                        MergeCells(sheet, Index[0, 0], i, Index[Index.GetLength(0) - 1, 0], i);
                    }
                }

                void ReplacePlaceholdersWithEmpty(ISheet sheet)
                {
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                            {
                                ICell cell = row.GetCell(colIndex);
                                if (cell != null && cell.CellType == CellType.String)
                                {
                                    string cellValue = cell.StringCellValue;
                                    if (Regex.IsMatch(cellValue, @"\$\w+\$"))
                                    {
                                        // Replace placeholder with empty string
                                        cell.SetCellValue(string.Empty);
                                    }
                                }
                            }
                        }
                    }
                }

                void AutoSizeRows(ISheet sheet)
                {
                    // 取得工作表的列數
                    int rowCount = sheet.PhysicalNumberOfRows;
                    // 遍歷每一列，調整每列的寬度以符合內容
                    for (int i = 0; i < rowCount; i++)
                    {
                        sheet.AutoSizeRow(i);
                    }
                }

                void AdjustRowHeight(ISheet sheet, string placeholder)
                {
                    var firstDataString = (string)variables[placeholder][0];
                    int count = firstDataString.Split(new string[] { "\n" }, StringSplitOptions.None).Length;
                    int[,] cellIndex = GetCellIndexByPlaceholder(sheet, placeholder);
                    IRow row = sheet.GetRow(cellIndex[0, 0] - 1);
                    row.HeightInPoints = 17 * count;
                }

                void AdjustProjNameRowHeight(ISheet sheet)
                {
                    string placeholder = "$PROJECT_NAME$";
                    int[,] cellIndex = GetCellIndexByPlaceholder(sheet, placeholder.ToLower());
                    IRow row = ws.GetRow(cellIndex[0, 0] - 1);
                    ICell cell = row.GetCell(cellIndex[0, 1] - 1);
                    ICellStyle cellStyle = wb.CreateCellStyle();
                    IFont cellFont = wb.CreateFont();
                    cellStyle.BorderTop = BorderStyle.Medium;
                    cellStyle.WrapText = true;
                    cellFont.FontName = "標楷體";
                    cellFont.FontHeightInPoints = 12;
                    cellStyle.SetFont(cellFont);

                    string projNameStr = (string)variables[placeholder.ToLower()][0];
                    int rowCount = projNameStr.Length / 21 + 1;
                    if (projNameStr.Length <= 21)
                    {
                        row.HeightInPoints = 17 * 2;
                        cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    }
                    else
                    {
                        row.HeightInPoints = 17 * rowCount;
                        cellStyle.VerticalAlignment = VerticalAlignment.Top;
                    }
                    cell.CellStyle = cellStyle;
                }

                // 依資料長度調整列數
                AdjustRowNums(ws, "$ch1_col1$");
                AdjustRowNums(ws, "$ch1ex_col1$");
                AdjustRowNums(ws, "$ch2_col1$");
                AdjustRowNums(ws, "$ch3_col1$", "$ch3_col4$");
                AdjustRowNums(ws, "$ch5$");
                AdjustRowNums(ws, "$ch6$");
                AdjustRowNums(ws, "$ch7$");
                AdjustRowNums(ws, "$ch8$");

                // 依資料字串中\n個數調整列高
                AdjustRowHeight(ws, "$ch5$");
                AdjustRowHeight(ws, "$ch6$");
                AdjustRowHeight(ws, "$ch7$");
                AdjustRowHeight(ws, "$ch8$");

                // 替換資料
                ImportDataIntoPlaceholder(ws);

            }

            // 不留實體檔案，直接串流到前端
            using var ms = new MemoryStream();
            wb.Write(ms);
            return File(ms.ToArray(), FileMime.Excel);
        }
        public class B0C_PROJECT_INFO
        {
            public string? PROJECT_NAME { get; set; }
            public string? PROJECT_ID { get; set; }
            public string? ORGANIZATION_ID { get; set; }
            public string? BES_ORD_NO { get; set; }
            public string? START_DATE { get; set; }
            public string? END_DATE { get; set; }
            public string? LAST_UPDATE_DATE { get; set; }
            public string? CALENDAR_DATE { get; set; }
            public string? WEATHER_AM { get; set; }
            public string? WEATHER_PM { get; set; }
            public string? WORK_HOUR { get; set; }
            public string? WORKDAY { get; set; }
            public string? EXP_PERCENT { get; set; }
            public string? ACT_PERCENT { get; set; }
            public string? ACT_SUM { get; set; }
            public string? NOCAL_DAY { get; set; }
            public string? EXTEND_DAY { get; set; }
            public string? SPREAD_DAY { get; set; }
            public string? V_DAY { get; set; }
            public string? T_DAY { get; set; }
            public string? S_DAY { get; set; }
            public string? M50 { get; set; }
            public string? M51 { get; set; }
            public string? M52 { get; set; }
            public string? M53 { get; set; }
            public string? M54 { get; set; }
        }
        public class B0C_PROJECT_MATERIAL
        {
            public string? PBG_CODE { get; set; }
            public string? ROW_NUMBER { get; set; }
            public string? NAME { get; set; }
            public string? UNIT { get; set; }
            public float? TODAY_QTY { get; set; }
            public float? SUM_QTY { get; set; }
        }
        public class B0C_PROJECT_PEOPLE
        {
            public string? PBG_CODE { get; set; }
            public string? NAME { get; set; }
            public float? TODAY_QTY { get; set; }
            public float? SUM_QTY { get; set; }
        }
        public class B0C_PROJECT_NOTE
        {
            public string? NOTES_A { get; set; }
            public string? NOTES_B { get; set; }
            public string? NOTES_C { get; set; }
            public string? NOTES { get; set; }
        }
        public class B0C_PROJECT_CONSTRUCTION_OVERVIEW
        {
            public string? OWNITEM_NO { get; set; }
            public string? NAME { get; set; }
            public string? UNIT { get; set; }
            public float? QUANTITY { get; set; }
            public float? NOW_EDR_QUANTITY { get; set; }
            public float? SUM_EDR_QUANTITY { get; set; }
        }
    }
}
