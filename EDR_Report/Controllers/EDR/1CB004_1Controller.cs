using EDR_Report.Commons;
using EDR_Report.Interfaces;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using System.Collections.Specialized;

namespace EDR_Report.Controllers
{
    public partial class ReportController
    {
        /// <summary>
        /// 報表模板內的自定義變數
        /// </summary>
        private class TemplateVars
        {
            /// <summary>
            /// 工程名稱
            /// </summary>
            public static string prj_ord_ch = "$prj_ord_ch$";
            public static string note_extended_days = "$note_extended_days$";
            /// <summary>
            /// 開工日期
            /// </summary>
            public static string prj_kickoff_date = "$prj_kickoff_date$";
            /// <summary>
            /// 完工日期
            /// </summary>
            public static string prj_close_date = "$prj_close_date$";
            /// <summary>
            /// 工程金額(加代辦管線) 要轉億元單位
            /// </summary>
            public static string prj_cur_nt_add = "$prj_cur_nt_add$";
            /// <summary>
            /// 工程金額(要轉億元單位
            /// </summary>
            public static string prj_cur_nt = "$prj_cur_nt$";
            /// <summary>
            /// 預定進度(累計)
            /// </summary>
            public static string note_exp_percent = "$note_exp_percent$";
            /// <summary>
            /// 實際進度(累計)
            /// </summary>
            public static string note_act_percent = "$note_act_percent$";
            /// <summary>
            /// 早上天氣
            /// </summary>
            public static string note_weather_am = "$note_weather_am$";
            /// <summary>
            /// 下午天氣
            /// </summary>
            public static string note_weather_pm = "$note_weather_pm$";
            /// <summary>
            /// 填表日期
            /// </summary>
            public static string note_rptDate = "$rptDate$";
            /// <summary>
            /// 核定工期
            /// </summary>
            public static string prj_confirmed_days = "$prj_confirmed_days$";
            /// <summary>
            /// 累計工期
            /// </summary>
            public static string prj_accumulated_days = "$note_accumulated_days$";
            /// <summary>
            /// 剩餘工期
            /// </summary>
            public static string prj_remained_days = "$note_remained_days$";
            /// <summary>
            ///  施工項目
            ///  資料為清單列表，所以該變數為excel模板內自定義標記，要新增列，再塞值到Cells
            /// </summary>
            public static string workItems = "$workItems$";
            /// <summary>
            /// 營造業專業工程特定施工項目
            /// 資料為清單列表，所以該變數為excel模板內自定義標記，要新增列，再塞值到Cells
            /// </summary>
            public static string speical_item = "$special_items$";
            /// <summary>
            /// 工地材料列表
            /// 資料為清單列表，所以該變數為excel模板內自定義標記，要新增列，再塞值到Cells
            /// </summary>
            public static string materials = "$materials$";
            /// <summary>
            /// 工地人員
            /// 資料為清單列表，所以該變數為excel模板內自定義標記，要新增列，再塞值到Cells
            /// </summary>
            public static string engineers = "$engineers$";
            /// <summary>
            /// 機具
            /// </summary>
            public static string equipments = "$equipments$";
            /// <summary>
            /// 第四大項
            /// </summary>
            public static string notes_item4 = "$notesA_item4$";
            /// <summary>
            /// 其他事項
            /// </summary>
            public static string notes_misc = " $notes_misc$";
            /// <summary>
            /// 第六大項 試驗紀錄
            /// </summary>
            public static string notes_experiment = "$notesA_item6$";
            public static string notes_item7 = "$notesB$";
            /// <summary>
            /// 第八大項 重要事項
            /// </summary>
            public static string notes_item8 = "$notesC$";
        }
        /// <summary>
        /// B0C 欣台工務所
        /// <para>Project ID: 5400</para>
        /// <para>工令: 1CB004</para>
        /// <para>業主：公路局</para>
        /// </summary>
        /// <param name="cd"></param>
        /// <returns></returns>
        IActionResult RPT1CB004_1(string cd)
        {
            _logger.Log(LogLevel.Information, "Entered RPT1CB004_1");
            try
            {
                // ordNum = "1CB0D4";
                //excel報表模板名稱
                string rptTempalte = "rpt1cb004_1.xlsx";
                //報表路徑
                string templateFullpath = $"{BasePath.TrimEnd('\\').TrimEnd('/')}\\{rptTempalte}";
                //string ErrorMessage = ""; //we declare the var that will contain the error message in case of error.
                if (!System.IO.File.Exists(templateFullpath)) templateFullpath = $"report/{rptTempalte}";
                if (!System.IO.File.Exists(templateFullpath))
                    return BadRequest($"<{templateFullpath}>報表模板不存在!");

                ///驗測excel格子內的公式，以及更多用途
                IFormulaEvaluator ObjectIFormulaEvaluator;
                ////
                ///回傳.xlsx or .xls workbook, 
                ///xlsx是微軟將所有Office 文件改成用xml 自定義metadata xml tag, 
                ///因應通用資料交換的格式, 
                ///xls是舊版二進位格式
                ///.xlsx 對應XSSFWorkbook物件
                ///.xls  對應HSSFWorkbook物件
                IWorkbook wb = _rptHelper.OpenExcelTemplate(templateFullpath, out ObjectIFormulaEvaluator);
                ISheet ws = wb.GetSheetAt(0);
                //string CellContentString = ""; //We'll use this to store the string of the cells when we'll iterate thorought them

                #region 表頭與專案進度
                /*
                +------------------------------------------------------------------------------------------------------------------------------+
                |                                                       公共工程施工日誌                                                         |
                +-------------------------------+---------------------------------------------------+------------------------------+-----------+
                | 本日天氣：上午：$weather_am$     | 下午：$weather_pm$                                | 填表日期：                     | $rptDate$ |
                +---------------+---------------+-------------------------------------------+-------+-------------+----------------+-----------+
                |    工程名稱    |                     $prj_ord_ch$                          |     承攬廠商名稱      |     中華工程股份有限公司      |
                +---------------+---------------+--------+----------+------------+----------+---------------------+----------------+-----------+
                |    核定工期    | 1108          | 日曆天  | 累計工期  | 769 日曆天  | 剩餘工期  |      339 日曆天      |  工期展延天數    |   138天   |
                +---------------+---------------+--------+----------+------------+----------+--------------+------+----------------+-----------+
                |    開工日期    |                      $prj_start_date$                     |   完工日期    |           $prj_end_date$          |
                +---------------+------------------------+----------------------------------+--------------+----------------+------------------+
                |      $prj_edr_cur_nt$億預定進度(%)       |               40.50              | $prj_edr_cur_nt$億實際進度(%)  |       39.52      |
                +----------------------------------------+----------------------------------+-------------------------------+------------------+
                |       27億(含代辦管線)預定進度(%)         |               40.05              | 27億(含代辦管線)實              |       39.17      |
                +----------------------------------------+----------------------------------+-------------------------------+------------------+
                
                 ****/


                var project_id = UserInfo.ProjectID!.Value;//TODO
                var yy = int.Parse(cd[..4]);
                var mm = int.Parse(cd[4..6]);
                var dd = int.Parse(cd[6..8]);
                var report_date = $"{yy}/{mm}/{dd}";//TODO
                var workProgress = _dbcontext.QueryWorkProgress(project_id, report_date).ToList();
                var dailyNotes = _dbcontext.QueryDailyWorkNotes(project_id, report_date).ToList();
                try
                {


                    /////正常應該能查到資料，若無資料回傳，可能是資料或查詢條件有問題
                    //if (workProgress.Count == 0)
                    //    return BadRequest($"project:{project_id}, date:{report_date} 沒有專案進度資料");
                    //if (dailyNotes.Count == 0)
                    //    return BadRequest($"project:{project_id}, date:{report_date} 沒有工地記事資料");

                    //有資料回傳，才能繼續將回傳資料，貼到excel
                    //本日天氣
                    var thedate = string.Empty;
                    _rptHelper.Replace(ws, TemplateVars.note_weather_am, dailyNotes[0].WEATHER_AM, 0, 0, 1, 1);
                    _rptHelper.Replace(ws, TemplateVars.note_weather_pm, dailyNotes[0].WEATHER_PM, 0, 0, 1, 1);
                    //報表日期
                    string rptDate = _rptHelper.TaiwanDateTime((string)dailyNotes[0].CALENDAR_DATE.ToString("yyyy/MM/dd"), TaiwanDateFormatEnum.民國年);
                    _rptHelper.Replace(ws, TemplateVars.note_rptDate, rptDate, 0, 0, 1, 1);
                    //專案名稱
                    _rptHelper.Replace(ws, TemplateVars.prj_ord_ch, workProgress[0].ORD_CH, 0, 0, 2, 2);
                    //核定工期
                    _rptHelper.Replace(ws, TemplateVars.prj_confirmed_days, workProgress[0].T_DAY.ToString(), 0, 0, 3, 3);
                    //累計工期
                    _rptHelper.Replace(ws, TemplateVars.prj_accumulated_days, (string)workProgress[0].S_DAY.ToString(), 0, 0, 3, 3);
                    //剩餘工期
                    _rptHelper.Replace(ws, TemplateVars.prj_remained_days, workProgress[0].V_DAY.ToString(), 0, 0, 3, 3);
                    //展延工期
                    _rptHelper.Replace(ws, TemplateVars.note_extended_days, dailyNotes[0].EXTEND_DAY.ToString(), 0, 0, 3, 3);
                    //開工工期
                    _rptHelper.TaiwanDateTimeReform(workProgress[0].ABDATE.ToString(), "yyyMMdd", out thedate);
                    _rptHelper.Replace(ws, TemplateVars.prj_kickoff_date, thedate, 0, 0, 4, 4);
                    //完工日期
                    _rptHelper.TaiwanDateTimeReform(workProgress[0].BEDATE.ToString(), "yyyMMdd", out thedate);
                    _rptHelper.Replace(ws, TemplateVars.prj_close_date, thedate, 0, 0, 4, 4);
                    //工程金額(要轉億元單位
                    var number_digits = Math.Round(workProgress[0].CUR_NT / 100000000);
                    _rptHelper.Replace(ws, TemplateVars.prj_cur_nt, $"{number_digits}", 0, 0, 5, 5);
                    //預定進度%
                    number_digits = dailyNotes[0].EXP_PERCENT == null ? 0.00 : dailyNotes[0].EXP_PERCENT;
                    number_digits = Math.Round(number_digits, 2);
                    _rptHelper.Replace(ws, TemplateVars.note_exp_percent, $"{number_digits}", 0, 0, 5, 5);
                    //實際進度%
                    number_digits = dailyNotes[0].ACT_PERCENT == null ? 0.00 : dailyNotes[0].ACT_PERCENT;
                    number_digits = Math.Round(number_digits, 2);
                    _rptHelper.Replace(ws, TemplateVars.note_act_percent, $"{number_digits}", 0, 0, 5, 5);
                    //N億(加代辦管線) 轉億元單位
                    number_digits = workProgress[0].CUR_NT_ADD == null ? 0.00 : workProgress[0].CUR_NT_ADD;
                    number_digits = Math.Round(number_digits / 100000000);
                    _rptHelper.Replace(ws, TemplateVars.prj_cur_nt_add, $"{number_digits}", 0, 0, 6, 6);
                    //N億(加代辦管線) 預定進度%
                    //TODO
                    //N億(加代辦管線) 實際進度%
                    //TODO
                }
                catch (Exception ex)
                {
                    return BadRequest(ex);
                }


                #endregion

                #region 工程項目
                /*******
                 * 
                +------------------------------------------------------------------------+--------+----------+
                | 一、依施工計畫書執行按圖施工概況（含約定之重要施工項目及完成數量等）：           |        |          |
                +-----------------------------------------+---------+--------------------+-----+--+---+------+
                | 營造業專業工程特定施工項目                  |         |                    |     |      |      |
                +-----------------------------------------+---------+--------------------+-----+------+------+
                | A.                                      |         |                    |     |      |      |
                +-----------------------------------------+---------+--------------------+-----+------+------+
                | 1.第1-1工區:里程1K+580-990路基土方回填               |                    |     |      |      |
                +---------------------------------------------------+--------------------+-----+------+------+
                | 2.第1-2工區:里程2K+020-2K+100 地質改良 A區 鑽孔與背填CB漿                                       |
                +--------------------------------------------------------------------------------------------+
                | 3.第1-2工區:里程2K+220-570路基土方回填                                                        |
                +--------------------------------------------------------------------------------------------+
                | 4.第1-2工區:里程2K+560三號大排基礎位置                                                         |
                +--------------------------------------------------------------------------------------------+
                | 5.第2工區:里程26K+600乙工洗車台鋼板復原                                                        |
                +--------------------------------------------------------------------------------------------+
                | 6.第2工區:里程26K+700 管溝開挖及油管預製(中油)                                                  |
                +--------------------------------------------------------------------------------------------+
                | 7.第2工區:里程27K+080海方厝A1橋台基礎開挖作業                                                   |
                +--------------------------------------------------------------------------------------------+
                | 8.第3工區:里程27K+289-28k+194管溝佈管澆置作業(台電)                                            |
                +--------------------------------------------------------------------------------------------+
                | 9.第3工區:里程27K+960-560施工便道及B5類剃除清運                                                |
                +--------------------------------------------------------------------------------------------+
                | 10.第3工區:里程27K+960-560施工便道及B5類剃除清運                                               |
                +--------------------------------------------------------------------------------------------+
                | 11.第3工區:里程28K+420-500(RT側)引水渠道開挖位置高程放樣                                        |
                +--------------------------------------------------------------------------------------------+


                +-------------------------------------------------------------------------------------------------------------------------------------------------+
                | 台4線2K+020~220段(第1-2工區)因路權內之廢棄物及垃圾無法繼續施工須辦理局部停工一案，交通部公路局北區公路新建工程分局同意辦理，本公司依據發文字號: |
                +-----------------------------------------------+------------+------------------+-----------------+-----------------------+-----------------------+
                | B.                                            |            |                  |                 |                       |                       |
                +------------+----------------------------------+------------+------------------+-----------------+-----------------------+-----------------------+
                |    項次     |             施工項目              |    單位    |     契約數量       |   本日完成數量    |      累計完成數量      |          備註          |
                +------------+----------------------------------+------------+------------------+-----------------+-----------------------+-----------------------+
                |            |                                  |            |                  |                 |                       |                       |
                +------------+----------------------------------+------------+------------------+-----------------+-----------------------+-----------------------+
                
                 * **/
                //處理特定施工項目表格
                var specItems = _dbcontext.QueryWorkItems(project_id, WorkItemsEnum.特定施工項目).ToList();
                try
                {




                    if (specItems == null || specItems.Count == 0)
                    {
                        //若資料庫空值，將模板變數清空
                        _rptHelper.Replace(ws, TemplateVars.speical_item, "", 0, 0, 10, 10);
                    }
                    else
                    {

                        //在excel模板找好變數所在的row 當成新增row的起始參考
                        _rptHelper.InsertRows(ws, 10, specItems.Count);
                        //將新增的特定施工項目rows, 將目前列的cells合併成一個cell
                        //TODO:要觀察 因InsertRow 已經處理cellstyle 所以合併時不處理cellstyle
                        for (int i = 1; i < specItems.Count; i++)
                        {
                            //從第12列開始逐一MERGE CELLS
                            IRow row = ws.GetRow(10 + i);
                            _rptHelper.MergeCellsOnRows(row, 1, 0, row.LastCellNum);
                            //將資料庫回傳的目前列轉成Dictionary<string, object>型態
                            Dictionary<string, object> dic = (Dictionary<string, object>)specItems[i].ToObject<Dictionary<string, object>>();
                            //將dic的所有Value 轉成List<string>
                            var vals = dic.Values.OfType<string>().ToList();
                            //將該列所有cells 設定值，本報表的特定施工項目只有一個cell有值
                            _rptHelper.SetCellValues(row, vals);
                        }

                    }

                    //處理施工項目
                    //report_date = "2024/10/24";
                    var workItems = _dbcontext.QueryWorkItems(project_id, WorkItemsEnum.施工項目).ToList();
                    ///正常應該能查到資料，若無資料回傳，可能是資料或查詢條件有問題
                    if (workItems.Count == 0)
                    {
                        //若資料庫空值，將模板變數清空
                        _rptHelper.Replace(ws, TemplateVars.workItems, "", 0, 0, 14, ws.LastRowNum);
                    }
                    else
                    {
                        //找到template 變數位置
                        IRow row = _rptHelper.FindRowByCellValue(ws, TemplateVars.workItems, 0, 10, 14, ws.LastRowNum);
                        int currentrow = row.RowNum;
                        //塞列嶼合併列儲存格
                        _rptHelper.InsertRows(ws, currentrow, workItems.Count - 1);
                        //_rptHelper.CopyRowDown(ws, row.RowNum, workItems.Count);

                        for (int i = 0; i < workItems.Count; i++)
                        {
                            //以儲存格位置為Key, 資料庫回傳資料為值
                            //int[] valuecells = new int[] { 0, 1, 6, 7, 10, 13, 17 }; 準備欲塞值的儲存格位置
                            var cellvalues = new ListDictionary();
                            cellvalues.Add(0, workItems[i].OWNITEM_NO);
                            cellvalues.Add(1, workItems[i].NAME);
                            cellvalues.Add(6, workItems[i].UNIT);
                            cellvalues.Add(7, workItems[i].QUANTITY);
                            cellvalues.Add(10, workItems[i].NOW_QUANTITY);
                            cellvalues.Add(13, workItems[i].SUM_QUANTITY);
                            //TODO 備註欄cellvalues.Add(17, workItems[i].);
                            row = ws.GetRow(currentrow + i);
                            //塞值
                            _rptHelper.SetRow(ws, currentrow + i, cellvalues);

                        }
                    }


                }
                catch (Exception ex)
                {

                    return BadRequest(ex);
                }
                #endregion

                #region 工程資源
                /****
                +-------------------------------------------------------------------------------+
                | 二、工地材料管理概況（含約定之重要材料使用狀況及數量等）：                            |
                +---------------------+---------+----------+--------------+--------------+------+
                |       材料名稱       |   單位   | 契約數量  | 本日使用數量   | 累計使用數量   | 備註 |
                +-----+---------------+---------+----------+--------------+--------------+------+
                |     |               |    T    |          |              |              |      |
                +-----+---------------+---------+----------+--------------+--------------+------+
                |     |               |    T    |          |              |              |      |
                +-----+---------------+---------+----------+--------------+--------------+------+
                |     |               |    M3   |          |              |              |      |
                +-----+---------------+---------+----------+--------------+--------------+------+
                |     |               |    M3   |          |              |              |      |
                +-----+---------------+---------+----------+--------------+--------------+------+
                
                +---------------------------------------------------------------------------------+
                | 三、工地人員及機具管理（含約定之出工人數及機具使用情形及數量）：                         |
                +--------------+----------+----------+--------------+--------------+--------------+
                |     工別      | 本日人數  | 累計人數  |   機具名稱    | 本日使用數量   | 累計使用數量  |
                +--------------+----------+----------+--------------+--------------+--------------+
                |    工程師     |     3    |     3    |    掃路機     |       2      |       2      |
                +--------------+----------+----------+--------------+--------------+--------------+
                |    工程師     |     3    |     1    |    掃路機     |       2      |       1      |
                +--------------+----------+----------+--------------+--------------+--------------+
                |              |          |          |    吊卡車     |      23      |      23      |
                +--------------+----------+----------+--------------+--------------+--------------+
                |              |          |          |    吊卡車     |      23      |       1      |
                +--------------+----------+----------+--------------+--------------+--------------+
                |              |          |          |     吊車      |       1      |       1      |
                +--------------+----------+----------+--------------+--------------+--------------+
                |              |          |          |     吊車      |       1      |       1      |
                +--------------+----------+----------+--------------+--------------+--------------+
                |              |          |          |   機具租金    |      54      |      54      |
                +--------------+----------+----------+--------------+--------------+--------------+
                |              |          |          |   機具租金    |      54      |       1      |
                +--------------+----------+----------+--------------+--------------+--------------+
                 ****/
                //處理工地材料
                //report_date = "2023/08/14";//TODO
                //project_id = 4800; //TODO
                var materials = _dbcontext.QueryResourceItems(project_id, report_date, ResourceClassEnum.材料).ToList();
                try
                {

                    ///
                    if (materials.Count == 0)
                    {
                        //若資料庫空值，將模板變數清空
                        _rptHelper.Replace(ws, TemplateVars.materials, "", 0, 0, 14, ws.LastRowNum);
                    }
                    else
                    {
                        //找到template 變數位置
                        IRow row = _rptHelper.FindRowByCellValue(ws, TemplateVars.materials, 0, 10, 14, ws.LastRowNum);
                        int currentrow = row.RowNum;
                        //塞列與合併列儲存格
                        _rptHelper.InsertRows(ws, currentrow, materials.Count - 1);

                        for (int i = 0; i < materials.Count; i++)
                        {
                            //以儲存格位置為Key, 資料庫回傳資料為值
                            var cellvalues = new ListDictionary();
                            cellvalues.Add(0, materials[i].NAME);
                            cellvalues.Add(6, materials[i].UNIT);
                            cellvalues.Add(7, materials[i].QUANTITY);
                            cellvalues.Add(10, materials[i].TODAY_QTY);
                            cellvalues.Add(13, materials[i].SUM_QTY);
                            //TODO 備註欄cellvalues.Add(17, materials[i].EDR_RES_RMK);
                            row = ws.GetRow(currentrow + i);
                            //塞值
                            _rptHelper.SetRow(ws, currentrow + i, cellvalues);

                        }
                    }

                    //using var ms1 = new MemoryStream();
                    //wb.Write(ms1);
                    //wb.Close();

                    //_logger.Log(LogLevel.Information, "Exiting RPT1CB004_1");
                    //return File(ms1.ToArray(), FileMime.Excel);
                    //處理工地人員與機具管理
                    //report_date = "2023/08/14";//TODO
                    //project_id = 4800; //TODO
                    var engineers = _dbcontext.QueryResourceItems(project_id, report_date, ResourceClassEnum.人工).ToList();
                    var equipments = _dbcontext.QueryResourceItems(project_id, report_date, ResourceClassEnum.機具器材).ToList();

                    //取工地人員與機具的資料庫回傳最大筆數，因為要新增row要以最大值為準
                    int maxItemCount = engineers.Count >= equipments.Count ? engineers.Count : equipments.Count;

                    if (engineers.Count == 0)
                    {
                        //若資料庫空值，將模板變數清空
                        _rptHelper.Replace(ws, TemplateVars.engineers, "", 0, 0, 14, ws.LastRowNum);
                    }
                    if (equipments.Count == 0)
                    {
                        //若資料庫空值，將模板變數清空
                        _rptHelper.Replace(ws, TemplateVars.equipments, "", 0, 0, 14, ws.LastRowNum);
                    }

                    if (maxItemCount > 0)
                    {
                        //因工地人員與機具管理可能生成很多rows, 故模版變數要重新查找到所在的row 
                        IRow row = _rptHelper.FindRowByCellValue(ws, TemplateVars.engineers, 0, 10, 14, ws.LastRowNum);
                        int currentrow = row.RowNum;
                        //塞列嶼合併列儲存格
                        _rptHelper.InsertRows(ws, currentrow, maxItemCount - 1);
                        for (int i = 0; i < engineers.Count; i++)
                        {
                            //以儲存格位置為Key, 資料庫回傳資料為值
                            //int[] valuecells = new int[] { 0, 1, 6, 7, 10, 13, 17 }; 準備欲塞值的儲存格位置
                            var cellvalues = new ListDictionary();
                            if (i < engineers.Count)
                            {
                                cellvalues.Add(0, engineers[i].NAME);
                                cellvalues.Add(3, engineers[i].TODAY_QTY);
                                cellvalues.Add(6, engineers[i].SUM_QTY);
                            }
                            if (i < equipments.Count)
                            {
                                cellvalues.Add(10, engineers[i].NAME);
                                cellvalues.Add(13, engineers[i].TODAY_QTY);
                                cellvalues.Add(17, engineers[i].SUM_QTY);
                            }
                            //TODO 備註欄cellvalues.Add(17, workItems[i].);
                            row = ws.GetRow(currentrow + i);
                            //塞值
                            _rptHelper.SetRow(ws, currentrow + i, cellvalues);

                        }
                        //using var ms1 = new MemoryStream();
                        //wb.Write(ms1);
                        //wb.Close();

                        //_logger.Log(LogLevel.Information, "Exiting RPT1CB004_1");
                        //return File(ms1.ToArray(), FileMime.Excel);

                        //List<int> cellnums1 = new List<int>();
                        //List<int> cellnums2 = new List<int>();
                        //for (int i = 0; i < maxItemCount; i++)
                        //{
                        //    //由於工地人員與機具管理每一列row將需要多個merged cells，需各別處理
                        //    //從第n列開始
                        //    row = ws.GetRow(row.RowNum);
                        //    //在該列merging cells，並回傳每個Mergedcells的第一格，作為存入資料用
                        //    cellnums1 = MergeCellsEngineer(row);
                        //    //將資料庫回傳的目前一列轉成Dictionary<string, object>型態
                        //    Dictionary<string, object> dic1 = (Dictionary<string, object>)engineers[i].ToObject<Dictionary<string, object>>();
                        //    //將dic1的所有Value 轉成List<string>
                        //    var vals1 = dic1.Values.OfType<string>().ToList();
                        //    //將該列所有cells 設定值
                        //    _rptHelper.SetCellValues(row, vals1, cellnums1);

                        //    cellnums2 = MergeCellsEquipment(row);
                        //    //將資料庫回傳的目前一列轉成Dictionary<string, object>型態
                        //    Dictionary<string, object> dic2 = (Dictionary<string, object>)equipments[i].ToObject<Dictionary<string, object>>();
                        //    //將dic2的所有Value 轉成List<string>
                        //    var vals2 = dic2.Values.OfType<string>().ToList();
                        //    //將該列所有cells 設定值
                        //    _rptHelper.SetCellValues(row, vals2, cellnums2);
                        //}
                    }

                }
                catch (Exception ex)
                {
                    return BadRequest(ex);
                }
                #endregion

                #region 工地記事
                /*
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 四、本日施工項目是否有須依「營造業專業工程特定施工項目應置之技術士種類、比率或人數標準表」規定應設置技術士之專業工程：☐有☐無   (此項若勾選有"，則應填寫後附「公共工程施工日誌之技術士簽章表」) |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 五、工地職業安全衛生事項之督導、公共環境與安全之維護及其他工地行政事務：                                                                                                                                  |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | (ㄧ)施工前檢查事項：                                                                                                                                                                             |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 1.實施勤前教育(含工地預防災變及危害告知)：☐有☐無                                                                                                                                                    |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：☐有 ☐無 ☐無新進勞工                                                                                                            |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 3.檢查勞工個人防護具：☐有 ☐無                                                                                                                                                                    |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+-----------+
                | (二)其他事項：                                                                                                                                                                       |           |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+-----------+
                |                                                                                                                                                                                                |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 六、施工取樣試驗紀錄：                                                                                                                                                                            |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                |                                                                                                                                                                                                |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 七、通知分包廠商辦理事項：                                                                                                                                                                         |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 無。                                                                                                                                                                                           |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 八、重要事項記錄：                                                                                                                                                                               |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 無。                                                                                                                                                                                           |
                +------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+ 
                 */
                //工地記事
                //data from dailyNotes


                if (dailyNotes.Count == 0)
                {
                    _rptHelper.Replace(ws, TemplateVars.notes_item4, "", 0, 0, 1, 1);
                    _rptHelper.Replace(ws, TemplateVars.notes_misc, "", 0, 0, 1, 1);
                    _rptHelper.Replace(ws, TemplateVars.notes_experiment, "", 0, 0, 1, 1);
                    _rptHelper.Replace(ws, TemplateVars.notes_item7, "", 0, 0, 1, 1);
                    _rptHelper.Replace(ws, TemplateVars.notes_item8, "", 0, 0, 1, 1);




                }
                else
                {
                    //四、
                    var str = dailyNotes[0].NOTES_A == null ? "" : dailyNotes[0].NOTES_A;
                    var notesA_item4 = str.Contains("勞工安全查核事項") ? "◼有 □無" : "□有 ◼無";
                    _rptHelper.Replace(ws, TemplateVars.notes_item4, notesA_item4, 0, 0, 1, 1);
                    //五、(二)其他事項
                    str = dailyNotes[0].NOTES == null ? "" : dailyNotes[0].NOTES;
                    var notes_misc = str;
                    _rptHelper.Replace(ws, TemplateVars.notes_misc, dailyNotes[0].NOTES, 0, 0, 1, 1);

                    //六、施工取樣試驗紀錄：

                    string notesA_item6 = string.Empty;
                    var test = (string)dailyNotes[0].NOTES_A;
                    if (test.Contains("試體") || test.Contains("取樣") || test.Contains("檢驗"))
                    {
                        notesA_item6 = test;
                    }
                    else
                    {
                        notesA_item6 = "";
                    }
                    _rptHelper.Replace(ws, TemplateVars.notes_experiment, dailyNotes[0].NOTES, 0, 0, 1, 1);
                    //七、通知分包廠商辦理事項：    
                    var notes_item7 = dailyNotes[0].NOTES_B == null ? "" : dailyNotes[0].NOTES_B;
                    _rptHelper.Replace(ws, TemplateVars.notes_item7, notes_item7, 0, 0, 1, 1);
                    //八、重要事項記錄：
                    var notes_item8 = dailyNotes[0].NOTES_C == null ? "" : dailyNotes[0].NOTES_C;
                    _rptHelper.Replace(ws, TemplateVars.notes_item8, notes_item8, 0, 0, 1, 1);
                }
                #endregion

                using var ms = new MemoryStream();
                wb.Write(ms);
                wb.Close();

                _logger.Log(LogLevel.Information, "Exiting RPT1CB004_1");
                return File(ms.ToArray(), FileMime.Excel);
            }
            catch (Exception ex)
            {
                var error = $"[RPT1CB004_1][Error][Message]:{ex.Message}[InnerException]:{ex.InnerException}";
                _logger.Log(LogLevel.Error, error);
                return BadRequest(error);
            }


        }
        private List<int> MergeCellsWorkItem(IRow row)
        {
            List<int> cellnum = new List<int>();
            //項次 直接塞值
            cellnum.Add(0);
            //TODO:要觀察 因InsertRow 已經處理cellstyle 所以合併時不處理cellstyle
            //施工項目
            _rptHelper.MergeCellsOnRows(row, 1, 1, 5);
            cellnum.Add(1);
            //單位 直接塞值
            cellnum.Add(6);
            //契約數量
            _rptHelper.MergeCellsOnRows(row, 1, 7, 9);
            cellnum.Add(7);
            //本日完成數量
            _rptHelper.MergeCellsOnRows(row, 1, 10, 12);
            cellnum.Add(11);
            //累計完成數量
            _rptHelper.MergeCellsOnRows(row, 1, 13, 16);
            cellnum.Add(13);
            //備註
            _rptHelper.MergeCellsOnRows(row, 1, 17, 19);
            cellnum.Add(17);

            return cellnum;
        }

        private List<int> MergeCellsMaterial(IRow row)
        {
            List<int> cellnum = new List<int>();
            //材料名稱
            _rptHelper.MergeCellsOnRows(row, 1, 0, 5);
            cellnum.Add(0);
            //單位 直接塞值
            cellnum.Add(6);
            //契約數量
            _rptHelper.MergeCellsOnRows(row, 1, 7, 9);
            cellnum.Add(7);
            //本日完成數量
            _rptHelper.MergeCellsOnRows(row, 1, 10, 12);
            cellnum.Add(11);
            //累計完成數量
            _rptHelper.MergeCellsOnRows(row, 1, 13, 16);
            cellnum.Add(13);
            //備註
            _rptHelper.MergeCellsOnRows(row, 1, 17, 19);
            cellnum.Add(17);

            return cellnum;
        }

        private List<int> MergeCellsEngineer(IRow row)
        {
            List<int> cellnum = new List<int>();
            //工別
            _rptHelper.MergeCellsOnRows(row, 1, 0, 2);
            cellnum.Add(0);
            //本日人數
            _rptHelper.MergeCellsOnRows(row, 1, 3, 5);
            cellnum.Add(3);
            //累計人數
            _rptHelper.MergeCellsOnRows(row, 1, 6, 9);
            cellnum.Add(6);

            return cellnum;
        }

        private List<int> MergeCellsEquipment(IRow row)
        {
            List<int> cellnum = new List<int>();
            //機具名稱
            _rptHelper.MergeCellsOnRows(row, 1, 10, 12);
            cellnum.Add(10);
            //本日使用數量
            _rptHelper.MergeCellsOnRows(row, 1, 13, 16);
            cellnum.Add(13);
            //累計使用數量
            _rptHelper.MergeCellsOnRows(row, 1, 17, 19);
            cellnum.Add(17);


            return cellnum;
        }
    }
}
