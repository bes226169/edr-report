using Dapper;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Text.RegularExpressions;

namespace EDR_Report.Controllers
{
    public partial class ReportController
    {
        
        IActionResult RPTLAYOUT(string? state, string cd)
        {
            
            // 報表參數
            var projectId = UserInfo.ProjectID;
            var myuserId = UserInfo.UserID;
            var calendarDate = new DateTime(int.Parse(cd[..4]), int.Parse(cd.Substring(4, 2)), int.Parse(cd.Substring(6, 2)));
            string filePath= $"{BasePath.TrimEnd('\\').TrimEnd('/')}\\";
            string workOrder;

            ////////////
            // 編輯區1//
            ////////////
            if (projectId == 4341)
            {
                /// <summary>
                /// A7A 台大健康
                /// <para>Project ID: 4341</para>
                /// <para>工令: 1BA701</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1ba701";
            }
            else if (projectId == 4760)
            {
                /// <summary>
                /// A8B 八德
                /// <para>Project ID: 4760</para>
                /// <para>工令: 1CA802</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1ca802";
            }
            else if (projectId == 4780)
            {
                /// <summary>
                /// A8C 鳥嘴潭所
                /// <para>Project ID: 4780</para>
                /// <para>工令: 1CA803</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1ca803";
            }
            else if (projectId == 5320)
            {
                /// <summary>
                /// B0A 新展
                /// <para>Project ID: 5320</para>
                /// <para>工令: 1CB001</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb001";
            }
            else if (projectId == 5460)
            {
                /// <summary>
                /// A7A 台大健康
                /// <para>Project ID: 5460</para>
                /// <para>工令: 1BB006</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1bb006";
            }
            else if (projectId == 5520)
            {
                /// <summary>
                /// A6E 大林蒲
                /// <para>Project ID: 5520</para>
                /// <para>工令: 1BB101</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1bb101";
            }
            else if (projectId == 5541)
            {
                /// <summary>
                /// A8F 大埔
                /// <para>Project ID: 5541</para>
                /// <para>工令: 1CB102</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb102";
            }
            else if (projectId == 5621)
            {
                /// <summary>
                /// B1C 寶科
                /// <para>Project ID: 5621</para>
                /// <para>工令: 1CB105</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb105";
            }
            else if (projectId == 5781)
            {
                /// <summary>
                /// B1F 橋科
                /// <para>Project ID: 5781</para>
                /// <para>工令: 1CB110</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb110";
            }
            else if (projectId == 5782)
            {
                /// <summary>
                /// A6E 大林蒲
                /// <para>Project ID: 5782</para>
                /// <para>工令: 1CB108</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb108";
            }
            else if (projectId == 5820)
            {
                /// <summary>
                /// B1H 欣科
                /// <para>Project ID: 5820</para>
                /// <para>工令: 1CB111</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb111";
            }
            else if (projectId == 5900)
            {
                /// <summary>
                /// B2A 通宵
                /// <para>Project ID: 5900</para>
                /// <para>工令: 1CB202</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb202";
            }
            else if (projectId == 6000)
            {
                /// <summary>
                /// B2D 松湖工務所
                /// <para>Project ID: 6000</para>
                /// <para>工令: 1BB206</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1bb206";
            }
            else if (projectId == 6020)
            {
                /// <summary>
                /// B1H 欣科
                /// <para>Project ID: 6020</para>
                /// <para>工令: 1CB207</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb207";
            }
            else if (projectId == 6040)
            {
                /// <summary>
                /// B2D 博誠工務所
                /// <para>Project ID: 6040</para>
                /// <para>工令: 1CB208</para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb208";
            }
            else if (projectId == 5780)
            {
                /// <summary>
                /// B1E 華桂聯合工務所
                /// <para>Project ID: 5780</para>
                /// <para>工令: 1CB109/para>
                /// </summary>
                /// <param name="state"></param>
                /// <param name="cd"></param>
                /// <returns></returns>
                workOrder = "1cb109";
            }
            else
            {
                workOrder = "";
            }

            filePath = string.Concat(filePath, $"edr{workOrder}.xlsx");
            if (!System.IO.File.Exists(filePath)) filePath = $"report/edr{workOrder}.xlsx";
            if (!System.IO.File.Exists(filePath)) return NotFound("找不到套表檔");

            DBFunc db = new();
            // 讀入資料
            ////////////
            // 編輯區2//
            ////////////
            var projInfo = GetProjInfo(db, projectId, calendarDate);
            var projConsOverview = projectId switch
            {
                6040 => GetProjConsOverview(db, projectId, calendarDate, myuserId, "1", false), // 1:所有, 2:本日
                5541 => GetProjConsOverview(db, projectId, calendarDate, myuserId, "1", false),
                5782 => GetProjConsOverview(db, projectId, calendarDate, myuserId, "1", false),
                4760 => GetProjConsOverview(db, projectId, calendarDate, myuserId, "1", false),
                _ => GetProjConsOverview(db, projectId, calendarDate, myuserId, state, false)
            };
            var projConsOverviewSpec = GetProjConsOverview(db, projectId, calendarDate, myuserId, state, true);
            var projMan = projectId switch
            {
                5782 => GetProjManMachineMaterial(db, projectId, calendarDate, "1", "man"), //1:所有使用,2:本日使用,3:所有項目
                6040 => GetProjManMachineMaterial(db, projectId, calendarDate, "3", "man"), //1:所有使用,2:本日使用,3:所有項目
                _ => GetProjManMachineMaterial(db, projectId, calendarDate, "2", "man") //1:所有使用,2:本日使用,3:所有項目
            };
            var projMachine = projectId switch
            {
                5782 => GetProjManMachineMaterial(db, projectId, calendarDate, "1", "machine"), //1:所有使用,2:本日使用,3:所有項目
                6040 => GetProjManMachineMaterial(db, projectId, calendarDate, "3", "machine"), //1:所有使用,2:本日使用,3:所有項目
                _ => GetProjManMachineMaterial(db, projectId, calendarDate, "2", "machine") //1:所有使用,2:本日使用,3:所有項目
            };
            var projMaterial = projectId switch
            {
                5782 => GetProjManMachineMaterial(db, projectId, calendarDate, "1", "material"), //1:所有使用,2:本日使用,3:所有項目
                6040 => GetProjManMachineMaterial(db, projectId, calendarDate, "3", "material"), //1:所有使用,2:本日使用,3:所有項目
                _ => GetProjManMachineMaterial(db, projectId, calendarDate, "2", "material") //1:所有使用,2:本日使用,3:所有項目
            };
            var projNote = GetProjNote(db, projectId, calendarDate);
            var projMilestone = GetProjMilestone(db, projectId);
            // 待確認資料來源
            projInfo["CONSTRUCTOR"] = new object[] { "中華工程股份有限公司" };
            // 建立資料字典檔(只增不刪)
            Dictionary<string, object[]> variables = new()
            {
                { "$report_serial_t1$", projInfo["REPORT_SERIAL_T1"]},     // '112A040C009-' + 4碼
                { "$report_serial_t2$", projInfo["REPORT_SERIAL_T2"]},     // 4碼
                { "$report_serial_t3$", projInfo["REPORT_SERIAL_T3"]},     // 3碼
                { "$weather_am$", projInfo["WEATHER_AM"] },
                { "$weather_pm$", projInfo["WEATHER_PM"] },
                { "$calendar_date_t1$", projInfo["CALENDAR_DATE_T1"] },    // 112 年 01 月 01 日  星期一
                { "$calendar_date_t1_2$", projInfo["CALENDAR_DATE_T1"] },  // 112 年 01 月 01 日  星期一
                { "$calendar_date_t2$", projInfo["CALENDAR_DATE_T2"] },    // 112年01月01日(星期一)
                { "$calendar_date_t3$", projInfo["CALENDAR_DATE_T3"] },    // 112年01月01日 星期一
                { "$calendar_date$", projInfo["CALENDAR_DATE"] },
                { "$calendar_date_weekday$", projInfo["CALENDAR_DATE_WEEKDAY"] },
                { "$project_name$", projInfo["PROJECT_NAME"] },
                { "$project_own_ch$", projInfo["PROJECT_OWN_CH"] },
                { "$contractor$", projInfo["CONSTRUCTOR"]  },
                { "$t_day$", projInfo["TDAY"] },                           //1,234天
                { "$construction_period$", projInfo["CONSTRUCTION_PERIOD"]},
                { "$s_day$", projInfo["SDAY"] },                           //1,234天
                { "$v_day$", projInfo["VDAY"] },                           //1,234天
                { "$t_day_t2$", projInfo["TDAY_T2"] },                     //1,234日曆天
                { "$s_day_t2$", projInfo["SDAY_T2"] },                     //1,234日曆天
                { "$s_day_t3$", projInfo["SDAY_T3"] },                     //1,234.0日曆天
                { "$v_day_t2$", projInfo["VDAY_T2"] },                     //1,234日曆天
                { "$v_day_sub$", projInfo["VDAY_SUB"] },
                { "$vday_sub_t2$", projInfo["VDAY_SUB_T2"] },               //1,234.0日曆天
                { "$vday_sub_num$", projInfo["VDAY_SUB_NUM"] },
                { "$spread_day$", projInfo["SPREAD_DAY"] },
                { "$spread_day_t2$", projInfo["SPREAD_DAY_T2"] },
                { "$spread_day_num$", projInfo["SPREAD_DAY_NUM"] },
                { "$state_date$", projInfo["START_DATE"] },
                { "$end_date$", projInfo["END_DATE"] },
                { "$state_date_wkd$", projInfo["START_DATE_WKD"] },
                { "$end_date_wkd$", projInfo["END_DATE_WKD"] },
                { "$exp_percent$", projInfo["EXP_PERCENT"] },
                { "$exp_percent_t2$", projInfo["EXP_PERCENT_T2"] },
                { "$exp_percent_t3$", projInfo["EXP_PERCENT_T3"] },         // 0.01
                { "$act_sum$", projInfo["ACT_SUM"] },
                { "$act_sum_t2$", projInfo["ACT_SUM_T2"] },
                { "$act_sum_t3$", projInfo["ACT_SUM_T3"] },                 // 0.01
                { "$act_percent_t3$",projInfo["ACT_PERCENT_T3"] },           // 0.01
                { "$diff_percent$",projInfo["DIFF_PERCENT"] },
                { "$nocal_day_t2$",projInfo["NOCAL_DAY_T2"]},                //1,234.0日曆天
                { "$m50$", projInfo["M50"] },
                { "$m51$", projInfo["M51"] },
                { "$m52$", projInfo["M52"] },
                { "$m53$", projInfo["M53"] },
                { "$m54$", projInfo["M54"] },

                { "$period1$", projMilestone["MILESTONE_NO_SUB"] },
                { "$period2$", projMilestone["MILESTONE_NO_SUB"] },
                { "$period3$", projMilestone["MILESTONE_NO_SUB"] },
                { "$period4$", projMilestone["MILESTONE_NO_SUB"] },
                { "$period5$", projMilestone["MILESTONE_NO_SUB"] },
                { "$period6$", projMilestone["MILESTONE_NO_SUB"] },
                { "$period7$", projMilestone["MILESTONE_NO_SUB"] },
                { "$period8$", projMilestone["MILESTONE_NO_SUB"] },
                { "$wk_day$", projMilestone["MILESTONE_WK_DAY"] },

                { "$co_col1$", projConsOverview["OWNITEM_NO"]},
                { "$co_col2$", projConsOverview["NAME"]},
                {"$construction_item$", projConsOverview["CONSTRUCTION_ITEM"]},
                { "$co_col3$", projConsOverview["UNIT"]},
                { "$co_col4$", projConsOverview["QUANTITY"]},
                { "$co_col5$", projConsOverview["NOW_EDR_QUANTITY"]},
                { "$co_col6$", projConsOverview["SUM_EDR_QUANTITY"]},
                { "$co_col4_2d$", projConsOverview["QUANTITY_2DECI"]},
                { "$co_col5_2d$", projConsOverview["NOW_EDR_QUANTITY_2DECI"]},
                { "$co_col6_2d$", projConsOverview["SUM_EDR_QUANTITY_2DECI"]},
                { "$co_col4_4d$", projConsOverview["QUANTITY_4DECI"]},
                { "$co_col5_4d$", projConsOverview["NOW_EDR_QUANTITY_4DECI"]},
                { "$co_col6_4d$", projConsOverview["SUM_EDR_QUANTITY_4DECI"]},
                //{ "$co_col7$", projConsOverview["REMARK"]},

                { "$cos_col1$", projConsOverviewSpec["OWNITEM_NO"]},
                { "$cos_col2$", projConsOverviewSpec["NAME"]},
                { "$cos_col3$", projConsOverviewSpec["UNIT"]},
                { "$cos_col4$", projConsOverviewSpec["QUANTITY"]},
                { "$cos_col5$", projConsOverviewSpec["NOW_EDR_QUANTITY"]},
                { "$cos_col6$", projConsOverviewSpec["SUM_EDR_QUANTITY"]},
                //{ "$cos_col7$", projConsOverviewSpec["REMARK"]},

                { "$material_col1$", projMaterial["SEQUENCE_NO"]},
                { "$material_col2$", projMaterial["NAME"]},
                { "$material_col3$", projMaterial["UNIT"]},
                { "$material_col4$", projMaterial["QUANTITY"]},
                { "$material_col5$", projMaterial["TODAY_QTY"]},
                { "$material_col6$", projMaterial["SUM_QTY"]},
                { "$material_col5_2d$", projMaterial["TODAY_QTY_2DECI"]},
                { "$material_col6_2d$", projMaterial["SUM_QTY_2DECI"]},
                //{ "$material_col7$", projMaterial["REMARK"]},

                { "$man_col1$", projMan["NAME"]},
                { "$man_col2$", projMan["TODAY_QTY"]},
                { "$man_col3$", projMan["SUM_QTY"]},
                { "$man_col2_2d$", projMan["TODAY_QTY_2DECI"]},
                { "$man_col3_2d$", projMan["SUM_QTY_2DECI"]},

                { "$machine_col1$", projMachine["NAME"]},
                { "$machine_col2$", projMachine["TODAY_QTY"]},
                { "$machine_col3$", projMachine["SUM_QTY"]},
                { "$machine_col2_2d$", projMachine["TODAY_QTY_2DECI"]},
                { "$machine_col3_2d$", projMachine["SUM_QTY_2DECI"]},

                { "$note$", projNote["NOTES"]},
                { "$note_a$", projNote["NOTES_A"]},
                { "$note_b$", projNote["NOTES_B"] },
                { "$note_c$", projNote["NOTES_C"] },
                { "$note_d$", projNote["NOTES_D"] },
                { "$note_e$", projNote["NOTES_E"] },
                { "$note_f$", projNote["NOTES_F"] },
                { "$note_g$", projNote["NOTES_G"] },
            };

            IWorkbook wb;
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // 讀取xlsx檔案
                wb = new XSSFWorkbook(fileStream);
                ISheet ws = wb.GetSheetAt(0);

                ////////////
                // 編輯區3//
                ////////////
                if (projectId == 4780)
                {
                    /// <summary>
                    /// A8C 鳥嘴潭所
                    /// <para>Project ID: 4780</para>
                    /// <para>工令: 1CA803</para>
                    /// </summary>
                    /// <param name="state"></param>
                    /// <param name="cd"></param>
                    /// <returns></returns>
                    // 依資料長度調整列數
                    AdjustRowNums(ws, variables, "$co_col2$");
                    AdjustRowNums(ws, variables, "$material_col2$");
                    AdjustRowNums(ws, variables, "$man_col1$", "$machine_col1$");
                    // 調整高度，要先調列數再調高度不然後面shift上去的row高度會改為預設高度
                    AdjustRowHeight(ws, variables, "$co_col2$");
                    AdjustRowHeight(ws, variables, "$material_col2$");
                    AdjustRowHeight(ws, variables, "$man_col1$", "$machine_col1$");
                    AdjustRowHeight(ws, variables, "$note$");
                    AdjustRowHeight(ws, variables, "$note_a$");
                    AdjustRowHeight(ws, variables, "$note_b$");
                    AdjustRowHeight(ws, variables, "$note_c$");

                }
                else if (projectId == 4760)
                {
                    /// <summary>
                    /// A8B 八德
                    /// <para>Project ID: 4760</para>
                    /// <para>工令: 1CA802</para>
                    /// </summary>
                    /// <param name="state"></param>
                    /// <param name="cd"></param>
                    /// <returns></returns>
                    AdjustRowNums(ws, variables, "$co_col2$");
                    //AdjustRowNums(ws, variables, "$material_col2$");
                    AdjustRowNums(ws, variables, "$man_col1$", "$machine_col1$");
                    // 調整高度，要先調列數再調高度不然後面shift上去的row高度會改為預設高度
                    AdjustRowHeight(ws, variables, "$project_name$");
                    AdjustRowHeight(ws, variables, "$co_col2$");
                    //AdjustRowHeight(ws, variables, "$material_col2$");
                    AdjustRowHeight(ws, variables, "$man_col1$", "$machine_col1$");
                    AdjustRowHeight(ws, variables, "$note_a$");
                    AdjustRowHeight(ws, variables, "$note_c$");
                    AdjustRowHeight(ws, variables, "$note_d$");
                }
                else if (projectId == 5520)
                {

                    AdjustRowNums(ws, variables, "$construction_item$");

                    AdjustRowNums(ws, variables, "$man_col1$", "$machine_col1$", "$material_col2$");
                    // 調整高度，要先調列數再調高度不然後面shift上去的row高度會改為預設高度
                    //AdjustRowHeight_1BB101(ws, variables, "$project_name$");
                    AdjustRowHeight(ws, variables, "$construction_item$");
                    AdjustRowHeight(ws, variables, "$machine_col1$", "$material_col2$", "$material_col2$");
                    AdjustRowHeight(ws, variables, "$note_a$");
                    AdjustRowHeight(ws, variables, "$note_c$");
                    AdjustRowHeight(ws, variables, "$note_d$");
                }
                else if (projectId == 5541)
                {
                    /// <summary>
                    /// A8F 大埔
                    /// <para>Project ID: 5541</para>
                    /// <para>工令: 1CB102</para>
                    /// </summary>
                    /// <param name="state"></param>
                    /// <param name="cd"></param>
                    /// <returns></returns>
                    AdjustRowNums(ws, variables, "$co_col2$");
                    AdjustRowNums(ws, variables, "$material_col2$");
                    AdjustRowNums(ws, variables, "$man_col1$", "$machine_col1$");
                    // 調整高度，要先調列數再調高度不然後面shift上去的row高度會改為預設高度
                    AdjustRowHeight(ws, variables, "$project_name$");
                    AdjustRowHeight(ws, variables, "$co_col2$");
                    AdjustRowHeight(ws, variables, "$material_col2$");
                    AdjustRowHeight(ws, variables, "$man_col1$", "$machine_col1$");
                    AdjustRowHeight(ws, variables, "$note$");
                    AdjustRowHeight(ws, variables, "$note_a$");
                    AdjustRowHeight(ws, variables, "$note_b$");
                    AdjustRowHeight(ws, variables, "$note_c$");
                    AdjustRowHeight(ws, variables, "$note_d$");
                    AdjustRowHeight(ws, variables, "$note_g$");
                }
                else if (projectId == 5782)
                {
                    /// <summary>
                    /// A6E 大林蒲
                    /// <para>Project ID: 5782</para>
                    /// <para>工令: 1CB108</para>
                    /// </summary>
                    /// <param name="state"></param>
                    /// <param name="cd"></param>
                    /// <returns></returns>
                    AdjustRowNums(ws, variables, "$co_col2$");
                    AdjustRowNums(ws, variables, "$man_col1$", "$machine_col1$", "$material_col2$");
                    // 調整高度，要先調列數再調高度不然後面shift上去的row高度會改為預設高度
                    AdjustRowHeight(ws, variables, "$project_name$");
                    AdjustRowHeight(ws, variables, "$co_col2$");
                    AdjustRowHeight(ws, variables, "$man_col1$", "$machine_col1$", "$material_col2$");
                    AdjustRowHeight(ws, variables, "$note_a$");
                    AdjustRowHeight(ws, variables, "$note_c$");
                    AdjustRowHeight(ws, variables, "$note_d$");
                    AdjustRowHeight(ws, variables, "$note_f$");
                }
                else if (projectId == 6000)
                {
                    /// <summary>
                    /// B2D 松湖工務所
                    /// <para>Project ID: 6000</para>
                    /// <para>工令: 1BB206</para>
                    /// </summary>
                    /// <param name="state"></param>
                    /// <param name="cd"></param>
                    /// <returns></returns>
                    //AdjustRowNumsAndHeight(ws, variables, "$period1$");
                    //AdjustRowNumsAndHeight(ws, variables, "$period5$");
                    //VerticalMergeRows(ws, "$period1$", new int[] { 1, 9, 17, 25 });
                    //VerticalMergeRows(ws, "$period5$", new int[] { 1, 9, 17, 25 });
                }
                else if (projectId == 6040)
                {
                    /// <summary>
                    /// B2D 博誠工務所
                    /// <para>Project ID: 6040</para>
                    /// <para>工令: 1CB208</para>
                    /// </summary>
                    /// <param name="state"></param>
                    /// <param name="cd"></param>
                    /// <returns></returns>
                    // 依資料長度調整列數
                    AdjustRowNums(ws, variables, "$co_col2$");
                    AdjustRowNums(ws, variables, "$material_col2$");
                    AdjustRowNums(ws, variables, "$man_col1$", "$machine_col1$");
                    // 調整高度，要先調列數再調高度不然後面shift上去的row高度會改為預設高度
                    AdjustRowHeight(ws, variables, "$project_name$");
                    AdjustRowHeight(ws, variables, "$co_col2$");
                    AdjustRowHeight(ws, variables, "$material_col2$");
                    AdjustRowHeight(ws, variables, "$man_col1$", "$machine_col1$");
                    AdjustRowHeight(ws, variables, "$note$");
                    AdjustRowHeight(ws, variables, "$note_a$");
                    AdjustRowHeight(ws, variables, "$note_b$");
                    AdjustRowHeight(ws, variables, "$note_c$");
                    AdjustRowHeight(ws, variables, "$note_d$");
                    AdjustRowHeight(ws, variables, "$note_g$");
                }
                else if (projectId == 5780)
                {
                    /// <summary>
                    /// B1E 華桂聯合工務所
                    /// <para>Project ID: 5780</para>
                    /// <para>工令: 1CB109</para>
                    /// </summary>
                    /// <param name="state"></param>
                    /// <param name="cd"></param>
                    /// <returns></returns>
                    // 依資料長度調整列數
                    AdjustRowNums(ws, variables, "$material_col2$");
                    AdjustRowNums(ws, variables, "$man_col1$", "$machine_col1$");
                    // 調整高度，要先調列數再調高度不然後面shift上去的row高度會改為預設高度
                    AdjustRowHeight(ws, variables, "$project_name$");
                    AdjustRowHeight(ws, variables, "$material_col2$");
                    AdjustRowHeight(ws, variables, "$man_col1$", "$machine_col1$");
                    AdjustRowHeight(ws, variables, "$note$");
                    AdjustRowHeight(ws, variables, "$note_a$");
                    AdjustRowHeight(ws, variables, "$note_b$");
                    AdjustRowHeight(ws, variables, "$note_c$");
                    AdjustRowHeight(ws, variables, "$note_g$");
                    AdjustRowHeight(ws, variables, "$note_f$");
                }

                // 替換資料
                ImportDataIntoPlaceholder(ws, variables);
            }
            // 不留實體檔案，直接串流到前端
            using var ms = new MemoryStream();
            wb.Write(ms);
            return File(ms.ToArray(), FileMime.Excel);
        }
        // 一、資料模型
        public class PROJECT_INFO_MODEL
        {
            public string? REPORT_SERIAL_T1 { get; set; }
            public string? REPORT_SERIAL_T2 { get; set; }
            public string? REPORT_SERIAL_T3 { get; set; }
            public string? PROJECT_NAME { get; set; }
            public string? PROJECT_ID { get; set; }
            public string? PROJECT_OWN_CH { get; set; }
            public string? ORGANIZATION_ID { get; set; }
            public string? LOCATION { get; set; }
            public string? BES_ORD_NO { get; set; }
            public string? START_DATE { get; set; }
            public string? END_DATE { get; set; }
            public string? START_DATE_WKD { get; set; }
            public string? END_DATE_WKD { get; set; }
            public string? LAST_UPDATE_DATE { get; set; }
            public string? CALENDAR_DATE_T1 { get; set; }
            public string? CALENDAR_DATE_T2 { get; set; }
            public string? CALENDAR_DATE_T3 { get; set; }
            public string? CALENDAR_DATE { get; set; }
            public string? CALENDAR_DATE_WEEKDAY { get; set; }
            public string? WEATHER_AM { get; set; }
            public string? WEATHER_PM { get; set; }
            public string? WORK_HOUR { get; set; }
            public string? WORKDAY { get; set; }
            public string? EXP_PERCENT { get; set; }
            public string? EXP_PERCENT_T2 { get; set; }
            public string? EXP_PERCENT_T3 { get; set; }
            public string? ACT_PERCENT { get; set; }
            public string? ACT_PERCENT_T3 { get; set; }
            public string? ACT_SUM { get; set; }
            public string? ACT_SUM_T2 { get; set; }
            public string? ACT_SUM_T3 { get; set; }
            public string? DIFF_PERCENT { get; set; }
            public string? NOCAL_DAY { get; set; }
            public string? NOCAL_DAY_T2 { get; set; }
            public string? EXTEND_DAY { get; set; }
            public string? SPREAD_DAY { get; set; }
            public string? SPREAD_DAY_T2 { get; set; }
            public string? SPREAD_DAY_T3 { get; set; }
            public string? SPREAD_DAY_NUM { get; set; }
            public string? VDAY { get; set; }
            public string? VDAY_T2 { get; set; }
            public string? VDAY_SUB { get; set; }
            public string? VDAY_SUB_T2 { get; set; }
            public string? VDAY_SUB_NUM { get; set; }
            public string? TDAY { get; set; }
            public string? CONSTRUCTION_PERIOD { get; set; }
            public string? TDAY_T2 { get; set; }
            public string? SDAY { get; set; }
            public string? SDAY_T2 { get; set; }
            public string? SDAY_T3 { get; set; }
            public string? M49 { get; set; }
            public string? M50_ { get; set; }
            public string? M50 { get; set; }
            public string? M51_ { get; set; }
            public string? M51 { get; set; }
            public string? M52_ { get; set; }
            public string? M521_ { get; set; }
            public string? M52 { get; set; }
            public string? M53 { get; set; }
            public string? M54_ { get; set; }
            public string? M541_ { get; set; }
            public string? M54 { get; set; }
        }
        public class PROJECT_MMM_MODEL
        {
            public string? PROJECT_ID { get; set; }
            public string? RESOURCE_CLASS { get; set; }
            public string? RESOURCE_ID { get; set; }
            public string? SEQUENCE_NO { get; set; }
            public string? PBG_CODE { get; set; }
            public string? COST_CODE { get; set; }
            public string? NAME { get; set; }
            public string? UNIT { get; set; }
            public string? UNIT_PRICE { get; set; }
            public string? QUANTITY { get; set; }
            public string? TODAY_QTY { get; set; }
            public string? SUM_QTY { get; set; }
            public string? TODAY_QTY_2DECI { get; set; }
            public string? SUM_QTY_2DECI { get; set; }
            public string? FILTER_ID { get; set; }
        }
        public class PROJECT_NOTE_MODEL
        {
            public string? NOTES_A { get; set; }
            public string? NOTES_B { get; set; }
            public string? NOTES_C { get; set; }
            public string? NOTES_D { get; set; }
            public string? NOTES_E { get; set; }
            public string? NOTES_F { get; set; }
            public string? NOTES_G { get; set; }
            public string? NOTES { get; set; }
            public string? NOTES_A_NEWLINE_COUNT { get; set; }
            public string? NOTES_B_NEWLINE_COUNTT { get; set; }
            public string? NOTES_C_NEWLINE_COUNT { get; set; }
            public string? NOTES_D_NEWLINE_COUNT { get; set; }
            public string? NOTES_E_NEWLINE_COUNT { get; set; }
            public string? NOTES_F_NEWLINE_COUNT { get; set; }
            public string? NOTES_G_NEWLINE_COUNT { get; set; }
            public string? NOTES_NEWLINE_COUNT { get; set; }
        }
        public class PROJECT_CONSTRUCTION_OVERVIEW_MODEL
        {
            public string? OWNITEM_NO { get; set; }
            public string? NAME { get; set; }
            public string? UNIT { get; set; }
            public string? CONSTRUCTION_ITEM { get; set; }
            public string? OWN_CODE { get; set; }
            public string? OWN_CONTROL_ITEM { get; set; }
            public string? QUANTITY { get; set; }
            public string? NOW_EDR_QUANTITY { get; set; }
            public string? SUM_EDR_QUANTITY { get; set; }
            public string? QUANTITY_2DECI { get; set; }
            public string? NOW_EDR_QUANTITY_2DECI { get; set; }
            public string? SUM_EDR_QUANTITY_2DECI { get; set; }
            public string? QUANTITY_4DECI { get; set; }
            public string? NOW_EDR_QUANTITY_4DECI { get; set; }
            public string? SUM_EDR_QUANTITY_4DECI { get; set; }
        }
        public class PROJECT_MILESTONE_MODEL
        {
            public string? PROJECT_ID { get; set; }
            public string? MILESTONE_NO_SUB { get; set; }
            public string? MILESTONE_WK_DAY { get; set; }
        }
        public class FORM_DATA_MODEL
        {
            public int SelectedOption { get; set; }
            public DateTime CalendarDate { get; set; }
        }
        // 二、資料字典
        static Dictionary<string, object[]> ConvertListToDictionary<T>(List<T> list)
        {
            // 使用 DataTable 作為中介，用於方便的轉換操作
            DataTable dataTable = new();
            // 為 DataTable 新增欄位，以物件 T 的屬性名稱為欄位名稱
            typeof(T)
                .GetProperties()
                .ToList()
                .ForEach(property => dataTable.Columns.Add(property.Name));
            // 將物件清單的每個元素轉換為 DataRow
            list.ForEach(info =>
            {
                var row = dataTable.NewRow();
                // 將物件 T 的每個屬性值填入對應的 DataRow
                typeof(T)
                    .GetProperties()
                    .ToList()
                    .ForEach(property => row[property.Name] = property.GetValue(info));
                // 將填滿資料的 DataRow 加入 DataTable
                dataTable.Rows.Add(row);
            });
            // 將 DataTable 資料轉換為字典
            Dictionary<string, object[]> dict = new();
            // 遍歷 DataTable 的每個欄位
            foreach (DataColumn column in dataTable.Columns)
            {
                // 取得欄位名稱
                string columnName = column.ColumnName;
                // 取得欄位的所有資料值
                object[] columnData = new object[dataTable.Rows.Count];
                for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                {
                    columnData[rowIndex] = dataTable.Rows[rowIndex][columnName];
                }
                // 將欄位名稱轉為大寫，並將欄位名稱及對應的資料值附加到字典
                dict[columnName.ToUpper()] = columnData;
            }
            return dict;
        }
        static Dictionary<string, object[]> GetProjInfo(DBFunc db, int? projectId, DateTime calendarDate)
        {
            string calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
            var list = db.query<PROJECT_INFO_MODEL>("edr", @"
            SELECT
                '112A040C009-' || 
                    SUBSTR('0000' || TO_CHAR(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - PROJ.START_DATE + 1), -4) AS REPORT_SERIAL_T1
                , SUBSTR('0000' || TO_CHAR(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - PROJ.START_DATE + 1), -4) AS REPORT_SERIAL_T2
                , SUBSTR('0000' || TO_CHAR(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - PROJ.START_DATE + 1), -3) AS REPORT_SERIAL_T3
                , PROJ.PROJECT_NAME_SHORT_N AS PROJECT_NAME
                , PROJ.PROJECT_ID AS PROJECT_ID
                , NVL(PROJ.PROJECT_OWN_CH, ' ') AS PROJECT_OWN_CH
                , PROJ.ORGANIZATION_ID AS ORGANIZATION_ID
                , PROJ.LOCATION AS LOCATION
                , PROJ.BES_ORD_NO AS BES_ORD_NO
                , TO_CHAR(PROJ.START_DATE, 
                    'yyy""年""MM""月""dd""日""',
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS START_DATE
                , TO_CHAR(PROJ.END_DATE, 
                    'yyy""年""MM""月""dd""日""', 
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS END_DATE
                , TO_CHAR(PROJ.START_DATE, 
                    'yyy""年""MM""月""dd""日""(fmDay)',
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS START_DATE_WKD
                , TO_CHAR(PROJ.END_DATE, 
                    'yyy""年""MM""月""dd""日""(fmDay)', 
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS END_DATE_WKD
                , PROJ.LAST_UPDATE_DATE AS LAST_UPDATE_DATE
                , TO_CHAR(to_date(:calendarDateStr, 'yyyy/MM/dd'), 
                    'yyy"" 年 ""MM"" 月 ""dd"" 日 "" fmDay', 
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS CALENDAR_DATE_T1
                , TO_CHAR(to_date(:calendarDateStr, 'yyyy/MM/dd'), 
                    'yyy""年""MM""月""dd""日""(fmDay)', 
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS CALENDAR_DATE_T2
                , TO_CHAR(to_date(:calendarDateStr, 'yyyy/MM/dd'), 
                    'yyy""年""MM""月""dd""日"" fmDay', 
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS CALENDAR_DATE_T3
                , TO_CHAR(to_date(:calendarDateStr, 'yyyy/MM/dd'), 
                    'yyy""年""MM""月""dd""日""', 
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS CALENDAR_DATE
                , TO_CHAR(to_date(:calendarDateStr, 'yyyy/MM/dd'), 
                    'fmDay', 
                    'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS CALENDAR_DATE_WEEKDAY
                , NVL(NOTE.WEATHER_AM, ' ') AS WEATHER_AM                                                               -- 天氣(上午)
                , NVL(NOTE.WEATHER_PM, ' ') AS WEATHER_PM                                                               -- 天氣(下午)
                , NVL(NOTE.WORK_HOUR, -1) AS WORK_HOUR                                                                  -- 工作時數
                , NVL(NOTE.WORKDAY, -1 ) AS WORKDAY                                                                     -- 工作天(Y/N)
                , TO_CHAR(NVL(NOTE.EXP_PERCENT,-1) * 100.0, 'FM999,999.0000') || '%' AS EXP_PERCENT                     -- 預定進度(%) (至本日累計預定進度)
                , TO_CHAR(NVL(NOTE.EXP_PERCENT,-1) * 1.0, 'FM999,999.000') || '%' AS EXP_PERCENT_T2 
                , TO_CHAR(NVL(NOTE.EXP_PERCENT,-1) * 1.0, 'FM999,999.00')  AS EXP_PERCENT_T3
                , NVL(NOTE.ACT_PERCENT, -1) AS ACT_PERCENT                                                              -- 本日實際進度
                , TO_CHAR(NVL(NOTE.ACT_PERCENT, -1) * 1.0, 'FM999,990.00') AS ACT_PERCENT_T3
                , TO_CHAR(NVL(NOTE.ACT_SUM, -1) * 100.0, 'FM999,999.0000') || '%' AS ACT_SUM                            -- 實際進度(%) (至本日累計實際進度)
                , TO_CHAR(NVL(NOTE.ACT_SUM, -1) * 1.0, 'FM999,999.000') || '%' AS ACT_SUM_T2
                , TO_CHAR(NVL(NOTE.ACT_SUM,-1) * 1.0, 'FM999,999.00')  AS ACT_SUM_T3
                , TO_CHAR(TO_NUMBER(NOTE.ACT_SUM) - TO_NUMBER(NOTE.EXP_PERCENT), 'FM9990.000') || '%' AS DIFF_PERCENT   -- 超前或落後(%)
                , NVL(NOTE.NOCAL_DAY, -1) AS NOCAL_DAY                                                                  -- 免計工期(天
                , TO_CHAR(NVL(NOTE.NOCAL_DAY, 0) * 1.0 , 'FM999,990.0') || '日曆天'  AS NOCAL_DAY_T2                    -- 免計工期(日曆天)
                , NVL(NOTE.EXTEND_DAY, -1) AS EXTEND_DAY                                                                -- 展延工期(天)
                , TO_CHAR(NVL(PROJ.SPREAD_DAY, 0), 'FM999,999,999,999') || '天' AS SPREAD_DAY                           -- 展延天數
                , TO_CHAR(NVL(PROJ.SPREAD_DAY, 0), 'FM999,999,999,999') || '日曆天' AS SPREAD_DAY_T2                    -- 展延天數
                , TO_CHAR(NVL(PROJ.SPREAD_DAY, 0), 'FM999,999,999,999')  AS SPREAD_DAY_NUM                              -- 展延天數(純數字)
                , TO_CHAR(NVL(PROJ.PROJ_SPREAD_DATE - 
                    TO_DATE(:calendarDateStr, 'yyyy/MM/dd'), 0), 'FM999,999,999,999') || '天' AS VDAY                   -- 剩餘工期
                , TO_CHAR(NVL(PROJ.PROJ_SPREAD_DATE - 
                    TO_DATE(:calendarDateStr, 'yyyy/MM/dd'), 0), 'FM999,999,999,999') || '日曆天' AS VDAY_T2            -- 剩餘工期
                , TO_CHAR((PROJ.ORIGINAL_DAY - 
                    TO_NUMBER(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - 
                    PROJ.START_DATE + 1)), 'FM999,999,999,999') || '天' AS VDAY_SUB                                     -- 剩餘工期 核定減累計
                , TO_CHAR((PROJ.ORIGINAL_DAY - 
                    TO_NUMBER(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - 
                    PROJ.START_DATE + 1)), 'FM999,999,990.0') || '日曆天' AS VDAY_SUB_T2                                     -- 剩餘工期(日曆天)
                , TO_CHAR((PROJ.ORIGINAL_DAY - 
                    TO_NUMBER(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - 
                    PROJ.START_DATE + 1)), 'FM999,999,999,999')  AS VDAY_SUB_NUM           -- 剩餘工期(純數字)
                , TO_CHAR(NVL(PROJ.ORIGINAL_DAY, 0), 'FM999,999,999,999') || '天' AS TDAY                               -- 核定工期
                , TO_CHAR(NVL(PROJ.ORIGINAL_DAY, 0), 'FM999,999,999,999') || '日曆天' AS TDAY_T2                        -- 核定工期
                , TO_NUMBER(NVL(PROJ.ORIGINAL_DAY, 0)) + TO_NUMBER(NVL(PROJ.SPREAD_DAY, 0)) || '天' AS CONSTRUCTION_PERIOD  --工期
                , TO_CHAR(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - 
                    PROJ.START_DATE + 1, 'FM999,999,999,999') || '天' AS SDAY                                           -- 累計工期
                , TO_CHAR(TO_DATE(:calendarDateStr, 'yyyy/MM/dd') - 
                    PROJ.START_DATE + 1, 'FM999,999,999,999') || '日曆天' AS SDAY_T2                                    -- 累計工期
                , NOTE.M49                                                                                              -- 以下全部由紙本勾選
                , NOTE.M50 AS M50_
                , CASE
                    WHEN NOTE.M50 = 'Y' 
                        THEN '▓有 □無'
                    WHEN NOTE.M50 = 'N' 
                        THEN '□有 ▓無'
                    ELSE     '□有 □無'
                    END AS M50
                , NOTE.M51 AS M51_
                , CASE
                    WHEN NOTE.M51 = 'Y' 
                        THEN '1.實施勤前教育(含工地預防災變及危害告知)：▓有 □無'
                    WHEN NOTE.M51 = 'N' 
                        THEN '1.實施勤前教育(含工地預防災變及危害告知)：□有 ▓無'
                    ELSE     '1.實施勤前教育(含工地預防災變及危害告知)：□有 □無'
                    END AS M51
                , NOTE.M52 AS M52_
                , NOTE.M521 AS M521_
                , CASE
                    WHEN NOTE.M52 = '1' AND NOTE.M521 = 'Y' 
                        THEN '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：▓有 □無 ▓無新進勞工'
                    WHEN NOTE.M52 = '1' AND (NOTE.M521 = 'N' OR NOTE.M521 IS NULL)
                        THEN '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：▓有 □無 □無新進勞工'
                    WHEN NOTE.M52 = '2' AND NOTE.M521 = 'Y' 
                        THEN '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：□有 ▓無 ▓無新進勞工'
                    WHEN NOTE.M52 = '2' AND (NOTE.M521 = 'N' OR NOTE.M521 IS NULL)
                        THEN '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：□有 ▓無 □無新進勞工'   
                    WHEN NOTE.M52 IS NULL AND NOTE.M521 = 'Y' 
                        THEN '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：□有 □無 ▓無新進勞工'
                    WHEN NOTE.M52 IS NULL AND (NOTE.M521 = 'N' OR NOTE.M521 IS NULL)
                        THEN '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：□有 □無 □無新進勞工' 
                    ELSE     '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：□有 □無 □無新進勞工'
                    END AS M52
                , NOTE.M53 AS M53_
                , CASE
                    WHEN NOTE.M53 = 'Y' 
                        THEN '3.檢查勞工個人防護具：▓有 □無'
                    WHEN NOTE.M53 = 'N' 
                        THEN '3.檢查勞工個人防護具：□有 ▓無'
                    ELSE     '3.檢查勞工個人防護具：□有 □無'
                    END AS M53
                , NOTE.M54 AS M54_
                , NOTE.M541 AS M541_
                , CASE
                    WHEN NOTE.M54 = '1' AND NOTE.M541 = 'Y' 
                        THEN '1.自主檢查工地實際進用移工與核准名冊相符：▓有 □無 ▓無外籍移工'
                    WHEN NOTE.M54 = '1' AND (NOTE.M541 = 'N' OR NOTE.M541 IS NULL)
                        THEN '1.自主檢查工地實際進用移工與核准名冊相符：▓有 □無 □無外籍移工'
                    WHEN NOTE.M54 = '2' AND NOTE.M541 = 'Y'
                        THEN '1.自主檢查工地實際進用移工與核准名冊相符：□有 ▓無 ▓無外籍移工'
                    WHEN NOTE.M54 = '2' AND (NOTE.M541 = 'N' OR NOTE.M541 IS NULL)
                        THEN '1.自主檢查工地實際進用移工與核准名冊相符：□有 ▓無 □無外籍移工'
                    WHEN NOTE.M54 IS NULL AND NOTE.M541 = 'Y'
                        THEN '1.自主檢查工地實際進用移工與核准名冊相符：□有 □無 ▓無外籍移工'  
                    WHEN NOTE.M54 IS NULL AND (NOTE.M541 = 'N' OR NOTE.M541 IS NULL)
                        THEN '1.自主檢查工地實際進用移工與核准名冊相符：□有 □無 □無外籍移工'   
                    ELSE     '1.自主檢查工地實際進用移工與核准名冊相符：□有 □無 □無外籍移工'
                    END AS M54
            FROM VSUSER.VS_PROJECTS PROJ
            LEFT JOIN (SELECT * FROM VSUSER.BES_EDR_NOTES WHERE CALENDAR_DATE = TO_DATE(:calendarDateStr, 'yyyy/MM/dd')) NOTE
            ON PROJ.PROJECT_ID = NOTE.PROJECT_ID
            WHERE PROJ.PROJECT_ID = :projectId
                ", new
            {
                projectId,
                calendarDateStr
            }).ToList();
            if (list.Count == 0)
            {
                list = db.query<PROJECT_INFO_MODEL>("edr", @"
                    SELECT
                          ' ' AS PROJECT_NAME
                        , ' ' AS PROJECT_ID
                        , ' ' AS PROJECT_OWN_CH
                        , ' ' AS ORGANIZATION_ID
                        , ' ' AS LOCATION
                        , ' ' AS BES_ORD_NO
                        , ' ' AS START_DATE
                        , ' ' AS END_DATE
                        , ' ' AS LAST_UPDATE_DATE
                        , TO_CHAR(to_date(:calendarDateStr,'yyyy/MM/dd'), 
                            'yyy""年""MM""月""dd""日""(fmDay)', 
                            'NLS_CALENDAR= ''ROC Official''NLS_DATE_LANGUAGE=''TRADITIONAL CHINESE''') AS CALENDAR_DATE
                        , ' ' AS WEATHER_AM
                        , ' ' AS WEATHER_PM
                        , -1 AS WORK_HOUR
                        , -1 AS WORKDAY
                        , ' ' AS EXP_PERCENT
                        , -1 AS ACT_PERCENT                                     
                        , ' ' AS ACT_SUM              
                        , -1 AS NOCAL_DAY                 
                        , -1 AS EXTEND_DAY                         
                        , -1 AS SPREAD_DAY                           
                        , -1 AS VDAY     
                        , -1 AS VDAY_SUB   
                        , -1 AS TDAY                               
                        , -1 AS SDAY          
                        , ' ' AS M49
                        , ' ' AS M50_
                        , '□有 □無' AS M50
                        , ' ' AS M51_
                        , '1.實施勤前教育(含工地預防災變及危害告知)：□有 □無' AS M51
                        , ' ' AS M52_
                        , ' ' AS M521_
                        , '2.確認新進勞工是否提報勞工保險(或其他商業保險)資料及安全衛生教育訓練紀錄：□有 □無 □無新進勞工' AS M52
                        , '3.檢查勞工個人防護具：□有 □無' AS M53
                        , ' ' AS M54_
                        , ' ' AS M541_
                        , '1.自主檢查工地實際進用移工與核准名冊相符：□有 □無 □無外籍移工' AS M54
                    FROM dual
                    ", new
                {
                    calendarDateStr
                }).ToList();
            }
            return ConvertListToDictionary(list);
        }
        static Dictionary<string, object[]> GetProjConsOverview(DBFunc db, int? projectId, DateTime calendarDate, int? myuserId, string myState, bool specific = false)
        {
            string query2 = "vsuser.bes_edr_rep1";
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
                    , NVL(TRIM(UNIT), ' ') AS UNIT
                    , OWNITEM_NO || '、' || NAME AS CONSTRUCTION_ITEM
                    , OWN_CODE
                    , OWN_CONTROL_ITEM
                    , TO_CHAR(NVL(QUANTITY, NULL), 'FM999,999,999,999') AS QUANTITY
                    , TO_CHAR(NVL(NOW_EDR_QUANTITY, NULL), 'FM999,999,999,999') AS NOW_EDR_QUANTITY
                    , TO_CHAR(NVL(SUM_EDR_QUANTITY, NULL), 'FM999,999,999,999') AS SUM_EDR_QUANTITY

                    , CASE WHEN QUANTITY = 0 THEN '-' 
                        ELSE TO_CHAR(NVL(QUANTITY, NULL), 'FM999,999,999,990.00') 
                        END AS QUANTITY_2DECI
                    , CASE WHEN NOW_EDR_QUANTITY = 0 THEN '-'
                        ELSE TO_CHAR(NVL(NOW_EDR_QUANTITY, NULL), 'FM999,999,999,990.00') 
                        END AS NOW_EDR_QUANTITY_2DECI
                    , CASE WHEN SUM_EDR_QUANTITY = 0 THEN '-'
                        ELSE TO_CHAR(NVL(SUM_EDR_QUANTITY, NULL), 'FM999,999,999,990.00') 
                        END AS SUM_EDR_QUANTITY_2DECI

                    , CASE WHEN QUANTITY = 0 THEN '-' 
                        ELSE TO_CHAR(NVL(QUANTITY, NULL), 'FM999,999,999,990.0000') 
                        END AS QUANTITY_4DECI
                    , CASE WHEN NOW_EDR_QUANTITY = 0 THEN '-'
                        ELSE TO_CHAR(NVL(NOW_EDR_QUANTITY, NULL), 'FM999,999,999,990.0000') 
                        END AS NOW_EDR_QUANTITY_4DECI
                    , CASE WHEN SUM_EDR_QUANTITY = 0 THEN '-'
                        ELSE TO_CHAR(NVL(SUM_EDR_QUANTITY, NULL), 'FM999,999,999,990.0000') 
                        END AS SUM_EDR_QUANTITY_4DECI
                FROM (
                SELECT
                      OWNITEM_NO
                    , NAME
                    , UNIT
                    , CASE WHEN OWN_CONTROL_ITEM = 'S' THEN 0 ELSE QUANTITY END AS QUANTITY
                    , CASE WHEN OWN_CONTROL_ITEM = 'S' THEN 0 ELSE NOW_EDR_QUANTITY END AS NOW_EDR_QUANTITY
                    , CASE WHEN OWN_CONTROL_ITEM = 'S' THEN 0 ELSE SUM_EDR_QUANTITY END AS SUM_EDR_QUANTITY
                    , OWN_CODE
                    , PROJECT_ID
                    , CREATED_BY
                    , OWN_CONTROL_ITEM
                FROM VSUSER.BES_EDR_TEMP_QUANTITY
                )
                WHERE PROJECT_ID = {projectId}
                AND CREATED_BY = {myuserId}
                --AND ROWNUM <= 10
                ORDER BY OWN_CODE
                {query3_}
                ";
            var list = db.query<PROJECT_CONSTRUCTION_OVERVIEW_MODEL>("edr", query3).ToList();
            return ConvertListToDictionary(list);
        }
        static Dictionary<string, object[]> GetProjManMachineMaterial(DBFunc db, int? projectId, DateTime calendarDate, string state, string resource)
        {
            string calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
            string resourceClass;
            if (resource == "man")
            {
                resourceClass = "('2')";
            }
            else if (resource == "machine")
            {
                resourceClass = "('3', '5')";
            }
            else if (resource == "material")
            {
                resourceClass = "('4')";
            }
            else
            {
                resourceClass = "";
            }
            string filterId;
            if (state == "1")
            {
                filterId ="('1', '2')";
            }
            else if (state == "2")
            {
                filterId ="('2')";
            }
            else if (state == "3")
            {
                filterId ="('1', '2', '3')";
            }
            else
            {
                filterId = "";
            }
            var list = db.query<PROJECT_MMM_MODEL>("edr", $@"
                SELECT
                    PROJECT_ID
                    , RESOURCE_CLASS
                    , RESOURCE_ID
                    , SEQUENCE_NO
                    , PBG_CODE
                    , COST_CODE
                    , NAME
                    , UNIT
                    , UNIT_PRICE
                    , TO_CHAR(QUANTITY, 'FM999,999,999,999') AS QUANTITY
                    , TODAY_QTY
                    , SUM_QTY
                    , FILTER_ID --1:所有,2:本日,3:歷史所有
                FROM (
                SELECT 
                    a.PROJECT_ID
                    , a.RESOURCE_CLASS
                    , a.RESOURCE_ID
                    , a.SEQUENCE_NO
                    , a.PBG_CODE
                    , a.COST_CODE
                    , TRIM(a.NAME) AS NAME
                    , TRIM(b.NAME) AS UNIT
                    , a.UNIT_PRICE
                    , a.QUANTITY
                    , NVL(c.TODAY_QTY, '0') AS TODAY_QTY
                    , NVL(c.SUM_QTY, '0') AS SUM_QTY
                    , NVL(FILTER_ID, '3') AS FILTER_ID --1:所有,2:本日,3:歷史所有
                FROM vsuser.BES_PROJECT_RESOURCES a
                LEFT JOIN (SELECT UOM_ID, NAME FROM vsuser.VS_UOMS) b
                ON a.UOM_ID=b.UOM_ID
                LEFT JOIN(
                SELECT
                    PROJECT_ID
                    , RESOURCE_CLASS
                    , RESOURCE_ID
                    , TODAY_QTY
                    , SUM_QTY
                    , CASE WHEN TODAY_QTY='0' THEN '1' ELSE '2' END AS FILTER_ID
                FROM (
                    SELECT       
                        PROJECT_ID
                        , RESOURCE_CLASS
                        , RESOURCE_ID
                        , TO_CHAR(SUM(CASE WHEN DATA_DATE = TO_DATE(:calendarDateStr, 'yyyy/MM/dd') 
                            THEN NVL(QUANTITY, 0) ELSE 0 END), 'FM999,999,999,999') AS TODAY_QTY
                        , TO_CHAR(SUM(CASE WHEN DATA_DATE <= TO_DATE(:calendarDateStr, 'yyyy/MM/dd') 
                            THEN NVL(QUANTITY, 0) ELSE 0 END), 'FM999,999,999,999') AS SUM_QTY
                        , TO_CHAR(SUM(CASE WHEN DATA_DATE = TO_DATE(:calendarDateStr, 'yyyy/MM/dd') 
                            THEN NVL(QUANTITY, 0) ELSE 0 END), 'FM999,999,999,990.00') AS TODAY_QTY_2DECI
                        , TO_CHAR(SUM(CASE WHEN DATA_DATE <= TO_DATE(:calendarDateStr, 'yyyy/MM/dd') 
                            THEN NVL(QUANTITY, 0) ELSE 0 END), 'FM999,999,999,990.00') AS SUM_QTY_2DECI
                    FROM vsuser.bes_edr_resqty_v 
                    GROUP BY PROJECT_ID, RESOURCE_CLASS, RESOURCE_ID
                    )
                )c
                ON a.PROJECT_ID=c.PROJECT_ID 
                AND a.RESOURCE_CLASS=c.RESOURCE_CLASS 
                AND a.RESOURCE_ID=c.RESOURCE_ID
                WHERE a.PROJECT_ID = :projectId
                )
                WHERE RESOURCE_CLASS IN {resourceClass} --2:人,3,5:機,4:料
                AND FILTER_ID IN {filterId}
                ORDER BY PROJECT_ID, RESOURCE_CLASS, SEQUENCE_NO
                ", new
            {
                projectId,
                calendarDateStr
            }).ToList();
            return ConvertListToDictionary(list);
        }
        static Dictionary<string, object[]> GetProjNote(DBFunc db, int? projectId, DateTime calendarDate)
        {
            string calendarDateStr = calendarDate.ToString("yyyy/MM/dd");
            var list = db.query<PROJECT_NOTE_MODEL>("edr", $@"
                SELECT
                    project_id
                  , NVL(notes_a, ' ') AS NOTES_A                                      --工地記事 (六、施工取樣試驗紀錄：)
                  , NVL(notes_b, ' ') AS NOTES_B                                      --工地記事 (七、通知協力廠商辦理事項：)
                  , NVL(notes_c, ' ') AS NOTES_C                                      --工地記事 (八、重要事項記錄：)
                  , NVL(notes_d, ' ') AS NOTES_D 
                  , NVL(notes_e, ' ') AS NOTES_E 
                  , NVL(notes_f, ' ') AS NOTES_F 
                  , NVL(notes_g, ' ') AS NOTES_G 
                  , NVL(notes, ' ') AS NOTES                                          --其他事項
                  , REGEXP_COUNT(NVL(notes_a, ' '), '\n') AS NOTES_A_NEWLINE_COUNT
                  , REGEXP_COUNT(NVL(notes_b, ' '), '\n') AS NOTES_B_NEWLINE_COUNT
                  , REGEXP_COUNT(NVL(notes_c, ' '), '\n') AS NOTES_C_NEWLINE_COUNT
                  , REGEXP_COUNT(NVL(notes_d, ' '), '\n') AS NOTES_D_NEWLINE_COUNT
                  , REGEXP_COUNT(NVL(notes_e, ' '), '\n') AS NOTES_E_NEWLINE_COUNT
                  , REGEXP_COUNT(NVL(notes_f, ' '), '\n') AS NOTES_F_NEWLINE_COUNT
                  , REGEXP_COUNT(NVL(notes_g, ' '), '\n') AS NOTES_G_NEWLINE_COUNT
                  , REGEXP_COUNT(NVL(notes, ' '), '\n') AS NOTES_NEWLINE_COUNT
                FROM vsuser.bes_edr_notes
                WHERE project_id = :projectId
                AND calendar_date = to_date(:calendarDateStr, 'yyyy/MM/dd')
                ", new
            {
                projectId,
                calendarDateStr
            }).ToList();
            if (list.Count == 0)
            {
                list = db.query<PROJECT_NOTE_MODEL>("edr", $@"
                    SELECT
                          :projectId AS project_id
                        , ' ' AS NOTES_A                                      --工地記事 (六、施工取樣試驗紀錄：)
                        , ' ' AS NOTES_B                                      --工地記事 (七、通知協力廠商辦理事項：)
                        , ' ' AS NOTES_C                                      --工地記事 (八、重要事項記錄：)
                        , ' ' AS NOTES_D
                        , ' ' AS NOTES_E
                        , ' ' AS NOTES_F
                        , ' ' AS NOTES_G
                        , ' ' AS NOTES                                        --其他事項
                        , 0 AS NOTES_A_NEWLINE_COUNT
                        , 0 AS NOTES_B_NEWLINE_COUNT
                        , 0 AS NOTES_C_NEWLINE_COUNT
                        , 0 AS NOTES_D_NEWLINE_COUNT
                        , 0 AS NOTES_E_NEWLINE_COUNT
                        , 0 AS NOTES_F_NEWLINE_COUNT
                        , 0 AS NOTES_G_NEWLINE_COUNT
                        , 0 AS NOTES_NEWLINE_COUNT
                    FROM dual
                    ", new
                {
                    projectId,
                    calendarDateStr
                }).ToList();
            }
            return ConvertListToDictionary(list);
        }
        static Dictionary<string, object[]> GetProjMilestone(DBFunc db, int? projectId)
        {
            var list = db.query<PROJECT_MILESTONE_MODEL>("edr", $@"
                SELECT 
                      PROJECT_ID
                    , REPLACE(REPLACE(REPLACE(MILESTONE_NO_SUB, '第', ''), '期', ''), ' ', '') AS MILESTONE_NO_SUB
                    , MILESTONE_DATE_START		--開工日期，原契約開工日期
                    , MILESTONE_DATE			--預定完工日期，里程碑完工日期
                    , MILESTONE_WK_DAY			--核定工期，工期(日曆天/工作天)
                    , MILESTONE_M_P				--核定工期，佔契約金額百分比
                    , POSTPONE_E_DATE			--展延後完工日期
                    , POSTPONE_DAY				--展延天數(日曆天/工作天)
                    , MILESTONE_ITEM			--完成交付項目
                    , OVERDUE_FINE				--逾期罰款
                    , OVERDUE_FINE_D			--預估逾期罰款(天數/金額萬元)
                    , REMARK					--備註
                FROM vsuser.bes_proj_milestone
                WHERE PROJECT_ID = :projectId
                ORDER BY MILESTONE_NO_SUB
                ", new
            {
                projectId
            }).ToList();
            return ConvertListToDictionary(list);
        }
        // 三、FUN
        // 將cellvalue寫進placeholder位置
        static void WriteCellValueByPlaceholder(ISheet sheet, string placeholder, object cellValue)
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
        // 將cellvalue寫進cell位置，index從1開始
        static void WriteCellValueByIndex(ISheet sheet, int rowIndex, int colIndex, object cellValue)
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
        // 返回placeholder位置，index[0,0]表示第一筆placeholder的row位置，index[0,1]表示第一筆placeholder的column位置，index從1開始
        static int[,] GetCellIndexByPlaceholder(ISheet sheet, string placeholder)
        {
            List<int[]> indexesList = new();
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
        // 移除rowIndexStart到rowIndexEnd範圍內的rows，同時上移表格，index從1開始
        static void RemoveRangeRows(ISheet sheet, int rowIndexStart, int rowIndexEnd)
        {
            rowIndexStart--;
            rowIndexEnd--;
            int rowCount = rowIndexEnd - rowIndexStart + 1;
            for (int i = 0; i < rowCount; i++)
            {
                IRow row = sheet.GetRow(rowIndexStart);
                if (row != null)
                {
                    sheet.RemoveRow(row);
                    sheet.ShiftRows(rowIndexStart + 1, sheet.LastRowNum, -1);
                }
            }
        }
        // 將rowIndexStart的row內容往下複製repeatTimes次
        static void CopyRowDown(ISheet sheet, int rowIndexStart, int repeatTimes)
        {
            rowIndexStart--;
            for (int i = 0; i < repeatTimes; i++)
            {
                IRow sourceRow = sheet.GetRow(rowIndexStart);
                if (sourceRow != null)
                {
                    int lastRowNum = sheet.LastRowNum;
                    sheet.ShiftRows(rowIndexStart + 1, lastRowNum, 1, true, false);
                    IRow newRow = sheet.CreateRow(rowIndexStart + 1);
                    newRow.Height = sourceRow.Height;
                    for (int m = 0; m < sheet.NumMergedRegions; m++)
                    {
                        var mergedRegion = sheet.GetMergedRegion(m);
                        if (mergedRegion.FirstRow == rowIndexStart && mergedRegion.LastRow == rowIndexStart)
                        {
                            int newFirstRow = rowIndexStart + 1;
                            int newLastRow = rowIndexStart + 1 + (mergedRegion.LastRow - mergedRegion.FirstRow);
                            int newFirstCol = mergedRegion.FirstColumn;
                            int newLastCol = mergedRegion.LastColumn;
                            sheet.AddMergedRegion(new CellRangeAddress(newFirstRow, newLastRow, newFirstCol, newLastCol));
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
                    rowIndexStart++;
                }
            }
        }
        // 依範圍合併儲存格，index從1開始
        static void MergeCells(ISheet sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet), "Sheet cannot be null");
            }
            if (startRow < 0 || startColumn < 0 || endRow < 0 || endColumn < 0)
            {
                throw new ArgumentException("Row and column indexes cannot be negative");
            }
            sheet.AddMergedRegion(new CellRangeAddress(startRow - 1, endRow - 1, startColumn - 1, endColumn - 1));
        }
        // 回傳出現次數1次(times = 1)的placeholder或出現次數大於1次(times = 2)的placeholder
        static string[] FindPlaceholderAppearingTimes(ISheet sheet, int times)
        {
            Dictionary<string, int> textCounts = new();
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
                            Regex regex = new(@"\$\w+\$");
                            MatchCollection matches = regex.Matches(cellValue);
                            foreach (System.Text.RegularExpressions.Match match in matches.Cast<System.Text.RegularExpressions.Match>())
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
        // 取代資料所有出現1次的placeholder
        static void ReplaceAllSinglePlaceholder(ISheet sheet, Dictionary<string, object[]> data)
        {
            string[] uniquePlaceholder = FindPlaceholderAppearingTimes(sheet, 1);
            foreach (string placeholder in uniquePlaceholder)
            {
                if (data.ContainsKey(placeholder))
                {
                    if (data[placeholder].Length == 1)
                    {
                        WriteCellValueByPlaceholder(sheet, placeholder, data[placeholder][0]);
                    }
                    else if (data[placeholder].Length == 0)
                    {
                        WriteCellValueByPlaceholder(sheet, placeholder, "");
                    }
                }
            }
        }
        // 取代資料出現1次以上的placeholder
        static void ReplaceMultiplePlaceholder(ISheet sheet, string placeholder, object[] data)
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
        // 取代資料所有出現1次以上的placeholder
        static void ReplaceAllMultipleVariable(ISheet sheet, Dictionary<string, object[]> data)
        {
            string[] uniqueVariables = FindPlaceholderAppearingTimes(sheet, 2);
            foreach (string placeholder in uniqueVariables)
            {
                if (data.ContainsKey(placeholder))
                {
                    ReplaceMultiplePlaceholder(sheet, placeholder, (object[])data[placeholder]);
                }
            }
        }
        // 依placeholder1(或與placeholder2比較)，依資料筆數調整row數
        static void AdjustRowNums(ISheet sheet, Dictionary<string, object[]> data, string placeholder1, string? placeholder2 = null, string? placeholder3 = null)
        {
            string placeholder;
            if (!string.IsNullOrEmpty(placeholder2) && !string.IsNullOrEmpty(placeholder3))
            {
                if (data[placeholder1].GetLength(0) >= data[placeholder2].GetLength(0) && data[placeholder1].GetLength(0) >= data[placeholder3].GetLength(0))
                {
                    placeholder = placeholder1;
                }
                else if (data[placeholder2].GetLength(0) >= data[placeholder1].GetLength(0) && data[placeholder2].GetLength(0) >= data[placeholder3].GetLength(0))
                {
                    placeholder = placeholder2;
                }
                else
                {
                    placeholder = placeholder3;
                }
            }
            else if (!string.IsNullOrEmpty(placeholder2) && string.IsNullOrEmpty(placeholder3))
            {
                if (data[placeholder1].GetLength(0) >= data[placeholder2].GetLength(0))
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
            object[] d = (object[])data[placeholder];
            var x = indexesArray.GetLength(0);
            int rowNumDiff = d.GetLength(0) - indexesArray.GetLength(0);
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
        // 取代資料所有placeholder
        static void ImportDataIntoPlaceholder(ISheet sheet, Dictionary<string, object[]> data)
        {
            ReplaceAllSinglePlaceholder(sheet, data);
            ReplaceAllMultipleVariable(sheet, data);
            ReplacePlaceholdersWithEmpty(sheet);
        }
        // 水平合併mergeColRangeArray，依placeholder所在row位置重複執行
        static void HorizontalMergeCols(ISheet sheet, string placeholder, int[,] mergeColRangeArray)
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
        // 垂直合併placeholder所在的row，依mergeCol重複執行
        static void VerticalMergeRows(ISheet sheet, string placeholder, int[] mergeCol)
        {
            int[,] Index = GetCellIndexByPlaceholder(sheet, placeholder);
            foreach (int i in mergeCol)
            {
                MergeCells(sheet, Index[0, 0], i, Index[Index.GetLength(0) - 1, 0], i);
            }
        }
        // 將未被資料取代的placeholder清除
        static void ReplacePlaceholdersWithEmpty(ISheet sheet)
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
        // 回傳index所在位置合併儲存格的水平起點與終點
        static List<int> GetMergedRangeHor(ISheet sheet, int rowIndex, int colIndex)
        {
            rowIndex--;
            colIndex--;
            List<int> result = new();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress mergedRegion = sheet.GetMergedRegion(i);
                if (mergedRegion.IsInRange(rowIndex, colIndex))
                {
                    result.Add(mergedRegion.FirstColumn + 1);
                    result.Add(mergedRegion.LastColumn + 1);
                    break;
                }
            }
            if (result.Count == 0)
            {
                result.Add(colIndex + 1);
                result.Add(colIndex + 1);
            }
            return result;
        }
        // 依placeholder1(或與placeholder2比較)，依資料字數調整row高
        static void AdjustRowHeight(ISheet sheet, Dictionary<string, object[]> variables, string placeholder1, string? placeholder2 = null, string? placeholder3 = null)
        {
            string placeholder;
            if (!string.IsNullOrEmpty(placeholder1) && !string.IsNullOrEmpty(placeholder2) && !string.IsNullOrEmpty(placeholder3))
            {
                int[,] indexesArray1 = GetCellIndexByPlaceholder(sheet, placeholder1);
                int[,] indexesArray2 = GetCellIndexByPlaceholder(sheet, placeholder2);
                int[,] indexesArray3 = GetCellIndexByPlaceholder(sheet, placeholder3);
                object[] data1 = variables[placeholder1];
                object[] data2 = variables[placeholder2];
                object[] data3 = variables[placeholder3];
                var MergedRangeHor1 = GetMergedRangeHor(sheet, indexesArray1[0, 0], indexesArray1[0, 1]);
                var MergedRangeHor2 = GetMergedRangeHor(sheet, indexesArray2[0, 0], indexesArray2[0, 1]);
                var MergedRangeHor3 = GetMergedRangeHor(sheet, indexesArray3[0, 0], indexesArray3[0, 1]);
                int colIndexMergeStart1 = MergedRangeHor1[0];
                int colIndexMergeEnd1 = MergedRangeHor1[1];
                int colIndexMergeStart2 = MergedRangeHor2[0];
                int colIndexMergeEnd2 = MergedRangeHor2[1];
                int colIndexMergeStart3 = MergedRangeHor3[0];
                int colIndexMergeEnd3 = MergedRangeHor3[1];
                double totalCharNum1 = 0.0;
                double totalCharNum2 = 0.0;
                double totalCharNum3 = 0.0;
                for (int i = colIndexMergeStart1; i <= colIndexMergeEnd1; i++)
                {
                    totalCharNum1 += sheet.GetColumnWidth(i - 1) / 256 / 2.0625;
                }
                for (int i = colIndexMergeStart2; i <= colIndexMergeEnd2; i++)
                {
                    totalCharNum2 += sheet.GetColumnWidth(i - 1) / 256 / 2.0625;
                }
                for (int i = colIndexMergeStart3; i <= colIndexMergeEnd3; i++)
                {
                    totalCharNum3 += sheet.GetColumnWidth(i - 1) / 256 / 2.0625;
                }

                List<double> rowHeight1 = new();
                for (int i = 0; i < data1.Length; i++)
                {
                    string cellValue1 = (string)data1[i];
                    string[] lines1 = cellValue1.Split('\n');
                    double rowHeightCounts1 = 0;
                    foreach (var line in lines1)
                    {
                        rowHeightCounts1 += Math.Ceiling((double)line.Length / totalCharNum1);
                    }
                    rowHeight1.Add(rowHeightCounts1 * 20 * 17);
                }
                List<double> rowHeight2 = new();
                for (int i = 0; i < data2.Length; i++)
                {
                    string cellValue2 = (string)data2[i];
                    string[] lines2 = cellValue2.Split('\n');
                    double rowHeightCounts2 = 0;
                    foreach (var line in lines2)
                    {
                        rowHeightCounts2 += Math.Ceiling((double)line.Length / totalCharNum2);
                    }
                    rowHeight2.Add(rowHeightCounts2 * 20 * 17);
                }
                List<double> rowHeight3 = new();
                for (int i = 0; i < data3.Length; i++)
                {
                    string cellValue3 = (string)data3[i];
                    string[] lines3 = cellValue3.Split('\n');
                    double rowHeightCounts3 = 0;
                    foreach (var line in lines3)
                    {
                        rowHeightCounts3 += Math.Ceiling((double)line.Length / totalCharNum3);
                    }
                    rowHeight3.Add(rowHeightCounts3 * 20 * 17);
                }
                double[] rowHeight = Enumerable.Range(0, Math.Max(Math.Max(rowHeight1.Count, rowHeight2.Count), rowHeight3.Count))
                    .Select(i => Math.Max(Math.Max(rowHeight1.ElementAtOrDefault(i), rowHeight2.ElementAtOrDefault(i)), rowHeight3.ElementAtOrDefault(i)))
                    .ToArray();
                int startRow = indexesArray1[0, 0];
                for (int i = 0; i < rowHeight.Length; i++)
                {
                    var row = sheet.GetRow(startRow - 1 + i);
                    row.Height = (short)rowHeight[i];
                }
            }
            else if (!string.IsNullOrEmpty(placeholder1) && !string.IsNullOrEmpty(placeholder2) && string.IsNullOrEmpty(placeholder3))
            {
                int[,] indexesArray1 = GetCellIndexByPlaceholder(sheet, placeholder1);
                int[,] indexesArray2 = GetCellIndexByPlaceholder(sheet, placeholder2);
                object[] data1 = variables[placeholder1];
                object[] data2 = variables[placeholder2];
                var MergedRangeHor1 = GetMergedRangeHor(sheet, indexesArray1[0, 0], indexesArray1[0, 1]);
                var MergedRangeHor2 = GetMergedRangeHor(sheet, indexesArray2[0, 0], indexesArray2[0, 1]);
                int colIndexMergeStart1 = MergedRangeHor1[0];
                int colIndexMergeEnd1 = MergedRangeHor1[1];
                int colIndexMergeStart2 = MergedRangeHor2[0];
                int colIndexMergeEnd2 = MergedRangeHor2[1];
                int rowMin = Math.Min(data1.GetLength(0), data2.GetLength(0));
                int rowMax = Math.Max(data1.GetLength(0), data2.GetLength(0));
                int rowDiff = data1.GetLength(0) - data2.GetLength(0);
                double totalCharNum1 = 0.0;
                double totalCharNum2 = 0.0;
                if (rowMin != 0)
                {
                    for (int i = colIndexMergeStart1; i <= colIndexMergeEnd1; i++)
                    {
                        totalCharNum1 += sheet.GetColumnWidth(i - 1) / 256 / 2.0625;
                    }
                    for (int i = colIndexMergeStart2; i <= colIndexMergeEnd2; i++)
                    {
                        totalCharNum2 += sheet.GetColumnWidth(i - 1) / 256 / 2.0625;
                    }
                    for (int i = 0; i < rowMin; i++)
                    {
                        string cellValue1 = (string)data1[i];
                        string cellValue2 = (string)data2[i];
                        string[] lines1 = cellValue1.Split('\n');
                        string[] lines2 = cellValue2.Split('\n');
                        double rowHeightCounts1 = 0;
                        double rowHeightCounts2 = 0;
                        foreach (var line in lines1)
                        {
                            rowHeightCounts1 += Math.Ceiling((double)line.Length / totalCharNum1);
                        }
                        foreach (var line in lines2)
                        {
                            rowHeightCounts2 += Math.Ceiling((double)line.Length / totalCharNum2);
                        }
                        if (rowHeightCounts1 >= rowHeightCounts2)
                        {
                            var row = sheet.GetRow(indexesArray1[i, 0] - 1);
                            row.Height = (short)(rowHeightCounts1 * 20 * 17);
                        }
                        else
                        {
                            var row = sheet.GetRow(indexesArray2[i, 0] - 1);
                            row.Height = (short)(rowHeightCounts2 * 20 * 17);
                        }
                    }
                }
                if (rowDiff != 0)
                {
                    int colIndexMergeStart;
                    int colIndexMergeEnd;
                    object[] data;
                    int[,] indexesArray;
                    if (rowDiff >= 0)
                    {
                        colIndexMergeStart = colIndexMergeStart1;
                        colIndexMergeEnd = colIndexMergeEnd1;
                        data = data1;
                        indexesArray = indexesArray1;
                    }
                    else
                    {
                        colIndexMergeStart = colIndexMergeStart2;
                        colIndexMergeEnd = colIndexMergeEnd2;
                        data = data2;
                        indexesArray = indexesArray2;
                    }
                    double totalCharNum = 0.0;
                    for (int i = colIndexMergeStart; i <= colIndexMergeEnd; i++)
                    {
                        totalCharNum += sheet.GetColumnWidth(i - 1) / 256 / 2.0625;
                    }
                    for (int i = rowMin; i < rowMax; i++)
                    {
                        string cellValue = (string)data[i];
                        string[] lines = cellValue.Split('\n');
                        double rowHeightCounts = 0;
                        foreach (var line in lines)
                        {
                            rowHeightCounts += Math.Ceiling((double)line.Length / totalCharNum);
                        }
                        var row = sheet.GetRow(indexesArray[i, 0] - 1);
                        row.Height = (short)(rowHeightCounts * 20 * 17);
                    }
                }
            }
            else
            {
                placeholder = placeholder1;
                int[,] indexesArray = GetCellIndexByPlaceholder(sheet, placeholder);
                object[] data = variables[placeholder];
                var MergedRangeHor = GetMergedRangeHor(sheet, indexesArray[0, 0], indexesArray[0, 1]);
                int colIndexMergeStart = MergedRangeHor[0];
                int colIndexMergeEnd = MergedRangeHor[1];
                double totalCharNum = 0.0;
                for (int i = colIndexMergeStart; i <= colIndexMergeEnd; i++)
                {
                    totalCharNum += sheet.GetColumnWidth(i - 1) / 256 / 2.0625;
                }
                for (int i = 0; i < data.Length; i++)
                {
                    string cellValue = (string)data[i];
                    string[] lines = cellValue.Split('\n');
                    double rowHeightCounts = 0;
                    foreach (var line in lines)
                    {
                        rowHeightCounts += Math.Ceiling((double)line.Length / totalCharNum);
                    }
                    var row = sheet.GetRow(indexesArray[i, 0] - 1);
                    row.Height = (short)(rowHeightCounts * 20 * 17);
                }
            }

        }
        // 調整row數及row高
        static void AdjustRowNumsAndHeight(ISheet sheet, Dictionary<string, object[]> variables, string placeholder1, string? placeholder2 = null, string? placeholder3 = null)
        {
            AdjustRowNums(sheet, variables, placeholder1, placeholder2, placeholder3);
            AdjustRowHeight(sheet, variables, placeholder1, placeholder2, placeholder3);
        }
    }
}
