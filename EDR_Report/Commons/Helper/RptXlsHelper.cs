using EDR_Report.Interfaces;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Specialized;
using System.Collections;
using System.Globalization;

namespace EDR_Report.Commons
{
    public enum TaiwanDateFormatEnum
    {
        民國年 = 0,
        民國年加星期,
        民國年不顯示民國,
        民國年不顯示民國加星期,
        民國年斜線年月日

    }
    public enum WorkItemsEnum
    {
        施工項目 = 'N',
        特定施工項目 = 'Y'
    }
    public enum ResourceClassEnum
    {
        人工 = 2,
        機具 = 3,
        材料 = 4,
        器材 = 5,
        機具器材 = 35

    }

    public class RptXlsHelper : IRptXlsHelper
    {
        private bool _isXSS = true;

        public bool TaiwanDateTimeReform(string taiwan_datetime, string datetime_format, out string TaiDate)
        {
            TaiDate = null;
            bool ret = false;

            DateTime testResult = new DateTime();
            switch (datetime_format)
            {
                case "yyyMMdd":
                    if (DateTime.TryParse($"{int.Parse(taiwan_datetime.Substring(0, 3)) + 1911}/{taiwan_datetime.Substring(3, 2)}/{taiwan_datetime.Substring(5, 2)}", out testResult))
                    {
                        ret = true;
                        TaiDate = $"{taiwan_datetime.Substring(0, 3)}年{taiwan_datetime.Substring(3, 2)}月{taiwan_datetime.Substring(5, 2)}日";
                    }
                    break;
                case "yyy/MM/dd":
                    if (DateTime.TryParse($"{taiwan_datetime.Substring(0, 3) + 1911}/{taiwan_datetime.Substring(4, 2)}/{taiwan_datetime.Substring(7, 2)}", out testResult))
                    {
                        ret = true;
                        TaiDate = $"{taiwan_datetime.Substring(0, 3)}年{taiwan_datetime.Substring(3, 2)}月{taiwan_datetime.Substring(5, 2)}日";
                    }
                    break;
            }

            //var dateTxt = "二 5月 16 14:41:40 +0800 2023";
            //var format = "ddd M月 d HH:mm:ss zzz yyyy";
            //var parsedDate = DateTimeOffset.ParseExact(dateTxt, format, writeableClone);

            return ret;
        }

        public string TaiwanDateTime(string report_date, TaiwanDateFormatEnum dateformat)
        {
            string TaiDate = string.Empty;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-TW");
            Calendar twCalendar = new TaiwanCalendar();
            DateTime theDate;
            if (!DateTime.TryParse(report_date, out theDate))
                throw new Exception("報表日期錯誤");



            switch (dateformat)
            {
                case TaiwanDateFormatEnum.民國年:
                default:
                    TaiDate = string.Format("民國{0}年{1}月{2}日", twCalendar.GetYear(theDate), theDate.Month, theDate.Day);
                    break;
                case TaiwanDateFormatEnum.民國年加星期:
                    TaiDate = string.Format("民國{0}年{1}月{2}日({3})", twCalendar.GetYear(theDate), theDate.Month, theDate.Day, twCalendar.GetDayOfWeek(theDate));
                    break;
                case TaiwanDateFormatEnum.民國年不顯示民國:
                    TaiDate = string.Format("{0}年{1}月{2}日", twCalendar.GetYear(theDate), theDate.Month, theDate.Day);
                    break;
                case TaiwanDateFormatEnum.民國年不顯示民國加星期:
                    TaiDate = string.Format("{0}年{1}月{2}日({3})", twCalendar.GetYear(theDate), theDate.Month, theDate.Day, twCalendar.GetDayOfWeek(theDate));
                    break;
                case TaiwanDateFormatEnum.民國年斜線年月日:
                    TaiDate = string.Format("{0}/{1}/{2}", twCalendar.GetYear(theDate), theDate.Month, theDate.Day);
                    break;
            }


            return TaiDate;
        }
        public string GetFormattedCellValue(ICell cell, IFormulaEvaluator eval = null)
        {
            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;

                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            DateTime date = cell.DateCellValue;
                            ICellStyle style = cell.CellStyle;
                            // Excel uses lowercase m for month whereas .Net uses uppercase
                            string format = style.GetDataFormatString().Replace('m', 'M');
                            return date.ToString(format);
                        }
                        else
                        {
                            return cell.NumericCellValue.ToString();
                        }

                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";

                    case CellType.Formula:
                        if (eval != null)
                            return GetFormattedCellValue(cell, eval);
                        else
                            return cell.CellFormula;

                    case CellType.Error:
                        return FormulaError.ForInt(cell.ErrorCellValue).String;
                }
            }
            // null or blank cell, or unknown cell type
            return string.Empty;
        }
        public IWorkbook OpenExcelTemplate(string templateFullpath, out IFormulaEvaluator ObjectIFormulaEvaluator)
        {
            IWorkbook wb;// = new XSSFWorkbook(templateFullpath);
            string FileExtensionString = System.IO.Path.GetExtension(templateFullpath);
            using (var fs = new FileStream(templateFullpath, FileMode.Open, FileAccess.Read))
            {
                if (FileExtensionString.Equals(".xls"))
                {

                    //create the instance of our object ObjectWorkBook for XLS excel files
                    wb = new HSSFWorkbook(fs);
                    //create the instance of our object ObjectIFormulaEvaluator for XLS excel files
                    ObjectIFormulaEvaluator = new HSSFFormulaEvaluator(wb); //NOTE: If you're sure your sheets have no formulas you can delete or comment this line
                    this._isXSS = false;
                }
                else
                {
                    //NOTE: On next line we're really using XSSF for any file extension different from XLS (XLS but also PNG, DOCX, TXT...) Perhaps you should control
                    //it's really a XLSX and not any other
                    //create the instance of our object ObjectWorkBook for XLSX excel files
                    wb = new XSSFWorkbook(fs);
                    //create the instance of our object ObjectIFormulaEvaluator for XLSX excel files
                    ObjectIFormulaEvaluator = new XSSFFormulaEvaluator(wb); //NOTE: If you're sure your sheets have no formulas you can delete or comment this line
                    this._isXSS = true;
                }
            }
            return wb;
        }
        public string[] GetColumnValues(ISheet sheet, int columnNum)
        {

            ArrayList val = new ArrayList();

            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum; // This may not be valid row count.
                                                 // If first row is table head, i starts from 1
                for (int i = 1; i <= rowCount; i++)
                {
                    IRow curRow = sheet.GetRow(i);
                    // Works for consecutive data. Use continue otherwise 
                    if (curRow == null)
                    {
                        // Valid row count
                        rowCount = i - 1;
                        break;
                    }
                    // Get data from the 4th column (4th cell of each row)
                    var cellValue = curRow.GetCell(columnNum).StringCellValue.Trim();
                    val.Add(cellValue);
                }

            }
            return (string[])val.ToArray(typeof(string));

        }

        public IRow FindRowByCellValue(ISheet ws, string varName, int starCol, int endCol, int startRow, int endRow)
        {
            bool isFound = false;
            string CellContentString = "";
            IRow ObjectIRow = ws.GetRow(0);
            int endrow_num = (endCol == 0 ? ws.LastRowNum : endRow);
            for (int CountRow = starCol; CountRow <= endrow_num; CountRow++)
            {

                ObjectIRow = ws.GetRow(CountRow);

                int LastColumn = (endCol == 0 ? ObjectIRow.LastCellNum - 1 : endCol);
                for (int CountCol = 0; CountCol <= LastColumn; CountCol++)
                {

                    NPOI.SS.UserModel.ICell ICellObject = ObjectIRow.GetCell(CountCol);

                    CellContentString = GetFormattedCellValue(ICellObject);
                    Console.WriteLine(CellContentString);
                    if (CellContentString == varName)
                    {
                        isFound = true;
                        break;
                    }
                }
                if (isFound)
                    break;

            }
            return ObjectIRow;
        }
        public IRow[] FindRowByValue(ISheet ws, string varName, int starCol, int endCol, int startRow, int endRow)
        {
            bool isFound = false;
            var retRows = new List<IRow>();
            string CellContentString = "";
            IRow ObjectIRow = ws.GetRow(0);
            int endrow_num = (endCol == 0 ? ws.LastRowNum : endRow);
            for (int CountRow = starCol; CountRow <= endrow_num; CountRow++)
            {

                ObjectIRow = ws.GetRow(CountRow);

                int LastColumn = (endCol == 0 ? ObjectIRow.LastCellNum - 1 : endCol);
                for (int CountCol = 0; CountCol <= LastColumn; CountCol++)
                {

                    NPOI.SS.UserModel.ICell ICellObject = ObjectIRow.GetCell(CountCol);

                    CellContentString = GetFormattedCellValue(ICellObject);
                    Console.WriteLine(CellContentString);
                    if (CellContentString == varName)
                    {
                        isFound = true;
                        break;
                    }
                }
                if (isFound)
                {
                    retRows.Add(ObjectIRow);
                    break;
                }


            }
            return retRows.ToArray();
        }
        public void Replace(ISheet ws, string varName, string value, int StartCol, int EndCol, int StartRow, int EndRow)
        {
            string CellContentString = "";
            int endrow_num = (EndCol == 0 ? ws.LastRowNum : EndRow);
            for (int CountRow = StartCol; CountRow <= endrow_num; CountRow++)
            {

                IRow ObjectIRow = ws.GetRow(CountRow);

                int LastColumn = (EndCol == 0 ? ObjectIRow.LastCellNum - 1 : EndCol);
                for (int CountCol = 0; CountCol <= LastColumn; CountCol++)
                {

                    NPOI.SS.UserModel.ICell ICellObject = ObjectIRow.GetCell(CountCol);

                    CellContentString = GetFormattedCellValue(ICellObject);
                    Console.WriteLine(CellContentString);
                    if (CellContentString.Contains(varName))
                    {
                        ICellObject.SetCellValue(CellContentString.Replace(varName, value));
                        break;
                    }
                }

            }
        }
        public void ReplaceAll(ISheet ws, string varName, string value, int StartCol, int EndCol, int StartRow, int EndRow)
        {
            string CellContentString = "";
            int endrow_num = (EndCol == 0 ? ws.LastRowNum : EndRow);
            for (int CountRow = StartCol; CountRow <= endrow_num; CountRow++)
            {

                IRow ObjectIRow = ws.GetRow(CountRow);

                int LastColumn = (EndCol == 0 ? ObjectIRow.LastCellNum - 1 : EndCol);
                for (int CountCol = 0; CountCol <= LastColumn; CountCol++)
                {

                    NPOI.SS.UserModel.ICell ICellObject = ObjectIRow.GetCell(CountCol);

                    CellContentString = GetFormattedCellValue(ICellObject);
                    Console.WriteLine(CellContentString);
                    if (CellContentString == varName)
                    {
                        ICellObject.SetCellValue(value);
                    }
                }

            }
        }

        public NPOI.SS.UserModel.ICell MergedCell(NPOI.SS.UserModel.ICell cell)
        {
            if (cell.IsMergedCell)//是否是合并单元格
            {
                for (int i = 0; i < cell.Sheet.NumMergedRegions; i++)//遍历所有的合并单元格
                {
                    var cellRange = cell.Sheet.GetMergedRegion(i);
                    if (cell.ColumnIndex >= cellRange.FirstColumn && cell.ColumnIndex <= cellRange.LastColumn
                        && cell.RowIndex >= cellRange.FirstRow && cell.RowIndex <= cellRange.LastRow)//判断查询的单元格是否在合并单元格内
                    {
                        return cell.Sheet.GetRow(cellRange.FirstRow).GetCell(cellRange.FirstColumn);
                    }
                }
            }
            return cell;
        }
        public NPOI.SS.UserModel.ICell[] PrepareCells(IRow row)
        {
            List<NPOI.SS.UserModel.ICell> cells = new List<NPOI.SS.UserModel.ICell>();
            if (row != null)
            {
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    NPOI.SS.UserModel.ICell theFirstCellOfMerged = MergedCell(row.GetCell(j));
                    if (theFirstCellOfMerged != null)
                    {
                        //_rptHelper.SetCellValue(theFirstCellOfMerged, workItems[i].OWNITEM_NO);
                        cells.Add(theFirstCellOfMerged);
                        continue;
                    }

                }

            }

            return cells.ToArray();
        }

        public void SetRow(IRow row, ListDictionary cellvalues)
        {
            foreach (DictionaryEntry dict in cellvalues)
            {
                //int cellnum = (int)dict.Key;
                NPOI.SS.UserModel.ICell cell = row.GetCell((int)dict.Key);
                SetCellValue(cell, dict.Value);
            }

        }
        public void SetRow(ISheet ws, int rownum, ListDictionary cellvalues)
        {

            //foreach (DictionaryEntry dict in cellvalues)
            //{
            //    IRow row = ws.GetRow(rownum);
            //    NPOI.SS.UserModel.ICell cell = row.GetCell((int)dict.Key);
            //    //SetCellValue(cell, dict.Value);
            //    SetCellValue(cell, "test123456");
            //    var test = 0;
            //}
            IRow row = ws.GetRow(rownum);
            foreach (DictionaryEntry dict in cellvalues)
            {

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    if (i != (int)dict.Key)
                    { continue; }
                    NPOI.SS.UserModel.ICell cell = row.GetCell(i);
                    SetCellValue(cell, dict.Value);
                }
            }

        }
        public void MergeCellsOnRows(IRow row, int rowCount, int startCol, int endCol)
        {
            var ws = row.Sheet;
            //從下一個row開始merge: 將傳進來的row +1 , 幾列要merge ->rowCount
            NPOI.SS.Util.CellRangeAddress cra = new NPOI.SS.Util.CellRangeAddress(row.RowNum + 1, rowCount, startCol, endCol);
            ws.AddMergedRegion(cra);
        }
        public void CopyRow(ISheet ws, int sourceRowNum, int destinationRowNum, int shiftCounts)
        {
            int newRowIdx = sourceRowNum; //= sourceRowNum;

            //框選的rows全部往下移copyCount列
            ws.ShiftRows(destinationRowNum, ws.LastRowNum, shiftCounts, true, false);
            for (int i = 0; i < shiftCounts; i++)
            {
                IRow newRow = ws.CreateRow(newRowIdx);
                IRow sourceRow = ws.GetRow(newRowIdx);
                //newRowIdx = newRow.RowNum;


                if (sourceRow != null)
                {
                    try
                    {

                        // If there are are any merged regions in the source row, copy to new row
                        for (int j = 0; j < ws.NumMergedRegions; j++)
                        {
                            CellRangeAddress cellRangeAddress = ws.GetMergedRegion(j);
                            if (cellRangeAddress.FirstRow == newRowIdx && cellRangeAddress.LastRow == newRowIdx)
                            {
                                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRowIdx,
                                                                (newRowIdx + (cellRangeAddress.FirstRow - cellRangeAddress.LastRow)),
                                                                cellRangeAddress.FirstColumn,
                                                                cellRangeAddress.LastColumn);
                                ws.AddMergedRegion(newCellRangeAddress);
                            }
                        }
                        newRowIdx++;
                        // Loop through source columns to add to new row
                        for (int k = 0; k < sourceRow.LastCellNum; k++)
                        {

                            // Grab a copy of the old/new cell
                            NPOI.SS.UserModel.ICell sourceCell = sourceRow.GetCell(k);
                            NPOI.SS.UserModel.ICell newCell = newRow.CreateCell(k);

                            // If the old cell is null jump to next cell
                            if (sourceCell == null)
                            {
                                newCell = null;
                                continue;
                            }


                            if (_isXSS)
                            {
                                // Copy style from old cell and apply to new cell
                                ICellStyle newCellStyle = ws.Workbook.CreateCellStyle();
                                newCellStyle.CloneStyleFrom(sourceCell.CellStyle); ;
                                newCell.CellStyle = newCellStyle;

                                // If there is a cell comment, copy
                                if (newCell.CellComment != null) newCell.CellComment = sourceCell.CellComment;

                                // If there is a cell hyperlink, copy
                                if (sourceCell.Hyperlink != null) newCell.Hyperlink = sourceCell.Hyperlink;
                                var dataFormatter = new DataFormatter(CultureInfo.CurrentCulture);
                                //newCell.SetCellValue(dataFormatter.FormatCellValue(oldCell));
                                //newCell.SetCellValue(oldCell.StringCellValue);
                                switch (sourceCell.CellType)
                                {
                                    case CellType.Blank:
                                        newCell.SetCellValue(sourceCell.StringCellValue);
                                        break;
                                    case CellType.Boolean:
                                        newCell.SetCellValue(sourceCell.BooleanCellValue);
                                        break;
                                    case CellType.Error:
                                        newCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                                        break;
                                    case CellType.Formula:
                                        newCell.SetCellFormula(sourceCell.CellFormula);
                                        break;
                                    case CellType.Numeric:
                                        newCell.SetCellValue(sourceCell.NumericCellValue);
                                        break;
                                    case CellType.String:
                                        newCell.SetCellValue(sourceCell.RichStringCellValue);
                                        break;
                                    case CellType.Unknown:
                                        newCell.SetCellValue(sourceCell.StringCellValue);
                                        break;
                                }
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                }
            }


            //newRow = ws.GetRow(sourceRowNum);






        }
        /// <summary>
        /// 模擬手動新增列與合併儲存格
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="startRowNum"></param>
        /// <param name="shiftCounts"></param>
        public void InsertRows(ISheet ws, int startRowNum, int shiftCounts)
        {
            int curRowNum = 0;

            for (int i = 0; i < shiftCounts; i++)
            {
                curRowNum = startRowNum + i;
                //框選的rows逐次往下移1列
                ws.ShiftRows(curRowNum + 1, ws.LastRowNum, 1, true, false);

                ///NPOI CreateRow的定義怪怪的，應該是移動後要
                ///在空列新建一列ws.CreateRow(curRowNum)，
                ///結果是在空列下方建一列ws.CreateRow(curRowNum + 1)
                ///原本的位置已經往下移一列，
                ///應該是 ws.GetRow(curRowNum+1)
                ///結果是 ws.GetRow(curRowNum)
                IRow newRow = ws.CreateRow(curRowNum + 1);
                IRow sourceRow = ws.GetRow(curRowNum);
                Console.WriteLine("1.newRowIdx=" + curRowNum);
                //
                if (sourceRow == null)
                {
                    continue;
                }

                try
                {
                    // Loop through source columns to add to new row
                    for (int k = 0; k < sourceRow.LastCellNum; k++)
                    {

                        // Grab a copy of the old/new cell
                        NPOI.SS.UserModel.ICell sourceCell = sourceRow.GetCell(k);
                        NPOI.SS.UserModel.ICell newCell = newRow.CreateCell(k);

                        // If the old cell is null jump to next cell
                        if (sourceCell == null)
                        {
                            newCell = null;
                            continue;
                        }
                        // Copy style from old cell and apply to new cell
                        ICellStyle newCellStyle = ws.Workbook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(sourceCell.CellStyle); ;
                        newCell.CellStyle = newCellStyle;

                        // If there is a cell comment, copy
                        if (newCell.CellComment != null) newCell.CellComment = sourceCell.CellComment;

                        // If there is a cell hyperlink, copy
                        if (sourceCell.Hyperlink != null) newCell.Hyperlink = sourceCell.Hyperlink;
                        var dataFormatter = new DataFormatter(CultureInfo.CurrentCulture);
                        //newCell.SetCellValue(dataFormatter.FormatCellValue(oldCell));
                        //newCell.SetCellValue(oldCell.StringCellValue);
                        switch (sourceCell.CellType)
                        {
                            case CellType.Blank:
                                newCell.SetCellValue(sourceCell.StringCellValue);
                                break;
                            case CellType.Boolean:
                                newCell.SetCellValue(sourceCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                newCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                                break;
                            case CellType.Formula:
                                newCell.SetCellFormula(sourceCell.CellFormula);
                                break;
                            case CellType.Numeric:
                                newCell.SetCellValue(sourceCell.NumericCellValue);
                                break;
                            case CellType.String:
                                newCell.SetCellValue(sourceCell.RichStringCellValue);
                                break;
                            case CellType.Unknown:
                                newCell.SetCellValue(sourceCell.StringCellValue);
                                break;
                        }

                    }
                    // If there are are any merged regions in the source row, copy to new row
                    for (int j = 0; j < ws.NumMergedRegions; j++)
                    {
                        CellRangeAddress cellRangeAddress = ws.GetMergedRegion(j);
                        if (cellRangeAddress.FirstRow == curRowNum && cellRangeAddress.LastRow == curRowNum)
                        {
                            CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                                                            (newRow.RowNum + (cellRangeAddress.LastRow - cellRangeAddress.FirstRow)),
                                                            cellRangeAddress.FirstColumn,
                                                            cellRangeAddress.LastColumn);
                            ws.AddMergedRegion(newCellRangeAddress);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }


                //curRowNum++;
                Console.WriteLine("2.newRowIdx=" + curRowNum);
            }

        }
        //public void InsertRows(ISheet ws, int startRowNum, int shiftCounts)
        //{

        //    //框選的rows逐次往下移1列

        //    for (int i = 0; i < shiftCounts; i++)
        //    {
        //        ws.ShiftRows(startRowNum + 1, ws.LastRowNum, 1, true, false);
        //        IRow newRow = ws.CreateRow(startRowNum + 1);
        //        IRow sourceRow = ws.GetRow(startRowNum);

        //        if (sourceRow != null)
        //        {
        //            try
        //            {



        //                // Loop through source columns to add to new row
        //                for (int k = 0; k < sourceRow.LastCellNum; k++)
        //                {

        //                    // Grab a copy of the old/new cell
        //                    NPOI.SS.UserModel.ICell sourceCell = sourceRow.GetCell(k);
        //                    NPOI.SS.UserModel.ICell newCell = newRow.CreateCell(k);

        //                    // If the old cell is null jump to next cell
        //                    if (sourceCell == null)
        //                    {
        //                        newCell = null;
        //                        continue;
        //                    }


        //                    if (_isXSS)
        //                    {
        //                        // Copy style from old cell and apply to new cell
        //                        ICellStyle newCellStyle = ws.Workbook.CreateCellStyle();
        //                        newCellStyle.CloneStyleFrom(sourceCell.CellStyle); ;
        //                        newCell.CellStyle = newCellStyle;

        //                        // If there is a cell comment, copy
        //                        if (newCell.CellComment != null) newCell.CellComment = sourceCell.CellComment;

        //                        // If there is a cell hyperlink, copy
        //                        if (sourceCell.Hyperlink != null) newCell.Hyperlink = sourceCell.Hyperlink;
        //                        var dataFormatter = new DataFormatter(CultureInfo.CurrentCulture);
        //                        //newCell.SetCellValue(dataFormatter.FormatCellValue(oldCell));
        //                        //newCell.SetCellValue(oldCell.StringCellValue);
        //                        switch (sourceCell.CellType)
        //                        {
        //                            case CellType.Blank:
        //                                newCell.SetCellValue(sourceCell.StringCellValue);
        //                                break;
        //                            case CellType.Boolean:
        //                                newCell.SetCellValue(sourceCell.BooleanCellValue);
        //                                break;
        //                            case CellType.Error:
        //                                newCell.SetCellErrorValue(sourceCell.ErrorCellValue);
        //                                break;
        //                            case CellType.Formula:
        //                                newCell.SetCellFormula(sourceCell.CellFormula);
        //                                break;
        //                            case CellType.Numeric:
        //                                newCell.SetCellValue(sourceCell.NumericCellValue);
        //                                break;
        //                            case CellType.String:
        //                                newCell.SetCellValue(sourceCell.RichStringCellValue);
        //                                break;
        //                            case CellType.Unknown:
        //                                newCell.SetCellValue(sourceCell.StringCellValue);
        //                                break;
        //                        }
        //                    }

        //                }
        //                // If there are are any merged regions in the source row, copy to new row
        //                for (int j = 0; j < ws.NumMergedRegions; j++)
        //                {
        //                    CellRangeAddress cellRangeAddress = ws.GetMergedRegion(j);
        //                    if (cellRangeAddress.FirstRow == startRowNum && cellRangeAddress.LastRow == startRowNum)
        //                    {
        //                        CellRangeAddress newCellRangeAddress = new CellRangeAddress(startRowNum + 1,
        //                                                        (startRowNum + 1 + (cellRangeAddress.LastRow - cellRangeAddress.FirstRow)),
        //                                                        cellRangeAddress.FirstColumn,
        //                                                        cellRangeAddress.LastColumn);
        //                        ws.AddMergedRegion(newCellRangeAddress);
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                throw ex;
        //            }

        //        }

        //        startRowNum++;
        //    }

        //}
        //public void CopyRow(IWorkbook wb, ISheet ws, int sourceRowNum, int destinationRowNum, int copyRows)
        //{
        //    // Get the source / new row
        //    IRow newRow = ws.GetRow(destinationRowNum);
        //    int newRowIdx = sourceRowNum;
        //    IRow sourceRow = ws.GetRow(sourceRowNum);

        //    // If the row exist in destination, push down all rows by 1 else create a new row
        //    if (newRow != null)
        //    {
        //        ws.ShiftRows(destinationRowNum, ws.LastRowNum, 1);
        //    }
        //    else
        //    {
        //        newRow = ws.CreateRow(destinationRowNum);
        //    }

        //    // Loop through source columns to add to new row
        //    for (int i = 0; i < sourceRow.LastCellNum; i++)
        //    {
        //        // Grab a copy of the old/new cell
        //        NPOI.SS.UserModel.ICell oldCell = sourceRow.GetCell(i);
        //        NPOI.SS.UserModel.ICell newCell = newRow.CreateCell(i);

        //        // If the old cell is null jump to next cell
        //        if (oldCell == null)
        //        {
        //            newCell = null;
        //            continue;
        //        }



        //        if (_isXSS)
        //        {
        //            // Copy style from old cell and apply to new cell
        //            ICellStyle newCellStyle = wb.CreateCellStyle();
        //            newCellStyle.CloneStyleFrom(oldCell.CellStyle); ;
        //            newCell.CellStyle = newCellStyle;

        //            // If there is a cell comment, copy
        //            if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

        //            // If there is a cell hyperlink, copy
        //            if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;
        //            var dataFormatter = new DataFormatter(CultureInfo.CurrentCulture);
        //            //newCell.SetCellValue(dataFormatter.FormatCellValue(oldCell));
        //            newCell.SetCellValue(oldCell.StringCellValue);
        //            //switch (oldCell.CellType)
        //            //{
        //            //    case ICellStyle.BLANK:
        //            //        newCell.SetCellValue(oldCell.StringCellValue);
        //            //        break;
        //            //    case HSSFCellType.BOOLEAN:
        //            //        newCell.SetCellValue(oldCell.BooleanCellValue);
        //            //        break;
        //            //    case HSSFCellType.ERROR:
        //            //        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
        //            //        break;
        //            //    case HSSFCellType.FORMULA:
        //            //        newCell.SetCellFormula(oldCell.CellFormula);
        //            //        break;
        //            //    case HSSFCellType.NUMERIC:
        //            //        newCell.SetCellValue(oldCell.NumericCellValue);
        //            //        break;
        //            //    case HSSFCellType.STRING:
        //            //        newCell.SetCellValue(oldCell.RichStringCellValue);
        //            //        break;
        //            //    case HSSFCellType.Unknown:
        //            //        newCell.SetCellValue(oldCell.StringCellValue);
        //            //        break;
        //            //}
        //        } 

        //    }

        //    try
        //    {

        //        // If there are are any merged regions in the source row, copy to new row
        //        for (int i = 0; i < ws.NumMergedRegions; i++)
        //        {
        //            CellRangeAddress cellRangeAddress = ws.GetMergedRegion(i);
        //            if (cellRangeAddress.FirstRow == sourceRow.RowNum)
        //            {
        //                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
        //                                              (newRow.RowNum +
        //                                               (cellRangeAddress.FirstRow -
        //                                                cellRangeAddress.LastRow)),
        //                                              cellRangeAddress.FirstColumn,
        //                                              cellRangeAddress.LastColumn);
        //                ws.AddMergedRegion(newCellRangeAddress);
        //            }
        //        }
        //    } catch (Exception ex)
        //    {
        //        throw ex;
        //    }

        //}

        public void CreateCell(IRow currentRow, int col, string value, HSSFCellStyle hssfStyle = null, XSSFCellStyle xssStyle = null)
        {
            if (this._isXSS)
            {
                XSSFCell Cell = (XSSFCell)((XSSFRow)currentRow).CreateCell(col);
                Cell.SetCellValue(value);
                Cell.CellStyle = xssStyle;
            }
            else
            {
                HSSFCell Cell = (HSSFCell)((HSSFRow)currentRow).CreateCell(col);
                Cell.SetCellValue(value);
                Cell.CellStyle = hssfStyle;
            }
        }


        public void SetCellStyleOnRow(ISheet ws, int rowIdx, int colIndex, bool createRow = false, bool defaultStyle = false)
        {
            if (this._isXSS)
            {

                XSSFCellStyle style;
                XSSFRow row; //= (XSSFRow)ws.GetRow(rowIdx);
                if (createRow)
                    row = (XSSFRow)((XSSFSheet)ws).CreateRow(rowIdx + 1);
                else
                    row = (XSSFRow)ws.GetRow(rowIdx);

                if (defaultStyle)
                {
                    style = (XSSFCellStyle)ws.Workbook.CreateCellStyle();
                    style.WrapText = true;
                    style.Alignment = HorizontalAlignment.Left;
                    style.VerticalAlignment = VerticalAlignment.Top;
                }
                else
                {
                    row = (XSSFRow)ws.GetRow(rowIdx);
                    style = (XSSFCellStyle)ws.Workbook.CreateCellStyle();
                    style.WrapText = true;
                    style.FillForegroundColor = IndexedColors.LightBlue.Index;
                    style.Alignment = HorizontalAlignment.Center;
                }
                row.Cells[colIndex].CellStyle = style;
            }
            else
            {
                HSSFCellStyle style;
                HSSFRow row; //= (XSSFRow)ws.GetRow(rowIdx);
                if (createRow)
                    row = (HSSFRow)((HSSFSheet)ws).CreateRow(rowIdx + 1);
                else
                    row = (HSSFRow)ws.GetRow(rowIdx);

                if (defaultStyle)
                {
                    style = (HSSFCellStyle)ws.Workbook.CreateCellStyle();
                    style.WrapText = true;
                    style.Alignment = HorizontalAlignment.Left;
                    style.VerticalAlignment = VerticalAlignment.Top;
                }
                else
                {
                    row = (HSSFRow)ws.GetRow(rowIdx);
                    style = (HSSFCellStyle)ws.Workbook.CreateCellStyle();
                    style.WrapText = true;
                    style.FillForegroundColor = IndexedColors.LightBlue.Index;
                    style.Alignment = HorizontalAlignment.Center;
                }
                row.Cells[colIndex].CellStyle = style;
            }

        }


        public ICellStyle GetPreferredCellStyle(NPOI.SS.UserModel.ICell cell)
        {
            // a method to get the preferred cell style for a cell
            // this is either the already applied cell style
            // or if that not present, then the row style (default cell style for this row)
            // or if that not present, then the column style (default cell style for this column)
            ICellStyle cellStyle = cell.CellStyle;
            if (cellStyle.Index == 0) cellStyle = cell.Row.RowStyle;
            if (cellStyle == null) cellStyle = cell.Sheet.GetColumnStyle(cell.ColumnIndex);
            if (cellStyle == null) cellStyle = cell.CellStyle;
            return cellStyle;
        }
        public void SetMergedCellRegions(ISheet ws, ICellStyle style)
        {
            int mergedRegions = ws.NumMergedRegions;
            for (int regions = 0; regions < mergedRegions; regions++)
            {
                CellRangeAddress mergedRegionIndex = ws.GetMergedRegion(regions);
                for (int currentRegion = mergedRegionIndex.FirstRow; currentRegion < mergedRegionIndex.LastRow; currentRegion++)
                {
                    var currentRow = ws.GetRow(currentRegion);
                    for (int currentCell = mergedRegionIndex.FirstColumn; currentCell < mergedRegionIndex.LastColumn; currentCell++)
                    {
                        // sheet.SetDefaultColumnStyle(i, mandatoryCellStyle);
                        currentRow.Cells[currentCell].CellStyle = style;
                    }
                }
            }

        }
        //Get merged cell
        public NPOI.SS.UserModel.ICell GetFirstCellInMergedRegionContainingCell(NPOI.SS.UserModel.ICell cell)
        {
            if (cell != null && cell.IsMergedCell)
            {
                ISheet sheet = cell.Sheet;
                for (int i = 0; i < sheet.NumMergedRegions; i++)
                {
                    CellRangeAddress region = sheet.GetMergedRegion(i);
                    if (region.ContainsRow(cell.RowIndex) &&
                        region.ContainsColumn(cell.ColumnIndex))
                    {
                        IRow row = sheet.GetRow(region.FirstRow);
                        NPOI.SS.UserModel.ICell firstCell = row?.GetCell(region.FirstColumn);
                        return firstCell;
                    }
                }
                return null;
            }
            return cell;
        }

        public List<string> ReadRow(IRow row, int starCol, int endCol)
        {
            var ValuesInRow = new List<string>();
            // Since there would be no emtpy cell, // we start from the left-most cell and read until the end.
            for (int colIndex = starCol; colIndex < endCol; colIndex++)
            {
                var cell = row.GetCell(colIndex);
                // Get the value of this cell, whether it is part of a merged cell or not.
                var cellValue = this.GetCellValue(cell);
                ValuesInRow.Add(cellValue);
            }
            return ValuesInRow;


        }
        public void SetCellValues(IRow row, List<string> value, List<int> colIdx = null)
        {

            if (colIdx == null || colIdx.Count == 0)
            {
                for (int i = 0; i < value.Count; i++)
                {
                    row.GetCell(i).SetCellValue(value[i].ToString());
                }
            }
            else
            {
                for (int i = 0; i < colIdx.Count; i++)
                {
                    row.GetCell(colIdx[i]).SetCellValue(value[i].ToString());
                }

            }
        }
        public void SetCellValue(NPOI.SS.UserModel.ICell cell, object cellValue)
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


        public string GetCellValue(NPOI.SS.UserModel.ICell cell)
        {
            var dataFormatter = new DataFormatter(CultureInfo.CurrentCulture);
            // If this is not part of a merge cell,// just get this cell's value like normal.
            if (!cell.IsMergedCell)
            {
                return dataFormatter.FormatCellValue(cell);
            }
            // Otherwise, we need to find the value of this merged cell.
            else
            {
                // Get current sheet.
                var currentSheet = cell.Sheet;
                // Loop through all merge regions in this sheet.
                for (int i = 0; i < currentSheet.NumMergedRegions; i++)
                {
                    var mergeRegion = currentSheet.GetMergedRegion(i);
                    // If this merged region contains this cell.
                    if (mergeRegion.FirstRow <= cell.RowIndex && cell.RowIndex <= mergeRegion.LastRow
                                    && mergeRegion.FirstColumn <= cell.ColumnIndex && cell.ColumnIndex <= mergeRegion.LastColumn)
                    {
                        // Find the top-most and left-most cell in this region.
                        var firstRegionCell = currentSheet.GetRow(mergeRegion.FirstRow).GetCell(mergeRegion.FirstColumn);
                        // And return its value.
                        return dataFormatter.FormatCellValue(firstRegionCell);
                    }
                }
            }
            // This should never happen.
            throw new Exception("Cannot find this cell in any merged region");
        }


        public string ReplacePlaceHolders(string template, Dictionary<string, string> placeHoldersAndValues)
        {
            foreach (var placeHolderAndValue in placeHoldersAndValues)
            {
                if (template.Contains(placeHolderAndValue.Key))
                {
                    template = template.Replace(placeHolderAndValue.Key, placeHolderAndValue.Value);
                }
            }

            return template;
        }

        public void CopyRowDown(ISheet sheet, int rowIndexStart, int repeatTimes)
        {
            //rowIndexStart--; // 將 rowIndex 轉換為以 0 為基底的索引

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
                        NPOI.SS.UserModel.ICell sourceCell = sourceRow.GetCell(j);
                        if (sourceCell != null)
                        {
                            NPOI.SS.UserModel.ICell newCell = newRow.CreateCell(j);
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
        public void RemoveMergedCellRegions(ICellStyle style)
        {
            throw new NotImplementedException();
        }

        public bool IsMergeCell(ISheet ws, int rowIndex, int colIndex)
        {
            throw new NotImplementedException();
        }
    }
}
