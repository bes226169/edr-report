using EDR_Report.Commons;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Specialized;

namespace EDR_Report.Interfaces
{
    public interface IRptXlsHelper
    {
        public string GetFormattedCellValue(ICell cellObject, IFormulaEvaluator eval = null);
        public IWorkbook OpenExcelTemplate(string templateFullpath, out IFormulaEvaluator ObjectIFormulaEvaluator);
        public string TaiwanDateTime(string report_date, TaiwanDateFormatEnum dateformat);
        public bool TaiwanDateTimeReform(string taiwan_datetime, string datetime_format, out string TaiDate);

        public void MergeCellsOnRows(IRow row, int rowCount, int startCol, int endCol);
        public string[] GetColumnValues(ISheet sheet, int columnNum);
        public void CopyRowDown(ISheet sheet, int rowIndexStart, int repeatTimes);

        public IRow FindRowByCellValue(ISheet ws, string varName, int starCol, int endCol, int startRow, int endRow);

        public IRow[] FindRowByValue(ISheet ws, string varName, int starCol, int endCol, int startRow, int endRow);
        public void Replace(ISheet ws, string varName, string value, int StartCol, int EndCol, int StartRow, int EndRow);

        public void CopyRow(ISheet worksheet, int sourceRowNum, int destinationRowNum, int shiftCounts);
        public void InsertRows(ISheet ws, int startRow, int rowCount);
        public void SetCellValues(IRow row, List<string> value, List<int> colIdx = null);

        public void CreateCell(IRow currentRow, int col, string value, HSSFCellStyle hssfStyle = null, XSSFCellStyle xssStyle = null);

        public void SetCellStyleOnRow(ISheet ws, int rowIdx, int colIndex, bool createRow = false, bool defaultStyle = false);

        public ICellStyle GetPreferredCellStyle(ICell cell);

        public void SetMergedCellRegions(ISheet ws, ICellStyle style);
        public void RemoveMergedCellRegions(ICellStyle style);
        public ICell GetFirstCellInMergedRegionContainingCell(ICell cell);
        public List<string> ReadRow(IRow row, int starCol, int endCol);
        public string GetCellValue(ICell cell);
        public bool IsMergeCell(ISheet ws, int rowIndex, int colIndex);
        public string ReplacePlaceHolders(string template, Dictionary<string, string> placeHoldersAndValues);
        public ICell MergedCell(ICell cell);
        public void SetCellValue(ICell cell, object Value);

        public void SetRow(IRow row, ListDictionary cellvalues);
        public void SetRow(ISheet ws, int rownum, ListDictionary cellvalues);
    }
}
