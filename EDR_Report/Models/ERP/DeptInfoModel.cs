namespace EDR_Report.Models.ERP
{
    public class DeptInfoModel
    {
        /// <summary>
        /// 部門代碼
        /// </summary>
        public string? DEPT { get; set; }
        /// <summary>
        /// 部門名稱
        /// </summary>
        public string? DEPT_CH { get; set; }
        public string? DEPT_KND { get; set; }
        /// <summary>
        /// 職務名稱
        /// </summary>
        public string? JOB_CH { get; set; }
        public string? DEPT_TO_99 { get; set; }
        public string? VNDREMP { get; set; }
        public string? VNDR_TYPE { get; set; }
        public string? BPM_DEPT { get; set; }
    }
}