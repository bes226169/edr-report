namespace EDR_Report.Models.CPM
{
    public class CPM_PERSONALIZEModel
    {
        /// <summary>工號</summary>
        public string? EMPNO { get; set; }
        /// <summary>Layout Style</summary>
        public string? LAYOUT { get; set; }
        /// <summary>方向</summary>
        public string? ORIENTATION { get; set; }
        /// <summary>Header Style</summary>
        public string? HEAD { get; set; }
        /// <summary>Footer Style</summary>
        public string? FOOTER { get; set; }
        /// <summary>背景顏色</summary>
        public string? BACKGROUNDS { get; set; }
        /// <summary>Color Schemes</summary>
        public string? COLOR { get; set; }
        /// <summary>建立者(工號)</summary>
        public string? CREATED_BY { get; set; }
        /// <summary>更新日期</summary>
        public DateTime? UPDATE_DATE { get; set; }
        /// <summary>語系</summary>
        public string? LANGUAGE { get; set; }
    }
}
