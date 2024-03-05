namespace EDR_Report.Models.CPM
{
    public class CPM_PROGRAMME_LISTModel
    {
        /// <summary>ID</summary>
        public int? PRO_ID { get; set; }
        /// <summary>程式</summary>
        public string? PRO_NAME { get; set; }
        /// <summary>程式說明</summary>
        public string? PRO_DESC { get; set; }
        public string? URL { get; set; }
        /// <summary>流程圖用(日報作業),若是管理報表程式欄位需空白</summary>
        public string? GROUPKIND0 { get; set; }
        /// <summary>流程圖用(合約作業),若是管理報表程式欄位需空白</summary>
        public string? GROUPKIND1 { get; set; }
        /// <summary>營管用vs,管報為功能群組名稱,同一功能群組要一致,系統才會帶出</summary>
        public string? GROUPKIND2 { get; set; }
        /// <summary>層級,0,1,2/營管(0/1/2)3層,管報(0,1)2層</summary>
        public decimal? GROUPKIND_RANK2 { get; set; }
        /// <summary>顯示狀態(0:主MENU不顯示 1:主MENU要顯示 2:功能刪除或作廢)</summary>
        public string? FLAG { get; set; }
        /// <summary>備註</summary>
        public string? REMARK { get; set; }
        public DateTime? LAST_UPDATE_DATE { get; set; }
        public decimal? LAST_UPDATED_BY { get; set; }
        public DateTime? CREATION_DATE { get; set; }
        public decimal? CREATED_BY { get; set; }
        public string? TYPE { get; set; }
        public string? PROGRAM_NAME { get; set; }
        /// <summary>groupkind2管報表頭順序</summary>
        public decimal? GROUPKIND2_SORT { get; set; }
        public string? KIND { get; set; }
        /// <summary>0:預設內嵌 1:另開視窗 2:另開分頁</summary>
        public string? OPENKIND { get; set; }
        /// <summary>IIS站台名稱</summary>
        public string? DOMAINNAME { get; set; }
        public string? TRANS_FLAG { get; set; }
        /// <summary>程式menu路徑</summary>
        public string? TRANS1_NAME { get; set; }
        public string? ITCHECK { get; set; }
        public string? GROUPKIND2_SORT1 { get; set; }
    }
}
