namespace EDR_Report.Models.ERP
{
    public class VS_USERSModel
    {
        public int? USER_ID { get; set; }
        public string? USER_NAME { get; set; }
        public int? ORGANIZATION_ID { get; set; }
        public string? PASSWORD { get; set; }
        public int? CONTACT_ID { get; set; }
        public DateTime? LAST_UPDATE_DATE { get; set; }
        public int? LAST_UPDATED_BY { get; set; }
        public DateTime? CREATION_DATE { get; set; }
        public int? CREATED_BY { get; set; }
        public string? SYSTEM_DSN { get; set; }
        public string? CH_NAME { get; set; }
        public string? PORTALPAS { get; set; }
        public string? STOPFLAG { get; set; }
        public DateTime? STOPTIME { get; set; }
        public int? LOGINTIME { get; set; }
        public string? OLD_PASSWORD { get; set; }
        public DateTime? UPDATE_PASS_DATE { get; set; }

        public string? PORTALPAS_DEC { get; set; }
    }
}
