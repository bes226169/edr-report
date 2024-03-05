using EDR_Report.Commons;

namespace EDR_Report.Interfaces
{
    public interface I1CB004DbContext
    {
        public IEnumerable<dynamic> QueryWorkProgress(int project_id, string report_date);
        public IEnumerable<dynamic> QueryDailyWorkNotes(int project_id, string report_date);
        public IEnumerable<dynamic> QueryWorkItems(int project_id, WorkItemsEnum workitem);
        public IEnumerable<dynamic> QueryResourceItems(int project_id, string report_date, ResourceClassEnum resClass);
    }
}
