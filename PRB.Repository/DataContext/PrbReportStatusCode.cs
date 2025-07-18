using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbReportStatusCode
    {
        public PrbReportStatusCode()
        {
            PrbReportStatuses = new HashSet<PrbReportStatus>();
            PrbReportSummaries = new HashSet<PrbReportSummary>();
        }

        public string ReportStatusCode { get; set; } = null!;
        public string ReportStatusDesc { get; set; } = null!;

        public virtual ICollection<PrbReportStatus> PrbReportStatuses { get; set; }
        public virtual ICollection<PrbReportSummary> PrbReportSummaries { get; set; }
    }
}
