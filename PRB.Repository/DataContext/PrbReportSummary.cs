using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbReportSummary
    {
        public string ReportMonth { get; set; } = null!;
        public string ReportStatusCode { get; set; } = null!;
        public string FilePassword { get; set; } = null!;

        public virtual PrbReportStatusCode ReportStatusCodeNavigation { get; set; } = null!;
    }
}
