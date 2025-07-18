using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbReportType
    {
        public string Month { get; set; } = null!;
        public string ReportTypeCode { get; set; } = null!;
        public string? Commentary { get; set; }

        public virtual PrbReportTypeCode ReportTypeCodeNavigation { get; set; } = null!;
    }
}
