using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbReportStatus
    {
        public string ReportMonth { get; set; } = null!;
        public DateTime ReportDate { get; set; }
        public string RoleCode { get; set; } = null!;
        public string ReportStatusCode { get; set; } = null!;
        public string? Comments { get; set; }

        public virtual PrbReportStatusCode ReportStatusCodeNavigation { get; set; } = null!;
        public virtual PrbRoleCode RoleCodeNavigation { get; set; } = null!;
    }
}
