using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbReportTypeCode
    {
        public PrbReportTypeCode()
        {
            PrbReportTypes = new HashSet<PrbReportType>();
        }

        public string ReportTypeCode { get; set; } = null!;
        public string ReportTypeDesc { get; set; } = null!;

        public virtual ICollection<PrbReportType> PrbReportTypes { get; set; }
    }
}
