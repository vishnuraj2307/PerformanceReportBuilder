using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PRB.Domain.Model1
{
    public class UpdateReportStatus
    {
        public string? ReportMonth { get; set; }
        public DateTime? ReportDate { get; set; }
        public string? RoleCode { get; set; } 
        public string? ReportStatusCode { get; set; }
        public string? Comments { get; set; }
     
      
    }
}
