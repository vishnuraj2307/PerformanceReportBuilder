using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbRoleCode
    {
        public PrbRoleCode()
        {
            PrbReportStatuses = new HashSet<PrbReportStatus>();
            PrbUsers = new HashSet<PrbUser>();
        }

        public string RoleCode { get; set; } = null!;
        public string RoleDesc { get; set; } = null!;

        public virtual ICollection<PrbReportStatus> PrbReportStatuses { get; set; }
        public virtual ICollection<PrbUser> PrbUsers { get; set; }
    }
}
