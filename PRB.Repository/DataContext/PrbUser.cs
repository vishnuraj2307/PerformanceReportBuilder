using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbUser
    {
        public string UserMailId { get; set; } = null!;
        public string UserName { get; set; } = null!;
        public string Password { get; set; } = null!;
        public string RoleCode { get; set; } = null!;

        public virtual PrbRoleCode RoleCodeNavigation { get; set; } = null!;
    }
}
