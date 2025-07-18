using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbTicker
    {
        public PrbTicker()
        {
            PrbCompanyPrices = new HashSet<PrbCompanyPrice>();
            PrbHoldingDetails = new HashSet<PrbHoldingDetail>();
        }

        public string CompanyTicker { get; set; } = null!;
        public string CompanyName { get; set; } = null!;
        public string SectorCode { get; set; } = null!;
        public string CompanyDesc { get; set; } = null!;
        public int? TemplateId { get; set; }

        public virtual PrbSectorCode SectorCodeNavigation { get; set; } = null!;
        public virtual PrbTemplatePath? Template { get; set; }
        public virtual ICollection<PrbCompanyPrice> PrbCompanyPrices { get; set; }
        public virtual ICollection<PrbHoldingDetail> PrbHoldingDetails { get; set; }
    }
}
