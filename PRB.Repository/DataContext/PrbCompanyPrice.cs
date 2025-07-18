using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbCompanyPrice
    {
        public string CompanyTicker { get; set; } = null!;
        public DateTime ReportDate { get; set; }
        public decimal LastMarketPrice { get; set; }

        public virtual PrbTicker CompanyTickerNavigation { get; set; } = null!;
    }
}
