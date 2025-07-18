using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbCurrencyCode
    {
        public PrbCurrencyCode()
        {
            PrbHoldingDetails = new HashSet<PrbHoldingDetail>();
        }

        public string CurrencyCode { get; set; } = null!;
        public string CurrencyDesc { get; set; } = null!;

        public virtual ICollection<PrbHoldingDetail> PrbHoldingDetails { get; set; }
    }
}
