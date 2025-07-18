using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbTransactionTypeCode
    {
        public PrbTransactionTypeCode()
        {
            PrbHoldingDetails = new HashSet<PrbHoldingDetail>();
        }

        public string TransactionTypeCode { get; set; } = null!;
        public string TransactionDesc { get; set; } = null!;

        public virtual ICollection<PrbHoldingDetail> PrbHoldingDetails { get; set; }
    }
}
