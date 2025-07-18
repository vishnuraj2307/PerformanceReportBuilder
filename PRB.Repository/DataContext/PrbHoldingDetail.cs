using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbHoldingDetail
    {
        public string CompanyTicker { get; set; } = null!;
        public DateTime TransactionDate { get; set; }
        public string TransactionTypeCode { get; set; } = null!;
        public int Quantity { get; set; }
        public string CurrencyCode { get; set; } = null!;
        public decimal Amount { get; set; }

        public virtual PrbTicker CompanyTickerNavigation { get; set; } = null!;
        public virtual PrbCurrencyCode CurrencyCodeNavigation { get; set; } = null!;
        public virtual PrbTransactionTypeCode TransactionTypeCodeNavigation { get; set; } = null!;
    }
}
