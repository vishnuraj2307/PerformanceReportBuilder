using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PRB.Domain.Model1
{
    public class UpdateDetailedReportDatas
    {
        [Key]
        public string? CompanyTicker { get; set; }
        public DateTime TransactionDate { get; set; }
        public int Quantity { get; set; }
        public decimal Amount { get; set; }
        public string TransactionTypeCode { get; set; } = null!;
        public string CurrencyCode { get; set; } = null!;


    }
}
