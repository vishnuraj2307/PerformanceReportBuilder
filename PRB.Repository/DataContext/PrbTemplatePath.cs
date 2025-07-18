using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbTemplatePath
    {
        public PrbTemplatePath()
        {
            PrbTickers = new HashSet<PrbTicker>();
        }

        public int TemplateId { get; set; }
        public string? FileName { get; set; }
        public string? FilePath { get; set; }
        public DateTime ExpiryDate { get; set; }

        public virtual ICollection<PrbTicker> PrbTickers { get; set; }
    }
}
