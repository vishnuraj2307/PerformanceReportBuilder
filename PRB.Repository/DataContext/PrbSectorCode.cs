using System;
using System.Collections.Generic;

namespace PRB.Repository.DataContext
{
    public partial class PrbSectorCode
    {
        public PrbSectorCode()
        {
            PrbTickers = new HashSet<PrbTicker>();
        }

        public string SectorCode { get; set; } = null!;
        public string SectorName { get; set; } = null!;

        public virtual ICollection<PrbTicker> PrbTickers { get; set; }
    }
}
