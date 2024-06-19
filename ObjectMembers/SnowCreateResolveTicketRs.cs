using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.SNOW
{
    public class SnowCreateResolveTicketRs
    {
        public SnowCreateResolveTicketRsResult result { get; set; }
    }

    public class SnowCreateResolveTicketRsResult
    {
        public string TicketNumber { get; set; }
        public string Status { get; set; }
        public bool ExistingOA { get; set; }
    }
}
