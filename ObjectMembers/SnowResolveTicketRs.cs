using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.SNOW
{
    public class SnowResolveTicketRs
    {
        public SnowResolveTicketRsResult result { get; set; }
    }
    public class SnowResolveTicketRsResult
    {
        public ResolveTicket ticket { get; set; }
    }
    public class ResolveTicket
    {
        public string update { get; set; }
        public string number { get; set; }
        public string updated_on { get; set; }
        public string attachment_result { get; set; }
    }
}
