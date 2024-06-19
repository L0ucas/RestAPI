using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.SNOW
{
    public class SnowResolveTicketRq
    {
        public string number { get; set; }
        public string sys_id { get; set; }
        public string resolution_code { get; set; }
        public string resolution_notes { get; set; }
        public string comments { get; set; }
        public string assigned_to { get; set; }
        public string u_resolved_by { get; set; }
        public List<SnowResolveTicketAttachment> attachment { get; set; }

    }
    public class SnowResolveTicketAttachment
    {
        public string fileName { get; set; }
        public string contentType { get; set; }
        public string content { get; set; }

    }
}
