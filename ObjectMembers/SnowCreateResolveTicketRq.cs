using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.SNOW
{
    public class SnowCreateResolveTicketRq
    {
        public string category { get; set; }
        public string subcategory { get; set; }
        public int recordno { get; set; }
        public string description { get; set; }
        public SnowCreateResolveTicketRqResolution resolution { get; set; }
        public SnowCreateResolveTicketRqNotes notes { get; set; }
    }

    public class SnowCreateResolveTicketRqResolution
    {
        public string code { get; set; }
        public string notes { get; set; }
    }

    public class SnowCreateResolveTicketRqNotes
    {
        public string work { get; set; }
        public string comment { get; set; }
    }
}
