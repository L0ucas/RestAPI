using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.SNOW
{
    public class SnowTicketDataPumpRq
    {
        public string Number { get; set; }
        public string RaisedBy { get; set; }
        public UserData RaisedByInfo { get; set; }
        public string RaisedFor { get; set; }
        public UserData RaisedForInfo { get; set; }
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string CatalogItem { get; set; }
        public string ContactType { get; set; }
        public string CaseType { get; set; }
        public string ShortDescription { get; set; }
        public string Description { get; set; }
        public string Opened { get; set; }
        public DateTime OpenedDt { get; set; }
        public string State { get; set; }
        public string Priority { get; set; }
        public string AssignmentGroup { get; set; }
        public string AssignedTo { get; set; }
        public UserData AssignedToInfo { get; set; }
        public string DFMReferenceID { get; set; }
        public string SAPReferenceID { get; set; }
        public bool IsActionComplete { get; set; }
        public string ActionMsg { get; set; }
        public AttachmentData AttachInfo { get; set; }
    }

    public class UserData
    {
        public string UserID { get; set; }
        public string EmployeeNo { get; set; }
        public string Company { get; set; }
        public string APPCompanyCode { get; set; }
        public string EntityName { get; set; }
        public string PersonalArea { get; set; }
        public string Department { get; set; }
        public string Mill { get; set; }
        public string MillHeadID { get; set; }
        public string Manager { get; set; }
        public string ManagerID { get; set; }
        public string Email { get; set; }
        public bool VVIP { get; set; }
        public bool VIP { get; set; }
        public string CurrentPhysicalLocation { get; set; }
    }

    public enum UserType
    {
        RaisedBy = 0,
        RaisedFor = 1,
        AssignedTo = 2
    }

    public class AttachmentData
    {
        public string ContentType { get; set; }
        public string FileName { get; set; }
        public string Base64 { get; set; }
    }
}
