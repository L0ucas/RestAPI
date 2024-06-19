using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TMHelper.Common;
using TMHelper.Common.Config;

namespace TMHelper.SNOW
{
    public static class SNOWHelper
    {
        #region Constant
        public const string CONST_USP_SNOW_TICKET_DATA_UPD_INS = "UspSnowTicketDataUpdIns";
        public const string CONST_USP_SNOW_TICKET_USER_DATA_UPD_INS = "UspSnowTicketUserDataUpdIns";
        public const string CONST_USP_SNOW_TICKET_ATTACHMENT_DATA_UPD_INS = "UspSnowTicketAttachmentDataUpdIns";
        public const string CONST_USP_SNOW_TICKET_DATA_SEL = "UspSnowTicketDataSel";
        public const string CONST_USP_SNOW_TICKET_DATA_DEL = "UspSnowTicketDataDel";
        #endregion

        public static long UpsertSnowTicketData(List<SnowTicketDataPumpRq> requestDataList, out string strExceptionMessage)
        {
            long lngResult = 0;
            DataSet dsResult = null;

            strExceptionMessage = null;

            try
            {
                #region Prepare input parameter
                DataSet dsTVPTicketData = new DataSet();
                DataSet dsTVPUserData = new DataSet();
                DataSet dsTVPAttData = new DataSet();

                #region SNOW Ticket Data
                DataTable dtblTVPTicketData = dsTVPTicketData.Tables.Add("SnowTicketDataTable");
                dtblTVPTicketData.Columns.Add("Number", typeof(string));
                dtblTVPTicketData.Columns.Add("RaisedBy", typeof(string));
                dtblTVPTicketData.Columns.Add("RaisedById", typeof(string));
                dtblTVPTicketData.Columns.Add("RaisedFor", typeof(string));
                dtblTVPTicketData.Columns.Add("RaisedForId", typeof(string));
                dtblTVPTicketData.Columns.Add("Category", typeof(string));
                dtblTVPTicketData.Columns.Add("SubCategory", typeof(string));
                dtblTVPTicketData.Columns.Add("CatalogItem", typeof(string));
                dtblTVPTicketData.Columns.Add("ContactType", typeof(string));
                dtblTVPTicketData.Columns.Add("CaseType", typeof(string));
                dtblTVPTicketData.Columns.Add("ShortDescription", typeof(string));
                dtblTVPTicketData.Columns.Add("Description", typeof(string));
                dtblTVPTicketData.Columns.Add("Opened", typeof(DateTime));
                dtblTVPTicketData.Columns.Add("State", typeof(string));
                dtblTVPTicketData.Columns.Add("Priority", typeof(string));
                dtblTVPTicketData.Columns.Add("AssignmentGroup", typeof(string));
                dtblTVPTicketData.Columns.Add("AssignedTo", typeof(string));
                dtblTVPTicketData.Columns.Add("AssignedToId", typeof(string));
                dtblTVPTicketData.Columns.Add("DFMReferenceID", typeof(string));
                dtblTVPTicketData.Columns.Add("SAPReferenceID", typeof(string));
                dtblTVPTicketData.Columns.Add("IsActionComplete", typeof(bool));
                dtblTVPTicketData.Columns.Add("ActionMsg", typeof(string));
                #endregion

                #region SNOW Ticket User Data
                DataTable dtblTVPTUData = dsTVPUserData.Tables.Add("SnowTicketUserDataTable");
                dtblTVPTUData.Columns.Add("Id", typeof(string));
                dtblTVPTUData.Columns.Add("Number", typeof(string));
                dtblTVPTUData.Columns.Add("UserID", typeof(string));
                dtblTVPTUData.Columns.Add("UserType", typeof(string));
                dtblTVPTUData.Columns.Add("EmployeeNo", typeof(string));
                dtblTVPTUData.Columns.Add("Company", typeof(string));
                dtblTVPTUData.Columns.Add("APPCompanyCode", typeof(string));
                dtblTVPTUData.Columns.Add("EntityName", typeof(string));
                dtblTVPTUData.Columns.Add("PersonalArea", typeof(string));
                dtblTVPTUData.Columns.Add("Department", typeof(string));
                dtblTVPTUData.Columns.Add("Mill", typeof(string));
                dtblTVPTUData.Columns.Add("MillHeadID", typeof(string));
                dtblTVPTUData.Columns.Add("Manager", typeof(string));
                dtblTVPTUData.Columns.Add("ManagerID", typeof(string));
                dtblTVPTUData.Columns.Add("Email", typeof(string));
                dtblTVPTUData.Columns.Add("VVIP", typeof(string));
                dtblTVPTUData.Columns.Add("VIP", typeof(string));
                dtblTVPTUData.Columns.Add("CurrentPhysicalLocation", typeof(string));
                #endregion

                #region SNOW Ticket Attachment Data
                DataTable dtblTVPTAData = dsTVPAttData.Tables.Add("SnowTicketAttachmentDataTable");
                dtblTVPTAData.Columns.Add("Number", typeof(string));
                dtblTVPTAData.Columns.Add("ContentType", typeof(string));
                dtblTVPTAData.Columns.Add("FileName", typeof(string));
                dtblTVPTAData.Columns.Add("Base64", typeof(string));
                #endregion
                #endregion

                #region Assign data
                foreach (var requestData in requestDataList)
                {
                    DataRow drTVPTicketData = dtblTVPTicketData.NewRow();
                    drTVPTicketData["Number"] = requestData.Number;
                    drTVPTicketData["RaisedBy"] = requestData.RaisedBy;
                    drTVPTicketData["RaisedFor"] = requestData.RaisedFor;
                    drTVPTicketData["Category"] = requestData.Category;
                    drTVPTicketData["SubCategory"] = requestData.SubCategory;
                    drTVPTicketData["CatalogItem"] = requestData.CatalogItem;
                    drTVPTicketData["ContactType"] = requestData.ContactType;
                    drTVPTicketData["CaseType"] = requestData.CaseType;
                    drTVPTicketData["ShortDescription"] = requestData.ShortDescription;
                    drTVPTicketData["Description"] = requestData.Description;
                    drTVPTicketData["Opened"] = requestData.OpenedDt;
                    drTVPTicketData["State"] = requestData.State;
                    drTVPTicketData["Priority"] = requestData.Priority;
                    drTVPTicketData["AssignmentGroup"] = requestData.AssignmentGroup;
                    drTVPTicketData["AssignedTo"] = requestData.AssignedTo;
                    drTVPTicketData["DFMReferenceID"] = requestData.DFMReferenceID;
                    drTVPTicketData["SAPReferenceID"] = requestData.SAPReferenceID;
                    drTVPTicketData["IsActionComplete"] = requestData.IsActionComplete;
                    drTVPTicketData["ActionMsg"] = requestData.ActionMsg;

                    #region RaisedByInfo
                    if (requestData.RaisedByInfo != null)
                    {
                        string strRaisedByInfoId = Guid.NewGuid().ToString();

                        drTVPTicketData["RaisedById"] = strRaisedByInfoId;

                        DataRow drTVPTUData = dtblTVPTUData.NewRow();
                        drTVPTUData["Id"] = strRaisedByInfoId;
                        drTVPTUData["Number"] = requestData.Number;
                        drTVPTUData["UserID"] = requestData.RaisedByInfo.UserID;
                        drTVPTUData["UserType"] = UserType.RaisedBy.ToString();
                        drTVPTUData["EmployeeNo"] = requestData.RaisedByInfo.EmployeeNo;
                        drTVPTUData["Company"] = requestData.RaisedByInfo.Company;
                        drTVPTUData["APPCompanyCode"] = requestData.RaisedByInfo.APPCompanyCode;
                        drTVPTUData["EntityName"] = requestData.RaisedByInfo.EntityName;
                        drTVPTUData["PersonalArea"] = requestData.RaisedByInfo.PersonalArea;
                        drTVPTUData["Department"] = requestData.RaisedByInfo.Department;
                        drTVPTUData["Mill"] = requestData.RaisedByInfo.Mill;
                        drTVPTUData["MillHeadID"] = requestData.RaisedByInfo.MillHeadID;
                        drTVPTUData["Manager"] = requestData.RaisedByInfo.Manager;
                        drTVPTUData["ManagerID"] = requestData.RaisedByInfo.ManagerID;
                        drTVPTUData["Email"] = requestData.RaisedByInfo.Email;
                        drTVPTUData["VVIP"] = requestData.RaisedByInfo.VVIP;
                        drTVPTUData["VIP"] = requestData.RaisedByInfo.VIP;
                        drTVPTUData["CurrentPhysicalLocation"] = requestData.RaisedByInfo.CurrentPhysicalLocation;

                        dtblTVPTUData.Rows.Add(drTVPTUData);
                    }
                    #endregion

                    #region RaisedForInfo
                    if (requestData.RaisedForInfo != null)
                    {
                        string strRaisedForInfoId = Guid.NewGuid().ToString();

                        drTVPTicketData["RaisedForId"] = strRaisedForInfoId;

                        DataRow drTVPTUData = dtblTVPTUData.NewRow();
                        drTVPTUData["Id"] = strRaisedForInfoId;
                        drTVPTUData["Number"] = requestData.Number;
                        drTVPTUData["UserID"] = requestData.RaisedForInfo.UserID;
                        drTVPTUData["UserType"] = UserType.RaisedFor.ToString();
                        drTVPTUData["EmployeeNo"] = requestData.RaisedForInfo.EmployeeNo;
                        drTVPTUData["Company"] = requestData.RaisedForInfo.Company;
                        drTVPTUData["APPCompanyCode"] = requestData.RaisedForInfo.APPCompanyCode;
                        drTVPTUData["EntityName"] = requestData.RaisedForInfo.EntityName;
                        drTVPTUData["PersonalArea"] = requestData.RaisedForInfo.PersonalArea;
                        drTVPTUData["Department"] = requestData.RaisedForInfo.Department;
                        drTVPTUData["Mill"] = requestData.RaisedForInfo.Mill;
                        drTVPTUData["MillHeadID"] = requestData.RaisedForInfo.MillHeadID;
                        drTVPTUData["Manager"] = requestData.RaisedForInfo.Manager;
                        drTVPTUData["ManagerID"] = requestData.RaisedForInfo.ManagerID;
                        drTVPTUData["Email"] = requestData.RaisedForInfo.Email;
                        drTVPTUData["VVIP"] = requestData.RaisedForInfo.VVIP;
                        drTVPTUData["VIP"] = requestData.RaisedForInfo.VIP;
                        drTVPTUData["CurrentPhysicalLocation"] = requestData.RaisedForInfo.CurrentPhysicalLocation;

                        dtblTVPTUData.Rows.Add(drTVPTUData);
                    }
                    #endregion

                    #region AssignedToInfo
                    if (requestData.AssignedToInfo != null)
                    {
                        string strAssignedToInfoId = Guid.NewGuid().ToString();

                        drTVPTicketData["AssignedToId"] = strAssignedToInfoId;

                        DataRow drTVPTUData = dtblTVPTUData.NewRow();
                        drTVPTUData["Id"] = strAssignedToInfoId;
                        drTVPTUData["Number"] = requestData.Number;
                        drTVPTUData["UserID"] = requestData.AssignedToInfo.UserID;
                        drTVPTUData["UserType"] = UserType.AssignedTo.ToString();
                        drTVPTUData["EmployeeNo"] = requestData.AssignedToInfo.EmployeeNo;
                        drTVPTUData["Company"] = requestData.AssignedToInfo.Company;
                        drTVPTUData["APPCompanyCode"] = requestData.AssignedToInfo.APPCompanyCode;
                        drTVPTUData["EntityName"] = requestData.AssignedToInfo.EntityName;
                        drTVPTUData["PersonalArea"] = requestData.AssignedToInfo.PersonalArea;
                        drTVPTUData["Department"] = requestData.AssignedToInfo.Department;
                        drTVPTUData["Mill"] = requestData.AssignedToInfo.Mill;
                        drTVPTUData["MillHeadID"] = requestData.AssignedToInfo.MillHeadID;
                        drTVPTUData["Manager"] = requestData.AssignedToInfo.Manager;
                        drTVPTUData["ManagerID"] = requestData.AssignedToInfo.ManagerID;
                        drTVPTUData["Email"] = requestData.AssignedToInfo.Email;
                        drTVPTUData["VVIP"] = requestData.AssignedToInfo.VVIP;
                        drTVPTUData["VIP"] = requestData.AssignedToInfo.VIP;
                        drTVPTUData["CurrentPhysicalLocation"] = requestData.AssignedToInfo.CurrentPhysicalLocation;

                        dtblTVPTUData.Rows.Add(drTVPTUData);
                    }
                    #endregion

                    #region AttachInfo
                    if (requestData.AttachInfo != null)
                    {
                        DataRow drTVPTAData = dtblTVPTAData.NewRow();
                        drTVPTAData["Number"] = requestData.Number;
                        drTVPTAData["ContentType"] = requestData.AttachInfo.ContentType;
                        drTVPTAData["FileName"] = requestData.AttachInfo.FileName;
                        drTVPTAData["Base64"] = requestData.AttachInfo.Base64;

                        dtblTVPTAData.Rows.Add(drTVPTAData);
                    }
                    #endregion

                    dtblTVPTicketData.Rows.Add(drTVPTicketData);
                }
                #endregion

                #region Call store procedure
                DatabaseHelper dbHelper = new DatabaseHelper(@"Data Source = 10.89.67.19;Initial Catalog=APP_BPODB;User ID=APP_BPODB_SA;Password=3W*fY-)U;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

                if (lngResult == 0 && dsTVPTicketData != null && dsTVPTicketData.Tables != null && dsTVPTicketData.Tables.Count > 0 && dsTVPTicketData.Tables[0].Rows != null && dsTVPTicketData.Tables[0].Rows.Count > 0)
                {
                    lngResult = dbHelper.ExecuteStoreProcedureWithTVP(CONST_USP_SNOW_TICKET_DATA_UPD_INS, dsTVPTicketData, out dsResult, out strExceptionMessage);
                    if (lngResult != 0) { strExceptionMessage = $"Failed to update/insert SNOW ticket data. {strExceptionMessage}"; }
                }

                if (lngResult == 0 && dsTVPUserData != null && dsTVPUserData.Tables != null && dsTVPUserData.Tables.Count > 0 && dsTVPUserData.Tables[0].Rows != null && dsTVPUserData.Tables[0].Rows.Count > 0)
                {
                    lngResult = dbHelper.ExecuteStoreProcedureWithTVP(CONST_USP_SNOW_TICKET_USER_DATA_UPD_INS, dsTVPUserData, out dsResult, out strExceptionMessage);
                    if (lngResult != 0) { strExceptionMessage = $"Failed to update/insert SNOW ticket user data. {strExceptionMessage}"; }
                }

                if (lngResult == 0 && dsTVPAttData != null && dsTVPAttData.Tables != null && dsTVPAttData.Tables.Count > 0 && dsTVPAttData.Tables[0].Rows != null && dsTVPAttData.Tables[0].Rows.Count > 0)
                {
                    lngResult = dbHelper.ExecuteStoreProcedureWithTVP(CONST_USP_SNOW_TICKET_ATTACHMENT_DATA_UPD_INS, dsTVPAttData, out dsResult, out strExceptionMessage);
                    if (lngResult != 0) { strExceptionMessage = $"Failed to update/insert SNOW ticket attachment data. {strExceptionMessage}"; }
                }
                #endregion
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }

        public static long ExtractSnowTicketData(out List<SnowTicketDataPumpRq> listSnowTicketData, out string strExceptionMessage)
        {
            long lngResult = -1;
            DataSet dsResult = null;

            listSnowTicketData = null;
            strExceptionMessage = null;

            try
            {
                #region Call store procedure 
                DatabaseHelper dbHelper = new DatabaseHelper(@"Data Source = 10.89.67.19;Initial Catalog=APP_BPODB;User ID=APP_BPODB_SA;Password=3W*fY-)U;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
                lngResult = dbHelper.ExecuteStoreProcedure(CONST_USP_SNOW_TICKET_DATA_SEL, out dsResult, out strExceptionMessage);
                #endregion

                if (lngResult == 0)
                {
                    if (dsResult != null && dsResult.Tables != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows != null)
                    {
                        listSnowTicketData = new List<SnowTicketDataPumpRq>();
                        foreach (DataRow drResult in dsResult.Tables[0].Rows)
                        {
                            string strAttContentType = drResult["ContentType"].ToString();
                            string strAttFileName = drResult["FileName"].ToString();
                            string strAttBase64 = drResult["Base64"].ToString();

                            SnowTicketDataPumpRq dataReq = new SnowTicketDataPumpRq();
                            dataReq.Number = drResult["Number"].ToString();
                            dataReq.RaisedBy = drResult["RaisedBy"].ToString();
                            dataReq.RaisedFor = drResult["RaisedFor"].ToString();
                            dataReq.Category = drResult["Category"].ToString();
                            dataReq.SubCategory = drResult["SubCategory"].ToString();
                            dataReq.CatalogItem = drResult["CatalogItem"].ToString();
                            dataReq.ContactType = drResult["ContactType"].ToString();
                            dataReq.CaseType = drResult["CaseType"].ToString();
                            dataReq.ShortDescription = drResult["ShortDescription"].ToString();
                            dataReq.Description = drResult["Description"].ToString();
                            dataReq.Opened = drResult["Opened"].ToString();
                            dataReq.State = drResult["State"].ToString();
                            dataReq.Priority = drResult["Priority"].ToString();
                            dataReq.AssignmentGroup = drResult["AssignmentGroup"].ToString();
                            dataReq.AssignedTo = drResult["AssignedTo"].ToString();
                            dataReq.DFMReferenceID = drResult["DFMReferenceID"].ToString();
                            dataReq.SAPReferenceID = drResult["SAPReferenceID"].ToString();
                            dataReq.IsActionComplete = (bool)drResult["IsActionComplete"];
                            dataReq.ActionMsg = drResult["ActionMsg"].ToString();

                            #region AttachInfo
                            if (!string.IsNullOrWhiteSpace(strAttContentType) && !string.IsNullOrWhiteSpace(strAttFileName) && !string.IsNullOrWhiteSpace(strAttBase64))
                            {
                                dataReq.AttachInfo = new AttachmentData();
                                dataReq.AttachInfo.ContentType = strAttContentType;
                                dataReq.AttachInfo.FileName = strAttFileName;
                                dataReq.AttachInfo.Base64 = strAttBase64;
                            }
                            #endregion

                            #region UserData
                            if (dsResult.Tables.Count > 1 && dsResult.Tables[1] != null && dsResult.Tables[1].Rows != null && dsResult.Tables[1].Rows.Count > 0)
                            {
                                DataRow[] arrUserForThisNumber = dsResult.Tables[1].Select($"Number = '{dataReq.Number}'");
                                if (arrUserForThisNumber != null && arrUserForThisNumber.Length > 0)
                                {
                                    foreach (DataRow drUser in arrUserForThisNumber)
                                    {
                                        string strUserType = drUser["UserType"].ToString();
                                        if (!string.IsNullOrWhiteSpace(strUserType))
                                        {
                                            #region RaisedByInfo
                                            if (strUserType.Trim().Equals(UserType.RaisedBy.ToString()))
                                            {
                                                dataReq.RaisedByInfo = new UserData();
                                                dataReq.RaisedByInfo.UserID = drUser["UserID"].ToString();
                                                dataReq.RaisedByInfo.EmployeeNo = drUser["EmployeeNo"].ToString();
                                                dataReq.RaisedByInfo.Company = drUser["Company"].ToString();
                                                dataReq.RaisedByInfo.APPCompanyCode = drUser["APPCompanyCode"].ToString();
                                                dataReq.RaisedByInfo.EntityName = drUser["EntityName"].ToString();
                                                dataReq.RaisedByInfo.PersonalArea = drUser["PersonalArea"].ToString();
                                                dataReq.RaisedByInfo.Department = drUser["Department"].ToString();
                                                dataReq.RaisedByInfo.Mill = drUser["Mill"].ToString();
                                                dataReq.RaisedByInfo.MillHeadID = drUser["MillHeadID"].ToString();
                                                dataReq.RaisedByInfo.Manager = drUser["Manager"].ToString();
                                                dataReq.RaisedByInfo.ManagerID = drUser["ManagerID"].ToString();
                                                dataReq.RaisedByInfo.Email = drUser["Email"].ToString();
                                                dataReq.RaisedByInfo.VVIP = (bool)drUser["VVIP"];
                                                dataReq.RaisedByInfo.VIP = (bool)drUser["VIP"];
                                                dataReq.RaisedByInfo.CurrentPhysicalLocation = drUser["CurrentPhysicalLocation"].ToString();
                                            }
                                            #endregion

                                            #region RaisedForInfo
                                            if (strUserType.Trim().Equals(UserType.RaisedFor.ToString()))
                                            {
                                                dataReq.RaisedForInfo = new UserData();
                                                dataReq.RaisedForInfo.UserID = drUser["UserID"].ToString();
                                                dataReq.RaisedForInfo.EmployeeNo = drUser["EmployeeNo"].ToString();
                                                dataReq.RaisedForInfo.Company = drUser["Company"].ToString();
                                                dataReq.RaisedForInfo.APPCompanyCode = drUser["APPCompanyCode"].ToString();
                                                dataReq.RaisedForInfo.EntityName = drUser["EntityName"].ToString();
                                                dataReq.RaisedForInfo.PersonalArea = drUser["PersonalArea"].ToString();
                                                dataReq.RaisedForInfo.Department = drUser["Department"].ToString();
                                                dataReq.RaisedForInfo.Mill = drUser["Mill"].ToString();
                                                dataReq.RaisedForInfo.MillHeadID = drUser["MillHeadID"].ToString();
                                                dataReq.RaisedForInfo.Manager = drUser["Manager"].ToString();
                                                dataReq.RaisedForInfo.ManagerID = drUser["ManagerID"].ToString();
                                                dataReq.RaisedForInfo.Email = drUser["Email"].ToString();
                                                dataReq.RaisedForInfo.VVIP = (bool)drUser["VVIP"];
                                                dataReq.RaisedForInfo.VIP = (bool)drUser["VIP"];
                                                dataReq.RaisedForInfo.CurrentPhysicalLocation = drUser["CurrentPhysicalLocation"].ToString();
                                            }
                                            #endregion

                                            #region AssignedToInfo
                                            if (strUserType.Trim().Equals(UserType.AssignedTo.ToString()))
                                            {
                                                dataReq.AssignedToInfo = new UserData();
                                                dataReq.AssignedToInfo.UserID = drUser["UserID"].ToString();
                                                dataReq.AssignedToInfo.EmployeeNo = drUser["EmployeeNo"].ToString();
                                                dataReq.AssignedToInfo.Company = drUser["Company"].ToString();
                                                dataReq.AssignedToInfo.APPCompanyCode = drUser["APPCompanyCode"].ToString();
                                                dataReq.AssignedToInfo.EntityName = drUser["EntityName"].ToString();
                                                dataReq.AssignedToInfo.PersonalArea = drUser["PersonalArea"].ToString();
                                                dataReq.AssignedToInfo.Department = drUser["Department"].ToString();
                                                dataReq.AssignedToInfo.Mill = drUser["Mill"].ToString();
                                                dataReq.AssignedToInfo.MillHeadID = drUser["MillHeadID"].ToString();
                                                dataReq.AssignedToInfo.Manager = drUser["Manager"].ToString();
                                                dataReq.AssignedToInfo.ManagerID = drUser["ManagerID"].ToString();
                                                dataReq.AssignedToInfo.Email = drUser["Email"].ToString();
                                                dataReq.AssignedToInfo.VVIP = (bool)drUser["VVIP"];
                                                dataReq.AssignedToInfo.VIP = (bool)drUser["VIP"];
                                                dataReq.AssignedToInfo.CurrentPhysicalLocation = drUser["CurrentPhysicalLocation"].ToString();
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                            #endregion

                            if (!string.IsNullOrWhiteSpace(dataReq.Opened) && DateTime.TryParse(dataReq.Opened.Trim(), out DateTime dtOpened))
                                dataReq.OpenedDt = dtOpened;

                            listSnowTicketData.Add(dataReq);
                        }
                    }
                    else
                    {
                        throw new Exception("Dataset or datatable return from store procedure is null");
                    }
                }
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }

        public static long HousekeepSnowTicketData(int iMaxRetainDay, out string strExceptionMessage)
        {
            long lngResult = -1;
            DataSet dsResult = null;

            strExceptionMessage = null;

            try
            {
                #region Prepare input parameter
                DataTable dtblInput = new DataTable();
                dtblInput.Columns.Add("MaxRetainDay", typeof(int));

                DataRow drInput = dtblInput.NewRow();
                drInput["MaxRetainDay"] = iMaxRetainDay;
                dtblInput.Rows.Add(drInput);
                #endregion

                #region Call store procedure 
                DatabaseHelper dbHelper = new DatabaseHelper(@"Data Source = 10.89.67.19;Initial Catalog=APP_BPODB;User ID=APP_BPODB_SA;Password=3W*fY-)U;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
                lngResult = dbHelper.ExecuteStoreProcedureWithParameters(CONST_USP_SNOW_TICKET_DATA_DEL, dtblInput, out dsResult, out strExceptionMessage);
                #endregion
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }
    }
}
