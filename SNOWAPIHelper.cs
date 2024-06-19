using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using TMHelper.Common;
using TMHelper.Common.Config;

namespace TMHelper.SNOW
{
    public class SNOWAPIHelper
    {
        #region Properties
        public string BaseURL { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        #endregion

        #region Constant
        public const string CONST_CREATE_RESOLVE_TICKET_API = "api/aspup/hr_cases/create-resolve-ticket";
        public const string CONST_RESOLVE_TICKET_API = "api/aspup/fa_cases/resolved-ticket";
        #endregion

        #region Constructor
        public SNOWAPIHelper(string strBaseURL, string strUserName, string strPassword)
        {
            this.BaseURL = strBaseURL;
            this.UserName = strUserName;
            this.Password = strPassword;
        }
        #endregion

        public SNOWAPIHelper()
        {

        }

        public bool CreateResolveSnowTicket(SnowCreateResolveTicketRq requestData, out SnowCreateResolveTicketRs responseData, out string strExceptionMessage)
        {
            bool blnIsSuccess = false;
            string strJsonResMsg = null;

            responseData = null;
            strExceptionMessage = null;

            try
            {
                string strAPIUrl = CommonHelpers.AppendToURL(this.BaseURL, CONST_CREATE_RESOLVE_TICKET_API);

                Dictionary<string, string> dicHeaders = new Dictionary<string, string>();
                dicHeaders.Add(HttpRequestHeader.ContentType.ToString(), "application/json");
                dicHeaders.Add(HttpRequestHeader.Authorization.ToString(), "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes($"{this.UserName}:{this.Password}")));

                blnIsSuccess = WebServiceHelper.CallWebService(requestData, strAPIUrl, HttpMethod.Post, dicHeaders, null, out strJsonResMsg, out strExceptionMessage, true);
                if (blnIsSuccess)
                {
                    if (!string.IsNullOrWhiteSpace(strJsonResMsg))
                    {
                        responseData = WebServiceHelper.DeserializeJSon<SnowCreateResolveTicketRs>(strJsonResMsg);
                        if (responseData == null)
                        {
                            blnIsSuccess = false;
                            strExceptionMessage = string.Format("Failed to deserialize JSON reponse message return from {0}", strAPIUrl);
                        }
                    }
                    else
                    {
                        blnIsSuccess = false;
                        strExceptionMessage = string.Format("Empty JSON response message return from {0}", strAPIUrl);
                    }
                }
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }

            return blnIsSuccess;
        }

        public string ResolveSnowTicket(string TicketNumber, string systemID, string resolutionCode, string resolutionNote, string comments, string assignTo, string resolveBy, string FilePath)
        {

            SnowResolveTicketRs responseData = null;
            string strExceptionMsg = null;

            try
            {
                CommonConfigHelper comConfigHelp = new CommonConfigHelper();
                long lngRPAParamResult = comConfigHelp.ExtractRPAParameter(out List<RPAParameter> listParam, out strExceptionMsg, "RESOLVE_SNOW_TICKET");
                long lngRPACredResult = comConfigHelp.ExtractRPACredential("SNOW_API_CRED", out List<RPACredential> listCred, out strExceptionMsg);

                if (lngRPAParamResult != 0)
                    throw new Exception(strExceptionMsg);

                if (lngRPACredResult != 0)
                    throw new Exception(strExceptionMsg);

                if(!(listParam != null && listParam.Count > 0 ))
                    throw new Exception("listParam is Empty");

                if (!(listCred != null && listCred.Count > 0))
                    throw new Exception("listCred is Empty");
                   

                var SnowURLParam = listParam.FirstOrDefault(m => m.ParamKey.Equals("SNOW_URL", StringComparison.OrdinalIgnoreCase));
                if (!(SnowURLParam != null && SnowURLParam != default(RPAParameter)))
                    throw new Exception("ParamKey SNOW_URL Cannot Found");

                var SnowAPIUsername = listCred.FirstOrDefault(m => m.Key.Equals("RESOLVE_SNOW_API", StringComparison.OrdinalIgnoreCase));
                if (!(SnowAPIUsername != null && SnowAPIUsername != default(RPACredential)))
                    throw new Exception("Snow API Username Cannot Found");

                var SnowAPIPassword = listCred.FirstOrDefault(m => m.Key.Equals("RESOLVE_SNOW_API", StringComparison.OrdinalIgnoreCase));
                if (!(SnowAPIPassword != null && SnowAPIPassword != default(RPACredential)))
                    throw new Exception("Snow API Password Cannot Found");

                this.BaseURL = SnowURLParam.ParamValue.ToString();
                this.UserName = SnowAPIUsername.UserName.ToString();
                this.Password = CryptographyHelper.Decrypt(SnowAPIPassword.Password.ToString());


                SnowResolveTicketRq requestData = new SnowResolveTicketRq();


                requestData.number = TicketNumber;
                requestData.sys_id = systemID;
                requestData.resolution_code = resolutionCode;
                requestData.resolution_notes = resolutionNote;
                requestData.comments = comments;
                requestData.assigned_to = assignTo;
                requestData.u_resolved_by = resolveBy;

                requestData.attachment = new List<SnowResolveTicketAttachment>();

                if(!string.IsNullOrWhiteSpace(FilePath))
                {
                    string[] arrFilePath = FilePath.Trim().Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string strFilePath in arrFilePath)
                    {
                        if (File.Exists(strFilePath))
                        {
                            SnowResolveTicketAttachment att = new SnowResolveTicketAttachment();
                            att.fileName = Path.GetFileName(strFilePath);
                            att.contentType = "*/*";
                            
                            byte[] bytFile = File.ReadAllBytes(strFilePath);
                            att.content = Convert.ToBase64String(bytFile);

                            requestData.attachment.Add(att);
                        }
                    }

                }


                bool blnSuccess = ResolveSnowTicket(requestData, out responseData, out strExceptionMsg);

                

            }
            catch (Exception ex)
            {
                strExceptionMsg = ex.Message;
                
            }

            return strExceptionMsg;
        }

        public bool ResolveSnowTicket(SnowResolveTicketRq requestData, out SnowResolveTicketRs responseData, out string strExceptionMessage)
        {
            bool blnIsSuccess = false;
            string strJsonResMsg = null;

            responseData = null;
            strExceptionMessage = null;

            try
            {
                string strAPIUrl = CommonHelpers.AppendToURL(this.BaseURL, CONST_RESOLVE_TICKET_API);

                Dictionary<string, string> dicHeaders = new Dictionary<string, string>();
                dicHeaders.Add(HttpRequestHeader.ContentType.ToString(), "application/json");
                dicHeaders.Add(HttpRequestHeader.Authorization.ToString(), "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes($"{this.UserName}:{this.Password}")));

                blnIsSuccess = WebServiceHelper.CallWebService(requestData, strAPIUrl, HttpMethod.Post, dicHeaders, null, out strJsonResMsg, out strExceptionMessage, true);
                if (blnIsSuccess)
                {
                    if (!string.IsNullOrWhiteSpace(strJsonResMsg))
                    {
                        responseData = WebServiceHelper.DeserializeJSon<SnowResolveTicketRs>(strJsonResMsg);
                        if (responseData != null && responseData.result != null && responseData.result.ticket != null)
                        {
                            if(string.IsNullOrWhiteSpace(responseData.result.ticket.number))
                            {
                                strExceptionMessage = responseData.result.ticket.update;

                                if (strExceptionMessage.Equals("No Update is taken Due to the Ticket is Resolved Already")) //If Snow ticket already closed
                                {
                                    strExceptionMessage = string.Empty;
                                }
                                
                            }
                            else
                            {
                                strExceptionMessage = string.Empty;
                            }
                            
                        }
                        else
                        {

                            blnIsSuccess = false;
                            strExceptionMessage = string.Format("Failed to deserialize JSON reponse message return from {0}", strAPIUrl);
                        }
                    }
                    else
                    {
                        blnIsSuccess = false;
                        strExceptionMessage = string.Format("Empty JSON response message return from {0}", strAPIUrl);
                    }
                }
            }
            catch
            {

            }
            return blnIsSuccess;
        }

        


    }
}
