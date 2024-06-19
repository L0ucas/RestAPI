using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.Serialization.Json;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Script.Serialization;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Data;
using System.Collections;

namespace TMHelper.Common
{
    public static class WebServiceHelper
    {
        #region Constant
        public const string CONST_JSON_REQ_CONTENT_TYPE = "application/json;charset=UTF-8";
        public const string CONST_TEXT_XML_REQ_CONTENT_TYPE = "text/xml";

        public static class AuthType
        {
            public const string BEARER = "Bearer";
            public const string BASIC = "Basic";
        }

        public enum WebSeriveBodyType
        {
            None = 0,
            FormData = 1,
            xWWWFormUrlEncoded = 2,
            JSON = 3,
            Binary = 4
        }
        #endregion

        #region JSON serialization
        public static string SerializeJSon<T>(T t)
        {
            try
            {
                //MemoryStream stream = new MemoryStream();
                //DataContractJsonSerializer ds = new DataContractJsonSerializer(typeof(T));
                //DataContractJsonSerializerSettings s = new DataContractJsonSerializerSettings();
                //ds.WriteObject(stream, t);
                //string jsonString = Encoding.UTF8.GetString(stream.ToArray());
                //stream.Close();
                var json = new JavaScriptSerializer().Serialize(t);

                return json;
            }
            catch (Exception ex)
            {
                throw ex; //Exception handle by caller
            }
        }

        public static T DeserializeJson<T>(string strJson)
        {
            try
            {
                var serializer = new JavaScriptSerializer();
                T obj = serializer.Deserialize<T>(strJson);
                return obj;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static T DeserializeJSon<T>(string jsonString)
        {
            try
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T));
                MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
                T obj = (T)ser.ReadObject(stream);
                return obj;
            }
            catch (Exception ex)
            {
                throw ex; //Exception handle by caller
            }
        }

        public static string DataSetToJson(DataSet dsInput)
        {
            string strJson = null;
            Dictionary<string, object> dicDataSet = null;
            List<Dictionary<string, object>> listDataTable = null;
            Dictionary<string, object> dicDataRow = null;

            try
            {
                #region Pre-checking
                if (dsInput == null) return null;
                if (dsInput.Tables == null) return null;
                if (dsInput.Tables.Count == 0) return null;
                #endregion

                dicDataSet = new Dictionary<string, object>();

                foreach (DataTable dtblInput in dsInput.Tables)
                {
                    listDataTable = new List<Dictionary<string, object>>();
                    
                    foreach (DataRow drInput in dtblInput.Rows)
                    {
                        dicDataRow = new Dictionary<string, object>();

                        foreach (DataColumn dcInput in dtblInput.Columns)
                            dicDataRow.Add(dcInput.ColumnName, drInput[dcInput]);

                        listDataTable.Add(dicDataRow);
                    }

                    dicDataSet.Add(dtblInput.TableName, listDataTable);
                }

                strJson = SerializeJSon(dicDataSet);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strJson;
        }
        
        public static DataSet JsonToDataSet(string strJson)
        {
            DataSet dsJson = null;

            try
            {
                #region Pre-checking
                if (string.IsNullOrWhiteSpace(strJson)) return null;
                #endregion

                Dictionary<string, object> dicDataSet = DeserializeJson<Dictionary<string, object>>(strJson);
                if (dicDataSet != null && dicDataSet.Count > 0)
                {
                    dsJson = new DataSet();

                    foreach (var dsKeyPair in dicDataSet)
                    {
                        DataTable dtblJson = dsJson.Tables.Add(dsKeyPair.Key);
                        ArrayList arrList = (ArrayList)dsKeyPair.Value;

                        Dictionary<string, object> dicArrList = (Dictionary<string, object>)arrList[0];
                        var allColumnName = dicArrList.Select(m => m.Key).Distinct();
                        dtblJson.Columns.AddRange(allColumnName.Select(c => new DataColumn(c)).ToArray());

                        foreach (Dictionary<string, object> dicArr in arrList)
                        {
                            DataRow drJson = dtblJson.NewRow();

                            foreach (var rowKeyPair in dicArr)
                                drJson[rowKeyPair.Key] = (string)rowKeyPair.Value;

                            dtblJson.Rows.Add(drJson);
                        }

                        
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dsJson;
        }
        #endregion

        #region XML SOAP serialization
        public static string SerializeXMLSOAP<T>(T t)
        {
            string strXMLSOAP = string.Empty;
            try
            {
                XmlWriterSettings xWriterSetting = new XmlWriterSettings
                {
                    Indent = true,
                    OmitXmlDeclaration = false,
                    Encoding = Encoding.UTF8
                };

                using (MemoryStream ms = new MemoryStream())
                {
                    using (XmlWriter xWriter = XmlWriter.Create(ms, xWriterSetting))
                    {
                        xWriter.WriteStartElement("soap", "Envelope", "http://schemas.xmlsoap.org/soap/envelope/");
                        xWriter.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
                        xWriter.WriteAttributeString("xmlns", "xsd", null, "http://www.w3.org/2001/XMLSchema");

                        xWriter.WriteStartElement("soap", "Body", "http://schemas.xmlsoap.org/soap/envelope/");

                        var xSerializer = new XmlSerializer(typeof(T));
                        xSerializer.Serialize(xWriter, t);

                        xWriter.WriteFullEndElement();
                        xWriter.WriteFullEndElement();

                        ms.Position = 0;

                        strXMLSOAP = Encoding.UTF8.GetString(ms.ToArray());
                    }
                }

                strXMLSOAP += @"</soap:Body></soap:Envelope>";
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strXMLSOAP;
        }

        public static T DeserializeXMLSOAP<T>(string strXMLSOAP)
        {
            try
            {
                XDocument xDoc = XDocument.Parse(strXMLSOAP);
                XmlSerializer xSerializer = new XmlSerializer(typeof(T));

                XName xNmSOAPBody = XName.Get("Body", "http://schemas.xmlsoap.org/soap/envelope/");

                T obj = (T)xSerializer.Deserialize(xDoc
                                                   .Descendants(xNmSOAPBody)
                                                   .First()
                                                   .FirstNode
                                                   .CreateReader());

                return obj;
            }
            catch (Exception ex)
            {
                throw ex; //Exception handle by caller
            }
        }
        
        public static T DeserializeXML<T>(string strXML)
        {
            try
            {
                XDocument xDoc = XDocument.Parse(strXML);
                XmlSerializer xSerializer = new XmlSerializer(typeof(T));

                T obj = (T)xSerializer.Deserialize(xDoc.CreateReader());

                return obj;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string SerializeXML<T>(T t)
        {
            string strXMLSOAP = string.Empty;

            try
            {
                XmlSerializer xSerializer = new XmlSerializer(typeof(T));
                using (var strWrite = new StringWriter())
                {
                    using (XmlWriter xWriter = XmlWriter.Create(strWrite, new XmlWriterSettings { Indent = true, OmitXmlDeclaration = true }))
                    {
                        XmlSerializerNamespaces xNs = new XmlSerializerNamespaces();
                        xNs.Add("", "");

                        xSerializer.Serialize(xWriter, t, xNs);
                        strXMLSOAP = strWrite.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strXMLSOAP;
        }
        #endregion

        #region Call RESTful webservice
        public static string CallWebService(string webServiceRequest, string strWebServiceUrl, string Oauth_Token, bool blnByPassServerCertCheck = false)
        {
            string strJsonRequest = string.Empty;
            string strJsonResponse = string.Empty;
            //object webServiceResponse = null;

            try
            {
                if (blnByPassServerCertCheck)
                    ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true; //by pass certificate checking

                //strJsonRequest = SerializeJSon<object>(webServiceRequest);
                strJsonResponse = HttpPost(strWebServiceUrl, webServiceRequest);
                //webServiceResponse = DeserializeJSon<object>(strJsonResponse);
            }
            catch (Exception ex)
            {
                throw ex; //Exception handle by caller
            }

            return strJsonResponse;
        }

        public static bool CallWebService(object objRequest, string strWsUrl, HttpMethod wsHttpMethod, Dictionary<string, string> dicMsgHeader, string strAuthToken,
                                          out string strJsonResMsg, out string strExceptionMsg, bool blnByPassServerCertCheck = false, string strAuthType = AuthType.BEARER)
        {
            return CallWebService(objRequest, strWsUrl, wsHttpMethod, dicMsgHeader, strAuthToken, null, null, false, null, out strJsonResMsg, out strExceptionMsg, blnByPassServerCertCheck, strAuthType);
        }

        public static bool CallWebService_Bearer_Token(object objRequest, string strWsUrl, HttpMethod wsHttpMethod, Dictionary<string, string> dicMsgHeader, string strAuthToken,
                                          out string strJsonResMsg, out string strExceptionMsg, bool blnByPassServerCertCheck = false, string strAuthType = AuthType.BEARER, WebSeriveBodyType wsBodyType = WebSeriveBodyType.xWWWFormUrlEncoded,
                                           Dictionary<string, string> dicURLEncodeContentData = null)
        {
            return CallWebService(objRequest, strWsUrl, wsHttpMethod, dicMsgHeader, strAuthToken, null, null, false, null, out strJsonResMsg, out strExceptionMsg, blnByPassServerCertCheck, strAuthType, wsBodyType, dicURLEncodeContentData);
        }

        public static bool CallWebServiceToPostFile(string strWsUrl, Dictionary<string, string> dicMsgHeader, string strAuthToken, string strWsFileParamName,
                                                  string strFilePathToPost, out string strJsonResMsg, out string strExceptionMsg, 
                                                  bool blnByPassServerCertCheck = false, string strAuthType = AuthType.BEARER)
        {
            return CallWebService(null, strWsUrl, HttpMethod.Post, dicMsgHeader, strAuthToken, strWsFileParamName, strFilePathToPost, false, null, out strJsonResMsg, out strExceptionMsg, blnByPassServerCertCheck, strAuthType, WebSeriveBodyType.FormData);
        }

        public static bool CallWebServiceAndDownLoadFile(object objRequest, string strWsUrl, HttpMethod wsHttpMethod, Dictionary<string, string> dicMsgHeader, string strAuthToken,
                                                         string strResponseSaveFilePath, out string strJsonResMsg, out string strExceptionMsg, bool blnByPassServerCertCheck = false, string strAuthType = AuthType.BEARER)
        {
            return CallWebService(objRequest, strWsUrl, wsHttpMethod, dicMsgHeader, strAuthToken, null, null, true, strResponseSaveFilePath, out strJsonResMsg, out strExceptionMsg, blnByPassServerCertCheck, strAuthType);
        }

        public static bool CallWebServiceFormUrlEncoded(Dictionary<string, string> dicURLEncodeContentData, string strWsUrl, HttpMethod wsHttpMethod, Dictionary<string, string> dicMsgHeader, 
                                                        string strAuthToken, out string strJsonResMsg, out string strExceptionMsg, bool blnByPassServerCertCheck = false, string strAuthType = AuthType.BEARER)
        {
            return CallWebService(null, strWsUrl, wsHttpMethod, dicMsgHeader, strAuthToken, null, null, false, null, out strJsonResMsg, out strExceptionMsg, blnByPassServerCertCheck, strAuthType, WebSeriveBodyType.xWWWFormUrlEncoded, dicURLEncodeContentData);
        }

        public static bool CallWebServiceFormData(Dictionary<string, string> dicFormData, string strWsFileParamName, string strFilePathToPost, 
                                                  string strWsUrl, HttpMethod wsHttpMethod, Dictionary<string, string> dicMsgHeader,
                                                  string strAuthToken, out string strJsonResMsg, out string strExceptionMsg, 
                                                  bool blnByPassServerCertCheck = false, string strAuthType = AuthType.BEARER, string strFileContentType = null)
        {
            return CallWebService(null, strWsUrl, wsHttpMethod, dicMsgHeader, strAuthToken, strWsFileParamName, strFilePathToPost, false, null, out strJsonResMsg, out strExceptionMsg, blnByPassServerCertCheck, strAuthType, WebSeriveBodyType.FormData, null, dicFormData, strFileContentType);
        }

        private static bool CallWebService(object objRequest, string strWsUrl, HttpMethod wsHttpMethod, Dictionary<string, string> dicMsgHeader, string strAuthToken,
                                           string strWsFileParamName, string strFilePathToPost, bool blnIsSaveFileFromResponse,
                                           string strResponseSaveFilePath, out string strJsonResMsg, out string strExceptionMsg, bool blnByPassServerCertCheck = false,
                                           string strAuthType = AuthType.BEARER, WebSeriveBodyType wsBodyType = WebSeriveBodyType.JSON,
                                           Dictionary<string, string> dicURLEncodeContentData = null, Dictionary<string, string> dicFormData = null, string strFileContentType = null)
        {
            bool blnIsSuccess = true;
            HttpResponseMessage httpResMsg = null;
            string strJsonRequest = string.Empty;

            strExceptionMsg = string.Empty;
            strJsonResMsg = string.Empty;

            try
            {
                if (blnByPassServerCertCheck)
                {
                    ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true; //by pass certificate checking
                    ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                }

                using (var httpClient = new HttpClient())
                {
                    var httpReqMsg = new HttpRequestMessage();
                    httpReqMsg.Method = wsHttpMethod;
                    httpReqMsg.RequestUri = new Uri(strWsUrl);

                    //Add message headers
                    if (!string.IsNullOrWhiteSpace(strAuthToken))
                    {
                        httpReqMsg.Headers.Add(HttpRequestHeader.Authorization.ToString(), string.Format("{0} {1}", strAuthType, strAuthToken));
                    }
                    if (dicMsgHeader != null && dicMsgHeader.Count > 0)
                    {
                        foreach (var keyValPair in dicMsgHeader)
                        {
                            httpReqMsg.Headers.Add(keyValPair.Key, keyValPair.Value);
                        }
                    }

                    if (wsBodyType == WebSeriveBodyType.JSON) //If request object is not empty
                    {
                        if (objRequest != null)
                        {
                            //strJsonRequest = SerializeJSon<object>(objRequest);
                            strJsonRequest = new JavaScriptSerializer().Serialize(objRequest);
                            //objRequest.ToString();
                            //httpReqMsg.Headers.Add(HttpRequestHeader.Accept.ToString(), "*/*");
                            //httpReqMsg.Headers.Add("X-Version", "1");

                            httpReqMsg.Content = new StringContent(strJsonRequest.ToString(), Encoding.UTF8, "application/json");
                        }

                        httpResMsg = httpClient.SendAsync(httpReqMsg).GetAwaiter().GetResult(); //Send request to API
                    }

                    else if (wsBodyType == WebSeriveBodyType.xWWWFormUrlEncoded)
                    {
                        httpReqMsg.Content = new FormUrlEncodedContent(dicURLEncodeContentData);
                        httpResMsg = httpClient.SendAsync(httpReqMsg).GetAwaiter().GetResult(); //Send request to API
                    }
                    else if (wsBodyType == WebSeriveBodyType.FormData)
                    {
                        using (var form = new MultipartFormDataContent())
                        {
                            FileStream fs = null;
                            StreamContent streamContent = null;
                            ByteArrayContent fileContent = null;

                            if (!string.IsNullOrWhiteSpace(strFilePathToPost) && File.Exists(strFilePathToPost))
                            {
                                fs = File.OpenRead(strFilePathToPost);
                                streamContent = new StreamContent(fs);
                                fileContent = new ByteArrayContent(streamContent.ReadAsByteArrayAsync().GetAwaiter().GetResult());

                                if (!string.IsNullOrWhiteSpace(strFileContentType))
                                    fileContent.Headers.ContentType = new MediaTypeHeaderValue(strFileContentType);
                                else
                                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("multipart/form-data");

                                form.Add(fileContent, strWsFileParamName, Path.GetFileName(strFilePathToPost));
                            }

                            if (dicFormData != null && dicFormData.Count > 0)
                            {
                                foreach (var fData in dicFormData)
                                {
                                    StringContent sc = new StringContent(fData.Value);
                                    form.Add(sc, fData.Key);
                                }
                            }

                            httpReqMsg.Content = form;
                            httpResMsg = httpClient.SendAsync(httpReqMsg).GetAwaiter().GetResult(); //Send request to API

                            if (fs != null) { fs.Dispose(); }
                            if (streamContent != null) { streamContent.Dispose(); }
                            if (fileContent != null) { fileContent.Dispose(); }
                        }
                    }
                }

                if (httpResMsg != null)
                {
                    if (httpResMsg.StatusCode == HttpStatusCode.OK)
                    {
                        if (httpResMsg.Content != null)
                        {
                            if (blnIsSaveFileFromResponse)
                            {
                                byte[] byteFile = httpResMsg.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                                File.WriteAllBytes(strResponseSaveFilePath, byteFile);
                            }
                            else
                            {
                                strJsonResMsg = httpResMsg.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                            }
                        }
                        else
                        {
                            throw new Exception(string.Format("HTTP response content is null after call {0}", strWsUrl));
                        }
                    }
                    else if (httpResMsg.StatusCode == HttpStatusCode.BadRequest)
                    {
                        if (httpResMsg.Content != null)
                        {
                            if (blnIsSaveFileFromResponse)
                            {
                                byte[] byteFile = httpResMsg.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                                File.WriteAllBytes(strResponseSaveFilePath, byteFile);
                            }
                            else
                            {
                                strJsonResMsg = httpResMsg.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                            }
                        }
                        else
                        {
                            throw new Exception(string.Format("HTTP response content is null after call {0}", strWsUrl));
                        }
                    }
                    else
                    {
                        throw new Exception(string.Format("Fail to call {0}. HTTP status code is {1}", strWsUrl, httpResMsg.StatusCode.ToString()));
                    }
                }
                else
                {
                    throw new Exception(string.Format("HTTP response message is null after call {0}", strWsUrl));
                }
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMsg = ex.Message;
            }

            return blnIsSuccess;
        }
        #endregion

        #region POST to webservice
        public static string HttpPost(string Url, string postDataStr, string strReqContentType = "", bool blnIsAppendTLS = false)
        {
            try
            {
                if (blnIsAppendTLS)
                    ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                if (string.IsNullOrWhiteSpace(strReqContentType))
                    strReqContentType = CONST_JSON_REQ_CONTENT_TYPE;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                request.Method = "POST";
                request.ContentType = strReqContentType;
                request.ContentLength = postDataStr.Length;
                request.Timeout = 5000000;

                byte[] postdatabyte = Encoding.UTF8.GetBytes(postDataStr);
                request.ContentLength = postdatabyte.Length;
                request.AllowAutoRedirect = false;

                Stream stream;
                stream = request.GetRequestStream();
                stream.Write(postdatabyte, 0, postdatabyte.Length);
                stream.Close();

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                string encoding = response.ContentEncoding;
                if (encoding == null || encoding.Length < 1)
                {
                    encoding = "UTF-8"; //Default   
                }

                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(encoding));
                string retString = reader.ReadToEnd();
                return retString;
            }
            catch (Exception ex)
            {
                throw ex; //Exception handle by caller
            }
        }

      

        #endregion

        public static string HttpGet(string Url, string postDataStr, string strReqContentType = "", bool blnIsAppendTLS = false, WebProxy webProxy = null, NetworkCredential credential = null)
        {
            try
            {
                if (blnIsAppendTLS)
                    ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                if (string.IsNullOrWhiteSpace(strReqContentType))
                    strReqContentType = CONST_JSON_REQ_CONTENT_TYPE;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                request.Method = "GET";
                request.ContentType = strReqContentType;
                request.ContentLength = postDataStr.Length;

                if (webProxy != null)
                    request.Proxy = webProxy;

                if (credential != null)
                    request.Credentials = credential;

                byte[] postdatabyte = Encoding.UTF8.GetBytes(postDataStr);
                request.ContentLength = postdatabyte.Length;

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                string encoding = response.ContentEncoding;
                if (encoding == null || encoding.Length < 1)
                {
                    encoding = "UTF-8"; //Default   
                }

                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(encoding));
                string retString = reader.ReadToEnd();
                return retString;
            }
            catch (Exception ex)
            {
                throw ex; //Exception handle by caller
            }
        }

        #region Get Server SSL cert expiration date
        public static bool GetSSLCertExpirationDateString(string strSecureURL, out string strExpirationDate, out string strExceptionMsg)
        {
            #region Variable Declaration
            bool blnIsSuccess = false;
            Uri uriRes = null;
            X509Certificate servCert = null;
            X509Certificate2 servCert2 = null;
            HttpWebResponse httpWebRes = null;

            strExpirationDate = string.Empty;
            strExceptionMsg = string.Empty;
            #endregion

            try
            {
                #region Pre-checking
                if (string.IsNullOrWhiteSpace(strSecureURL)) { throw new Exception("Empty or null secure URL parse"); }
                if (!Uri.TryCreate(strSecureURL, UriKind.Absolute, out uriRes)) { throw new Exception("Secure URL format is invalid"); }
                if (uriRes == null) { throw new Exception("Secure URL format is invalid"); }
                if (uriRes.Scheme != Uri.UriSchemeHttps) { throw new Exception("Secure URL scheme is not HTTPS"); }
                #endregion

                #region Establish connection to the URL
                HttpWebRequest httpWebReq = (HttpWebRequest)WebRequest.Create(strSecureURL);
                httpWebRes = (HttpWebResponse)httpWebReq.GetResponse();
                servCert = httpWebReq.ServicePoint.Certificate;
                #endregion

                #region Retrieve cert expiration date string
                if (servCert != null)
                {
                    servCert2 = new X509Certificate2(servCert);
                    if (servCert2 != null)
                    {
                        strExpirationDate = servCert2.GetExpirationDateString();
                        blnIsSuccess = true;
                    }
                    else
                    {
                        throw new Exception("Server certificate2 is null");
                    }
                }
                else
                {
                    throw new Exception("Server certificate is null");
                }
                #endregion
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMsg = ex.Message;
            }
            finally
            {
                if (httpWebRes != null)
                {
                    httpWebRes.Close();
                    httpWebRes.Dispose();
                }
            }

            return blnIsSuccess;
        }
        #endregion

        #region Get HTML doc/string from URL
        public static bool GetHTMLStringFromURL(string strTargetURL, out string strHTML, out string strExceptionMessage, string strUserName = null, string strPassword = null)
        {
            #region Variable Declaration
            bool blnIsSuccess = true;
            Uri uriRes = null;
            strHTML = null;
            strExceptionMessage = null;
            #endregion

            try
            {
                #region Pre-checking
                if (string.IsNullOrWhiteSpace(strTargetURL)) { throw new Exception("Empty or null target URL"); }
                if (!Uri.TryCreate(strTargetURL, UriKind.Absolute, out uriRes)) { throw new Exception("Target URL format is invalid"); }
                #endregion

                #region Extact HTML string
                using (var wClient = new WebClient())
                {
                    if (!string.IsNullOrWhiteSpace(strUserName) && !string.IsNullOrWhiteSpace(strPassword))
                        wClient.Credentials = new NetworkCredential(strUserName, strPassword);

                    wClient.Encoding = Encoding.UTF8;
                    strHTML = wClient.DownloadString(strTargetURL);
                }
                #endregion
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }

            return blnIsSuccess;
        }
        #endregion
    }
}
