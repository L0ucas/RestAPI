using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Xml;

namespace TMHelper.Common
{
    public class ROCHelper
    {
        public ROCHelper(string strRocEndPointUrl)
        {
            ROCUrl = strRocEndPointUrl;
        }

        public ROCHelper()
        {
            ROCUrl = "https://app.visualfloors.aspiro.co/roc/roc.asmx/ROC"; //default URL
        }

        private string ROCUrl { get; set; }
        public string Deal { get; set; }
        public string Process { get; set; }
        public string Desktop { get; set; }
        public string TransactionId { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string Status { get; set; }
        public string Description { get; set; }
        public string DateFormat { get; set; }

        public string PostDataToRocService()
        {
            string resultPost = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(this.Desktop))
                    this.Desktop = Environment.MachineName;

                string dataString = string.Format("Data={0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|", this.Deal, this.Process, this.Desktop, this.TransactionId, this.StartTime, this.EndTime, this.Status, this.Description, this.DateFormat);
                var data = Encoding.UTF8.GetBytes(dataString);
                var lstrUri = new Uri(ROCUrl);
                resultPost = SendRequest(lstrUri, data, "application/x-www-form-urlencoded", "POST");
            }
            catch (Exception ex)
            {
                resultPost = string.Format("Exception occurred while post data to ROC service. Error = {0}", ex.Message);
            }

            return resultPost;
        }

        private string SendRequest(Uri uri, Byte[] DataBytes, string contentType, string method)
        {
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(uri);
                req.ContentType = contentType;
                req.Method = method;
                req.ContentLength = DataBytes.Length;

                var stream = req.GetRequestStream();
                stream.Write(DataBytes, 0, DataBytes.Length);
                stream.Close();

                var response = req.GetResponse().GetResponseStream();

                var reader = new StreamReader(response);
                var res = reader.ReadToEnd();

                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(res);

                reader.Close();
                response.Close();

                return xmlDoc.ChildNodes[1].LastChild.InnerText;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
