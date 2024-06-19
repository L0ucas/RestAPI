using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Xml.Serialization;
using System.Globalization;
using System.Reflection;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Net;
using TMHelper.Common.Config;

namespace TMHelper.Common
{
    public static class CommonHelpers
    {
        public static string GetCurrentProjectDirectoryPath()
        {
            string strPath = string.Empty;

            try
            {
                string strCurDirPath = Directory.GetCurrentDirectory();
                strPath = Directory.GetParent(Directory.GetParent(Directory.GetParent(strCurDirPath).FullName).FullName).FullName + "\\IAPIndoRPA";
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strPath;
        }

        public static string GetProjectResourceOrFolderPath(string strFolderName = null)
        {
            string strPath = string.Empty;
            try
            {
                strPath = GetCurrentProjectDirectoryPath() + "\\" + (string.IsNullOrWhiteSpace(strFolderName) ? "Resources" : strFolderName.StartsWith("\\") ? strFolderName : "\\" + strFolderName);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strPath;
        }

        public static DataSet ConvertXmlSettingsIntoDataSet(string strXmlSettingsFilePath)
        {
            DataSet dsOutput = null;
            System.IO.FileStream fileStream = null;
            try
            {
                if (string.IsNullOrWhiteSpace(strXmlSettingsFilePath))
                    throw new Exception("Xml settings file path is null or empty.");

                fileStream = new System.IO.FileStream(strXmlSettingsFilePath, System.IO.FileMode.Open);
                dsOutput = new DataSet();
                dsOutput.ReadXml(fileStream);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (fileStream != null)
                    fileStream.Close();
            }

            return dsOutput;
        }

        public static string GetConfig(string strFilePath, string strKey)
        {
            ConfigCollection configCollection;
            string path = strFilePath;
            XmlSerializer serializer = new XmlSerializer(typeof(ConfigCollection));
            StreamReader reader = new StreamReader(path);
            configCollection = (ConfigCollection)serializer.Deserialize(reader);
            reader.Close();

            Config config = configCollection.Config.Where(e => e.Key.Equals(strKey)).FirstOrDefault();

            if (config != null)
            {
                return config.Value;
            }
            return string.Empty;
        }

        public static DateTime ConvertDateTime(string strSource, string strFormatSource)
        {
            DateTime output = new DateTime();
            string strFailMsg = string.Format("Method {0} failed. ", MethodInfo.GetCurrentMethod().Name);
            try
            {
                if (strSource.Length != strFormatSource.Length)
                {
                    throw new Exception("Length of source string and format do not match.");
                }

                if (!(strSource.Length == 8 || strSource.Length == 12 || strSource.Length == 14))
                {
                    throw new Exception("Source date time does not in proper length.");
                }

                output = DateTime.ParseExact(strSource, strFormatSource, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                throw new Exception(strFailMsg + ex.Message);
            }
            return output;
        }

        public static bool ConvertDateTime(string strSource, string strFormatSource, out DateTime dtOutput, out string strErr)
        {
            bool blSuccess = false;

            dtOutput = new DateTime();
            strErr = string.Empty;
            string strFailMsg = string.Format("Method {0} failed. ", MethodInfo.GetCurrentMethod().Name);

            //if (strFormatSource.EndsWith(":m", StringComparison.InvariantCultureIgnoreCase))
            //{
            //    strSource = string.Format("{0}:0", strSource);
            //    strFormatSource = string.Format("{0}:s", strFormatSource);
            //}

            try
            {
                dtOutput = DateTime.ParseExact(strSource, strFormatSource, CultureInfo.InvariantCulture);
                blSuccess = true;
            }
            catch (Exception ex)
            {
                strErr = strFailMsg + ex.Message;
            }
            return blSuccess;
        }

        public static void SetLoggingMsg(string strFilePath, string strMsg)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(strMsg);
                string strPath = strFilePath.Substring(0, strFilePath.LastIndexOf("\\"));
                if (!Directory.Exists(strPath))
                {
                    Directory.CreateDirectory(strPath);
                }
                File.AppendAllText(strFilePath, sb.ToString());
            }
            catch (Exception ex) { }
        }

        public static string ShrinkInvoiceNumber(string strInvoiceNumber, string strVendorName, int iMaxInvoiceLength)
        {
            string strShrinkedInvNum = string.Empty;
            try
            {
                #region Input data checking
                if (string.IsNullOrWhiteSpace(strInvoiceNumber))
                    return strInvoiceNumber;
                if (strInvoiceNumber.Length <= iMaxInvoiceLength)
                    return strInvoiceNumber;
                #endregion

                #region First rule: remove all special characters
                strShrinkedInvNum = RemoveAllSpecialChar(strInvoiceNumber);
                #endregion

                if (!string.IsNullOrWhiteSpace(strShrinkedInvNum)
                    && strShrinkedInvNum.Length > iMaxInvoiceLength)
                {
                    #region Second rule: remove INV constant keyword
                    strShrinkedInvNum = ReplaceStringWithIgnoreCase(strShrinkedInvNum, @"INV", string.Empty);
                    #endregion

                    if (!string.IsNullOrWhiteSpace(strShrinkedInvNum)
                        && strShrinkedInvNum.Length > iMaxInvoiceLength)
                    {
                        #region Third rule: remove vendor name abbreviation
                        List<string> lstVendorNmAbbrev = GetVendorNameAbbreviationList(strVendorName);
                        if (lstVendorNmAbbrev != null && lstVendorNmAbbrev.Count > 0)
                        {
                            lstVendorNmAbbrev.ForEach(m => {
                                strShrinkedInvNum = ReplaceStringWithIgnoreCase(strShrinkedInvNum, m, string.Empty);
                            });
                        }
                        #endregion

                        if (!string.IsNullOrWhiteSpace(strShrinkedInvNum)
                            && strShrinkedInvNum.Length > iMaxInvoiceLength)
                        {
                            #region Fourth rule: change 4 digits year into 2 digits e.g. 2019 to 19
                            List<string> lstYear = GetAllYearInString(strShrinkedInvNum);
                            if (lstYear != null && lstYear.Count > 0)
                            {
                                lstYear.ForEach(m =>
                                {
                                    strShrinkedInvNum = strShrinkedInvNum.Replace(m, m.Substring(2));
                                });
                            }
                            #endregion

                            if (!string.IsNullOrWhiteSpace(strShrinkedInvNum)
                                && strShrinkedInvNum.Length > iMaxInvoiceLength)
                            {
                                #region Just trim the substring that could fits the maximum length
                                strShrinkedInvNum = strShrinkedInvNum.Substring(0, iMaxInvoiceLength);
                                #endregion
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strShrinkedInvNum;
        }

        public static string RemoveAllSpecialChar(string strOriginal)
        {
            string strOutput = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(strOriginal))
                    return strOriginal;

                Regex regexWithoutSpecialChar = new Regex(@"[^0-9a-zA-Z]+");
                strOutput = regexWithoutSpecialChar.Replace(strOriginal, string.Empty);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strOutput;
        }

        public static string ReplaceStringWithIgnoreCase(string strFullString, string strOriChar, string strNewChar)
        {
            string strOutput = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(strFullString))
                    return strFullString;

                strOutput = Regex.Replace(strFullString, strOriChar, strNewChar, RegexOptions.IgnoreCase);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strOutput;
        }

        public static List<string> GetVendorNameAbbreviationList(string strVendorName)
        {
            List<string> lstVendorNmAbbrev = null;

            try
            {
                if (string.IsNullOrWhiteSpace(strVendorName))
                    return null;

                lstVendorNmAbbrev = new List<string>();
                string strAbbreviation = string.Empty;

                #region Form abbreviation from first characters of every word
                Array.ForEach(strVendorName.Split(' '), s => strAbbreviation += s[0].ToString());
                int i = strAbbreviation.Length;
                do
                {
                    lstVendorNmAbbrev.Add(strAbbreviation.Substring(0, i));
                    i--;
                }
                while (i > 1); //Abbreviation must be at least 2 characters
                #endregion

                #region Form abbreviation from first word
                strAbbreviation = strVendorName.Split(' ')[0];
                int j = strAbbreviation.Length;
                do
                {
                    lstVendorNmAbbrev.Add(strAbbreviation.Substring(0, j));
                    j--;
                }
                while (j > 1); //Abbreviation must be at least 2 characters
                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lstVendorNmAbbrev;
        }

        public static List<string> GetAllYearInString(string strFullString)
        {
            List<string> lstYear = null;

            try
            {
                if (string.IsNullOrWhiteSpace(strFullString))
                    return null;

                //regex to match substring that start with 20 and follow by 2 digits behind.
                Regex regex = new Regex(@"(20)\d{2}");
                var yearMatch = regex.Matches(strFullString);

                if (yearMatch != null && yearMatch.Count > 0)
                {
                    lstYear = new List<string>();
                    foreach (Match m in yearMatch)
                    {
                        lstYear.Add(m.Groups[0].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lstYear;
        }

        public static bool IsNumeric(string s)
        {
            bool blIsNumeric = false;

            if (!string.IsNullOrEmpty(s))
            {
                blIsNumeric = s.All(char.IsDigit); 
                    //int.TryParse(s, out int n);
            }

            return blIsNumeric;
        }

        public static List<char> GetAllSpecialCharInString(string strFullString)
        {
            List<char> lstSpecialChar = null;
            try
            {
                if (string.IsNullOrWhiteSpace(strFullString))
                    return null;

                lstSpecialChar = new List<char>();

                Regex regexSpecialChar = new Regex(@"[^0-9a-zA-Z]+");
                MatchCollection matchCollection = regexSpecialChar.Matches(strFullString);

                if (matchCollection != null && matchCollection.Count > 0)
                {
                    foreach (Match match in matchCollection)
                    {
                        if (match != null
                            && match.Groups != null
                            && match.Groups.Count > 0
                            && char.TryParse(match.Groups[0].ToString().Trim(), out char cResult))
                        {
                            lstSpecialChar.Add(cResult);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lstSpecialChar;
        }

        public static string FormatCsvValue(string strValue)
        {
            if (string.IsNullOrWhiteSpace(strValue))
                return string.Empty;
            else if (IsNumeric(strValue))
                return string.Format("=\"" + strValue + "\"");
            else
                return string.Format("\"" + strValue + "\""); //value with comma able to print on same cell
        }

        public static string ReplaceInvalidFileNameChars(string strOriginal, char cReplacement)
        {
            if (string.IsNullOrWhiteSpace(strOriginal))
                return strOriginal;

            char[] arrInvalidChar = Path.GetInvalidFileNameChars();
            if (arrInvalidChar.Contains(cReplacement))
                cReplacement = '_';

            foreach (char c in arrInvalidChar)
                strOriginal = strOriginal.Replace(c, cReplacement);

            return strOriginal;
        }

        public static IEnumerable<IEnumerable<long>> GroupConsecutiveNumber(IEnumerable<long> NumberList)
        {
            var NumberGroup = new List<long>();
            foreach (var number in NumberList)
            {
                if (NumberGroup.Count == 0 || number - NumberGroup[NumberGroup.Count - 1] <= 1)
                    NumberGroup.Add(number);
                else
                {
                    yield return NumberGroup;
                    NumberGroup = new List<long> { number };
                }
            }
            yield return NumberGroup;
        }

        public static IEnumerable<IEnumerable<int>> GroupConsecutiveNumber(IEnumerable<int> NumberList)
        {
            var NumberGroup = new List<int>();
            foreach (var number in NumberList)
            {
                if (NumberGroup.Count == 0 || number - NumberGroup[NumberGroup.Count - 1] <= 1)
                    NumberGroup.Add(number);
                else
                {
                    yield return NumberGroup;
                    NumberGroup = new List<int> { number };
                }
            }
            yield return NumberGroup;
        }

        public static void SetLogging(bool blWriteLog, string strFilePathLog, string strContent)
        {
            if (blWriteLog)
                CommonHelpers.SetLoggingMsg(strFilePathLog, strContent + Environment.NewLine);
        }

        public static string GetNextFileName(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            string pathName = Path.GetDirectoryName(fileName);
            string fileNameOnly = Path.Combine(pathName, Path.GetFileNameWithoutExtension(fileName));
            int i = 0;
            // If the file exists, keep trying until it doesn't
            while (File.Exists(fileName))
            {
                i += 1;
                fileName = string.Format("{0}({1}){2}", fileNameOnly, i, extension);
            }
            return fileName;
        }

        public static double Rounding(double dblSource, int intDecimalPoint, int intThreshold)
        {
            if (dblSource < 0 || intDecimalPoint < 0 || !(intThreshold > 0 && intThreshold <= 9))
            {
                throw new Exception(string.Format("Unable to perform rounding with source number {0} or rounding decimal point {1} or threshold {2}. ",
                    dblSource.ToString(), intDecimalPoint.ToString(), intThreshold.ToString()));
            }
            double dblOutput = dblSource;
            string input_decimal_number = dblSource.ToString();
            string decimal_places = string.Empty;
            var regex = new System.Text.RegularExpressions.Regex("(?<=[\\.])[0-9]+");
            if (regex.IsMatch(input_decimal_number))
            {
                decimal_places = regex.Match(input_decimal_number).Value;
            }
            if (!string.IsNullOrWhiteSpace(decimal_places) && decimal_places.Length >= (intDecimalPoint + 1))
            {
                if (int.Parse(decimal_places.Substring(intDecimalPoint, 1)) >= intThreshold)
                {
                    dblOutput = dblSource + (1 / Math.Pow(10, intDecimalPoint));
                }
                dblOutput = double.Parse(dblOutput.ToString().Substring(0, dblOutput.ToString().IndexOf('.') + (intDecimalPoint + 1)));
            }
            return dblOutput;
        }

        public static string GetRPAConfig(string strConfigKey)
        {
            string strOutput = string.Empty;
            
            //DatabaseHelper databaseHelper = new DatabaseHelper(string.Format(strConn, GetConfig(strFilePath, strDatabaseIPKey)));
            DatabaseHelper databaseHelper = new DatabaseHelper(CommonConfigHelper.GetRPADBConnString());
            DataTable dtParam = new DataTable();
            DataSet dsOutput = new DataSet();
            string strErr = string.Empty;

            try
            {
                dtParam.Columns.Add(GlobalConstant.RPAMasterConfig.CONST_COL_NM_CONFIG_KEY);
                DataRow dr = dtParam.NewRow();
                dr[GlobalConstant.RPAMasterConfig.CONST_COL_NM_CONFIG_KEY] = strConfigKey;
                dtParam.Rows.Add(dr);
                databaseHelper.ExecuteStoreProcedureWithParameters("UspRPAMasterConfigSel", dtParam, out dsOutput, out strErr);
                if (dsOutput != null && dsOutput.Tables.Count > 0 && dsOutput.Tables[0].Rows.Count > 0)
                {
                    strOutput = dsOutput.Tables[0].Rows[0][GlobalConstant.RPAMasterConfig.CONST_COL_NM_CONFIG_VALUE].ToString();
                    if (string.IsNullOrWhiteSpace(strOutput))
                    {
                        throw new Exception(string.Format("RPA Master Config {0} is not setup. ", strConfigKey));
                    }
                }
                else
                {
                    throw new Exception(string.Format("Unable to get RPA Master Config {0}. ", strConfigKey));
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return strOutput;
        }

        public static string GetMachineIPv4Addr()
        {
            string strIPv4Addr = string.Empty;

            IPAddress[] arrMachineIP = GetMachineIPAddrList();
            if (arrMachineIP != null && arrMachineIP.Length > 0)
            {
                var specificIPv4Addr = arrMachineIP.FirstOrDefault(x => x.AddressFamily.Equals(System.Net.Sockets.AddressFamily.InterNetwork));
                if (specificIPv4Addr != null && specificIPv4Addr != default(IPAddress))
                    strIPv4Addr = specificIPv4Addr.ToString();
            }

            return strIPv4Addr;
        }

        public static string GetMachineIPv6Addr()
        {
            string strIPv6Addr = string.Empty;

            IPAddress[] arrMachineIP = GetMachineIPAddrList();
            if (arrMachineIP != null && arrMachineIP.Length > 0)
            {
                var specificIPv6Addr = arrMachineIP.FirstOrDefault(x => x.AddressFamily.Equals(System.Net.Sockets.AddressFamily.InterNetworkV6));
                if (specificIPv6Addr != null && specificIPv6Addr != default(IPAddress))
                    strIPv6Addr = specificIPv6Addr.ToString();
            }

            return strIPv6Addr;
        }

        public static IPAddress[] GetMachineIPAddrList()
        {
            IPAddress[] IPAddrArray = null;

            IPHostEntry ipHostEntry = Dns.GetHostEntry(Dns.GetHostName());
            if (ipHostEntry != null)
                IPAddrArray = ipHostEntry.AddressList;

            return IPAddrArray;
        }

        public static string GetPropertyValueByName<T>(T typeObject, string strPropertyName)
        {
            string strPropertyValue = string.Empty;

            try
            {
                Type type = typeObject.GetType();
                PropertyInfo propInfo = type?.GetProperty(strPropertyName);
                strPropertyValue = propInfo?.GetValue(typeObject)?.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strPropertyValue;
        }

        public static DataTable ListToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties by using reflection   
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {

                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        public static string AppendToURL(string baseURL, params string[] segments)
        {
            {
                return string.Join("/", new[] { baseURL.TrimEnd('/') }
                    .Concat(segments.Select(s => s.Trim('/'))));
            }
        }

        public static string FormatFileSizeInByte(long bytes)
        {
            string strFileSize = string.Empty;

            try
            {
                string[] arrSizeSuffix = { "Bytes", "KB", "MB", "GB", "TB", "PB" };

                int iLoopIndex = 0;
                decimal decBytes = (decimal)bytes;
                while (Math.Round(decBytes / 1024) >= 1)
                {
                    decBytes = decBytes / 1024;
                    iLoopIndex++;
                }

                strFileSize = string.Format("{0:n1}{1}", decBytes, arrSizeSuffix[iLoopIndex]);
            }
            catch (Exception ex)
            {
                strFileSize = bytes.ToString();
            }

            return strFileSize;
        }

        public static string Reverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        public static string[] DirectoryGetFiles(string strFolderPath, string strSearchPattern, SearchOption srchOpt)
        {
            string[] arrFiles = null;
            List<string> listFilePath = null;

            try
            {
                string[] arrSearchPatterns = strSearchPattern.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                if (arrSearchPatterns != null && arrSearchPatterns.Length > 0)
                {
                    listFilePath = new List<string>();
                    foreach (string strSP in arrSearchPatterns)
                        listFilePath.AddRange(System.IO.Directory.GetFiles(strFolderPath, strSP, srchOpt));

                    listFilePath.Sort();
                    arrFiles = listFilePath.ToArray();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return arrFiles;
        }

        public static void SetLogFailTransaction(string strFilePath, string strFileName, string strRqJSON, string strRsJSON)
        {
            string strText = string.Empty;

            strText = "---------- Request ------------" + Environment.NewLine + strRqJSON + Environment.NewLine + Environment.NewLine;
            strText = strText + "---------- Response ------------" + Environment.NewLine + strRsJSON + Environment.NewLine + Environment.NewLine;

            CommonHelpers.SetLoggingMsg(Path.Combine(strFilePath, @"LogFailed", strFileName), strText);
        }

        public static void SetLogSuccessTransaction(string strFilePath, string strFileName, string strHeader, string strRqJSON, string strRsJSON)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("[{0} : {1}]", DateTime.Now.ToString("yyyyMMdd HHmmss"), strHeader));
            sb.AppendLine("---------- Request ------------");
            sb.AppendLine(strRqJSON);
            sb.AppendLine(string.Empty);
            sb.AppendLine("---------- Response ------------");
            sb.AppendLine(strRsJSON);
            sb.AppendLine("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------");
            sb.AppendLine(string.Empty);
            sb.AppendLine(string.Empty);

            CommonHelpers.SetLoggingMsg(Path.Combine(strFilePath, @"LogSuccess\", strFileName), sb.ToString());
        }

        public static bool CheckVersion(string strDbParameter, string strId, double dbVersionNo)
        {
            bool blnIsVersionMatch = false;
            string strExceptionMessage = string.Empty;
            try
            {
                #region Prepare input parameter
                System.Data.DataTable dtblInput = new System.Data.DataTable();
                dtblInput.Columns.Add("Id", typeof(string));

                DataRow drInput = dtblInput.NewRow();
                drInput["Id"] = strId;

                dtblInput.Rows.Add(drInput);
                #endregion

                DatabaseHelper dbHelper = new DatabaseHelper(strDbParameter);
                long lngResult = dbHelper.ExecuteStoreProcedureWithParameters("UspVersionCheckSel", dtblInput, out DataSet dsResult, out strExceptionMessage);

                #region Extracting SP Result
                if (lngResult == 0)
                {
                    if (dsResult != null && dsResult.Tables != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows != null && dsResult.Tables[0].Rows.Count > 0)
                    {
                        DataRow dataRowResult = dsResult.Tables[0].Rows[0];

                        double dbVersionNoDB = Decimal.ToDouble((decimal)dataRowResult["VersionNo"]);

                        if(dbVersionNoDB == dbVersionNo)
                        {
                            blnIsVersionMatch = true;
                        }
                    }
                }
                else
                {
                    throw new Exception("Dataset or datatable return from store procedure is null, Reason: " + strExceptionMessage);
                }
                #endregion

            }
            catch (Exception e)
            {
                throw new Exception("There is some error while checking version, " + e.Message);
            }

            return blnIsVersionMatch;
        }

        public static decimal? convertStringWithEndNegativeToDouble(string number)
        {
            try
            {
                bool negative = false;

                if (number.Substring(number.Length - 1).Equals("-"))
                {
                    number = number.Substring(0, number.Length - 1);
                    negative = true;
                }

                decimal? parsedNumber = null;

                if (decimal.TryParse(number, out decimal result))
                {
                    parsedNumber = result; //交易金额

                    if (negative)
                    {
                        parsedNumber *= -1;
                    }
                }

                return parsedNumber;

            } catch (Exception e)
            {
                return null;
            }
        }

        [Serializable()]
        public class Config
        {
            [System.Xml.Serialization.XmlElement("Key")]
            public string Key { get; set; }
            [System.Xml.Serialization.XmlElement("Value")]
            public string Value { get; set; }
        }

        [Serializable()]
        [System.Xml.Serialization.XmlRoot("ConfigCollection")]
        public class ConfigCollection
        {
            [XmlArray("Configs")]
            [XmlArrayItem("Config", typeof(Config))]
            public Config[] Config { get; set; }
        }
    }
}
