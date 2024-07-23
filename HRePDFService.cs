using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using IndoHRePDF;
using TMHelper.Common;
using TMHelper.SNOW;

namespace WindowsService1
{
    public class HRePDFService
    {
        public bool StartProcess(out string strExceptionMessage)
        {
            #region Variable declaration
            bool blnIsSuccess = true;
            bool blnIsActionSuccess = false;
            long lngResult = -1;
            long lngDBAction = -1;
            string strDBException = null;
            List<SnowTicketDataPumpRq> ListSnowTicketData = null;
            List<EPDFConfig> ListEPDFConfig = null;
            List<FileFolderMapping> ListFileFolderMap = null;

            strExceptionMessage = null;
            #endregion

            try
            {
                #region Extract config data
                IndoHRePDFMain ePDFMain = new IndoHRePDFMain();
                lngResult = ePDFMain.ExtractConfigData(out strExceptionMessage);

                if (lngResult != 0) { throw new Exception($"Failed to extract config data. {strExceptionMessage}"); }

                ListEPDFConfig = ePDFMain.ListEPDFConfig;

                if (!(ListEPDFConfig != null && ListEPDFConfig.Count > 0)) { throw new Exception("Configuration data list is either empty or null"); }
                #endregion

                #region Extract folder mapping data
                string strWorkingFolderPath = ListEPDFConfig[0].WorkingFolderPath;
                blnIsActionSuccess = ExtractMappingData(strWorkingFolderPath, out ListFileFolderMap, out strExceptionMessage);

                if (!blnIsActionSuccess) { throw new Exception($"Failed to extract mapping file data. {strExceptionMessage}"); }
                if (!(ListFileFolderMap != null && ListFileFolderMap.Count > 0)) { throw new Exception("Folder mapping data list is either empty or null"); }
                #endregion

                #region Extract SNOW ticket data
                lngResult = SNOWHelper.ExtractSnowTicketData(out ListSnowTicketData, out strExceptionMessage);

                if (lngResult != 0) { throw new Exception($"Failed to extract SNOW ticket data. {strExceptionMessage}"); }
                if (!(ListSnowTicketData != null && ListSnowTicketData.Count > 0))
                {
                    blnIsSuccess = true;
                    strExceptionMessage = "No SNOW ticket data pending to process";

                    DoLogging("HRePDFService", blnIsSuccess, strExceptionMessage);
                    return blnIsSuccess;
                }
                #endregion

                #region Group by SNOW ticket category
                var SnowTicketCatGroup =
                    ListSnowTicketData
                    .GroupBy(m => new
                    {
                        m.Category,
                        m.SubCategory
                    },
                    (key, group) => new
                    {
                        key.Category,
                        key.SubCategory,
                        SnowTicketDataList = group.ToList()
                    });
                #endregion

                #region Form folder for each SNOW ticket
                if (SnowTicketCatGroup != null)
                {
                    foreach (var groupItem in SnowTicketCatGroup)
                    {
                        if (groupItem.SnowTicketDataList != null && groupItem.SnowTicketDataList.Count > 0)
                        {
                            var SpecificConfigData = ListEPDFConfig.FirstOrDefault(m => m.SnowCategory.Trim().Equals(groupItem.Category.Trim(), StringComparison.OrdinalIgnoreCase)
                                                                                        && m.SnowSubCategory.Trim().Equals(groupItem.SubCategory.Trim(), StringComparison.OrdinalIgnoreCase));

                            if (SpecificConfigData != null && SpecificConfigData != default(EPDFConfig))
                            {
                                foreach (var ticketData in groupItem.SnowTicketDataList)
                                {
                                    try
                                    {
                                        if (ticketData != null)
                                        {
                                            #region Pre-checking
                                            if (ticketData.RaisedByInfo == null)
                                            {
                                                ticketData.IsActionComplete = true;
                                                ticketData.ActionMsg = "SNOW raised by info is null";
                                                continue;
                                            }

                                            if (ticketData.AttachInfo == null)
                                            {
                                                ticketData.IsActionComplete = true;
                                                ticketData.ActionMsg = "There is no attachment";
                                                continue;
                                            }

                                            //if RaisedBy personal area is empty, then use RaisedFor personal area to check (SSC/Accenture OPS will help APP user to raise ticket
                                            string strTicketPICPersonalArea = !string.IsNullOrWhiteSpace(ticketData.RaisedByInfo.PersonalArea) ?
                                                                              ticketData.RaisedByInfo.PersonalArea :
                                                                              ticketData.RaisedForInfo != null && !string.IsNullOrWhiteSpace(ticketData.RaisedForInfo.PersonalArea) ? ticketData.RaisedForInfo.PersonalArea : string.Empty;

                                            if (string.IsNullOrWhiteSpace(strTicketPICPersonalArea))
                                            {
                                                ticketData.IsActionComplete = true;
                                                ticketData.ActionMsg = "Both RaisedBy or RaisedFor personal area is empty";
                                                continue;
                                            }
                                            #endregion

                                            var SpecificMapData = ListFileFolderMap.FirstOrDefault(m => !string.IsNullOrWhiteSpace(m.PersonalArea) 
                                                                                                        && m.PersonalArea.Trim().Equals(strTicketPICPersonalArea, StringComparison.OrdinalIgnoreCase));

                                            if (SpecificMapData != null && SpecificMapData != default(FileFolderMapping))
                                            {
                                                #region Form folder path and unzip file
                                                string strOutputFolderPath = Path.Combine(SpecificConfigData.WorkingFolderPath, ProjectConstant.FolderName.CONST_INPUT, SpecificConfigData.ProcessName, SpecificMapData.SAPInstance, SpecificMapData.FolderName, ticketData.Number);
                                                if (!Directory.Exists(strOutputFolderPath))
                                                    Directory.CreateDirectory(strOutputFolderPath);

                                                string strZIPFilePath = Path.Combine(strOutputFolderPath, ticketData.AttachInfo.FileName);
                                                byte[] arrFileByte = Convert.FromBase64String(ticketData.AttachInfo.Base64);

                                                if (File.Exists(strZIPFilePath))
                                                    File.Delete(strZIPFilePath);

                                                File.WriteAllBytes(strZIPFilePath, arrFileByte);

                                                if (File.Exists(strZIPFilePath))
                                                {
                                                    blnIsActionSuccess = FileZipHelper.UnZipFile(strZIPFilePath, strOutputFolderPath, out strExceptionMessage);
                                                    File.Delete(strZIPFilePath);

                                                    string[] arrUnzipFile = Directory.GetFiles(strOutputFolderPath);

                                                    if (arrUnzipFile[0].Contains(".zip"))
                                                    {
                                                        string strNewZIPFilePath = Path.Combine(strOutputFolderPath, arrUnzipFile[0]);
                                                        FileZipHelper.UnZipFile(strNewZIPFilePath, strOutputFolderPath, out strExceptionMessage);
                                                        File.Delete(strNewZIPFilePath);

                                                        string NewTempFolder = Path.Combine(strOutputFolderPath, arrUnzipFile[0].Replace(".zip","").Replace("0-",""));
                                                        
                                                        if(Directory.Exists(NewTempFolder))
                                                        {
                                                            bool DirectoryFull = IsDirectoryEmpty(NewTempFolder);

                                                            if (!DirectoryFull)
                                                            {
                                                                string[] NewFiles = Directory.GetFiles(NewTempFolder);
                                                                foreach (var files in NewFiles)
                                                                {
                                                                    File.Copy(files, strOutputFolderPath + @"\" + Path.GetFileName(files));

                                                                }
                                                            }

                                                            var dir = new DirectoryInfo(NewTempFolder);
                                                            dir.Delete(true);
                                                        }
                                                        
                                                    }

                                                    ticketData.IsActionComplete = blnIsActionSuccess && arrUnzipFile != null && arrUnzipFile.Length > 0;
                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                ticketData.ActionMsg = $"{strTicketPICPersonalArea} not found in configuration";
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ticketData.ActionMsg = ex.Message;
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }
            finally
            {
                lngDBAction = SNOWHelper.UpsertSnowTicketData(ListSnowTicketData, out strDBException);
                DoLogging("UpsertSnowTicketData", (lngDBAction == 0), strDBException);

                lngDBAction = SNOWHelper.HousekeepSnowTicketData(14, out strDBException);
                DoLogging("HousekeepSnowTicketData", (lngDBAction == 0), strDBException);
            }

            return blnIsSuccess;
        }


        public bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }

        private bool ExtractMappingData (string strWorkingFolderPath, out List<FileFolderMapping> listFileFolderMap, out string strExceptionMessage)
        {
            bool blnIsSuccess = true;
            listFileFolderMap = null;
            strExceptionMessage = null;

            try
            {
                #region Pre-checking
                if (string.IsNullOrWhiteSpace(strWorkingFolderPath)) { throw new Exception("Mapping root folder path is null or empty"); }
                if (!Directory.Exists(strWorkingFolderPath)) { throw new Exception("Mapping root folder is not exist"); }


                string strMappingFolderPath = Path.Combine(strWorkingFolderPath, ProjectConstant.FolderName.CONST_MAPPING);

                string[] arrExcelFilePath = Directory.GetFiles(strMappingFolderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                if (!(arrExcelFilePath != null && arrExcelFilePath.Length > 0)) { throw new Exception("There is no excel file exists in mapping folder"); }

                string strMappingFilePath = arrExcelFilePath.FirstOrDefault(m => m.ToUpper().Contains("MAPPING_FOLDER"));
                if (!(!string.IsNullOrWhiteSpace(strMappingFilePath) && strMappingFilePath != default(string))) { throw new Exception($"There is no excel file found for folder mapping"); }

                string strMappingFileName = Path.GetFileName(strMappingFilePath);

                string strUserTempFolderPath = Path.GetTempPath();
                string strTempMappingFolderPath = Path.Combine(strUserTempFolderPath, ProjectConstant.ROCProcess.CONST_HR_EPDF, ProjectConstant.FolderName.CONST_MAPPING);
                string strTempMappingFilePath = Path.Combine(strTempMappingFolderPath, strMappingFileName);

                if (!Directory.Exists(strMappingFolderPath)) { throw new Exception("Mapping folder is not exists in the root folder"); }
                if (!File.Exists(strMappingFilePath)) { throw new Exception("Mapping file doesn't exists in folder"); }

                if (!Directory.Exists(strTempMappingFolderPath)) { Directory.CreateDirectory(strTempMappingFolderPath); }
                if (File.Exists(strTempMappingFilePath)) { File.Delete(strTempMappingFilePath); }

                File.Copy(strMappingFilePath, strTempMappingFilePath);

                if (!File.Exists(strTempMappingFilePath)) { throw new Exception("Failed to copy mapping file to temp folder"); }
                #endregion

                #region Convert mapping excel file to dataset
                ExcelHelper exlHelper = new ExcelHelper();
                blnIsSuccess = exlHelper.ExcelToDataSet(strTempMappingFilePath, out DataSet dsMapping, out strExceptionMessage);
                #endregion

                #region Assign data
                if (blnIsSuccess)
                {
                    if (dsMapping != null && dsMapping.Tables != null && dsMapping.Tables.Count > 0)
                    {
                        if (dsMapping.Tables.Contains("Personal Area PIC"))
                        {
                            DataTable dtblFolderPAPIC = dsMapping.Tables["Personal Area PIC"];
                            if (dtblFolderPAPIC.Rows != null && dtblFolderPAPIC.Rows.Count > 0)
                            {
                                listFileFolderMap = new List<FileFolderMapping>();
                                foreach (DataRow drFolderPAPIC in dtblFolderPAPIC.Rows)
                                {
                                    FileFolderMapping ffMap = new FileFolderMapping();
                                    ffMap.PersonalArea = drFolderPAPIC["Personal Area"].ToString();
                                    ffMap.FolderName = drFolderPAPIC["Folder"].ToString();
                                    ffMap.SAPInstance = drFolderPAPIC["SAP Instance"].ToString();

                                    listFileFolderMap.Add(ffMap);
                                }
                            }
                        }
                        else
                        {
                            throw new Exception("Personal Area PIC sheet is not exists");
                        }
                    }
                    else
                    {
                        throw new Exception("Mapping datatable is empty or null");
                    }
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

        private void DoLogging(string strProcessNm, bool blnIsSuccess, string strMessage)
        {
            CommonHelpers.SetLoggingMsg(Path.Combine(ProjectConstant.CONST_CONFIG_DEFAULT_PROJECT_FOLDER_PATH, "Log", $"IndoHRePDFSnowServiceLog_{DateTime.Now.ToString("yyyyMMdd")}.txt"),
                                        string.Format("{0} : Process {1} is {2}. {3}",
                                                        DateTime.Now.ToString("yyyyMMdd HHmmss"),
                                                        strProcessNm,
                                                        blnIsSuccess ? GlobalConstant.ROCConstant.CONST_ROC_SUCCESS_STATUS : GlobalConstant.ROCConstant.CONST_ROC_FAILED_STATUS,
                                                        strMessage) + Environment.NewLine);
        }
    }
}
