using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.Common.Config
{
    public class CommonConfigHelper
    {
        #region Properties
        private string DbConnectionString { get; set; }
        #endregion

        #region Constructor
        public CommonConfigHelper(string strDBConnectionString)
        {
            this.DbConnectionString = strDBConnectionString;
        }

        public CommonConfigHelper()
        {
            this.DbConnectionString = GetRPADBConnString();
        }
        #endregion

        #region Accessible Method
        #region RPA Calendar Config
        public long ExtractRPACalendarConfig(string strCalendarGroupList, out List<RPACalendar> ListRPACalendar, out string strExceptionMessage)
        {
            long lngResult = -1;
            DataSet dsResult = null;

            ListRPACalendar = null;
            strExceptionMessage = string.Empty;

            try
            {
                #region Prepare input parameter
                DataTable dtblInput = new DataTable();
                dtblInput.Columns.Add("CalendarGroupList", typeof(string));

                DataRow drInput = dtblInput.NewRow();
                drInput["CalendarGroupList"] = strCalendarGroupList;

                dtblInput.Rows.Add(drInput);
                #endregion

                #region Call store procedure
                DatabaseHelper dbHelper = new DatabaseHelper(this.DbConnectionString);
                lngResult = dbHelper.ExecuteStoreProcedureWithParameters(GlobalConstant.StoreProcedureName.CONST_USP_RPA_CALENDAR_SEL, dtblInput, out dsResult, out strExceptionMessage);
                #endregion

                #region Assign output value
                if (lngResult == 0)
                {
                    if (dsResult != null && dsResult.Tables != null && dsResult.Tables.Count > 0 && dsResult.Tables [0].Rows != null)
                    {
                        ListRPACalendar = new List<RPACalendar>();
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            RPACalendar rpaCalendar = new RPACalendar();
                            rpaCalendar.StartDateTime = (DateTime)dr["StartDateTime"];
                            rpaCalendar.EndDateTime = (DateTime)dr["EndDateTime"];
                            rpaCalendar.CalendarGroupCd = dr["CalendarGroupCd"].ToString();

                            ListRPACalendar.Add(rpaCalendar);
                        }
                    }
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

        public long ExtractRPACalendarConfig(List<string> RPACalendarGroupCdList, out List<RPACalendar> ListRPACalendar, out string strExceptionMessage)
        {
            long lngResult = -1;

            ListRPACalendar = null;
            strExceptionMessage = string.Empty;

            try
            {
                if (RPACalendarGroupCdList == null || (RPACalendarGroupCdList != null && RPACalendarGroupCdList.Count == 0))
                {
                    lngResult = 0;
                    RPACalendarGroupCdList = null;
                    strExceptionMessage = string.Empty;
                }
                else
                {
                    string strCalendarGroupList = $";{string.Join(";", RPACalendarGroupCdList)};";
                    return ExtractRPACalendarConfig(strCalendarGroupList, out ListRPACalendar, out strExceptionMessage);
                }
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }
        
        public static bool CheckIsChinaHoliday(DateTime dtInput, List<RPACalendar> ListChinaPublicHoliday, List<RPACalendar> ListChinaSpecialWorkingDay)
        {
            bool blnIsChinaHoliday = false;

            try
            {
                blnIsChinaHoliday = (dtInput.DayOfWeek == DayOfWeek.Saturday
                                    || dtInput.DayOfWeek == DayOfWeek.Sunday
                                    || (ListChinaPublicHoliday != null && ListChinaPublicHoliday.Count > 0 && ListChinaPublicHoliday.Any(m => dtInput >= m.StartDateTime && dtInput <= m.EndDateTime)))
                                    && !(ListChinaSpecialWorkingDay != null && ListChinaSpecialWorkingDay.Count > 0 && ListChinaSpecialWorkingDay.Any(m => dtInput >= m.StartDateTime && dtInput <= m.EndDateTime));
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return blnIsChinaHoliday;
        }

        public static bool CheckIsChinaWorkingDay(DateTime dtInput, List<RPACalendar> ListChinaSpecialWorkingDay)
        {
            bool blnIsChinaWorkingDay = false;

            try
            {
                blnIsChinaWorkingDay = (dtInput.DayOfWeek != DayOfWeek.Saturday && dtInput.DayOfWeek != DayOfWeek.Sunday) 
                                       || (ListChinaSpecialWorkingDay != null && ListChinaSpecialWorkingDay.Count > 0 && ListChinaSpecialWorkingDay.Any(m => dtInput >= m.StartDateTime && dtInput <= m.EndDateTime));
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return blnIsChinaWorkingDay;
        }
        #endregion

        #region RPA Tool Config
        public long ExtractRPAToolConfig(string strToolId, out RPAToolConfig rpaToolConfig, out string strExceptionMessage)
        {
            long lngResult = -1;
            DataSet dsResult = null;

            rpaToolConfig = null;
            strExceptionMessage = string.Empty;

            try
            {
                #region Prepare input data to query database
                DataTable dtblInput = new DataTable();
                dtblInput.Columns.Add("ToolId", typeof(Guid));

                DataRow drInput = dtblInput.NewRow();
                drInput["ToolId"] = strToolId;
                dtblInput.Rows.Add(drInput);
                #endregion

                #region Call store procedure
                DatabaseHelper dbHelper = new DatabaseHelper(this.DbConnectionString);
                lngResult = dbHelper.ExecuteStoreProcedureWithParameters(GlobalConstant.StoreProcedureName.CONST_USP_RPA_TOOL_CONFIG_SEL, dtblInput, out dsResult, out strExceptionMessage);
                #endregion

                #region Assign output value
                if (lngResult == 0)
                {
                    if (dsResult != null && dsResult.Tables != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows != null && dsResult.Tables[0].Rows.Count > 0)
                    {
                        DataRow dr = dsResult.Tables[0].Rows[0];
                        rpaToolConfig = new RPAToolConfig();
                        rpaToolConfig.Id = dr["Id"].ToString();
                        rpaToolConfig.ToolName = dr["ToolName"].ToString();
                        rpaToolConfig.Description = dr["Description"].ToString();
                        rpaToolConfig.OutputFolderPath = dr["OutputFolderPath"].ToString();

                        if (DateTime.TryParse(dr["RunDate"].ToString(), out DateTime dtRun))
                            rpaToolConfig.RunDate = dtRun;
                        else
                            rpaToolConfig.RunDate = null;
                    }
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
        #endregion

        #region RPA Parameter Config
        public long ExtractRPAParameter(out List<RPAParameter> RPAParamList, out string strExceptionMessage, params string[] arrParamGroup)
        {
            return ExtractRPAParameter(arrParamGroup != null ? arrParamGroup.ToList() : null, out RPAParamList, out strExceptionMessage);
        }

        public long ExtractRPAParameter(List<string> RPAParamGroupCdList, out List<RPAParameter> RPAParamList, out string strExceptionMessage)
        {
            #region Variable declaration
            long lngResult = -1;
            DataSet dsResult = null;
            List<RPAParameter> MasterRPAParamList = null;

            strExceptionMessage = string.Empty;
            RPAParamList = null;
            #endregion

            try
            {
                if (RPAParamGroupCdList == null || (RPAParamGroupCdList != null && RPAParamGroupCdList.Count == 0))
                {
                    lngResult = 0;
                    RPAParamList = null;
                    strExceptionMessage = string.Empty;
                }
                else
                {
                    #region Prepare input parameter
                    DataSet dsTVP = new DataSet();

                    DataTable dtblTVP = dsTVP.Tables.Add("RPAParameterTVPTable");
                    dtblTVP.Columns.Add("ParamGroup", typeof(string));

                    foreach (string strParamGrp in RPAParamGroupCdList)
                    {
                        DataRow drTVP = dtblTVP.NewRow();
                        drTVP["ParamGroup"] = strParamGrp;
                        dtblTVP.Rows.Add(drTVP);
                    }
                    #endregion

                    #region Call store procedure
                    DatabaseHelper dbHelper = new DatabaseHelper(this.DbConnectionString);
                    lngResult = dbHelper.ExecuteStoreProcedureWithTVP(GlobalConstant.StoreProcedureName.CONST_USP_RPA_PARAMETER_SEL, dsTVP, out dsResult, out strExceptionMessage);
                    #endregion

                    #region Assign output value
                    if (lngResult == 0)
                    {
                        if (dsResult != null && dsResult.Tables != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows != null)
                        {
                            MasterRPAParamList = new List<RPAParameter>();
                            foreach (DataRow dr in dsResult.Tables[0].Rows)
                            {
                                RPAParameter rpaParam = new RPAParameter();
                                rpaParam.Id = dr["Id"].ToString();
                                rpaParam.ParamGroup = dr["ParamGroup"].ToString();
                                rpaParam.ParentId = dr["ParentId"].ToString();
                                rpaParam.ParamKey = dr["ParamKey"].ToString();
                                rpaParam.ParamValue = dr["ParamValue"].ToString();
                                rpaParam.Description = dr["Description"].ToString();

                                MasterRPAParamList.Add(rpaParam);
                            }
                        }
                    }
                    #endregion

                    #region Construct RPA Common Config List
                    if (MasterRPAParamList != null && MasterRPAParamList.Count > 0)
                    {
                        List<RPAParameter> ParentRPAParamList = MasterRPAParamList.Where(m => string.IsNullOrWhiteSpace(m.ParentId)).ToList(); //Get the config from most top hierarchy
                        if (ParentRPAParamList != null && ParentRPAParamList.Count > 0)
                        {
                            RPAParamList = new List<RPAParameter>();
                            foreach (var parentParam in ParentRPAParamList)
                            {
                                RPAParameter rpaCommonConf = new RPAParameter();
                                rpaCommonConf.Id = parentParam.Id;
                                rpaCommonConf.ParamGroup = parentParam.ParamGroup;
                                rpaCommonConf.ParentId = parentParam.ParentId;
                                rpaCommonConf.ParamKey = parentParam.ParamKey;
                                rpaCommonConf.ParamValue = parentParam.ParamValue;
                                rpaCommonConf.Description = parentParam.Description;
                                rpaCommonConf.RPAParamChildList = FormRPAParamChildList(parentParam.Id, MasterRPAParamList); //Search for children configuration using recursive method

                                RPAParamList.Add(rpaCommonConf);
                            }
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }

        public long UpdateRPAParameter(List<RPAParameter> RPAParamList, out string strExceptionMessage)
        {
            long lngResult = -1;
            DataSet dsResult = null;

            strExceptionMessage = string.Empty;

            try
            {
                if (RPAParamList != null && RPAParamList.Count > 0)
                {
                    #region Prepare input parameter
                    DataSet dsTVP = new DataSet();

                    DataTable dtblTVP = dsTVP.Tables.Add("RPAParamUpdTVPTable");
                    dtblTVP.Columns.Add("Id", typeof(string));
                    dtblTVP.Columns.Add("ParamValue", typeof(string));

                    foreach (var RPAParam in RPAParamList)
                    {
                        DataRow drTVP = dtblTVP.NewRow();
                        drTVP["Id"] = RPAParam.Id;
                        drTVP["ParamValue"] = RPAParam.ParamValue;

                        dtblTVP.Rows.Add(drTVP);
                    }
                    #endregion

                    #region Call store procedure
                    DatabaseHelper dbHelper = new DatabaseHelper(this.DbConnectionString);
                    lngResult = dbHelper.ExecuteStoreProcedureWithTVP(GlobalConstant.StoreProcedureName.CONST_USP_RPA_PARAMETER_UPD, dsTVP, out dsResult, out strExceptionMessage);
                    #endregion
                }
                else
                {
                    lngResult = 0;
                    strExceptionMessage = string.Empty;
                }
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }
        #endregion

        #region RPA Credential Config
        public long ExtractRPACredential(string strCredGroup, out List<RPACredential> RPACredList, out string strExceptionMessage)
        {
            long lngResult = -1;
            DataSet dsResult = null;

            RPACredList = null;
            strExceptionMessage = string.Empty;

            try
            {
                #region Prepare input parameter
                DataTable dtblReq = new DataTable();
                dtblReq.Columns.Add("Group", typeof(string));

                DataRow drReq = dtblReq.NewRow();
                drReq["Group"] = strCredGroup;
                dtblReq.Rows.Add(drReq);
                #endregion

                #region Call store procedure
                DatabaseHelper dbHelper = new DatabaseHelper(this.DbConnectionString);
                lngResult = dbHelper.ExecuteStoreProcedureWithParameters(GlobalConstant.StoreProcedureName.CONST_USP_RPA_CREDENTIAL_SEL, dtblReq, out dsResult, out strExceptionMessage);
                #endregion

                #region Assign output value
                if (lngResult == 0)
                {
                    if (dsResult != null && dsResult.Tables != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows != null && dsResult.Tables[0].Rows.Count > 0)
                    {
                        RPACredList = new List<RPACredential>();
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            RPACredential rpaCred = new RPACredential();
                            rpaCred.Id = dr["Id"].ToString();
                            rpaCred.UserName = dr["UserName"].ToString();
                            rpaCred.Password = dr["Password"].ToString();
                            rpaCred.Group = dr["Group"].ToString();
                            rpaCred.Key = dr["Key"].ToString();
                            rpaCred.Description = dr["Description"].ToString();
                            rpaCred.Active = (bool)dr["Active"];
                            rpaCred.Platform = dr["Platform"].ToString();
                            rpaCred.Instance = dr["Instance"].ToString();
                            rpaCred.Language = dr["Language"].ToString();
                            rpaCred.Region = dr["Region"].ToString();
                            rpaCred.Tower = dr["Tower"].ToString();
                            rpaCred.Is360 = (bool)dr["Is360"]; 


                            if (dr["CreateDateTime"] != DBNull.Value
                                && DateTime.TryParse(dr["CreateDateTime"].ToString(), out DateTime dtCreate))
                                rpaCred.CreateDateTime = dtCreate;
                            else
                                rpaCred.CreateDateTime = null;

                            if (dr["UpdateDateTime"] != DBNull.Value
                                && DateTime.TryParse(dr["UpdateDateTime"].ToString(), out DateTime dtUpdate))
                                rpaCred.UpdateDateTime = dtUpdate;
                            else
                                rpaCred.UpdateDateTime = null;

                            RPACredList.Add(rpaCred);
                        }
                    }
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

        public long UpdateRPACredential(List<RPACredential> RPACredList, out string strExceptionMessage)
        {
            long lngResult = -1;
            DataSet dsResult = null;

            strExceptionMessage = string.Empty;

            try
            {
                if (RPACredList != null && RPACredList.Count > 0)
                {
                    #region Prepare input parameter
                    DataSet dsTVP = new DataSet();

                    DataTable dtblTVP = dsTVP.Tables.Add("RPACredentialUpdTable");
                    dtblTVP.Columns.Add("Id", typeof(string));
                    dtblTVP.Columns.Add("UserName", typeof(string));
                    dtblTVP.Columns.Add("Password", typeof(string));

                    foreach (var RPACred in RPACredList)
                    {
                        DataRow drTVP = dtblTVP.NewRow();
                        drTVP["Id"] = RPACred.Id;
                        drTVP["UserName"] = RPACred.UserName;
                        drTVP["Password"] = RPACred.Password;

                        dtblTVP.Rows.Add(drTVP);
                    }
                    #endregion

                    #region Call store procedure
                    DatabaseHelper dbHelper = new DatabaseHelper(this.DbConnectionString);
                    lngResult = dbHelper.ExecuteStoreProcedureWithTVP(GlobalConstant.StoreProcedureName.CONST_USP_RPA_CREDENTIAL_UPD, dsTVP, out dsResult, out strExceptionMessage);
                    #endregion
                }
                else
                {
                    lngResult = 0;
                    strExceptionMessage = string.Empty;
                }
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }
        #endregion

        #region RPA Email Template Config
        public long ExtractRPAEmailTemplate(out List<RPAEmailTemplate> RPAEmlTemplateList, out string strExceptionMessage, params string[] arrTemplateGroup)
        {
            return ExtractRPAEmailTemplate(arrTemplateGroup != null ? arrTemplateGroup.ToList() : null, out RPAEmlTemplateList, out strExceptionMessage);
        }

        public long ExtractRPAEmailTemplate(List<string> TemplateGrpCdList, out List<RPAEmailTemplate> RPAEmlTemplateList, out string strExceptionMessage)
        {
            #region Variable declaration
            long lngResult = -1;
            DataSet dsResult = null;

            RPAEmlTemplateList = null;
            strExceptionMessage = null;
            #endregion

            try
            {
                if (TemplateGrpCdList == null || (TemplateGrpCdList != null && TemplateGrpCdList.Count == 0))
                {
                    lngResult = 0;
                    RPAEmlTemplateList = null;
                    strExceptionMessage = string.Empty;
                }
                else
                {
                    #region Prepare input parameter
                    DataSet dsTVP = new DataSet();

                    DataTable dtblTVP = dsTVP.Tables.Add("RPAEmlTemplateTVPTable");
                    dtblTVP.Columns.Add("EmailTemplateGroup", typeof(string));

                    foreach (string strTemplateGrp in TemplateGrpCdList)
                    {
                        DataRow drTVP = dtblTVP.NewRow();
                        drTVP["EmailTemplateGroup"] = strTemplateGrp;
                        dtblTVP.Rows.Add(drTVP);
                    }
                    #endregion

                    #region Call store procedure
                    DatabaseHelper dbHelper = new DatabaseHelper(this.DbConnectionString);
                    lngResult = dbHelper.ExecuteStoreProcedureWithTVP(GlobalConstant.StoreProcedureName.CONST_USP_RPA_EMAIL_TEMPLATE_SEL, dsTVP, out dsResult, out strExceptionMessage);
                    #endregion

                    #region Assign output value
                    if (lngResult == 0)
                    {
                        if (dsResult != null && dsResult.Tables != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows != null)
                        {
                            RPAEmlTemplateList = new List<RPAEmailTemplate>();

                            foreach (DataRow drResult in dsResult.Tables[0].Rows)
                            {
                                RPAEmailTemplate emailTemplate = new RPAEmailTemplate();
                                emailTemplate.EmailTemplateCode = drResult["EmailTemplateCd"].ToString();
                                emailTemplate.EmailSubject = drResult["EmailSubject"].ToString();
                                emailTemplate.EmailSubjectParam = drResult["EmailSubjectParam"].ToString();
                                emailTemplate.EmailBody = drResult["EmailBody"].ToString();
                                emailTemplate.EmailBodyParam = drResult["EmailBodyParam"].ToString();
                                emailTemplate.EmailTemplateGroup = drResult["EmailTemplateGroup"].ToString();

                                RPAEmlTemplateList.Add(emailTemplate);

                                if (dsResult.Tables.Count > 1 && dsResult.Tables[1].Rows != null && dsResult.Tables[1].Rows.Count > 0)
                                {
                                    DataRow[] arrOverride = dsResult.Tables[1].Select(string.Format("EmailTemplateCd = '{0}'", emailTemplate.EmailTemplateCode));
                                    if (arrOverride != null && arrOverride.Length > 0)
                                    {
                                        emailTemplate.OverrideEmlTemplateList = new List<RPAEmailTemplateOverride>();

                                        foreach (DataRow drOverride in arrOverride)
                                        {
                                            RPAEmailTemplateOverride emlTempOverride = new RPAEmailTemplateOverride();
                                            emlTempOverride.EmailTemplateCode = drOverride["EmailTemplateCd"].ToString();
                                            emlTempOverride.EmailSubject = drOverride["EmailSubject"].ToString();
                                            emlTempOverride.EmailSubjectParam = drOverride["EmailSubjectParam"].ToString();
                                            emlTempOverride.EmailBody = drOverride["EmailBody"].ToString();
                                            emlTempOverride.EmailBodyParam = drOverride["EmailBodyParam"].ToString();
                                            emlTempOverride.EmailTemplateGroup = drOverride["EmailTemplateGroup"].ToString();
                                            emlTempOverride.OverrideParam = drOverride["OverrideParam"].ToString();

                                            emailTemplate.OverrideEmlTemplateList.Add(emlTempOverride);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                lngResult = -99;
                strExceptionMessage = ex.Message;
            }

            return lngResult;
        }

        public static string FormatRPAEmailTemplate<T>(List<RPAEmailTemplate> ListEmailTemplateConfig, RPAEmailTemplateType emailTemplateType, string strEmailTemplateCode, T typeObject, Dictionary<string, string> dicExtraParam = null, string strOverrideParam = null)
        {
            string strOutput = null;

            try
            {
                if (ListEmailTemplateConfig != null && ListEmailTemplateConfig.Count > 0 && !string.IsNullOrWhiteSpace(strEmailTemplateCode) && typeObject != null)
                {
                    var SpecificEmailTemplateConfig = ListEmailTemplateConfig.FirstOrDefault(m => m.EmailTemplateCode.Equals(strEmailTemplateCode, StringComparison.OrdinalIgnoreCase));
                    if (SpecificEmailTemplateConfig != null && SpecificEmailTemplateConfig != default(RPAEmailTemplate))
                        strOutput = FormatRPAEmailTemplate(SpecificEmailTemplateConfig, emailTemplateType, typeObject, dicExtraParam, strOverrideParam);
                }
            }
            catch (Exception ex)
            {
                strOutput = string.Empty;
            }

            return strOutput;
        }

        public static string FormatRPAEmailTemplate(List<RPAEmailTemplate> ListEmailTemplateConfig, RPAEmailTemplateType emailTemplateType, string strEmailTemplateCode, DataRow drData, Dictionary<string, string> dicExtraParam = null, string strOverrideParam = null)
        {
            string strOutput = null;

            try
            {
                if (ListEmailTemplateConfig != null && ListEmailTemplateConfig.Count > 0 && !string.IsNullOrWhiteSpace(strEmailTemplateCode) && drData != null)
                {
                    var SpecificEmailTemplateConfig = ListEmailTemplateConfig.FirstOrDefault(m => m.EmailTemplateCode.Equals(strEmailTemplateCode, StringComparison.OrdinalIgnoreCase));
                    if (SpecificEmailTemplateConfig != null && SpecificEmailTemplateConfig != default(RPAEmailTemplate))
                    {
                        string strEmailTemplate = (emailTemplateType == RPAEmailTemplateType.Subject) ? SpecificEmailTemplateConfig.EmailSubject : SpecificEmailTemplateConfig.EmailBody;
                        string strEmailTemplateParam = (emailTemplateType == RPAEmailTemplateType.Subject) ? SpecificEmailTemplateConfig.EmailSubjectParam : SpecificEmailTemplateConfig.EmailBodyParam;

                        #region Override template (if have)
                        if (!string.IsNullOrWhiteSpace(strOverrideParam)
                            && SpecificEmailTemplateConfig.OverrideEmlTemplateList != null && SpecificEmailTemplateConfig.OverrideEmlTemplateList.Count > 0)
                        {
                            var overrideEmlTempConfig = SpecificEmailTemplateConfig.OverrideEmlTemplateList.FirstOrDefault(m => !string.IsNullOrWhiteSpace(m.OverrideParam) && m.OverrideParam.Trim().Equals(strOverrideParam.Trim()));
                            if (overrideEmlTempConfig != null && overrideEmlTempConfig != default(RPAEmailTemplateOverride))
                            {
                                if (emailTemplateType == RPAEmailTemplateType.Subject && !string.IsNullOrWhiteSpace(overrideEmlTempConfig.EmailSubject))
                                {
                                    strEmailTemplate = overrideEmlTempConfig.EmailSubject;
                                    strEmailTemplateParam = overrideEmlTempConfig.EmailSubjectParam;
                                }
                                else if (emailTemplateType == RPAEmailTemplateType.Body && !string.IsNullOrWhiteSpace(overrideEmlTempConfig.EmailBody))
                                {
                                    strEmailTemplate = overrideEmlTempConfig.EmailBody;
                                    strEmailTemplateParam = overrideEmlTempConfig.EmailBodyParam;
                                }
                            }
                        }
                        #endregion

                        strOutput = strEmailTemplate;
                        if (dicExtraParam != null && dicExtraParam.Count > 0)
                        {
                            foreach (var keyValPair in dicExtraParam)
                                strOutput = strOutput.Replace(keyValPair.Key, keyValPair.Value);
                        }

                        string[] arrSubjectParamPerConfig = strEmailTemplateParam.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                        if (arrSubjectParamPerConfig != null && arrSubjectParamPerConfig.Length > 0)
                        {
                            foreach (string strPerConfig in arrSubjectParamPerConfig)
                            {
                                string[] arrKeyValuePair = strPerConfig.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                                if (arrKeyValuePair != null && arrKeyValuePair.Length > 1)
                                {
                                    string strKey = arrKeyValuePair[0];
                                    string strValueColumnNm = arrKeyValuePair[1];

                                    if (drData.Table.Columns.Contains(strValueColumnNm))
                                        strOutput = strOutput.Replace("{" + strKey + "}", drData[strValueColumnNm].ToString());
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strOutput;
        }

        public static string FormatRPAEmailTemplate<T>(RPAEmailTemplate SpecificEmailTemplateConfig, RPAEmailTemplateType emailTemplateType, T typeObject, Dictionary<string, string> dicExtraParam = null, string strOverrideParam = null)
        {
            string strOutput = null;

            try
            {
                if (typeObject != null)
                {
                    string strEmailTemplate = (emailTemplateType == RPAEmailTemplateType.Subject) ? SpecificEmailTemplateConfig.EmailSubject : SpecificEmailTemplateConfig.EmailBody;
                    string strEmailTemplateParam = (emailTemplateType == RPAEmailTemplateType.Subject) ? SpecificEmailTemplateConfig.EmailSubjectParam : SpecificEmailTemplateConfig.EmailBodyParam;

                    #region Override template (if have)
                    if (!string.IsNullOrWhiteSpace(strOverrideParam) 
                        && SpecificEmailTemplateConfig.OverrideEmlTemplateList != null && SpecificEmailTemplateConfig.OverrideEmlTemplateList.Count > 0)
                    {
                        var overrideEmlTempConfig = SpecificEmailTemplateConfig.OverrideEmlTemplateList.FirstOrDefault(m => !string.IsNullOrWhiteSpace(m.OverrideParam) && m.OverrideParam.Trim().Equals(strOverrideParam.Trim()));
                        if (overrideEmlTempConfig != null && overrideEmlTempConfig != default(RPAEmailTemplateOverride))
                        {
                            if (emailTemplateType == RPAEmailTemplateType.Subject && !string.IsNullOrWhiteSpace(overrideEmlTempConfig.EmailSubject))
                            {
                                strEmailTemplate = overrideEmlTempConfig.EmailSubject;
                                strEmailTemplateParam = overrideEmlTempConfig.EmailSubjectParam;
                            }
                            else if (emailTemplateType == RPAEmailTemplateType.Body && !string.IsNullOrWhiteSpace(overrideEmlTempConfig.EmailBody))
                            {
                                strEmailTemplate = overrideEmlTempConfig.EmailBody;
                                strEmailTemplateParam = overrideEmlTempConfig.EmailBodyParam;
                            }
                        }
                    }
                    #endregion

                    strOutput = strEmailTemplate;
                    if (dicExtraParam != null && dicExtraParam.Count > 0)
                    {
                        foreach (var keyValPair in dicExtraParam)
                            strOutput = strOutput.Replace(keyValPair.Key, keyValPair.Value);
                    }

                    string[] arrSubjectParamPerConfig = strEmailTemplateParam.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    if (arrSubjectParamPerConfig != null && arrSubjectParamPerConfig.Length > 0)
                    {
                        foreach (string strPerConfig in arrSubjectParamPerConfig)
                        {
                            string[] arrKeyValuePair = strPerConfig.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                            if (arrKeyValuePair != null && arrKeyValuePair.Length > 1)
                            {
                                string strKey = arrKeyValuePair[0];
                                string strValueNm = arrKeyValuePair[1];

                                strOutput = strOutput.Replace("{" + strKey + "}", CommonHelpers.GetPropertyValueByName(typeObject, strValueNm));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                strOutput = string.Empty;
            }

            return strOutput;
        }
        #endregion

        #region RPA Common APP Config File
        public static string GetRPACommonAppConfigFilePath()
        {
            
            try
            {
                #if DEBUG
                    string strAppConfigFilePath = @"C:\innersource\common\TMFramework\DotNet\TMHelper\TMHelper.Common\ConfigFile\RPAMaster\AppConfig.xml";
                #else
                    string strAppConfigFilePath = @"\\globalnet\aspiro\RPA\RPA_Working\Config\RPAMaster\AppConfig.xml";
                #endif

                return strAppConfigFilePath;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        public static string GetRPADBConnString(bool? blnIsTestInstance = null)
        {
            try
            {
                bool blnIsDebug = false;
                
                #if DEBUG
                    blnIsDebug = true;
                #else
                    blnIsDebug = false;
                #endif

                blnIsDebug = blnIsTestInstance != null && blnIsTestInstance.HasValue ? blnIsTestInstance.Value : blnIsDebug;

                #region Initialize config key
                string strDBServerIPConfigKey = !blnIsDebug ? "DatabaseIP" : "DatabaseIP_Staging";
                string strDBInitCatConfigKey = !blnIsDebug ? "DatabaseInitCat" : "DatabaseInitCat_Staging";
                string strDBUIDConfigKey = !blnIsDebug ? "DatabaseUID" : "DatabaseUID_Staging";
                string strDBPassConfigKey = !blnIsDebug ? "DatabasePassword" : "DatabasePassword_Staging";
                #endregion

                #region Retrieve config value from RPA common APP config file path
                string strAppConfigFilePath = GetRPACommonAppConfigFilePath();

                string strConfigDBServerIP = CommonHelpers.GetConfig(strAppConfigFilePath, strDBServerIPConfigKey);
                string strConfigDBInitCatalog = CommonHelpers.GetConfig(strAppConfigFilePath, strDBInitCatConfigKey);
                string strConfigDBUID = CommonHelpers.GetConfig(strAppConfigFilePath, strDBUIDConfigKey);
                string strConfigDBPassword = CommonHelpers.GetConfig(strAppConfigFilePath, strDBPassConfigKey);
                #endregion

                #region Construct DB connection string
                string strDBConnString = $"Data Source={strConfigDBServerIP};Initial Catalog={strConfigDBInitCatalog};User ID={CryptographyHelper.Decrypt(strConfigDBUID)};Password={CryptographyHelper.Decrypt(strConfigDBPassword)};Connect Timeout=40;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
                #endregion

                return strDBConnString;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
        #endregion

        #region Helper
        //Recursive Method
        private List<RPAParameter> FormRPAParamChildList(string strParentId, List<RPAParameter> MasterRPAParamList)
        {
            List<RPAParameter> RPAParamChildList = null;

            try
            {
                if (MasterRPAParamList != null && MasterRPAParamList.Count > 0)
                {
                    List<RPAParameter> specificMasterParamList = MasterRPAParamList.Where(m => !string.IsNullOrWhiteSpace(m.ParentId)
                                                                                               && m.ParentId.Trim().Equals(strParentId.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();

                    if (specificMasterParamList != null && specificMasterParamList.Count > 0)
                    {
                        RPAParamChildList = new List<RPAParameter>();

                        foreach (var masterData in specificMasterParamList)
                        {
                            RPAParameter rpaCommonConf = new RPAParameter();
                            rpaCommonConf.Id = masterData.Id;
                            rpaCommonConf.ParamGroup = masterData.ParamGroup;
                            rpaCommonConf.ParentId = masterData.ParentId;
                            rpaCommonConf.ParamKey = masterData.ParamKey;
                            rpaCommonConf.ParamValue = masterData.ParamValue;
                            rpaCommonConf.Description = masterData.Description;
                            rpaCommonConf.RPAParamChildList = FormRPAParamChildList(masterData.Id, MasterRPAParamList); //Call method recursively to find child configuration

                            RPAParamChildList.Add(rpaCommonConf);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return RPAParamChildList;
        }
        #endregion
    }
}
