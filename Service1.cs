using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using System.Threading.Tasks;
using System.Configuration;
using TMHelper.Common;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = null;
        EventLog eventLog = null;

        public Service1()
        {
            InitializeComponent();

            timer = new Timer();
            eventLog = new EventLog();

            if (!EventLog.SourceExists("IndoHRePDFService"))
            {
                EventLog.CreateEventSource("IndoHRePDFService", "");
            }

            eventLog.Source = "IndoHRePDFService";
            eventLog.Log = "";
        }

        protected override void OnStart(string[] args)
        {
            eventLog.WriteEntry("Indo HR EPDF Window Service Is Started");

            int iSerRunPerMs = 60000;
            string strSerRunPerMs = ConfigurationManager.AppSettings["ServiceRunPerMilliSec"].ToString();
            if (!string.IsNullOrWhiteSpace(strSerRunPerMs) && int.TryParse(strSerRunPerMs, out int iConfigVal))
                iSerRunPerMs = iConfigVal;

            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = iSerRunPerMs;
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            eventLog.WriteEntry("Indo HR EPDF Window Service Has Stop");
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            #region Roc Initialization
            string strMessage = string.Empty;

            ROCHelper rocHelper = new ROCHelper();
            rocHelper.Deal = GlobalConstant.ROCConstant.CONST_ROC_FA_DEAL;
            rocHelper.Process = "188435-app-hr-id-npayroll-epdf_automation-WindowService";
            rocHelper.TransactionId = "HR EPDF Automation - Window Serivce";
            rocHelper.StartTime = DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss tt");
            rocHelper.EndTime = string.Empty;
            rocHelper.Status = string.Empty;
            rocHelper.Description = string.Empty;
            rocHelper.DateFormat = "dd/MM/yyyy";
            #endregion

            try
            {
                #region Start process
                HRePDFService hrEPDFServ = new HRePDFService();
                bool blnIsSuccess = hrEPDFServ.StartProcess(out string strExceptionMessage);
                #endregion

                #region Logging
                if (!blnIsSuccess)
                    strMessage = string.Format("Error(s) occurred while execute spoke kiosk service. {0}", strExceptionMessage);
                else
                    strMessage = "Successfully execute spoke kiosk service";

                eventLog.WriteEntry(strMessage, blnIsSuccess ? EventLogEntryType.Information : EventLogEntryType.Error);

                #region Assign ROC value
                rocHelper.EndTime = DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss tt");
                rocHelper.Status = (blnIsSuccess) ? GlobalConstant.ROCConstant.CONST_ROC_SUCCESS_STATUS : GlobalConstant.ROCConstant.CONST_ROC_FAILED_STATUS;
                rocHelper.Description = string.Format("Total record : 1. {0}", strMessage);
                #endregion
                #endregion
            }
            catch (Exception ex)
            {
                #region Logging
                strMessage = string.Format("Exception occurred while execute HR EPDF service. {0}", ex.Message);
                eventLog.WriteEntry(strMessage, EventLogEntryType.Error);

                #region Assign ROC value
                rocHelper.EndTime = DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss tt");
                rocHelper.Status = GlobalConstant.ROCConstant.CONST_ROC_FAILED_STATUS;
                rocHelper.Description = string.Format("Total record : 0. {0}", strMessage);
                #endregion
                #endregion
            }
            finally
            {
                #region Log to ROC
                rocHelper.PostDataToRocService();
                #endregion
            }
        }
    }
}
