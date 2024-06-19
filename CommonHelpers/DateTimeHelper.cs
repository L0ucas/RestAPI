using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.Common
{
    public static class DateTimeHelper
    {
        public class TimeZoneId
        {
            public const string SEAsiaStandardTime = "SE Asia Standard Time"; //Indonesia time zone
        }

        public static DateTime GetCurrentDateTimeOfDifferentTimeZone(string strTimeZoneId)
        {
            DateTime dtResult = DateTime.Now;

            try
            {
                dtResult = ConvertDateTimeToDifferentTimeZone(DateTime.Now, strTimeZoneId);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dtResult;
        }

        public static DateTime ConvertDateTimeToDifferentTimeZone(DateTime dtInput, string strTimeZoneId)
        {
            DateTime dtResult = DateTime.Now;

            try
            {
                var timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(strTimeZoneId);
                dtResult = TimeZoneInfo.ConvertTime(dtInput, TimeZoneInfo.Local, timeZoneInfo);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dtResult;
        }

        public static string GetLastDateOfTheMonth(string strDate)
        {
            string strResult = string.Empty;

            DateTime dtResult = new DateTime();

            try
            {
                if (DateTime.TryParse(strDate, out dtResult))
                {
                    dtResult = dtResult.AddMonths(1);
                    dtResult = new DateTime(dtResult.Year, dtResult.Month, 1);
                    dtResult = dtResult.AddDays(-1);
                    strResult = dtResult.ToString("yyyy/MM/dd");
                }
            }
            catch (Exception)
            {
                strResult = string.Empty;
            }

            return strResult;
        }

        public static string GetDateDiff(string strDateFrom, string strDateTo)
        {
            string strResult = string.Empty;

            DateTime dtFrom = new DateTime();
            DateTime dtTo = new DateTime();

            try
            {
                if (DateTime.TryParse(strDateFrom, out dtFrom) && DateTime.TryParse(strDateTo, out dtTo))
                {
                    strResult = (dtTo - dtFrom).TotalDays.ToString();
                }
            }
            catch (Exception)
            {
                strResult = string.Empty;
            }

            return strResult;
        }

        public static int GetTotalMonthDiff(DateTime dtDateOne, DateTime dtDateTwo, bool blnIsIncludeCurMonth = true)
        {
            try
            {
                DateTime dtEarly = (dtDateOne > dtDateTwo) ? dtDateTwo.Date : dtDateOne.Date;
                DateTime dtLate = (dtDateOne > dtDateTwo) ? dtDateOne.Date : dtDateTwo.Date;

                int iMonthDiff = 0;
                DateTime dtNew = dtEarly.AddMonths(iMonthDiff);

                while (dtNew <= dtLate || (dtNew.Month == dtLate.Month && dtNew.Year == dtLate.Year))
                {
                    iMonthDiff++;
                    dtNew = dtEarly.AddMonths(iMonthDiff);
                }

                return blnIsIncludeCurMonth ? iMonthDiff : iMonthDiff - 1;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
