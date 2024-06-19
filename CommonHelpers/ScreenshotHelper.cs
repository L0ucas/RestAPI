using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using HWND = System.IntPtr;

namespace TMHelper.Common
{
    public static class ScreenShotHelper
    {
        public static bool CaptureToFile(string strSaveFilePath, out string strExceptionMessage)
        {
            bool blnIsSuccess = false;
            strExceptionMessage = string.Empty;

            try
            {
                Bitmap bmp = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(0, 0, 0, 0, Screen.PrimaryScreen.Bounds.Size);
                    bmp.Save(strSaveFilePath); 
                }

                blnIsSuccess = true;
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;

                blnIsSuccess = false;
                strExceptionMessage = string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message);
            }

            return blnIsSuccess;
        }

        public static bool BringWindowFrontAndCaptureToFile(string strWinNameToBringFront, string strSaveFilePath, out string strExceptionMessage)
        {
            bool blnIsSuccess = false;
            strExceptionMessage = string.Empty;

            try
            {
                ProcessesAndWindowsHelper.BringWindowToFront(strWinNameToBringFront); //Bring specific window to front
                blnIsSuccess = CaptureToFile(strSaveFilePath, out strExceptionMessage); //Capture screenshot and save to file
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;

                blnIsSuccess = false;
                strExceptionMessage = string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message);
            }

            return blnIsSuccess;
        }

        public static MemoryStream CaptureToStream(out string strExceptionMessage)
        {
            return CaptureToStream(ImageFormat.Png, out strExceptionMessage);
        }

        public static MemoryStream CaptureToStream(ImageFormat imageFormat, out string strExceptionMessage)
        {
            MemoryStream ms = null;
            strExceptionMessage = string.Empty;

            try
            {
                if (imageFormat == null) imageFormat = ImageFormat.Png; //default to png

                Bitmap bmp = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    ms = new MemoryStream();

                    g.CopyFromScreen(0, 0, 0, 0, Screen.PrimaryScreen.Bounds.Size);
                    bmp.Save(ms, imageFormat);

                    ms.Position = 0;
                }
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                if (ms != null) ms.Dispose();
                strExceptionMessage = string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message);
            }

            return ms;
        }

        public static MemoryStream BringWindowFrontAndCaptureToStream(string strWinNameToBringFront, out string strExceptionMessage)
        {
            return BringWindowFrontAndCaptureToStream(strWinNameToBringFront, ImageFormat.Png, out strExceptionMessage);
        }

        public static MemoryStream BringWindowFrontAndCaptureToStream(string strWinNameToBringFront, ImageFormat imageFormat, out string strExceptionMessage)
        {
            MemoryStream ms = null;
            strExceptionMessage = string.Empty;
            
            try
            {
                if (imageFormat == null) imageFormat = ImageFormat.Png; //default to png
                ProcessesAndWindowsHelper.BringWindowToFront(strWinNameToBringFront); //Bring specific window to front
                ms = CaptureToStream(imageFormat, out strExceptionMessage); //capture screenshot and assign to stream
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                if (ms != null) ms.Dispose();
                strExceptionMessage = string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message);
            }

            return ms;
        }
    }
}
