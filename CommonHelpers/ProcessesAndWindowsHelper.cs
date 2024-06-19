using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using HWND = System.IntPtr;

namespace TMHelper.Common
{
    public static class ProcessesAndWindowsHelper
    {
        #region Constant
        public class ControlAction
        {
            public const uint WM_SETTEXT = 0x000C;
            public const uint WM_CLOSE = 0x0010;
            public const uint BTN_CLICK = 0x00F5;
        }

        public class ControlName
        {
            public const string CONST_COMBO_BOX_EDITABLE = "ComboBoxEx32";
            public const string CONST_COMBO_BOX = "ComboBox";
            public const string CONST_BUTTON = "Button";
        }

        public class AppFileAccessMode
        {
            public const int READ = 0;
            public const int WRITE = 1;
            public const int READWRITE = 2;
        }

        public class AppFileSharedAccessMode
        {
            public const int SHARE_COMPAT = 0x0; //Allow open by other instance multiple times
            public const int SHARE_EXCLUSIVE = 0x10;  //Only one instance is able to open
            public const int SHARE_DENY_WRITE = 0x20; //Allow to be open by other instance, but write is denied
            public const int SHARE_DENY_READ = 0x30;  //Not allow to be read/open by other instance
            public const int SHARE_DENY_NONE = 0x40;  //Allow open by other instance and able to read/write
        }
        #endregion

        public static IDictionary<HWND, string> GetAllOpenedWindows()
        {
            Dictionary<HWND, string> windows;
            try
            {
                HWND shellWindow = GetShellWindow();
                windows = new Dictionary<HWND, string>();

                EnumWindows(delegate (HWND hWnd, int lParam)
                {
                    if (hWnd == shellWindow) return true;
                    if (!IsWindowVisible(hWnd)) return true;

                    int length = GetWindowTextLength(hWnd);
                    if (length == 0) return true;

                    StringBuilder builder = new StringBuilder(length);
                    GetWindowText(hWnd, builder, length + 1);

                    windows[hWnd] = builder.ToString();
                    return true;

                }, 0);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return windows;
        }

        public static IDictionary<HWND, string> GetAllOpenedWindowsByName(string strWindowTitle, bool blnIsExactWindowTitle = false)
        {
            IDictionary<HWND, string> dicAllWindows, dicSpecificWindows;
            try
            {
                dicAllWindows = new Dictionary<HWND, string>();
                dicSpecificWindows = new Dictionary<HWND, string>();

                dicAllWindows = GetAllOpenedWindows();

                if (dicAllWindows != null && dicAllWindows.Count > 0)
                {
                    var objSpecificWindow = dicAllWindows.Where(m => blnIsExactWindowTitle ? m.Value.Equals(strWindowTitle) : m.Value.ToUpper().Contains(strWindowTitle.ToUpper()));

                    if (objSpecificWindow != null && objSpecificWindow.Count() > 0)
                        dicSpecificWindows = objSpecificWindow.ToDictionary(n => n.Key, n => n.Value);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dicSpecificWindows;
        }

        public static IDictionary<HWND, WindowsInfo> GetAllOpenedWindowsInfo()
        {
            IDictionary<HWND, WindowsInfo> windows;
            try
            {
                HWND shellWindow = GetShellWindow();
                windows = new Dictionary<HWND, WindowsInfo>();

                EnumWindows(delegate (HWND hWnd, int lParam)
                {
                    if (hWnd == shellWindow) return true;
                    if (!IsWindowVisible(hWnd)) return true;

                    int length = GetWindowTextLength(hWnd);
                    //if (length == 0) return true;

                    StringBuilder sbWinText = new StringBuilder(length);
                    GetWindowText(hWnd, sbWinText, length + 1);

                    StringBuilder sbWinClass = new StringBuilder(1024);
                    GetClassName(hWnd, sbWinClass, 1024);

                    windows[hWnd] = new WindowsInfo()
                    {
                        WindowTitle = sbWinText.ToString(),
                        WindowClass = sbWinClass.ToString()
                    };

                    return true;

                }, 0);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return windows;
        }
        
        public static bool KillProcessByWindowHandled(HWND hwnd, out string strExceptionMessage)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                Process[] allProcesses = Process.GetProcesses();

                if (allProcesses != null && allProcesses.Count() > 0)
                {
                    Process[] specificProcess = allProcesses.Where(m => m.MainWindowHandle.Equals(hwnd)).ToArray();

                    if (specificProcess != null && specificProcess.Count() > 0)
                    {
                        foreach (var process in specificProcess)
                        {
                            if (process != null && !process.HasExited)
                                process.Kill();
                        }
                    }
                    else
                    {
                        throw new Exception("Unable to kill process as there is no process matches with the main window handle.");
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

        public static HWND GetObjectInWindow(HWND hwndWindow, string ControlName, string strWindowTitle = null)
        {
            HWND hwndWindowObject = HWND.Zero;
            try
            {
                hwndWindowObject = FindWindowEx(hwndWindow, HWND.Zero, ControlName, strWindowTitle ?? string.Empty);
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                throw new Exception(string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message));
            }

            return hwndWindowObject;
        }

        public static int ControlObjectInWindow(HWND hwndObject, uint ControlAction, string strParam = null)
        {
            int iControlResult = 0;

            try
            {
                iControlResult = SendMessage(hwndObject, ControlAction, HWND.Zero, strParam ?? string.Empty);
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                throw new Exception(string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message));
            }

            return iControlResult;
        }

        public static int BringWindowToFront(string strWindowTitle, bool blnIsExactWindowTitle = false)
        {
            int iResult = 0;
            try
            {
                IDictionary<HWND, string> dicWindow = GetAllOpenedWindowsByName(strWindowTitle, blnIsExactWindowTitle);
                if (dicWindow != null && dicWindow.Count > 0)
                    iResult = SetForegroundWindow(dicWindow.First().Key);
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                throw new Exception(string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message));
            }

            return iResult;
        }

        public static bool WaitForWindowToExists(string strWindowTitle, int iMaxSecondsToWait = 10, bool blnIsExactWindowTitle = false)
        {
            int iLoopCount = 1;
            bool blnIsSuccess = false;
            bool blnIsWindowExists = false;

            try
            {
                IDictionary<HWND, string> dicWindow = null;
                while (!blnIsWindowExists)
                {
                    if (iLoopCount > iMaxSecondsToWait)
                    {
                        blnIsSuccess = false;
                        break;
                    }

                    dicWindow = GetAllOpenedWindowsByName(strWindowTitle, blnIsExactWindowTitle);
                    if (dicWindow != null && dicWindow.Count > 0)
                    {
                        blnIsWindowExists = true;
                        blnIsSuccess = true;
                    }
                    else
                    {
                        System.Threading.Thread.Sleep(1000);
                    }

                    iLoopCount++;
                }
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                throw new Exception(string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message));
            }

            return blnIsSuccess;
        }

        public static bool WaitForWindowToExists(string strWindowTitle, string strWindowClassName, int iMaxSecondsToWait = 10)
        {
            int iLoopCount = 1;
            bool blnIsSuccess = false;
            bool blnIsWindowExists = false;

            try
            {
                IDictionary<HWND, WindowsInfo> dicAllWindows = null;
                while (!blnIsWindowExists)
                {
                    if (iLoopCount > iMaxSecondsToWait)
                    {
                        blnIsSuccess = false;
                        break;
                    }

                    dicAllWindows = GetAllOpenedWindowsInfo();
                    if (dicAllWindows != null && dicAllWindows.Count > 0)
                    {
                        var objSpecificWindow = dicAllWindows.Where(m => m.Value != null
                                                                         && m.Value.WindowClass.Equals(strWindowClassName)
                                                                         && m.Value.WindowTitle.Equals(strWindowTitle));

                        if (objSpecificWindow != null && objSpecificWindow.Count() > 0)
                        {
                            blnIsSuccess = true;
                            blnIsWindowExists = true;
                        }
                        else
                        {
                            System.Threading.Thread.Sleep(1000);
                        }
                    }
                    else
                    {
                        System.Threading.Thread.Sleep(1000);
                    }

                    iLoopCount++;
                }
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                throw new Exception(string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message));
            }

            return blnIsSuccess;
        }

        public static bool WaitForWindowToClose(string strWindowTitle, int iMaxSecondsToWait = 10, bool blnIsExactWindowTitle = false)
        {
            int iLoopCount = 1;
            bool blnIsSuccess = false;
            bool blnIsWindowClosed = false;

            try
            {
                IDictionary<HWND, string> dicWindow = null;
                while (!blnIsWindowClosed)
                {
                    if (iLoopCount > iMaxSecondsToWait)
                    {
                        blnIsSuccess = false;
                        break;
                    }

                    dicWindow = GetAllOpenedWindowsByName(strWindowTitle, blnIsExactWindowTitle);
                    if (dicWindow != null && dicWindow.Count > 0)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
                        blnIsWindowClosed = true;
                        blnIsSuccess = true;
                    }

                    iLoopCount++;
                }
            }
            catch (Exception ex)
            {
                string strCurNamespaceAndMethod = MethodInfo.GetCurrentMethod().DeclaringType.FullName + "." + MethodInfo.GetCurrentMethod().Name;
                throw new Exception(string.Format("Error source = {0}, Message = {1}", strCurNamespaceAndMethod, ex.Message));
            }

            return blnIsSuccess;
        }

        public static bool IsAppOrFileOpen(string strFilePath)
        {
            bool blnIsSuccess = false;

            try
            {
                if (string.IsNullOrWhiteSpace(strFilePath)) { return false; }

                string strFileDirPath = Path.GetDirectoryName(strFilePath);
                string strFileName = Path.GetFileNameWithoutExtension(strFilePath).ToLower();

                Process[] arrProcess = Process.GetProcessesByName(strFileName);
                if (arrProcess != null && arrProcess.Length > 0)
                {
                    foreach (Process proc in arrProcess)
                    {
                        if (proc.MainModule.FileName.StartsWith(strFileDirPath, StringComparison.InvariantCultureIgnoreCase))
                        {
                            blnIsSuccess = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
            }

            return blnIsSuccess;
        }

        public static bool IsAppOrFileMatchAccessMode(string strFilePath, int iAccessMode)
        {
            bool blnIsSuccess = false;

            try
            {
                IntPtr vHandle = _lopen(strFilePath, iAccessMode);
                if (vHandle == new HWND(-1))
                {
                    blnIsSuccess = true;
                }
                else
                {
                    CloseHandle(vHandle);
                    blnIsSuccess = false;
                }
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
            }

            return blnIsSuccess;
        }

        public static bool IsAppOrFileAllowReadWrite(string strFilePath)
        {
            return IsAppOrFileMatchAccessMode(strFilePath, AppFileAccessMode.READWRITE | AppFileSharedAccessMode.SHARE_EXCLUSIVE | AppFileSharedAccessMode.SHARE_DENY_NONE);
        }

        public static bool IsAppOrFileReadOnly(string strFilePath)
        {
            return IsAppOrFileMatchAccessMode(strFilePath, AppFileAccessMode.READ | AppFileSharedAccessMode.SHARE_DENY_WRITE);
        }

        #region Window API DLL import
        private delegate bool EnumWindowsProc(HWND hWnd, int lParam);

        [DllImport("KERNEL32.DLL")]
        private static extern HWND _lopen(string lpPathName, int iReadWrite);

        [DllImport("KERNEL32.DLL")]
        public static extern bool CloseHandle(HWND hObject);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(HWND hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(HWND hWnd);

        [DllImport("USER32.DLL")]
        public static extern int GetClassName(HWND hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern bool IsWindowVisible(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern HWND GetShellWindow();

        [DllImport("USER32.DLL", SetLastError = true)]
        private static extern HWND FindWindowEx(HWND parentHandle, HWND childAfter, string className, string windowTitle);

        [DllImport("USER32.DLL")]
        private static extern int SendMessage(HWND hWnd, uint wMsg, HWND wParam, string lParam);

        [DllImport("USER32.DLL")]
        private static extern int SetForegroundWindow(HWND hWnd);
        #endregion
    }

    public class WindowsInfo
    {
        public string WindowTitle { get; set; }
        public string WindowClass { get; set; }
    }
}
