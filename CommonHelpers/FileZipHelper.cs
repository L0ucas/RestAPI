using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using ICSharpCode.SharpZipLib.Zip;

namespace TMHelper.Common
{
    public static class FileZipHelper
    {
        public static bool ZipFolder(string strFolderPathToZip, string strOutputZipFilePath, out string strExceptionMessage, bool blnIsReplaceZipFile = true)
        {
            bool blnIsSuccess = true;
            string strOutputDirName = string.Empty;
            strExceptionMessage = string.Empty;

            try
            {
                #region Initial checking
                if (string.IsNullOrWhiteSpace(strFolderPathToZip) || string.IsNullOrWhiteSpace(strOutputZipFilePath))
                    return blnIsSuccess;
                if (!Path.GetExtension(strOutputZipFilePath).Equals(".zip", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("Invalid output zip path. Path not end with .zip");
                if (!Directory.Exists(strFolderPathToZip))
                    return blnIsSuccess;
                if (!blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    return blnIsSuccess;
                if (blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    File.Delete(strOutputZipFilePath);

                strOutputDirName = Path.GetDirectoryName(strOutputZipFilePath);
                if (!Directory.Exists(strOutputDirName))
                    Directory.CreateDirectory(strOutputDirName);
                #endregion

                #region Zip folder
                System.IO.Compression.ZipFile.CreateFromDirectory(strFolderPathToZip, strOutputZipFilePath, CompressionLevel.Optimal, true);
                #endregion
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }

            return blnIsSuccess;
        }

        public static bool UnZipFile(string strInputZipFilePath, string strOutputFolderPath, out string strExceptionMessage)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                #region Initial checking
                if (string.IsNullOrWhiteSpace(strInputZipFilePath) || string.IsNullOrWhiteSpace(strOutputFolderPath))
                    throw new Exception("Input Zip File Path or Output Folder Path is null or empty");
                if (!Path.GetExtension(strInputZipFilePath).Equals(".zip", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("Invalid input zip path. Path not end with .zip");
                if (!File.Exists(strInputZipFilePath))
                    throw new Exception("Input Zip File Path is not exists");
                if (!Directory.Exists(strOutputFolderPath))
                    Directory.CreateDirectory(strOutputFolderPath);
                #endregion

                #region UnZip file
                System.IO.Compression.ZipFile.ExtractToDirectory(strInputZipFilePath, strOutputFolderPath);
                #endregion
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }

            return blnIsSuccess;
        }

        public static bool ZipFiles(string[] arrFilePathToZip, string strOutputZipFilePath, out string strExceptionMessage, bool blnIsReplaceZipFile = true)
        {
            strExceptionMessage = string.Empty;
            if (arrFilePathToZip != null && arrFilePathToZip.Length > 0)
                return ZipFiles(arrFilePathToZip.ToList(), strOutputZipFilePath, out strExceptionMessage, blnIsReplaceZipFile);
            else
                return false;
        }

        public static bool ZipFiles(List<string> ListFilePathToZip, string strOutputZipFilePath, out string strExceptionMessage, bool blnIsReplaceZipFile = true)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                #region Initial checking
                if (ListFilePathToZip == null || (ListFilePathToZip != null && ListFilePathToZip.Count == 0))
                    return blnIsSuccess;
                if (string.IsNullOrWhiteSpace(strOutputZipFilePath))
                    return blnIsSuccess;
                if (!Path.GetExtension(strOutputZipFilePath).Equals(".zip", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("Invalid output zip path. Path not end with .zip");
                if (!blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    return blnIsSuccess;
                if (blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    File.Delete(strOutputZipFilePath);
                #endregion

                #region Zip files
                foreach (string strFilePath in ListFilePathToZip)
                {
                    if (!string.IsNullOrWhiteSpace(strFilePath) && File.Exists(strFilePath))
                    {
                        using (ZipArchive archive = System.IO.Compression.ZipFile.Open(strOutputZipFilePath, ZipArchiveMode.Update))
                        {
                            archive.CreateEntryFromFile(strFilePath, Path.GetFileName(strFilePath), CompressionLevel.Optimal);
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

            return blnIsSuccess;
        }

        //Zip using IO compression - select compression level (Optimal, Fastest, NoCompression), Optimal being the highest compression
        public static bool ZipFilesWithCompressLevel(List<string> ListFilePathToZip, string strOutputZipFilePath, CompressionLevel compressLevel, out string strExceptionMessage, bool blnIsReplaceZipFile = true)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                #region Initial checking
                if (ListFilePathToZip == null || (ListFilePathToZip != null && ListFilePathToZip.Count == 0))
                    return blnIsSuccess;
                if (string.IsNullOrWhiteSpace(strOutputZipFilePath))
                    return blnIsSuccess;
                if (!Path.GetExtension(strOutputZipFilePath).Equals(".zip", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("Invalid output zip path. Path not end with .zip");
                if (!blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    return blnIsSuccess;
                if (blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    File.Delete(strOutputZipFilePath);
                #endregion

                #region Zip files
                foreach (string strFilePath in ListFilePathToZip)
                {
                    if (!string.IsNullOrWhiteSpace(strFilePath) && File.Exists(strFilePath))
                    {
                        using (ZipArchive archive = System.IO.Compression.ZipFile.Open(strOutputZipFilePath, ZipArchiveMode.Update))
                        {
                            archive.CreateEntryFromFile(strFilePath, Path.GetFileName(strFilePath), compressLevel);
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

            return blnIsSuccess;
        }

        //Zip using SharpZipLib - select compress level from 0-9, 9 being the highest compression
        public static bool ZipFilesWithCompressLevel(List<string> ListFilePathToZip, string strOutputZipFilePath, int iCompressLevel, out string strExceptionMessage, bool blnIsReplaceZipFile = true)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                #region Initial checking
                if (ListFilePathToZip == null || (ListFilePathToZip != null && ListFilePathToZip.Count == 0))
                    return blnIsSuccess;
                if (string.IsNullOrWhiteSpace(strOutputZipFilePath))
                    return blnIsSuccess;
                if (!Path.GetExtension(strOutputZipFilePath).Equals(".zip", StringComparison.OrdinalIgnoreCase))
                    throw new Exception("Invalid output zip path. Path not end with .zip");
                if (!blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    return blnIsSuccess;
                if (blnIsReplaceZipFile && File.Exists(strOutputZipFilePath))
                    File.Delete(strOutputZipFilePath);
                #endregion

                #region Zip files
                using (ZipOutputStream zOS = new ZipOutputStream(File.Create(strOutputZipFilePath)))
                {
                    zOS.SetLevel(iCompressLevel); // 0-9, 9 being the highest compression

                    byte[] buffer = new byte[4096];

                    foreach (string strFilePath in ListFilePathToZip)
                    {
                        ZipEntry zipEntry = new ZipEntry(Path.GetFileName(strFilePath));
                        zipEntry.DateTime = DateTime.Now;

                        zOS.PutNextEntry(zipEntry);

                        using (FileStream fs = File.OpenRead(strFilePath))
                        {
                            int sourceBytes;

                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                zOS.Write(buffer, 0, sourceBytes);
                            }
                            while (sourceBytes > 0);
                        }
                    }

                    zOS.Finish();
                    zOS.Close();
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
    }
}
