using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.Common
{
    public class SFTPHelper
    {
        private string host { get; set; }
        private string username { get; set; }
        private string password { get; set; }
        private int port { get; set; }
        public SFTPHelper() { }
        public SFTPHelper(string host, int port, string username, string password)
        {
            this.host = host;
            this.port = port;
            this.username = username;
            this.password = password;
        }

        public bool Send(string ftpDirectory, string source, string destFileName, out string errMsg)
        {
            bool Success = true;
            string strFileName = string.Empty;
            string folderName = string.Empty;
            errMsg = string.Empty;

            try
            {
                #region Pre-checking
                if (!string.IsNullOrWhiteSpace(ftpDirectory))
                    ftpDirectory = ftpDirectory.Replace(@"\", "/").Replace(@"//", "/");
                else
                    throw new Exception("SFTP directory is null or empty");
                #endregion

                // Upload File
                using (var sftp = new SftpClient(this.host, this.port, this.username, this.password))
                {
                    sftp.Connect();

                    if (!sftp.Exists(ftpDirectory))
                        CreateAllDirectories(sftp, ftpDirectory);
                    //sftp.CreateDirectory(ftpDirectory);
                    sftp.ChangeDirectory(ftpDirectory);
                    using (var uplfileStream = System.IO.File.OpenRead(source))
                    {
                        uplfileStream.Position = 0;
                        sftp.BufferSize = 4 * 1024; // bypass Payload error large files
                        sftp.UploadFile(uplfileStream, destFileName, true);
                    }

                    sftp.Disconnect();
                }
            }
            catch (Exception ex)
            {
                Success = false;
                errMsg = string.Format("Error message: {0}. {1}", ex.Message,
                    !string.IsNullOrWhiteSpace(strFileName) ? $"Error transfer file name: {folderName}-{strFileName}" : string.Empty);
            }
            return Success;
        }

        public void CreateAllDirectories(SftpClient sftpClient, string strSftpDir)
        {
            try
            {
                // Consistent forward slashes
                strSftpDir = strSftpDir.Replace(@"\", "/").Replace(@"//", "/");
                foreach (string strDir in strSftpDir.Split('/'))
                {
                    // Ignoring leading/ending/multiple slashes
                    if (!string.IsNullOrWhiteSpace(strDir))
                    {
                        if (!sftpClient.Exists(strDir))
                            sftpClient.CreateDirectory(strDir);

                        sftpClient.ChangeDirectory(strDir);
                    }
                }
                // Going back to default directory
                sftpClient.ChangeDirectory("/");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
