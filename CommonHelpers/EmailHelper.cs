using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using Microsoft.Exchange.WebServices.Data;
using Domino;
using System.IO;
using System.Data;
using System.Collections;
using System.Net;
using Microsoft.Graph;
using Microsoft.Identity;
using System.Net.Http.Headers;
using Microsoft.Identity.Client;

namespace TMHelper.Common
{
    public class EmailHelper
    {
        #region Constant
        private const string CONST_ASPIRO_EMAIL_HOST_IP_ADDRESS = "relay.aspiro.co"; //"10.89.4.100";
        private const int CONST_ASPIRO_EMAIL_SMTP_PORT = 25;
        private const string CONST_ASPIRO_NETWORK_USERNAME = "globalnet\aspiro_rpa";
        private const string CONST_ASPIRO_NETWORK_PASSWORD_MD5_ENCRYPTED = "075D3699C2E745157E1B8AB32C4A3845";
        // Vendor Test Environment
        //private const string CONST_RPA_APP_TENANT_ID = "6a59cc44-ce76-405e-b2f5-23c1dc274661";
        //private const string CONST_RPA_APP_CLIENT_ID = "5f6b1e6e-fd62-436a-b158-bbf07829e940";
        //private const string CONST_RPA_APP_CLIENT_SECRET = "uZg8Q~5fJzlnlzJUFGAiTBzL.lpYM5J0l5Q_qbwm";
        // Aspiro Prod Environment
        private static readonly string CONST_RPA_APP_TENANT_ID = CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_RPA_MS_GRAPH_API_TENANT_ID);
        private static readonly string CONST_RPA_APP_CLIENT_ID = CryptographyHelper.SHA512Decrypt(CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_RPA_MS_GRAPH_API_CLIENT_ID));
        private static readonly string CONST_RPA_APP_CLIENT_SECRET = CryptographyHelper.SHA512Decrypt(CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_RPA_MS_GRAPH_API_CLIENT_SECRET));

        private static readonly string CONST_RPA_APP_AUTORITY_URL = "https://login.microsoftonline.com/" + CONST_RPA_APP_TENANT_ID;
        private static readonly string CONST_RPA_APP_REDIRECT_AUTORITY_URL = "https://login.microsoftonline.com";
        private static readonly string CONST_MS_GRAPH_API_URL = "https://graph.microsoft.com/v1.0";

        // The Group.Read.All permission is an admin-only scope, so authorization will fail if you 
        // want to sign in with a non-admin account. Remove that permission and comment out the group operations in 
        // the UserMode() method if you want to run this sample with a non-admin account.
        public static string[] CONS_RPA_APP_SCOPES = {
                                           "Mail.Send",
                                           "Mail.ReadWrite",
                                           "Mail.ReadWrite.Shared",
                                           "Files.ReadWrite",
                                            // Group.Read.All is an admin-only scope. It allows you to read Group details.
                                            // Uncomment this scope if you want to run the application with an admin account
                                            // and perform the group operations in the UserMode class.
                                            // You'll also need to uncomment the UserMode.UserModeRequests.GetDetailsForGroups() method.
                                            //"Group.Read.All" 
                                        };
        #endregion

        #region Properties
        private string EmailHostNameOrIP { get; set; }
        private int EmailSMTPPort { get; set; }
        private string NetworkCredentialUserName { get; set; }
        private string NetworkCredentialPassword { get; set; }

        private static ConfidentialClientApplication IdentityAppOnlyApp = null;

        //private string RPAMailUserId = CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_EXCHANGE_USERNAME);
        //private string RPAFromEmailAddr = "svaspiro3@nomlab.online";
        private static readonly string RPAFromEmailAddr = CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_RPA_MS_GRAPH_API_USERNAME);

        #endregion

        #region Constructor
        public EmailHelper(string strEmailHostNameOrIP, int strEmailSMTPPort, string strNetworkUID, string strNetworkPassword)
        {
            EmailHostNameOrIP = strEmailHostNameOrIP;
            EmailSMTPPort = strEmailSMTPPort;
            NetworkCredentialUserName = strNetworkUID;
            NetworkCredentialPassword = CryptographyHelper.MD5Decrypt(strNetworkPassword);
        }

        public EmailHelper(string strEmailHostNameOrIP, int strEmailSMTPPort)
        {
            EmailHostNameOrIP = strEmailHostNameOrIP;
            EmailSMTPPort = strEmailSMTPPort;
            NetworkCredentialUserName = CONST_ASPIRO_NETWORK_USERNAME;
            NetworkCredentialPassword = CryptographyHelper.MD5Decrypt(CONST_ASPIRO_NETWORK_PASSWORD_MD5_ENCRYPTED);
        }

        public EmailHelper()
        {
            EmailHostNameOrIP = CONST_ASPIRO_EMAIL_HOST_IP_ADDRESS;
            EmailSMTPPort = CONST_ASPIRO_EMAIL_SMTP_PORT;
            NetworkCredentialUserName = CONST_ASPIRO_NETWORK_USERNAME;
            NetworkCredentialPassword = CryptographyHelper.MD5Decrypt(CONST_ASPIRO_NETWORK_PASSWORD_MD5_ENCRYPTED);
        }
        #endregion

        private ExchangeService InitMsExchange()
        {
            ExchangeService exchangeService = null;
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            string strMailUserName = CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_EXCHANGE_USERNAME);
            string strMailPassword = CryptographyHelper.Decrypt(CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_EXCHANGE_PWD));
            string strExchangeServiceUrl = CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_EXCHANGE_SVC_URL);
            string strFromEmail = CommonHelpers.GetRPAConfig(GlobalConstant.RPAMasterConfig.CONST_CONFIG_KEY_EXCHANGE_FROM_EMAIL_ADDR);

            exchangeService = new ExchangeService();
            exchangeService.Credentials = new WebCredentials(strMailUserName, strMailPassword);

            if (!string.IsNullOrWhiteSpace(strExchangeServiceUrl))
                exchangeService.Url = new Uri(strExchangeServiceUrl);
            else
                exchangeService.AutodiscoverUrl(strFromEmail, ExchangeServiceRedirectCallback);

            return exchangeService;
        }

        private ExchangeService InitMsExchange(string strMailUserName, string strMailPassword, string strExchangeServiceUrl, string strFromEmail)
        {
            ExchangeService exchangeService = null;
            exchangeService = new ExchangeService();
            exchangeService.Credentials = new WebCredentials(strMailUserName, strMailPassword);

            if (!string.IsNullOrWhiteSpace(strExchangeServiceUrl))
                exchangeService.Url = new Uri(strExchangeServiceUrl);
            else
                exchangeService.AutodiscoverUrl(strFromEmail, ExchangeServiceRedirectCallback);

            return exchangeService;
        }

        private GraphServiceClient InitGraphServiceForApp()
        {
            //To include TLS 1.1 & 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            // Create Microsoft Graph client.
            IdentityAppOnlyApp = new ConfidentialClientApplication(CONST_RPA_APP_CLIENT_ID, CONST_RPA_APP_AUTORITY_URL, CONST_RPA_APP_REDIRECT_AUTORITY_URL, new ClientCredential(CONST_RPA_APP_CLIENT_SECRET), new TokenCache(), new TokenCache());
            GraphServiceClient graphClient = null;
            graphClient = new GraphServiceClient(
                CONST_MS_GRAPH_API_URL,
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var token = await GetTokenForAppAsync();
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        // This header has been added to identify our sample in the Microsoft Graph service.  If extracting this code for your project please remove.
                        requestMessage.Headers.Add("Prefer", "outlook.body-content-type='text'");

                    }));
            return graphClient;
        }

        private GraphServiceClient InitGraphServiceForApp(string clientIdForApp, string authorityUri, string redirectAuthorityUri, string clientSecret)
        {
            // Create Microsoft Graph client.
            IdentityAppOnlyApp = new ConfidentialClientApplication(clientIdForApp, authorityUri, redirectAuthorityUri, new ClientCredential(clientSecret), new TokenCache(), new TokenCache());
            GraphServiceClient graphClient = null;
            graphClient = new GraphServiceClient(
                CONST_MS_GRAPH_API_URL,
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var token = await GetTokenForAppAsync();
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        // This header has been added to identify our sample in the Microsoft Graph service.  If extracting this code for your project please remove.
                        requestMessage.Headers.Add("Prefer", "outlook.body-content-type='text'");

                    }));
            return graphClient;
        }
        public static async Task<string> GetTokenForAppAsync()
        {
            AuthenticationResult authResult;
            authResult = await IdentityAppOnlyApp.AcquireTokenForClientAsync(new string[] { "https://graph.microsoft.com/.default" });
            return authResult.AccessToken;
        }

        public bool SendMailViaSMTP(List<string> ToEmailAddrList, string FromEmailAddr, string strMailSubject, string strMailBody, out string strExceptionMessage, List<System.Net.Mail.Attachment> AttachmentList = null, List<string> CCEmailAddrList = null, List<string> BCCEmailAddrList = null, bool blnIsHtmlEmailBody = true)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                using (var mailMsg = new MailMessage())
                {
                    #region Construct basic mail details

                    if (ToEmailAddrList != null && ToEmailAddrList.Count > 0)
                    {
                        foreach (string strToEmail in ToEmailAddrList)
                            mailMsg.To.Add(strToEmail);
                    }

                    if (CCEmailAddrList != null && CCEmailAddrList.Count > 0)
                    {
                        foreach (string strCCEmail in CCEmailAddrList)
                            mailMsg.CC.Add(strCCEmail);
                    }

                    if (BCCEmailAddrList != null && BCCEmailAddrList.Count > 0)
                    {
                        foreach (string strBCCEmail in BCCEmailAddrList)
                            mailMsg.Bcc.Add(strBCCEmail);
                    }

                    if (AttachmentList != null && AttachmentList.Count > 0)
                    {
                        foreach (var attachment in AttachmentList)
                            mailMsg.Attachments.Add(attachment);
                    }

                    mailMsg.From = new MailAddress(FromEmailAddr);
                    mailMsg.Subject = strMailSubject;
                    mailMsg.Body = strMailBody;
                    mailMsg.IsBodyHtml = blnIsHtmlEmailBody;
                    #endregion

                    #region Construct smtp details and send email
                    SmtpClient smtpClient = new SmtpClient(EmailHostNameOrIP);
                    smtpClient.Port = EmailSMTPPort;
                    smtpClient.Credentials = new System.Net.NetworkCredential(NetworkCredentialUserName, NetworkCredentialPassword);

                    smtpClient.Send(mailMsg);
                    mailMsg.Dispose();
                    blnIsSuccess = true;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }

            return blnIsSuccess;
        }

        public bool SendMailViaMsGraph(string strToEmail, string strCCEmail,
            string strBCCEmail, string strMailSubject, string strMailBody, List<string> AttachmentList,
            out string strExceptionMessage, string strMailBox = null, bool blnIsHtmlEmailBody = true, bool blSaveSentItem = true)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                GraphServiceClient graphClient = InitGraphServiceForApp();
                string userName = string.IsNullOrWhiteSpace(strMailBox) ? RPAFromEmailAddr : strMailBox;

                Recipient fromRecipient = new Recipient();
                ItemBody mailBody = new ItemBody();
                List<Recipient> toRecipient = new List<Recipient>();
                List<Recipient> ccRecipient = new List<Recipient>();
                List<Recipient> bccRecipient = new List<Recipient>();
                MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();

                mailBody = new ItemBody
                {
                    ContentType = (blnIsHtmlEmailBody) ? Microsoft.Graph.BodyType.Html : Microsoft.Graph.BodyType.Text,
                    Content = strMailBody
                };

                if (!string.IsNullOrWhiteSpace(strToEmail))
                {
                    string[] arrToEmail = strToEmail.Replace(';', ',').Split(',');
                    foreach (string strEmail in arrToEmail)
                    {
                        toRecipient.Add(new Recipient
                        {
                            EmailAddress = new Microsoft.Graph.EmailAddress
                            {
                                Address = strEmail
                            }
                        });
                    }
                }

                if (!string.IsNullOrWhiteSpace(strCCEmail))
                {
                    string[] arrCCEmail = strCCEmail.Replace(';', ',').Split(',');
                    foreach (string strEmail in arrCCEmail)
                    {
                        ccRecipient.Add(new Recipient
                        {
                            EmailAddress = new Microsoft.Graph.EmailAddress
                            {
                                Address = strEmail
                            }
                        });
                    }
                }

                if (!string.IsNullOrWhiteSpace(strBCCEmail))
                {
                    string[] arrBCCEmail = strBCCEmail.Replace(';', ',').Split(',');
                    foreach (string strEmail in arrBCCEmail)
                    {
                        bccRecipient.Add(new Recipient
                        {
                            EmailAddress = new Microsoft.Graph.EmailAddress
                            {
                                Address = strEmail
                            }
                        });
                    }
                }

                if (AttachmentList != null && AttachmentList.Count > 0)
                {
                    foreach (string attFile in AttachmentList)
                    {
                        string fileName = Path.GetFileName(attFile);
                        string contentType = System.Web.MimeMapping.GetMimeMapping(fileName);
                        byte[] fileBytes = System.IO.File.ReadAllBytes(attFile);

                        attachments.Add(new Microsoft.Graph.FileAttachment
                        {
                            Name = fileName,
                            ContentType = contentType,
                            ContentBytes = fileBytes
                        });
                    }
                }

                var message = new Message
                {
                    Subject = strMailSubject.Trim(),
                    Body = mailBody,
                    ToRecipients = toRecipient,
                    CcRecipients = ccRecipient,
                    BccRecipients = bccRecipient,
                    Attachments = attachments
                };

                graphClient.Users[userName]
                      .SendMail(message, blSaveSentItem)
                      .Request()
                      .PostAsync().Wait();
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
            }

            return blnIsSuccess;
        }

        private bool ExchangeServiceRedirectCallback(string strUrl)
        {
            return !string.IsNullOrWhiteSpace(strUrl) ? strUrl.ToLower().StartsWith("https://") : false;
        }

        public IList<Mailing> GetMailBySubject(string subject)
        {
            try
            {
                IList<Mailing> emails = GetMailBySubjectV2(subject);
                return emails;
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to get email. {e.Message}", e);
            }
        }

        public IList<Mailing> GetMailBySubject(string strSubject, string strSender, DateTime dtDateSent, string downloadPath, bool saveAtt)
        {
            try
            {
                IList<Mailing> emails = GetMailBySubjectV2(strSubject, strSender, dtDateSent, downloadPath, saveAtt);
                return emails;
            }
            catch (Exception e)
            {
                throw new Exception("Failed to get email.", e);
            }

        }

        public bool SendMail(string mailTo, string ccTo, string bccTo, string mailSubject, string htmlBody, string att)
        {
            bool blSuccess = false;
            try
            {
                EmailHelper emailHelper = new EmailHelper();
                blSuccess = emailHelper.SendMailV2(mailTo, ccTo, bccTo, mailSubject, htmlBody, att);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return blSuccess;
        }

        public IList<Mailing> GetMailBySubjectV2(string strSubject)
        {
            IList<Mailing> emails = new List<Mailing>();
            string strErr = string.Empty;
            try
            {
                emails = GetMailBySubjectV2(strSubject, string.Empty, new DateTime(), string.Empty, false);
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to get email by subject. {e.Message}", e);
            }
            return emails;
        }

        public IList<Mailing> GetMailBySubjectV2(string strSubject, string strSender, DateTime dtDateSent, string downloadPath, bool saveAtt, bool blnIsSearchExactSubject, string strMailBox = null,
            DateTime DateCreatedFrom = new DateTime(), DateTime DateCreatedTo = new DateTime())
        {
            IList<Mailing> emails = new List<Mailing>();
            string strErr = string.Empty;
            string userName = string.IsNullOrWhiteSpace(strMailBox) ? RPAFromEmailAddr : strMailBox;
            List<QueryOption> options = new List<QueryOption>();
            List<string> searchFilterCollection = new List<string>();

            try
            {

                if (!string.IsNullOrWhiteSpace(strSubject))
                {
                    if (blnIsSearchExactSubject)
                        searchFilterCollection.Add(string.Format("Subject eq '{0}'", strSubject));
                    else
                        searchFilterCollection.Add(string.Format("contains(Subject,'{0}')", strSubject));
                }
                if (!string.IsNullOrWhiteSpace(strSender))
                {
                    searchFilterCollection.Add(string.Format("(contains(sender/emailAddress/address,'{0}') and startsWith(sender/emailAddress/address,'{0}'))", strSender));
                }

                if (dtDateSent != new DateTime())
                {
                    DateTime dtFrom = dtDateSent;
                    DateTime dtTo = dtDateSent.AddDays(1);

                    searchFilterCollection.Add(string.Format("CreatedDateTime ge {0}", new DateTimeOffset(DateTime.SpecifyKind(dtFrom, DateTimeKind.Local)).UtcDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")));
                    searchFilterCollection.Add(string.Format("CreatedDateTime lt {0}", new DateTimeOffset(DateTime.SpecifyKind(dtTo, DateTimeKind.Local)).UtcDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")));
                }

                if (DateCreatedFrom != new DateTime() && DateCreatedTo != new DateTime())
                {
                    DateTime dtFrom = DateCreatedFrom;
                    DateTime dtTo = DateCreatedTo.AddDays(1);
                    searchFilterCollection.Add(string.Format("CreatedDateTime ge {0}", new DateTimeOffset(DateTime.SpecifyKind(dtFrom, DateTimeKind.Local)).UtcDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")));
                    searchFilterCollection.Add(string.Format("CreatedDateTime lt {0}", new DateTimeOffset(DateTime.SpecifyKind(dtTo, DateTimeKind.Local)).UtcDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")));
                }

                options = createQueryOption(searchFilterCollection);

                emails = GetMailsViaGraphApi(saveAtt, userName, options, downloadPath);
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to get email by subject. {e.Message}", e);
            }
            return emails;
        }
        public IList<Mailing> GetMailBySubjectV2(string strSubject, string strSender, DateTime? dtDateSentFrom, DateTime? dtDateSentTo, string strDownloadFolderPath, bool blnIsSaveAtt, bool blnIsSearchExactSubject, string strMailBox = null)
        {
            IList<Mailing> emails = new List<Mailing>();
            string strErr = string.Empty;
            string userName = string.IsNullOrWhiteSpace(strMailBox) ? RPAFromEmailAddr : strMailBox;
            List<QueryOption> options = new List<QueryOption>();
            List<string> searchFilterCollection = new List<string>();

            try
            {
                #region Subject search filter
                if (!string.IsNullOrWhiteSpace(strSubject))
                {
                    if (blnIsSearchExactSubject)
                        searchFilterCollection.Add(string.Format("Subject eq '{0}'", strSubject));
                    else
                        searchFilterCollection.Add(string.Format("contains(Subject,'{0}')", strSubject));
                }
                #endregion

                #region Sender search filer
                if (!string.IsNullOrWhiteSpace(strSender))
                {
                    searchFilterCollection.Add(string.Format("(contains(sender/emailAddress/address,'{0}') and startsWith(sender/emailAddress/address,'{0}'))", strSender));
                }
                #endregion

                #region DateTime Created search filter
                if (dtDateSentFrom != null && dtDateSentFrom.HasValue)
                    searchFilterCollection.Add(string.Format("CreatedDateTime ge {0}", new DateTimeOffset(DateTime.SpecifyKind(dtDateSentFrom.Value, DateTimeKind.Local)).UtcDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")));
                if (dtDateSentTo != null && dtDateSentTo.HasValue)
                    searchFilterCollection.Add(string.Format("CreatedDateTime lt {0}", new DateTimeOffset(DateTime.SpecifyKind(dtDateSentTo.Value, DateTimeKind.Local)).UtcDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")));
                #endregion

                options = createQueryOption(searchFilterCollection);

                emails = GetMailsViaGraphApi(blnIsSaveAtt, userName, options, strDownloadFolderPath);
            }
            catch (Exception e)
            {
                throw new Exception($"Failed to get email by subject. {e.Message}", e);
            }

            return emails;
        }


        private List<QueryOption> createQueryOption(List<string> filterQueries)
        {
            List<QueryOption> output = new List<QueryOption>();

            #region Filter Query
            if (filterQueries != null & filterQueries.Count > 0)
            {
                string filterString = string.Empty;
                foreach (string filter in filterQueries)
                {
                    if (string.IsNullOrWhiteSpace(filterString))
                    {
                        filterString = string.Format(filter);
                    }
                    else
                    {
                        filterString = string.Format("{0} AND {1}", filterString, filter);
                    }
                }
                output.Add(new QueryOption("$filter", filterString));
            }
            #endregion

            return output;
        }

        private IList<Mailing> GetMailsViaGraphApi(bool saveAtt, string userName, List<QueryOption> options, string downloadPath, string fromFolder = null, string toFolder = null, bool moveToFolder = false)
        {
            GraphServiceClient graphClient = InitGraphServiceForApp();
            List<Message> messages = new List<Message>();
            List<Mailing> emails = new List<Mailing>();

            string mailFolder = string.IsNullOrWhiteSpace(fromFolder) ? "inbox" : fromFolder;
            string mailFromFolderId = GetFolderId(userName, mailFolder);
            string mailToFolderId = string.Empty;

            if (!string.IsNullOrWhiteSpace(toFolder))
                mailToFolderId = GetFolderId(userName, toFolder);

            if (saveAtt)
            {
                if (string.IsNullOrWhiteSpace(downloadPath))
                    throw new Exception("Attachment download path is empty");
                if (!System.IO.Directory.Exists(downloadPath))
                    System.IO.Directory.CreateDirectory(downloadPath);
            }

            if (saveAtt)
            {
                messages = graphClient.Users[userName].MailFolders[mailFromFolderId].Messages.Request(options).Expand("Attachments")
                    .Select("subject, sender, Body, CreatedDateTime, ReceivedDateTime, HasAttachments, Attachments")
                    .Top(1000)
                    .GetAsync().Result.ToList()
                    .OrderByDescending(e => e.ReceivedDateTime)
                    .ToList();
            }
            else
            {
                messages = graphClient.Users[userName].MailFolders[mailFromFolderId].Messages.Request(options)
                    .Select("subject, sender, Body, CreatedDateTime, ReceivedDateTime, HasAttachments, Attachments")
                    .Top(1000)
                    .GetAsync().Result.ToList()
                    .OrderByDescending(e => e.ReceivedDateTime)
                    .ToList();

            }

            if (messages != null && messages.Count > 0)
            {
                foreach (Message message in messages)
                {
                    Mailing email = new Mailing();
                    email.MailSubject = message.Subject.ToString();
                    email.From = message.Sender.EmailAddress.Address;
                    email.MailBody = message.Body.Content;
                    email.Date = message.CreatedDateTime.Value.LocalDateTime.ToString("yyyyMMdd HH:mm:ss");
                    email.DateCreated = message.CreatedDateTime.Value.LocalDateTime;
                    email.DateReceived = message.ReceivedDateTime.Value.LocalDateTime;
                    email.Id = message.Id;
                    email.FolderId = mailFromFolderId;

                    if (saveAtt && message.HasAttachments != null && message.HasAttachments.HasValue && message.HasAttachments.Value && message.Attachments != null && message.Attachments.Count > 0 && !string.IsNullOrWhiteSpace(downloadPath))
                    {
                        email.Attachments = new List<MailingAttachment>();

                        foreach (var graphAtt in message.Attachments)
                        {
                            /* 
                             * There are three types of graph attachment 
                             * 1) File attachment
                             * 2) Item attachment (contact, event or message)
                             * 3) Reference attachment (a link to a file)
                             */

                            //Only file attachment need to download
                            if (graphAtt.ODataType == "#microsoft.graph.fileAttachment") 
                            {
                                Microsoft.Graph.FileAttachment att = (Microsoft.Graph.FileAttachment)graphAtt;

                                MailingAttachment mailingAttachment = new MailingAttachment();
                                string strFileName = CommonHelpers.GetNextFileName(Path.Combine(downloadPath, att.Name));
                                System.IO.File.WriteAllBytes(strFileName, att.ContentBytes); // Requires System.IO
                                mailingAttachment.FileName = strFileName.Replace(string.Format("{0}\\", Path.GetDirectoryName(strFileName)), "");
                                mailingAttachment.FilePath = string.Format("{0}\\", Path.GetDirectoryName(strFileName));
                                mailingAttachment.isSaved = true;
                                mailingAttachment.ContentType = att.ContentType;
                                email.Attachments.Add(mailingAttachment);
                            }
                        }
                    }

                    if (moveToFolder && !string.IsNullOrWhiteSpace(mailToFolderId))
                    {
                        // Move message to folder
                        Task<Message> tMsg = System.Threading.Tasks.Task.Run(async() => await graphClient.Users[userName].MailFolders[mailFromFolderId].Messages[message.Id].Move(mailToFolderId).Request().PostAsync());
                        graphClient.Users[userName].MailFolders[mailFromFolderId].Messages[message.Id].Request().UpdateAsync(tMsg.GetAwaiter().GetResult());
                    }
                    emails.Add(email);
                }
            }

            return emails;
        }
        public IList<Mailing> GetMailBySubjectV2(string strSubject, string strSender, DateTime dtDateSent, string downloadPath, bool saveAtt, string strMailBox = null)
        {
            return GetMailBySubjectV2(strSubject, strSender, dtDateSent, downloadPath, saveAtt, false, strMailBox);
        }

        public IList<Mailing> GetAllMails(string strMailBox, string downloadPath = null, string FolderNm = null)
        {
            IList<Mailing> emails = new List<Mailing>();
            string strErr = string.Empty;
            string userName = string.IsNullOrWhiteSpace(strMailBox) ? RPAFromEmailAddr : strMailBox;
            List<QueryOption> options = new List<QueryOption>();
            List<string> searchFilterCollection = new List<string>();

            try
            {
                emails = GetMailsViaGraphApi(true, userName, options, downloadPath, string.Empty, FolderNm, false);
            }
            catch (Exception e)
            {
                throw new Exception("Failed to get email by subject. ", e);
            }
            return emails;
        }
        public string GetFolderId(string strMailBox, string FolderNm)
        {
            GraphServiceClient graphClient = InitGraphServiceForApp();

            string mailFolderId = string.Empty;
            Task<IUserMailFoldersCollectionPage> tUMailFolderCP = System.Threading.Tasks.Task.Run(async () => await graphClient.Users[strMailBox].MailFolders.Request().Top(100).GetAsync());
            var folderList = tUMailFolderCP.Result.ToList();

            foreach (Microsoft.Graph.MailFolder item in folderList)
            {
                if (item.DisplayName.Equals(FolderNm, StringComparison.InvariantCultureIgnoreCase))
                    mailFolderId = item.Id;
            }

            return mailFolderId;
        }
        public string GetFolderId(string strFolderNm)
        {
            string output = string.Empty;
            string userName = RPAFromEmailAddr;
            try
            {
                output = GetFolderId(userName, strFolderNm);
            }
            catch (Exception ex)
            {
                output = string.Empty;
            }

            return output;
        }

        public IList<Mailing> GetAllFolder(string strMailBox)
        {
            GraphServiceClient graphClient = InitGraphServiceForApp();
            List<Mailing> FolderList = new List<Mailing>();

            string mailFolderId = string.Empty;
            var folderList = graphClient.Users[strMailBox].MailFolders.Request().Top(100).GetAsync().Result.ToList();

            foreach (Microsoft.Graph.MailFolder item in folderList)
            {
                Mailing email = new Mailing();
                email.FolderId = item.Id.ToString();
                email.DisplayName = item.DisplayName;
                FolderList.Add(email);
            }

            return FolderList;
        }

        public bool SendMailV2(string mailTo, string ccTo, string bccTo, string mailSubject, string htmlBody, string att, string strMailBox = null)
        {
            bool blSuccess = false;
            string strErr = string.Empty;
            List<string> lstFileName = null;

            try
            {
                if (!string.IsNullOrWhiteSpace(att))
                {
                    lstFileName = att.Split(',').ToList();
                }
                blSuccess = SendMailViaMsGraph(mailTo, ccTo, bccTo,
                    mailSubject, htmlBody, lstFileName, out strErr, strMailBox);

                if (!blSuccess)
                    throw new Exception($"Failed to send email. {strErr}");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return blSuccess;
        }

        public IList<Mailing> GetMailBoxEmailList(string strSender, string strSubject, bool blnIsSaveAtt, string strAttSaveFolderPath, string strMailBoxName = null, bool blnIsSearchExactSubject = true)
        {
            IList<Mailing> mailingList = null;
            string strExceptionMsg = string.Empty;
            string userName = string.IsNullOrWhiteSpace(strMailBoxName) ? RPAFromEmailAddr : strMailBoxName;
            List<QueryOption> options = new List<QueryOption>();
            List<string> searchFilterCollection = new List<string>();

            try
            {
                string[] arrSender = strSender.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (arrSender != null && arrSender.Length > 0)
                {
                    string emailFromList = string.Empty;
                    foreach (string strCurSender in arrSender)
                    {
                        if (!string.IsNullOrWhiteSpace(emailFromList))
                        {
                            emailFromList = string.Format("{0} or ", emailFromList);
                        }
                        emailFromList = string.Format("{0}(contains(sender/emailAddress/address,'{1}') and startsWith(sender/emailAddress/address,'{1}'))", emailFromList, strCurSender);
                    }

                    if (!string.IsNullOrWhiteSpace(emailFromList))
                        searchFilterCollection.Add(string.Format("({0})", emailFromList));
                }

                if (!string.IsNullOrWhiteSpace(strSubject))
                {
                    if (blnIsSearchExactSubject)
                        searchFilterCollection.Add(string.Format("Subject eq '{0}'", strSubject));
                    else
                        searchFilterCollection.Add(string.Format("contains(Subject,'{0}')", strSubject));
                }

                options = createQueryOption(searchFilterCollection);

                mailingList = GetMailsViaGraphApi(blnIsSaveAtt, userName, options, strAttSaveFolderPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mailingList;
        }

        /*
        public bool ReplyEmail(Mailing targetEmail, string strReplyMailBody, out string strExceptionMessage, bool blnIsReplyAll = true, List<string> AddEmailToList = null, List<string> AddEmailCCList = null, List<string> AddEmailBCCList = null, bool blnIsSaveCopy = true)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;

            try
            {
                if (targetEmail == null)
                {
                    throw new Exception("Target email object is null");
                }
                else if (targetEmail.MsExchangeEmailMessage == null)
                {
                    throw new Exception("Microsoft Exchange email message object is null");
                }
                else
                {
                    ResponseMessage resMsg = targetEmail.MsExchangeEmailMessage.CreateReply(blnIsReplyAll);
                    resMsg.BodyPrefix = strReplyMailBody;

                    #region Add additional email To recipient
                    if (AddEmailToList != null && AddEmailToList.Count > 0)
                    {
                        foreach (string emailTo in AddEmailToList)
                        {
                            Microsoft.Exchange.WebServices.Data.EmailAddress emailAddr = new Microsoft.Exchange.WebServices.Data.EmailAddress();
                            emailAddr.Address = emailTo;

                            resMsg.ToRecipients.Add(emailAddr);
                        }
                    }
                    #endregion

                    #region Add additional email CC recipient
                    if (AddEmailCCList != null && AddEmailCCList.Count > 0)
                    {
                        foreach (string emailCC in AddEmailCCList)
                        {
                            Microsoft.Exchange.WebServices.Data.EmailAddress emailAddr = new Microsoft.Exchange.WebServices.Data.EmailAddress();
                            emailAddr.Address = emailCC;

                            resMsg.ToRecipients.Add(emailAddr);
                        }
                    }
                    #endregion

                    #region Add additional email BCC recipient
                    if (AddEmailBCCList != null && AddEmailBCCList.Count > 0)
                    {
                        foreach (string emailBCC in AddEmailBCCList)
                        {
                            Microsoft.Exchange.WebServices.Data.EmailAddress emailAddr = new Microsoft.Exchange.WebServices.Data.EmailAddress();
                            emailAddr.Address = emailBCC;

                            resMsg.ToRecipients.Add(emailAddr);
                        }
                    }
                    #endregion

                    #region Send the reply mail
                    if (blnIsSaveCopy)
                        resMsg.SendAndSaveCopy();
                    else
                        resMsg.Send();
                    #endregion
                }
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }

            return blnIsSuccess;
        }
        */

        public bool MoveMailToFolder(Mailing targetEmail, string strDestFolderName, out string strExceptionMessage, string strParentFolderName = null, string strMailBox = null)
        {
            bool blnIsSuccess = true;
            strExceptionMessage = string.Empty;
            string userName = string.IsNullOrWhiteSpace(strMailBox) ? RPAFromEmailAddr : strMailBox;
            try
            {
                if (targetEmail == null)
                {
                    throw new Exception("Target email object is null");
                }
                else
                {
                    string destMailFolderId = null;
                    string parentMailFolderId = null;

                    if (!string.IsNullOrWhiteSpace(strMailBox))
                    {
                        destMailFolderId = GetFolderId(strMailBox, strDestFolderName);
                        parentMailFolderId = GetFolderId(strMailBox, strParentFolderName);
                    }
                    else
                    {
                        destMailFolderId = GetFolderId(strDestFolderName);
                        parentMailFolderId = GetFolderId(strParentFolderName);
                    }

                    GraphServiceClient graphClient = InitGraphServiceForApp();

                    //Create new folder if not exists
                    if (string.IsNullOrWhiteSpace(destMailFolderId))
                    {
                        if (!string.IsNullOrWhiteSpace(parentMailFolderId) && parentMailFolderId != null)
                        {
                            var mailFolder = new MailFolder
                            {
                                DisplayName = strDestFolderName
                            };

                            graphClient.Me.MailFolders[parentMailFolderId].ChildFolders
                                .Request()
                                .AddAsync(mailFolder);
                        }
                        else
                        {
                            var driveItem = new DriveItem
                            {
                                Name = strDestFolderName,
                                Folder = new Microsoft.Graph.Folder
                                {
                                }
                            };

                            graphClient.Me.Drive.Root.Children
                                .Request()
                                .AddAsync(driveItem);
                        }

                        //Find destination folder id again
                        destMailFolderId = GetFolderId(strDestFolderName);
                    }

                    if (!string.IsNullOrWhiteSpace(destMailFolderId))
                    {
                        // Move message to folder
                        Task<Message> tMsg = System.Threading.Tasks.Task.Run(async () => await graphClient.Users[userName].MailFolders[targetEmail.FolderId].Messages[targetEmail.Id].Move(destMailFolderId).Request().PostAsync());
                        graphClient.Users[userName].MailFolders[targetEmail.FolderId].Messages[targetEmail.Id].Request().UpdateAsync(tMsg.GetAwaiter().GetResult());
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

        public bool MoveMailToFolder(Mailing targetEmail, string strDestFolderName, out string strExceptionMessage, string strParentFolderName = null)
        {
            return MoveMailToFolder(targetEmail, strDestFolderName, out strExceptionMessage, strParentFolderName, null);
        }
    }
}