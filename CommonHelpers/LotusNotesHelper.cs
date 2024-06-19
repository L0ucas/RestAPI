using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Domino;

namespace TMHelper.Common
{
    public class LotusNotesHelper
    {
        /// <summary>
        /// 发送邮件（HTML格式）
        /// </summary>
        /// <param name="param"></param>
        /// <returns></returns>
        public static bool SendMail(SendEmaiParam param)
        {
            try
            {
                NotesSession ns = new NotesSession();
                ns.Initialize(param.NotesPwd);
                NotesDatabase db = ns.GetDatabase(param.NotesServer, param.NotesDbName, false);
                NotesDocument doc = db.CreateDocument();
                doc.ReplaceItemValue("From", param.From);
                //收件人信息
                doc.ReplaceItemValue("SendTo", param.MailTo);
                //邮件标题
                doc.ReplaceItemValue("Subject", param.MailSubject);
                doc.SaveMessageOnSend = param.SaveMessageOnSend;
                NotesStream notesStream = ns.CreateStream();
                notesStream.WriteText(param.MailBody);//构建HTML邮件，可以在头和尾添加公司的logo和系统提醒语
                NotesMIMEEntity mine = doc.CreateMIMEEntity("Body");//构建邮件正文
                mine.SetContentFromText(notesStream, "text/html;charset=UTF-8", MIME_ENCODING.ENC_IDENTITY_BINARY);

                doc.AppendItemValue("Principal", param.Principal);//设置邮件的发件人昵称
                //发送邮件
                object obj = doc.GetItemValue("SendTo");
                doc.Send(false, ref obj);
                doc.CloseMIMEEntities();//关闭
                return true;
            }
            catch (Exception e)
            {
                throw new Exception("异常", e);
            }
        }

        /// <summary>
        /// 发送邮件及附件
        /// </summary>
        public static bool SendMail(SendEmaiAndAttachmentParam param)
        {
            try
            {
                NotesSession ns = new NotesSession();
                ns.Initialize(param.NotesPwd);
                //初始化NotesDatabase
                NotesDatabase db = ns.GetDatabase(param.NotesServer, param.NotesDbName, false);
                NotesDocument doc = db.CreateDocument();
                //发件人信息
                doc.ReplaceItemValue("From", param.From);
                //收件人信息
                doc.ReplaceItemValue("SendTo", param.MailTo);
                //邮件标题
                doc.ReplaceItemValue("Subject", param.MailSubject);
                doc.SaveMessageOnSend = param.SaveMessageOnSend;
                //邮件正文
                NotesRichTextItem nr = doc.CreateRichTextItem("Body");
                nr.AppendText(param.MailBody);
                if (param.AttachmentPaths.Length > 0)
                {

                    //邮件附件
                    NotesRichTextItem attachment = doc.CreateRichTextItem("attachment");
                    foreach (var attachmentPath in param.AttachmentPaths)
                    {
                        attachment.EmbedObject(EMBED_TYPE.EMBED_ATTACHMENT, "", attachmentPath, "attachment");
                    }
                }

                doc.AppendItemValue("Principal", param.Principal);//设置邮件的发件人昵称
                //发送邮件
                object obj = doc.GetItemValue("SendTo");
                doc.Send(false, ref obj);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }
        /// <summary>
        /// 获取邮件及附件
        /// </summary>
        /// <param name="filePath">附件保存的路径</param>
        /// <param name="folder"></param>
        /// <param name="notesPwd"></param>
        /// <returns></returns>
        public static IList<EMailAndAttachment> GetMail(string filePath, string folder, string notesPwd, string serverDomino)
        {
            try
            {
                IList<EMailAndAttachment> emails = new List<EMailAndAttachment>();
                NotesSession ns = new NotesSession();

                ns.Initialize(notesPwd);
                NotesDbDirectory dir = ns.GetDbDirectory(serverDomino);
                NotesDatabase notesDb = dir.OpenMailDatabase();

                var notesView = notesDb.GetView(folder);
                if (notesView == null)
                {
                    return emails;
                }

                var doc = notesView.GetFirstDocument();
                while (doc != null)
                {
                    EMailAndAttachment eMail = new EMailAndAttachment()
                    {
                        MailSubject = doc.GetFirstItem("Subject")?.Text,
                        From = doc.GetFirstItem("Form")?.Text,
                        MailBody = doc.GetFirstItem("Body")?.Text,
                        Date = doc.GetFirstItem("PostedDate")?.Text,
                    };
                    //附件
                    var items = (NotesItem[])doc.Items;
                    foreach (var item in items)
                    {
                        if (item.Name == "$FILE")
                        {
                            string fileName = ((object[])item.Values)[0].ToString();
                            var attachment = doc.GetAttachment(fileName);
                            attachment.ExtractFile(Path.Combine(filePath, fileName));
                            eMail.Attachments.Add(new LotusNotesAttachment() { FileName = fileName, FilePath = filePath });
                        }
                    }
                    emails.Add(eMail);
                    //查找下一份邮件
                    doc = notesView.GetNextDocument(doc);
                }

                return emails;
            }
            catch (Exception e)
            {
                throw new Exception("获取邮件失败！", e);
            }
        }

        /// <summary>
        /// 获取指定日期下指定标题的邮件
        /// </summary>
        /// <param name="date">开始时间</param>
        /// <param name="subject"></param>
        /// <param name="notesPwd"></param>
        /// <returns></returns>
        public static IList<EMail> GetMailBySubject(string subject, string date, string notesPwd, string serverDomino)
        {
            try
            {
                IList<EMail> emails = new List<EMail>();
                NotesSession ns = new NotesSession();

                ns.Initialize(notesPwd);

                NotesDbDirectory dir = ns.GetDbDirectory(serverDomino);
                NotesDatabase notesDb = dir.OpenMailDatabase();
                var dt = ns.CreateDateTime(date);

                var dc = notesDb.Search($"@Contains(Subject;\"{subject}\")", dt, 0);
                var doc = dc.GetFirstDocument();
                while (doc != null)
                {
                    emails.Add(new EMail()
                    {
                        MailSubject = doc.GetFirstItem("Subject")?.Text,
                        From = doc.GetFirstItem("From")?.Text,
                        MailBody = doc.GetFirstItem("Body")?.Text,
                        Date = doc.GetFirstItem("PostedDate")?.Text,
                    });
                    //查找下一份邮件
                    doc = dc.GetNextDocument(doc);
                }
                return emails;
            }
            catch (Exception e)
            {
                throw new Exception("获取邮件失败！", e);
            }

        }

        /// <summary>
        /// 获取指定标题的邮件
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="notesPwd"></param>
        /// <returns></returns>
        public static IList<EMail> GetMailBySubject(string subject, string notesPwd, string serverDomino)
        {
            try
            {
                IList<EMail> emails = new List<EMail>();
                NotesSession ns = new NotesSession();

                ns.Initialize(notesPwd);

                NotesDbDirectory dir = ns.GetDbDirectory(serverDomino);
                NotesDatabase notesDb = dir.OpenMailDatabase();

                var dc = notesDb.Search($"@Contains(Subject;\"{subject}\")", null, 0);//0表示匹配所有邮件
                var doc = dc.GetFirstDocument();
                while (doc != null)
                {
                    emails.Add(new EMail()
                    {
                        MailSubject = doc.GetFirstItem("Subject")?.Text,
                        From = doc.GetFirstItem("From")?.Text,
                        MailBody = doc.GetFirstItem("Body")?.Text,
                        Date = doc.GetFirstItem("PostedDate")?.Text,
                    });
                    //查找下一份邮件
                    doc = dc.GetNextDocument(doc);
                }
                return emails;
            }
            catch (Exception e)
            {
                throw new Exception("获取邮件失败！", e);
            }

        }


        /// <summary>
        /// 列出指定文档的所有邮件
        /// </summary>
        /// <param name="folder"></param>
        /// /// <param name="notesPwd"></param>
        /// <returns></returns>
        public static IList<EMail> GetMailByFolder(string folder, string notesPwd, string serverDomino)
        {
            IList<EMail> emails = new List<EMail>();
            NotesSession ns = new NotesSession();

            ns.Initialize(notesPwd);

            NotesDbDirectory dir = ns.GetDbDirectory(serverDomino);
            NotesDatabase notesDb = dir.OpenMailDatabase();

            NotesView notesView = notesDb.GetView(folder);

            if (notesView == null)
            {
                return emails;
            }

            var doc = notesView.GetFirstDocument();
            while (doc != null)
            {
                emails.Add(new EMail()
                {
                    MailSubject = doc.GetFirstItem("Subject")?.Text,
                    From = doc.GetFirstItem("From")?.Text,
                    MailBody = doc.GetFirstItem("Body")?.Text,
                    Date = doc.GetFirstItem("PostedDate")?.Text,
                });
                //查找下一份邮件
                doc = notesView.GetNextDocument(doc);
            }
            return emails;
        }

        /// <summary>
        /// 列出当前邮件数据库所有的视图
        /// </summary>
        /// <param name="notesPwd"></param>
        /// <returns></returns>
        public static IList<View> GetAllViews(string notesPwd, string serverDomino)
        {
            IList<View> views = new List<View>();
            NotesSession ns = new NotesSession();

            ns.Initialize(notesPwd);

            NotesDbDirectory dir = ns.GetDbDirectory(serverDomino);
            NotesDatabase notesDb = dir.OpenMailDatabase();

            foreach (var view in notesDb.Views)
            {
                views.Add(new View()
                {
                    Name = view.Name,
                    IsFolder = view.IsFolder
                });
            }

            return views;
        }

        public static bool IBMNotes_SendMail(string IBMPwd, string mailTo, string mailSubject, string htmlBody)
        {
            bool flag;
            try
            {
                NotesSession ns = new NotesSession();
                ns.Initialize(IBMPwd);
                string fileName = ns.URLDatabase.FileName;
                NotesDocument doc = ns.GetDatabase("", fileName, false).CreateDocument();
                doc.ReplaceItemValue("SendTo", mailTo.Split(new char[] { ',' }));
                doc.ReplaceItemValue("Subject", mailSubject);
                doc.SaveMessageOnSend = true;
                NotesStream notesStream = ns.CreateStream();
                notesStream.WriteText(htmlBody, EOL_TYPE.EOL_NONE);
                doc.CreateMIMEEntity("Body").SetContentFromText(notesStream, "text/html;charset=UTF-8", MIME_ENCODING.ENC_IDENTITY_BINARY);
                object itemValue = doc.GetItemValue("SendTo");
                doc.Send(false, ref itemValue);
                doc.CloseMIMEEntities(false, "Body");
                flag = true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                flag = false;
            }
            return flag;
        }

        public static bool IBMNotes_SendMail(string IBMPwd, string mailTo, string ccTo, string bccTo, string mailSubject, string htmlBody, string att)
        {
            try
            {
                NotesSession ns = new NotesSession();
                ns.Initialize(IBMPwd);
                string fileName = ns.URLDatabase.FileName;
                NotesDocument doc = ns.GetDatabase("", fileName, false).CreateDocument();
                doc.ReplaceItemValue("SendTo", mailTo.Split(new char[] { ',' }));
                doc.ReplaceItemValue("CopyTo", ccTo.Split(new char[] { ',' }));
                doc.ReplaceItemValue("BlindCopyTo", bccTo.Split(new char[] { ',' }));
                doc.ReplaceItemValue("Subject", mailSubject);
                doc.SaveMessageOnSend = true;
                NotesStream notesStream = ns.CreateStream();
                notesStream.WriteText(htmlBody, EOL_TYPE.EOL_NONE);
                doc.CreateMIMEEntity("Body").SetContentFromText(notesStream, "text/html;charset=UTF-8", MIME_ENCODING.ENC_IDENTITY_BINARY);

                if (att.Length > 0)
                {
                    string[] arrAtt = att.Split(',');
                    //邮件附件
                    NotesRichTextItem attachment = doc.CreateRichTextItem("attachment");
                    foreach (string strPath in arrAtt)
                    {
                        attachment.EmbedObject(EMBED_TYPE.EMBED_ATTACHMENT, "", strPath, "attachment");
                    }
                }

                object itemValue = doc.GetItemValue("SendTo");
                doc.Send(false, ref itemValue);
                doc.CloseMIMEEntities(false, "Body");
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex.InnerException);
            }
        }

        private static IList<EMailAndAttachment> GetMailBySubject(string strFormula, DateTime dtCutoffDate, string notesPwd, 
            string serverDomino, string downloadPath, bool saveAtt)
        {
            try
            {
                string date = dtCutoffDate.ToString("yyyy-MMM-dd HH:mm:ss");
                IList<EMailAndAttachment> emails = new List<EMailAndAttachment>();
                NotesSession ns = new NotesSession();
                ns.Initialize(notesPwd);
                NotesDbDirectory dir = ns.GetDbDirectory(serverDomino);
                NotesDatabase notesDb = dir.OpenMailDatabase();

                var dt = ns.CreateDateTime(date);                            //sample format "08/18/95 01:36:22 PM" 
                var dc = notesDb.Search(strFormula, dt , 0);//0表示匹配所有邮件
                var doc = dc.GetLastDocument();             // always get the latest received email
                while (doc != null)
                {
                    EMailAndAttachment email = new EMailAndAttachment()
                    {
                        MailSubject = doc.GetFirstItem("Subject")?.Text,
                        From = doc.GetFirstItem("From")?.Text,
                        MailBody = doc.GetFirstItem("Body")?.Text,
                        Date = doc.GetFirstItem("PostedDate")?.Text,
                    };

                    foreach (var item in doc.Items)
                    {
                        if (item.Name == "$FILE")
                        {
                            string strFileName = ((object[])item.Values)[0].ToString();
                            string strDownloadFolder = Path.Combine(downloadPath, DateTime.Today.ToString("yyyyMMdd"));
                            string strSavedFileName = $"{strFileName.Substring(0, strFileName.LastIndexOf('.'))}_{dtCutoffDate.ToString("yyyyMMddHHmmss")}_{DateTime.Now.ToString("yyyyMMddHHmmss")}{strFileName.Substring(strFileName.LastIndexOf('.'))}";

                            var attachment = doc.GetAttachment(strFileName);
                            if (saveAtt)
                            {
                                if (!Directory.Exists(strDownloadFolder))
                                {
                                    Directory.CreateDirectory(strDownloadFolder);
                                }
                                attachment.ExtractFile(Path.Combine(strDownloadFolder, strSavedFileName));
                            }
                            email.Attachments.Add(new LotusNotesAttachment() { FileName = strSavedFileName, FilePath = strDownloadFolder, isSaved = saveAtt });
                        }
                    }
                    emails.Add(email);
                    doc = dc.GetPrevDocument(doc);
                }
                return emails;
            }
            catch (Exception e)
            {
                throw new Exception("获取邮件失败！", e);
            }

        }

        public static IList<EMailAndAttachment> GetMailBySubject(string sender, string subject, DateTime dtReceived, DateTime dtCutOffDate, string notesPwd,
            string serverDomino, string downloadPath, bool saveAtt, bool caseSensitive)
        {
            try
            {
                string date = dtReceived.ToString("yyyy-MMM-dd HH:mm:ss");
                IList<EMailAndAttachment> emails = new List<EMailAndAttachment>();
                //NotesSession ns = new NotesSession();
                //ns.Initialize(notesPwd);
                //NotesDbDirectory dir = ns.GetDbDirectory(serverDomino);
                //NotesDatabase notesDb = dir.OpenMailDatabase();

                //var dt = ns.CreateDateTime(date);                            //sample format "08/18/95 01:36:22 PM" 
                var dtYear = dtReceived.Year;
                var dtMonth = dtReceived.Month;
                var dtDay = dtReceived.Day;
                string strFormula = "@Contains(" + ((caseSensitive) ? "Subject" : "@LowerCase(Subject)") + "; \"" +
                    ((caseSensitive) ? subject : subject.ToLower()) + "\")";
                if (sender != string.Empty)
                {
                    strFormula = strFormula + " & @Contains(" + ((caseSensitive) ? "From" : "@LowerCase(From)") + ";\"" +
                        ((caseSensitive) ? sender : sender.ToLower()) + "\")";              // sample "CN=Aspiro Rpa/OU=MHQ/O=ASP_MY" OR "xxx.xxx@accenture.com";
                }

                if (date != string.Empty)
                {
                    strFormula = strFormula + " & @Date(PostedDate) = @Date(" + dtYear.ToString() + ";" + dtMonth.ToString() + ";" + dtDay.ToString() + ")";
                }

                emails = GetMailBySubject(strFormula, dtCutOffDate, notesPwd, serverDomino, downloadPath, saveAtt);

                return emails;
            }
            catch (Exception e)
            {
                throw new Exception("获取邮件失败！", e);
            }

        }
    }
}
