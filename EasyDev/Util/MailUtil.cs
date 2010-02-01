using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.IO;
using Microsoft.Office.Interop.Outlook;

namespace EasyDev.Util
{
        /// <summary>
        /// Contains fields a email needs.
        /// </summary>
        [Serializable]
        public class MailPackage
        {
                /// <summary>
                /// 
                /// </summary>
                public string From { get; set; }

                /// <summary>
                /// 
                /// </summary>
                public string To { get; set; }

                /// <summary>
                /// Carbon Copy
                /// </summary>
                public string CC { get; set; }

                /// <summary>
                /// 
                /// </summary>
                public string Bcc { get; set; }

                /// <summary>
                /// 
                /// </summary>
                public string Subject { get; set; }

                /// <summary>
                /// 
                /// </summary>
                public string Body { get; set; }

                /// <summary>
                /// 
                /// </summary>
                public string[] Attachments { get; set; }

                public MailPackage(string from, string to, string subject, string body)
                {
                        From = from;
                        To = to;
                        Subject = subject;
                        Body = body;
                }
        }

        /// <summary>
        /// 
        /// </summary>
        public class MailUtil
        {
                /// <summary>
                /// 
                /// </summary>
                /// <param name="pkg"></param>
                public static void SimpleSendMail(MailPackage pkg)
                {
                        MailAddress fromAddress = new MailAddress(pkg.From);
                        MailMessage msg = new MailMessage();

                        msg.From = fromAddress;

                        string[] toAddressList = pkg.To.Split(new char[] { ';' });
                        foreach (string toAddress in toAddressList)
                        {
                                if (!string.IsNullOrEmpty(toAddress.Trim()))
                                {
                                        msg.To.Add(new MailAddress(toAddress));
                                }
                        }

                        msg.Subject = pkg.Subject;
                        msg.Body = pkg.Body;
                        msg.Priority = MailPriority.High;

                        if (pkg.Attachments != null)
                        {
                                foreach (string filename in pkg.Attachments)
                                {
                                        if (File.Exists(filename))
                                        {
                                                msg.Attachments.Add(new System.Net.Mail.Attachment(filename));
                                        }
                                }
                        }

                        SmtpClient client = new SmtpClient();
                        client.Send(msg);
                }

                /// <summary>
                /// 
                /// </summary>
                /// <param name="pkg"></param>
                public static void SendMailByLoalOutlook(MailPackage pkg)
                {
                        Application ol = new Application();
                        MailItem item = ol.CreateItem(0) as MailItem;
                        if (pkg.CC != null && pkg.CC.Length > 0)
                        {
                                item.CC = pkg.CC;
                        }

                        for (int i = 0; i < pkg.Attachments.Length; i++)
                        {
                                FileInfo info = new FileInfo(pkg.Attachments[i]);
                                item.Attachments.Add(pkg.Attachments[i], OlAttachmentType.olByValue, 1, info.Name);
                        }

                        item.To = pkg.To;
                        item.Subject = pkg.Subject;
                        item.Body = pkg.Body;
                        item.Send();
                }
        }
}