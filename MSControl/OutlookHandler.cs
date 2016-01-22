using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MSController
{
    /// <summary>
    /// Handles Microsoft Outlook, can send emails with multiple attachments.
    /// </summary>
    public class OutlookHandler
    {
        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipient">Single recipient of the email</param>
        public void sendMail(string subject, string body, string recipient)
        {
            sendMail(subject, body, new List<string>() { recipient }, new List<string>());
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        public void sendMail(string subject, string body, List<string> recipients)
        {
            sendMail(subject, body, recipients, new List<string>());
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipeint">Single recipient of the email.</param>
        /// <param name="attachmentPath">Single attachment path of the email.</param>
        public void sendMail(string subject, string body, string recipeint, string attachmentPath)
        {
            sendMail(subject, body, new List<string>() { recipeint }, new List<string>() { attachmentPath });
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="attachmentPath">Single attachment path of the email.</param>
        public void sendMail(string subject, string body, List<string> recipients, string attachmentPath)
        {
            sendMail(subject, body, recipients, new List<string>() { attachmentPath });
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipient">Single recipient of the email.</param>
        /// <param name="attachmentPaths">List of attachment paths of the email.</param>
        public void sendMail(string subject, string body, string recipient, List<string> attachmentPaths)
        {
            sendMail(subject, body, new List<string>() { recipient }, attachmentPaths);
        }

        /// <summary>
        /// Sends an email using Outlook.
        /// </summary>
        /// <param name="subject">Subject line of the email.</param>
        /// <param name="body">Body text of the email.</param>
        /// <param name="recipients">List of recipients of the email.</param>
        /// <param name="attachmentPaths">List of the attachment paths of the email.</param>
        public void sendMail(string subject, string body, List<string> recipients, List<string> attachmentPaths)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = subject;
                mailItem.Body = body;
                mailItem.To = String.Join("; ", recipients.ToArray());

                foreach (string attachment in attachmentPaths)
                    mailItem.Attachments.Add(attachment);

                mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
                mailItem.Display(false);

                ((Outlook._MailItem)mailItem).Send();
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}
