using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Utils;

namespace MailLib
{
    public enum ServerSecurity
    {
        None = SecureSocketOptions.None,
        Auto = SecureSocketOptions.Auto,
        Ssl = SecureSocketOptions.SslOnConnect,
        Tls = SecureSocketOptions.StartTls,
        TlsIfAvailable = SecureSocketOptions.StartTlsWhenAvailable
    }

    public class EmailSender
    {
        private readonly MimeMessage _message = new MimeMessage();
        private readonly BodyBuilder _bodyBuilder = new BodyBuilder();
        private readonly List<string> _embeddedImageIds = new List<string>();

        /// <summary>
        /// Gets or sets the mail server host name or IP address.
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// Gets or sets the port to use for mail transport.
        /// </summary>
        public int Port { get; set; }

        /// <summary>
        /// Gets or sets security type used for mail transport.
        /// </summary>
        public ServerSecurity ConnectionType { get; set; } = ServerSecurity.Auto;

        /// <summary>
        /// Gets or sets the username for mail server authentication, if applicable.
        /// </summary>
        public string UserName { get;set; }

        /// <summary>
        /// Gets or sets the password for mail server authentication, if applicable.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets the Sender: address for the message.
        /// </summary>
        public string Sender
        {
            get { return _message.Sender.Address; }
            set { _message.Sender = new MailboxAddress("", value); }
        }
        
        /// <summary>
        /// Gets or sets the subject of the message.
        /// </summary>
        public string Subject
        {
            get { return _message.Subject; }
            set { _message.Subject = value; }
        }

        /// <summary>
        /// Gets or sets the text body contents of the message.
        /// </summary>
        public string TextBody
        {
            get { return _bodyBuilder.TextBody; }
            set { _bodyBuilder.TextBody = value; }
        }

        /// <summary>
        /// Gets or sets the HTML body contents of the message.
        /// Changing this may break embedded images added with <see cref="AppendEmbeddedImage"/>.
        /// </summary>
        public string HtmlBody
        {
            get { return _bodyBuilder.HtmlBody; }
            set { _bodyBuilder.HtmlBody = value; }
        }

        /// <summary>
        /// Appends text to the message body.
        /// </summary>
        /// <param name="content">The text content to append</param>
        public void AppendToTextBody(
            string content)
        {
            TextBody += content;
        }

        /// <summary>
        /// Appends HTML to the message body.
        /// </summary>
        /// <param name="content">The HTML content to append</param>
        public void AppendToHtmlBody(
            string content)
        {
            HtmlBody += content;
        }

        /// <summary>
        /// Appends text from a file to the message body.
        /// </summary>
        /// <param name="filePath">Full or relative path to the file</param>
        public void AppendToTextBodyFromFile(
            string filePath)
        {
            TextBody += File.ReadAllText(filePath);
        }

        /// <summary>
        /// Appends HTML from a file to the message body.
        /// </summary>
        /// <param name="filePath">Full or relative path to the file</param>
        public void AppendToHtmlBodyFromFile(
            string filePath)
        {
            HtmlBody += File.ReadAllText(filePath);
        }

        /// <summary>
        /// Adds a To: address to the message.
        /// To add multiple To: recipients, call this once for each recipient email address.
        /// </summary>
        /// <param name="email">Email address of the recipient</param>
        /// <param name="name">Name of the recipient</param>
        public void AddToAddress(
            string email,
            string name = null)
        {
            if (!string.IsNullOrEmpty(name))
                _message.To.Add(new MailboxAddress(name, email));
            else
                _message.To.Add(new MailboxAddress("", email));
        }

        /// <summary>
        /// Adds a CC address to the message.
        /// To add multiple CC: recipients, call this once for each recipient email address.
        /// </summary>
        /// <param name="email">Email address of the recipient</param>
        /// <param name="name">Name of the recipient</param>
        public void AddCcAddress(
            string email,
            string name = null)
        {
            if (!string.IsNullOrEmpty(name))
                _message.Cc.Add(new MailboxAddress(name, email));
            else
                _message.Cc.Add(new MailboxAddress("", email));
        }

        /// <summary>
        /// Adds a BCC address to the message.
        /// To add multiple BCC: recipients, call this once for each recipient email address.
        /// </summary>
        /// <param name="email">Email address of the recipient</param>
        /// <param name="name">Name of the recipient</param>
        public void AddBccAddress(
            string email,
            string name = null)
        {
            if (!string.IsNullOrEmpty(name))
                _message.Bcc.Add(new MailboxAddress(name, email));
            else
                _message.Bcc.Add(new MailboxAddress("", email));
        }

        /// <summary>
        /// Adds a From address for the message.
        /// </summary>
        /// <param name="email">Email address of the sender</param>
        /// <param name="name">Name of the sender</param>
        public void AddFromAddress(string email,
                                   string name = null)
        {
            if (!string.IsNullOrEmpty(name))
                _message.From.Add(new MailboxAddress(name, email));
            else
                _message.From.Add(new MailboxAddress("", email));
        }
        
        /// <summary>
        /// Adds a Reply-To address for the message.
        /// To add multiple Reply-To addresses, call this once for each Reply-To email address.
        /// </summary>
        /// <param name="email">Email address of the reply-to</param>
        /// <param name="name">Name of the reply-to</param>
        public void AddReplyToAddress(string email,
                                        string name = null)
        {
            if (!string.IsNullOrEmpty(name))
                _message.ReplyTo.Add(new MailboxAddress(name, email));
            else
                _message.ReplyTo.Add(new MailboxAddress("", email));
        }

        /// <summary>
        /// Adds an attachment to the message.
        /// </summary>
        /// <param name="filePath">Full or relative path to the file</param>
        public void AddAttachment(
            string filePath)
        {
            _bodyBuilder.Attachments.Add(filePath);
        }

        /// <summary>
        /// Adds an embedded image that will appear at the very end of the message body.
        /// Setting <see cref="TextBody"/> or <see cref="HtmlBody"/> does NOT clear embedded
        /// images added with this method.
        /// </summary>
        /// <param name="imagePath">Full or relative path to the image file</param>
        public void AddEmbeddedImage(
            string imagePath)
        {
            var image = _bodyBuilder.LinkedResources.Add(imagePath);
            image.ContentId = MimeUtils.GenerateMessageId();

            _embeddedImageIds.Add(image.ContentId);
        }

        /// <summary>
        /// Adds an embedded image to the end of the current contents of the <see cref="HtmlBody"/>.
        /// </summary>
        /// <param name="imagePath">Full or relative path to the image file</param>
        public void AppendEmbeddedImage(
            string imagePath)
        {
            var image = _bodyBuilder.LinkedResources.Add(imagePath);
            image.ContentId = MimeUtils.GenerateMessageId();

            _bodyBuilder.HtmlBody += $"<img src='cid:{image.ContentId}'>";
        }

        /// <summary>
        /// Clears CC Addresses from _message
        /// </summary>
        public void ClearCcAddresses()
        {
            _message.Cc.Clear();
        }

        /// <summary>
        /// Clears BCC Addresses from _message
        /// </summary>
        public void ClearBccAddresses()
        {
            _message.Bcc.Clear();
        }

        /// <summary>
        /// Clears to Addresses from _message
        /// </summary>
        public void ClearToAddresses()
        {
            _message.To.Clear();
        }

        /// <summary>
        /// Sends the message.
        /// </summary>
        /// <returns>true if no error</returns>
        public bool Send()
        {
            try
            {
                using (var client = new SmtpClient())
                {
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    client.Connect(Host, Port, (SecureSocketOptions)(int)ConnectionType);

                    if (!string.IsNullOrEmpty(UserName)
                        && !string.IsNullOrEmpty(Password))
                    {
                        client.Authenticate(UserName, Password);
                    }

                    foreach (var embeddedImgId in _embeddedImageIds)
                        _bodyBuilder.HtmlBody += $"<img src='cid:{embeddedImgId}'>";

                    _message.Body = _bodyBuilder.ToMessageBody();

                    client.Send(_message);
                }
            } 
            catch (Exception exc)
            {
                // Try to log the error to the Windows Event Log
                try
                {
                    using (EventLog eventLog = new EventLog("Application"))
                    {
                        eventLog.Source = "Application";
                        eventLog.WriteEntry($"MailLib Error: {exc.Message}", EventLogEntryType.Error, 101, 1);
                    }
                }
                catch (Exception)
                { /* Can't really do anything. */ }

                return false;
            }

            return true;
        }
    }
}
