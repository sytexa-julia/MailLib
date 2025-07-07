using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Utils;
using System.Runtime.InteropServices;

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

//    [ComVisible(true)]
//    [Guid("17fbc6e7-2ceb-44b5-83da-b5da5387bc56")]
//    [ProgId("MailLib.EmailSender")]
    public class EmailSender
	{
		private readonly MimeMessage _message = new MimeMessage();
		private readonly BodyBuilder _bodyBuilder = new BodyBuilder();
		private readonly List<string> _embeddedImageIds = new List<string>();
		private System.Security.Authentication.SslProtocols _our_protocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13 | System.Security.Authentication.SslProtocols.Ssl2 | System.Security.Authentication.SslProtocols.Ssl3;
		private dynamic _emailConfig;
		private string _from;
		private string _to;
		private string _replyto;
		private string _cc;
		private string _bcc;
		private string _log_file;

		public string log_file { get => _log_file; set => _log_file = value; }

		/// <summary>
		/// Gets or sets the Reply-To address for the message.
		/// </summary>
		/// <value>The Reply-To address for the message.</value>
		public string ReplyTo
		{
			get => _replyto; set { Set_ReplyTo(value); }
		}

		/// <summary>
		/// Gets or sets the From address for the message.
		/// </summary>
		/// <value>The From address for the message.</value>
		public string From
		{
			get => _from; set { Set_From(value); }
		}

		/// <summary>
		/// Gets or sets the To addresses for the message.
		/// </summary>
		/// <value>The To addresses for the message.</value>
		public string To
		{
			get => _to; set { Set_To(value); }
		}

		/// <summary>
		/// Gets or sets the CC addresses for the message.
		/// </summary>
		/// <value>The CC addresses for the message.</value>
		public string Cc
		{
			get => _cc; set { Set_Cc(value); }
		}

		/// <summary>
		/// Gets or sets the BCC addresses for the message.
		/// </summary>
		/// <value>The BCC addresses for the message.</value>
		public string Bcc
		{
			get => _bcc; set { Set_Bcc(value); }
		}

		/// <summary>
		/// This is the CDO Configuration object
		/// It is used to set the SMTP server and credentials
		/// It comes from the ClassicASP World and makes this a drop in replacement for CDO.Message.Configuratoin    
		/// </summary>
		public dynamic Configuration
		{
			get => _emailConfig;
			set
			{
				_emailConfig = value;
				ProcessConfiguration();
			}
		}

		/// <summary>
		/// Method to process and extract needed configuration
		/// </summary>
		public void ProcessConfiguration()
		{
			if (Configuration == null)
			{
				throw new InvalidOperationException("CDO Configuration has not been set");
			}

			// Loop through all fields exactly as you wanted
			foreach (dynamic field in Configuration.Fields)
			{
				string name = field.Name;
				object value = field.Value;

				// Console.WriteLine($"{name} = {value}");

				// Extract values you care about 
                // We IGNORE "sendusing" because we only support (1) value: 1 = SMTP.
				switch (name)
				{
					case "http://schemas.microsoft.com/cdo/configuration/smtpserver":
						Host = value?.ToString();
						break;
					case "http://schemas.microsoft.com/cdo/configuration/smtpserverport":
						Port = Convert.ToInt32(value);
						break;
					case "http://schemas.microsoft.com/cdo/configuration/sendusername":
						UserName = value?.ToString();
						break;
					case "http://schemas.microsoft.com/cdo/configuration/sendpassword":
						Password = value?.ToString();
						break;
                    case "http://schemas.microsoft.com/cdo/configuration/from":     // We added these, so that replyto and from do not have to be set but once
                        Set_From( value?.ToString() );
                        // Console.WriteLine($"Debug: From: {_from}");
                        break;
                    case "http://schemas.microsoft.com/cdo/configuration/replyto":  // We added these, so that replyto and from do not have to be set but once
                        Set_ReplyTo(value?.ToString());
                        // Console.WriteLine($"Debug: ReplyTo: {_replyto}");
                        break;
                    case "http://schemas.microsoft.com/cdo/configuration/smtpusessl":
						bool useSsl = Convert.ToBoolean(value);
						ConnectionType = useSsl ? ServerSecurity.Ssl : ServerSecurity.Auto;
						break;
						// TODO: This can be used to set UseTls, and the sslMode (from these) [9] Total Parameters
				}
			}
		}

		/// <summary>
		/// Parses a comma-separated list of email addresses into a MailAddressCollection.
		/// This is used to parse the To, CC, BCC, and Reply-To addresses.
		/// It has the advantage of returning a Usable MailAddressCollection object,
		/// and not a just string.  Plus it properly parses COMPLEX Emails: ("John Doe" <john.doe@example.com>) 
		/// </summary>
		/// <param name="commaSeparatedEmails">Comma-separated list of email addresses</param>
		/// <returns>MailAddressCollection containing the parsed email addresses</returns>
		public static MailAddressCollection ParseEmailList(string commaSeparatedEmails)
		{
			var collection = new MailAddressCollection();
			if (string.IsNullOrWhiteSpace(commaSeparatedEmails))
				return collection;

			// Use MailAddressCollection's built-in parser (RFC 2822 compliant)
			collection.Add(commaSeparatedEmails);
			return collection;
		}

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
		public string UserName { get; set; }

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
		/// Appends an SSL protocol to the list of protocols to use.
		/// </summary>
		/// <param name="proto">The SSL protocol to append</param> 
		/// <remarks>  smtp.AppendSslProtocol(System.Security.Authentication.SslProtocols.Tls12);
		public void AppendSslProtocol(System.Security.Authentication.SslProtocols proto)
		{
			_our_protocols = _our_protocols | proto;
		}

		/// <summary>
		/// Appends text to the message body.
		/// </summary>
		/// <param name="content">The text content to append</param>
		public void AppendToTextBody(string content)
		{
			TextBody += content;
		}

		/// <summary>
		/// Appends HTML to the message body.
		/// </summary>
		/// <param name="content">The HTML content to append</param>
		public void AppendToHtmlBody(string content)
		{
			HtmlBody += content;
		}

		/// <summary>
		/// Appends text from a file to the message body.
		/// </summary>
		/// <param name="filePath">Full or relative path to the file</param>
		public void AppendToTextBodyFromFile(string filePath)
		{
			TextBody += File.ReadAllText(filePath);
		}

		/// <summary>
		/// Appends HTML from a file to the message body.
		/// </summary>
		/// <param name="filePath">Full or relative path to the file</param>
		public void AppendToHtmlBodyFromFile(string filePath)
		{
			HtmlBody += File.ReadAllText(filePath);
		}

		/// <summary>
		/// Adds a To: address to the message.
		/// To add multiple To: recipients, call this once for each recipient email address.
		/// </summary>
		/// <param name="email">Email address of the recipient</param>
		/// <param name="name">Name of the recipient</param>
		public void AddToAddress(string email, string name = null)
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
		public void AddCcAddress(string email, string name = null)
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
		public void AddBccAddress(string email, string name = null)
		{
			if (!string.IsNullOrEmpty(name))
				_message.Bcc.Add(new MailboxAddress(name, email));
			else
				_message.Bcc.Add(new MailboxAddress("", email));
		}

		/// <summary>
		/// Sets the Reply-To address for the message.
		/// </summary>
		/// <param name="email_list">Comma-separated list of email addresses for reply-to</param>
		void Set_ReplyTo(string email_list)
		{
			_replyto = email_list;
			var emails = ParseEmailList(email_list);
			_message.ReplyTo.Clear();
			foreach (var addr in emails)
			{
				_message.ReplyTo.Add(new MailboxAddress(addr.DisplayName ?? "", addr.Address));
				break;  // We only expect one
			}
		}

		/// <summary>
		/// Sets the To addresses for the message.
		/// </summary>
		/// <param name="email_list">Comma-separated list of email addresses for recipients</param>
		public void Set_To(string email_list)
		{
			_to = email_list;
			var emails = ParseEmailList(email_list);
			_message.To.Clear();
			foreach (var addr in emails)
				_message.To.Add(new MailboxAddress(addr.DisplayName ?? "", addr.Address));
		}

		/// <summary>
		/// Sets the BCC addresses for the message.
		/// </summary>
		/// <param name="email_list">Comma-separated list of email addresses for BCC recipients</param>
		public void Set_Bcc(string email_list)
		{
			_bcc = email_list;
			var emails = ParseEmailList(email_list);
			_message.Bcc.Clear();
			foreach (var addr in emails)
				_message.Bcc.Add(new MailboxAddress(addr.DisplayName ?? "", addr.Address));
		}

		/// <summary>
		/// Sets the CC addresses for the message.
		/// </summary>
		/// <param name="email_list">Comma-separated list of email addresses for CC recipients</param>
		public void Set_Cc(string email_list)
		{
			_cc = email_list;
			var emails = ParseEmailList(email_list);
			_message.Cc.Clear();
			foreach (var addr in emails)
				_message.Cc.Add(new MailboxAddress(addr.DisplayName ?? "", addr.Address));
		}

		/// <summary>
		/// Sets the From address for the message.
		/// </summary>
		/// <param name="email_list">Comma-separated list of email addresses for sender (only first address is used)</param>
		private void Set_From(string email_list)
		{
			_from = email_list;
			var emails = ParseEmailList(email_list);
			_message.From.Clear();
			foreach (var addr in emails)
			{
				_message.From.Add(new MailboxAddress(addr.DisplayName ?? "", addr.Address));
				break;  // We only expect one
			}
		}

		/// <summary>
		/// Adds a From address for the message.
		/// </summary>
		/// <param name="email">Email address of the sender</param>
		/// <param name="name">Name of the sender</param>
		public void AddFromAddress(string email, string name = null)
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
		public void AddReplyToAddress(string email, string name = null)
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
		public void AddAttachment(string filePath)
		{
			_bodyBuilder.Attachments.Add(filePath);
		}

		/// <summary>
		/// Adds an embedded image that will appear at the very end of the message body.
		/// Setting <see cref="TextBody"/> or <see cref="HtmlBody"/> does NOT clear embedded
		/// images added with this method.
		/// </summary>
		/// <param name="imagePath">Full or relative path to the image file</param>
		public void AddEmbeddedImage(string imagePath)
		{
			var image = _bodyBuilder.LinkedResources.Add(imagePath);
			image.ContentId = MimeUtils.GenerateMessageId();

			_embeddedImageIds.Add(image.ContentId);
		}

		/// <summary>
		/// Adds an embedded image to the end of the current contents of the <see cref="HtmlBody"/>.
		/// </summary>
		/// <param name="imagePath">Full or relative path to the image file</param>
		public void AppendEmbeddedImage(string imagePath)
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
		/// Creates an SmtpClient object (Based on LogFilePath being set.  If set, the client can DEBUG their connections!)
		/// </summary>
		/// <param name="logFilePath">Path to the log file</param>
		/// <returns>SmtpClient object</returns>
		private MailKit.Net.Smtp.SmtpClient CreateSmtpClient(string logFilePath)
		{
			return string.IsNullOrEmpty(logFilePath)
				? new MailKit.Net.Smtp.SmtpClient()
				: new MailKit.Net.Smtp.SmtpClient(new MailKit.ProtocolLogger(logFilePath));
		}

		/// <summary>
		/// Sends the message.  This has been tested with Google, Yahoo, and Outlook.
		/// It has been tested with TLS, SSL, and STARTTLS.
		/// It has also been tested using AhaSend.com as a SMTP server. (A great service we recommend!)
		/// </summary>
		/// <returns>true if no error</returns>
		public bool Send()
		{
			try
			{
				using (var client = CreateSmtpClient(_log_file))
				{
					// client.SslProtocols =  System.Security.Authentication.SslProtocols.Tls12; 
					client.SslProtocols = _our_protocols;
					Console.WriteLine("Protocols: " + client.SslProtocols);
					//client.ServerCertificateValidationCallback = (s, c, h, e) => true;

					client.Connect(Host, Port, (SecureSocketOptions)(int)ConnectionType);

					client.Authenticate(UserName, Password);

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
