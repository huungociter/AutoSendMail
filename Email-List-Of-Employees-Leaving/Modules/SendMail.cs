using MailKit.Security;
using MimeKit;

namespace EmployeeSendMailProcessor
{
    public class SendMail
    {
        string emailSMTPServer, emailSMTPUserName, emailSMTPPassword;
        int emailSMTPPort;

        public SendMail()
        {
            // Retrieve SMTP server settings from environment variables
            emailSMTPServer = DotNetEnv.Env.GetString("EMAIL_SMTP_SERVER");
            emailSMTPUserName = DotNetEnv.Env.GetString("EMAIL_SMTP_USERNAME");
            emailSMTPPassword = DotNetEnv.Env.GetString("EMAIL_SMTP_PASSWORD");
            emailSMTPPort = DotNetEnv.Env.GetInt("EMAIL_SMTP_PORT");
        }

        public void SendMailEmployeesLeavingTomorrow(List<string> lstReceiver, string mailSubject, string mailHtml, MemoryStream fileAttach, string fileName)
        {
            MimeMessage messageMail = new MimeMessage();
            InternetAddressList recipients = new();

            // Check if lstReceiver is not null and contains items
            if (lstReceiver != null && lstReceiver.Count > 0)
            {
                foreach (var mail in lstReceiver)
                {
                    recipients.Add(new MailboxAddress("", mail.ToString()));
                }
            }

            messageMail.To.AddRange(recipients);

            // Set the email subject and body
            messageMail.Subject = mailSubject;
            BodyBuilder bodyContents = new()
            {
                HtmlBody = mailHtml
            };

            // Handle MemoryStream attachment if no files are provided
            if (fileAttach != null && fileAttach.Length > 0)
            {
                bodyContents.Attachments.Add(fileName, fileAttach.ToArray());
            }

            // Set the sender's address, defaulting to the SMTP username if none is provided, check file env
            messageMail.From.Add(new MailboxAddress("", emailSMTPUserName));
            messageMail.Body = bodyContents.ToMessageBody();

            // Initialize the SMTP client
            MailKit.Net.Smtp.SmtpClient client = new();
            try
            {
                // Attempt to connect to the SMTP server with certificate validation
                client.CheckCertificateRevocation = false;
                client.Connect(emailSMTPServer, emailSMTPPort, SecureSocketOptions.Auto);
            }
            catch
            {
                // Fallback to connecting without certificate validation
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect(emailSMTPServer, 587, SecureSocketOptions.None);
            }

            // Authenticate and send the email
            client.Authenticate(emailSMTPUserName, emailSMTPPassword);
            client.Send(messageMail);
            client.Disconnect(true);
        }
    }
}
