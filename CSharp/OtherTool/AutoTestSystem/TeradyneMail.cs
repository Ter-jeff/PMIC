using AutoTestSystem.Model;
using CommonLib.Enum;
using NLog;
using System;
using System.IO;
using System.Net.Mail;

namespace AutoTestSystem
{
    public class TeradyneMail
    {
        public void SendMail(string mailTos, QueueFile queueFile)
        {
            var taiwanAutogenTeam = "TaiwanAutogenTeam@teradyne.com";
            var mailMessage = new MailMessage();
            foreach (var mailTo in mailTos.Split(';'))
                mailMessage.To.Add(new MailAddress(mailTo));
            mailMessage.To.Add(new MailAddress(taiwanAutogenTeam));
            mailMessage.From = new MailAddress(taiwanAutogenTeam);
            mailMessage.Subject = "Pattern Validation for " + Path.GetFileName(queueFile.InputFile) + " @ " +
                                  queueFile.TimeStamp;
            var body = "<span style = 'font-family:Calibri;'>";
            body += queueFile.Print();
            body += "</span>";
            mailMessage.Body = body;
            mailMessage.IsBodyHtml = true;
            if (queueFile.RunCondition != null)
            {
                if (!string.IsNullOrEmpty(queueFile.RunCondition.FinalOutputLog) &&
                    File.Exists(queueFile.RunCondition.FinalOutputLog))
                    mailMessage.Attachments.Add(new Attachment(queueFile.RunCondition.FinalOutputLog));
                if (!string.IsNullOrEmpty(queueFile.RunCondition.OutputReport) &&
                    File.Exists(queueFile.RunCondition.OutputReport))
                    mailMessage.Attachments.Add(new Attachment(queueFile.RunCondition.OutputReport));

            }

            if (!string.IsNullOrEmpty(queueFile.OutputProcessLog) &&
                File.Exists(queueFile.OutputProcessLog))
                mailMessage.Attachments.Add(new Attachment(queueFile.OutputProcessLog));
            if (!string.IsNullOrEmpty(queueFile.OutputIniFile) &&
                File.Exists(queueFile.OutputIniFile))
                mailMessage.Attachments.Add(new Attachment(queueFile.OutputIniFile));

            var client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Port = 25;
            client.Host = "sunfire.teradyne.com";
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = false;
            try
            {
                client.Send(mailMessage);
            }
            catch (Exception)
            {
                var logger = LogManager.GetCurrentClassLogger();
                logger.Error("[" + EnumNLogMessage.Input + "] " + "Send Mail Fialed !!! ");
            }
        }
    }
}