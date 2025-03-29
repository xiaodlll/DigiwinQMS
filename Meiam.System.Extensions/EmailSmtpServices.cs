using Aliyun.Acs.Core;
using Aliyun.Acs.Core.Http;
using Aliyun.Acs.Core.Profile;
using Meiam.System.Common;
using Meiam.System.Extensions.Dto;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.IO;

namespace Meiam.System.Extensions
{
    public class EmailSmtpServices : IEmailSmtpServices
    {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<EmailSmtpServices> _logger;

        public EmailSmtpServices(ILogger<EmailSmtpServices> logger)
        {
            _logger = logger;
        }
        
        public bool EmailSend(EmailSendDto parm)
        {
            // 收件人、抄送人、密送人列表
            string[] toAddresses = parm.ToAddress.Split(new string[] { ";", "；" }, StringSplitOptions.RemoveEmptyEntries);
            string[] ccAddresses = (parm.CcAddress ==null?null: parm.CcAddress.Split(new string[] { ";", "；" }, StringSplitOptions.RemoveEmptyEntries));
            string[] bccAddresses = (parm.BccAddress == null ? null : parm.BccAddress.Split(new string[] { ";", "；" }, StringSplitOptions.RemoveEmptyEntries));

            string subject = parm.Subject;
            string body = parm.Body;

            string sendResult = string.Empty;
            // 发送邮件
            bool isSuccess = false;
            try { 
                SendEmail(toAddresses, ccAddresses, bccAddresses, subject, body, parm.Attachments);
                sendResult = "成功！";
                isSuccess = true;
            } catch(Exception ex){
                sendResult = "失败！" + ex.ToString();
            }

            string logs = $"向用户: {parm.ToAddress}  发送邮件[{parm.Subject}] {sendResult}";

            _logger.LogInformation(logs);

            return isSuccess;
        }

        public static void SendEmail(string[] toAddresses, string[] ccAddresses, string[] bccAddresses, string subject, string body, EmailAttachmentDto[] attachments)
        {
            string smtpHost = AppSettings.Configuration["EMAIL_SMTP:Host"];
            string smtpPort = AppSettings.Configuration["EMAIL_SMTP:Port"];
            string smtpSSL = AppSettings.Configuration["EMAIL_SMTP:SSL"];
            string smtpUsername = AppSettings.Configuration["EMAIL_SMTP:Username"];
            string smtpPassword = AppSettings.Configuration["EMAIL_SMTP:Password"];

            // 初始化邮件对象
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(smtpUsername);

            // 添加多个收件人
            foreach (string toAddress in toAddresses)
            {
                mail.To.Add(toAddress);
            }

            // 添加抄送
            if (ccAddresses != null)
            {
                foreach (string ccAddress in ccAddresses)
                {
                    mail.CC.Add(ccAddress);
                }
            }

            // 添加密送
            if (bccAddresses != null)
            {
                foreach (string bccAddress in bccAddresses)
                {
                    mail.Bcc.Add(bccAddress);
                }
            }

            // 设置邮件主题和内容
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;

            if (attachments != null)
            {
                foreach (var attachmentDto in attachments)
                {
                    // 添加附件
                    if (File.Exists(attachmentDto.Path))
                    {
                        Attachment attachment = new Attachment(attachmentDto.Path);
                        attachment.Name = attachmentDto.Name;
                        mail.Attachments.Add(attachment);
                    }
                }
            }

            // 配置SMTP客户端
            SmtpClient smtpClient = new SmtpClient(smtpHost, int.Parse(smtpPort))
            {
                DeliveryMethod = SmtpDeliveryMethod.Network,
                EnableSsl = (smtpSSL == "1" ? true : false)  // 如果SMTP服务器需要SSL
            };
            if (!string.IsNullOrEmpty(smtpPassword))
            {
                smtpClient.Credentials = new NetworkCredential(smtpUsername, smtpPassword);
            }

            // 发送邮件
            smtpClient.Send(mail);
        }
    }
}
