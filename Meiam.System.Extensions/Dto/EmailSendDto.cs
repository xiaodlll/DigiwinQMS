using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Net.Mail;
using System.Text;

namespace Meiam.System.Extensions.Dto
{
    /// <summary>
    /// 邮件发送模板
    /// </summary>
    public class EmailSendDto
    {
        /// <summary>
        /// 收件人
        /// <summary>
        [Display(Name = "收件人")]
        public string ToAddress { get; set; }

        /// <summary>
        /// 抄送
        /// <summary>
        [Display(Name = "抄送")]
        public string CcAddress { get; set; }

        /// <summary>
        /// 密送
        /// <summary>
        [Display(Name = "密送")]
        public string BccAddress { get; set; }

        /// <summary>
        /// 主题
        /// <summary>
        [Display(Name = "主题")]
        public string Subject { get; set; }


        /// <summary>
        /// 内容
        /// <summary>
        [Display(Name = "内容")]
        public string Body { get; set; }

        /// <summary>
        /// 附件
        /// <summary>
        [Display(Name = "附件")]
        public EmailAttachmentDto[] Attachments { get; set; }

    }

    public class EmailAttachmentDto
    {
        /// <summary>
        /// 附件名
        /// <summary>
        [Display(Name = "附件名")]
        public string Name { get; set; }


        /// <summary>
        /// 地址
        /// <summary>
        [Display(Name = "地址")]
        public string Path { get; set; }

    }
}
