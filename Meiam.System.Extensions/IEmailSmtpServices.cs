using Aliyun.Acs.Core;
using Meiam.System.Extensions.Dto;
using System;
using System.Collections.Generic;
using System.Text;

namespace Meiam.System.Extensions
{
    public interface IEmailSmtpServices
    {
        bool EmailSend(EmailSendDto parm);
    }
}
