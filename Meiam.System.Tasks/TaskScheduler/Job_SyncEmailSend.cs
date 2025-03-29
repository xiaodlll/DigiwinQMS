/*
* ==============================================================================
*
* FileName: Job_SyncEmailSend.cs
* Created: 2020/6/19 9:11:44
* Author: Meiam
* Description: 
*
* ==============================================================================
*/
using Meiam.System.Extensions;
using Meiam.System.Extensions.Dto;
using Meiam.System.Interfaces;
using Meiam.System.Model;
using Microsoft.Extensions.Logging;
using Quartz;
using System;
using System.Threading.Tasks;

namespace Meiam.System.Tasks
{
    public class Job_SyncEmailSend : JobBase, IJob
    {
        private readonly ISysTasksQzService _tasksQzService;

        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<Job_SyncEmailSend> _logger;

        /// <summary>
        /// Email接口
        /// </summary>
        private readonly IEmailSmtpServices _emailSmtpServices;

        public Job_SyncEmailSend(ILogger<Job_SyncEmailSend> logger, ISysTasksQzService tasksQzService, IEmailSmtpServices emailSmtpServices)
        {
            _logger = logger;
            _tasksQzService = tasksQzService;
            _emailSmtpServices = emailSmtpServices;
        }

        private bool HasExcute = false;

        public async Task Execute(IJobExecutionContext context)
        {
            HasExcute = false;
            var executeLog = await ExecuteJob(context, async () => await Run(context));

            // 可以直接获取 JobDetail 的值
            var jobKey = context.JobDetail.Key;
            var jobId = jobKey.Name;
            //var param = context.MergedJobDataMap;

            // 也可以通过数据库配置，获取传递过来的参数
            JobDataMap data = context.JobDetail.JobDataMap;

            var model = _tasksQzService.GetId(jobId);

            if (model != null)
            {
                if (HasExcute)
                {
                    model.RunTimes += 1;
                    _logger.LogInformation(executeLog);
                    await Task.Delay(1000);
                }
                model.UpdateTime = DateTime.Now;
                model.Remark = logWrite(executeLog);
                _tasksQzService.Update(model);
            }

        }

        public static bool isRunning = false;

        public async Task Run(IJobExecutionContext context)
        {
            try
            {
                if (isRunning) 
                    return;
                isRunning = true;
                //遍历待发送邮件列表进行发送邮件
            }
            finally
            {
                isRunning = false;
            }
        }
    }
}
