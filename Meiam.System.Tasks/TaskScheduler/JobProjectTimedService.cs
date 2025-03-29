/*
* ==============================================================================
*
* FileName: JobProjectTimedService.cs
* Created: 2020/6/19 9:11:44
* Author: Meiam
* Description: 
*
* ==============================================================================
*/
using Meiam.System.Common;
using Meiam.System.Interfaces;
using Microsoft.Extensions.Logging;
using Quartz;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace Meiam.System.Tasks
{
    /// <summary>
    /// 添加项目任务超时提醒
    /// </summary>
    public class JobProjectTimedService : JobBase, IJob
    {
        private readonly ISysTasksQzService _tasksQzService;

        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<JobProjectTimedService> _logger;


        public JobProjectTimedService(ILogger<JobProjectTimedService> logger, ISysTasksQzService tasksQzService)
        {
            _logger = logger;
            _tasksQzService = tasksQzService;
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
                    await Task.Delay(1 * 60 * 1000);
                }
                model.UpdateTime = DateTime.Now;
                model.Remark = logWrite(executeLog);
                _tasksQzService.Update(model);
            }


        }
        public async Task Run(IJobExecutionContext context)
        {
            string projectExpireRemindTime = AppSettings.Configuration["AppSettings:ProjectExpireRemindTime"];
            if (string.IsNullOrEmpty(projectExpireRemindTime))
            {
                projectExpireRemindTime = "12:00";
            }
            if (DateTime.Now.ToString("HH:mm") == projectExpireRemindTime)
            {
                Execute();
                HasExcute = true;
            }
        }

        private void Execute()
        {
            try
            {
                _logger.LogInformation("执行 项目任务超时提醒 成功.");
            }
            catch (Exception ex)
            {
                _logger.LogError($"执行 项目任务超时提醒 失败.   错误信息 : {ex.ToString()}.");
            }
        }
    }
}
