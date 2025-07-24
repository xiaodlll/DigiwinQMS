using Microsoft.Extensions.Logging;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meiam.System.Interfaces.Service
{
    public class DataSyncService
    {
        private readonly ILogger<DataSyncService> _logger;
        private readonly ISqlSugarClient _sqlSugar;

        public DataSyncService(ILogger<DataSyncService> logger, ISqlSugarClient sqlSugar)
        {
            _logger = logger;
            _sqlSugar = sqlSugar;
        }

        public async Task SyncDataFromOracleToSqlServer()
        {
            try
            {
                _logger.LogInformation("开始从Oracle视图同步数据到SQL Server...");

                // 1. 从Oracle视图读取数据
                var oracleData = await _sqlSugar.Queryable<dynamic>()
                    .AS("V_YOUR_VIEW_NAME")  // 指定视图名称
                    .ToListAsync();

                // 2. 连接到SQL Server数据库
                using (var sqlServerDb = new SqlSugarScope(new ConnectionConfig()
                {
                    ConnectionString = "Your_SQLServer_ConnectionString",
                    DbType = DbType.SqlServer,
                    IsAutoCloseConnection = true
                }))
                {
                    // 3. 将数据插入SQL Server表
                    // 假设目标表名为 TargetTable
                    await sqlServerDb.Insertable(oracleData).AS("TargetTable").ExecuteCommandAsync();
                }

                _logger.LogInformation($"成功同步 {oracleData.Count} 条数据到SQL Server");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "同步数据时发生错误");
                throw; // Hangfire 会记录失败的任务
            }
        }
    }
}
