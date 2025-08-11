using Hangfire;
using Meiam.System.Common;
using Meiam.System.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using SqlSugar;
using System;

namespace Meiam.System.Hostd.Setup
{
    public static class SetupSqlSugar
    {
        public static void AddSqlsugarSetup(this IServiceCollection services)
        {
            if (services == null) throw new ArgumentNullException(nameof(services));
            services.AddScoped<ISqlSugarClient>(x =>
            {

                return new SqlSugarClient(new ConnectionConfig()
                {
                    ConnectionString = AppSettings.Configuration["DbConnection:ConnectionString"],
                    DbType = (DbType)Convert.ToInt32(AppSettings.Configuration["DbConnection:DbType"]),                  
                    IsAutoCloseConnection = true,
                    InitKeyType = InitKeyType.Attribute,
                    ConfigureExternalServices = new ConfigureExternalServices()
                    {
                        //DataInfoCacheService = new RedisCache()
                    },
                    MoreSettings = new ConnMoreSettings()
                    {
                        IsAutoRemoveDataCache = true
                    }
                });
            });

            #region SqlSugarHMD 配置
            string CustomID = AppSettings.Configuration["AppSettings:CustomID"];
            if (CustomID == "HMD") {
                services.AddSingleton<IOracleSqlSugarClient>(provider =>
                {
                    var config = new ConnectionConfig() {
                        ConfigId = "OracleDB",
                        ConnectionString = AppSettings.Configuration["ConnectionStrings:OracleConnection"],
                        DbType = DbType.Oracle,
                        IsAutoCloseConnection = true,
                        InitKeyType = InitKeyType.Attribute,
                        ConfigureExternalServices = new ConfigureExternalServices {
                            EntityService = (property, column) =>
                            {
                                if (!column.IsIgnore && !string.IsNullOrEmpty(column.DbColumnName)) {
                                    column.DbColumnName = column.DbColumnName.ToUpper(); // 列名转大写
                                }
                            }
                        },
                        MoreSettings = new ConnMoreSettings {
                            IsAutoRemoveDataCache = true
                            //OracleConnectionStringPull = true // 启用连接池优化
                        }
                    };

                    var oracleDb = new OracleSqlSugarClient(config, db =>
                    {
                        // AOP日志配置
                        db.Aop.OnLogExecuting = (sql, pars) =>
                        {
                            var logger = provider.GetRequiredService<ILogger<SqlSugarScope>>();
                            logger.LogDebug($"Oracle SQL: {UtilMethods.GetSqlString(DbType.Oracle, sql, pars)}");
                        };

                        db.Aop.OnError = ex =>
                        {
                            var logger = provider.GetRequiredService<ILogger<SqlSugarScope>>();
                            logger.LogError(ex, "Oracle数据库错误");
                        };
                    });

                    return oracleDb;
                });
            }
            #endregion
        }
    }
}
