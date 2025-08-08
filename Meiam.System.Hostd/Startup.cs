using Autofac;
using Autofac.Extras.DynamicProxy;
using Hangfire;
using Hangfire.Dashboard;
using Hangfire.SqlServer;
using MapsterMapper;
using Meiam.System.Core;
using Meiam.System.Extensions;
using Meiam.System.Hostd.Extensions;
using Meiam.System.Hostd.Global;
using Meiam.System.Hostd.Middleware;
using Meiam.System.Hostd.Setup;
using Meiam.System.Interfaces;
using Meiam.System.Interfaces.Service;
using Meiam.System.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpOverrides;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SqlSugar;
using System;
using System.Reflection;

namespace Meiam.System.Hostd
{
    public class Startup
    {
        public Startup(IConfiguration configuration, IWebHostEnvironment env)
        {
            Configuration = configuration;
            Env = env;
        }

        public IConfiguration Configuration { get; }
        public IWebHostEnvironment Env { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {

            #region 服务注入
            Console.WriteLine("Enable database linking");
            //启用数据库链接
            services.AddSqlsugarSetup();

            //Console.WriteLine("跨域设置");
            ////跨域设置
            //services.AddCorsSetup();

            Console.WriteLine("Swagger Setup");
            //说明文档
            services.AddSwaggerSetup();

            //注入Cap (EventBus) 按需引入
            //services.AddCapSetup();

            Console.WriteLine("Inject cache");
            //注入缓存
            services.AddMemoryCache();

            Console.WriteLine("Inject HTTPCONTEXT");
            //注入 HTTPCONTEXT
            services.AddHttpContextAccessor();

            //Console.WriteLine("Inject TokenManager");
            ////注入 TokenManager
            //services.AddScoped<TokenManager>();

            Console.WriteLine("Inject entity mapping service");
            //注入实体映射服务
            services.AddScoped<IMapper, ServiceMapper>();

            //Console.WriteLine("注入短信服务");
            ////注入短信服务
            //services.AddSingleton<IAliyunSmsServices, AliyunSmsServices>();

            //Console.WriteLine("Inject email service");
            ////注入邮件服务
            //services.AddSingleton<IEmailSmtpServices, EmailSmtpServices>();

            //Console.WriteLine("Register for REDIS services");
            ////注册REDIS 服务
            //RedisServer.Initalize();

            //Console.WriteLine("Open the scheduled task");
            ////开启计划任务
            //services.AddTaskSchedulers();

            Console.WriteLine("Global Settings");
            #endregion

            #region Hangfire配置
            Console.WriteLine("Configuring Hangfire...");
            services.AddHangfire(configuration => configuration
                .SetDataCompatibilityLevel(CompatibilityLevel.Version_170)
                .UseSimpleAssemblyNameTypeSerializer()
                .UseRecommendedSerializerSettings()
                .UseSqlServerStorage(Configuration["ConnectionStrings:HangfireConnection"], new SqlServerStorageOptions
                {
                    CommandBatchMaxTimeout = TimeSpan.FromMinutes(5),
                    SlidingInvisibilityTimeout = TimeSpan.FromMinutes(5),
                    QueuePollInterval = TimeSpan.Zero,
                    UseRecommendedIsolationLevel = true,
                    DisableGlobalLocks = true
                }));

            services.AddHangfireServer(options =>
            {
                options.ServerName = "HMD_Sync_Server";
                options.Queues = new[] { "sync" };
            });
            #endregion

            #region 全局设置
            //配置Json格式
            services.AddMvc().AddNewtonsoftJson(options =>
            {
                // 忽略循环引用
                options.SerializerSettings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;
                // 不使用驼峰
                //options.SerializerSettings.ContractResolver = new DefaultContractResolver();
                // 设置时间格式
                options.SerializerSettings.DateFormatString = "yyyy-MM-dd HH:mm:ss";
                // 如字段为null值，该字段不会返回到前端
                //options.SerializerSettings.NullValueHandling = NullValueHandling.Ignore;
            });

            //获取客户端 IP
            services.Configure<ForwardedHeadersOptions>(options =>
            {
                options.ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto;
                options.KnownNetworks.Clear();
                options.KnownProxies.Clear();
            });

            //注入全局异常过滤
            services.AddControllers(options =>
            {
                //全局异常过滤
                options.Filters.Add<GlobalExceptions>();
                //全局日志
                options.Filters.Add<GlobalActionMonitor>();

            })
            .ConfigureApiBehaviorOptions(options =>
            {
                //抑制系统自带模型验证
                options.SuppressModelStateInvalidFilter = true;
            });
            Console.WriteLine("Global settings completed!");
            services.AddCors(options =>
            {
                options.AddPolicy("AllowAll", builder =>
                {
                    builder.AllowAnyOrigin()
                           .AllowAnyMethod()
                           .AllowAnyHeader();
                });
            });

            #endregion

            #region 同步服务注册
            Console.WriteLine("Registering sync services...");
            services.AddScoped<IHMDService, HMDService>();
            #endregion

            #region SqlSugarHMD 配置
            Console.WriteLine("Configuring Oracle SqlSugar...");
            services.AddSingleton<ISqlSugarClient>(provider =>
            {
                var config = new ConnectionConfig()
                {
                    ConfigId = "OracleDB",
                    ConnectionString = Configuration.GetConnectionString("OracleConnection"),
                    DbType = DbType.Oracle,
                    IsAutoCloseConnection = true,
                    InitKeyType = InitKeyType.Attribute,
                    ConfigureExternalServices = new ConfigureExternalServices
                    {
                        EntityService = (property, column) =>
                        {
                            if (!column.IsIgnore && !string.IsNullOrEmpty(column.DbColumnName))
                            {
                                column.DbColumnName = column.DbColumnName.ToUpper(); // Oracle 列名大写
                            }
                        }
                    },
                    MoreSettings = new ConnMoreSettings
                    {
                        IsAutoRemoveDataCache = true
                        //OracleConnectionStringPull = true // 解决连接池问题
                    }
                };

                return new SqlSugarScope(config, db =>
                {
                    // 设置Oracle日期格式
                    db.Ado.ExecuteCommand("ALTER SESSION SET NLS_DATE_FORMAT='YYYY-MM-DD HH24:MI:SS'");
                    db.Ado.ExecuteCommand("ALTER SESSION SET NLS_TIMESTAMP_FORMAT='YYYY-MM-DD HH24:MI:SS'");

                    // AOP配置
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
            });
            #endregion

        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env, ISysTasksQzService tasksQzService,
            IRecurringJobManager recurringJobManager,
            IHMDService syncService)
        {

            #region 开发错误提示

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            };
            #endregion

            // 加上这句！
            app.UseCors("AllowAll");

            #region Hangfire中间件
            app.UseHangfireDashboard("/hangfire", new DashboardOptions
            {
                DashboardTitle = "数据同步监控",
                Authorization = new[] { new HangfireAuthorizationFilter() }
            });
            #endregion

            #region 初始化同步任务
            Console.WriteLine("Initializing sync jobs...");
            InitializeSyncJobs(recurringJobManager, syncService);
            #endregion

            #region 服务注入
            //// 跨域设置
            //app.UseCorsSetup();
            // 说明文档
            app.UseSwaggerSetup();
            // 请求日志监控
            app.UseMiddleware<RequestMiddleware>();
            // 使用静态文件
            app.UseForwardedHeaders();
            // 使用静态文件
            app.UseStaticFiles();
            // 使用cookie
            app.UseCookiePolicy();
            // 使用Routing
            app.UseRouting();
            // 使用相应缓存中间件
            app.UseResponseCaching();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
                endpoints.MapHangfireDashboard();
            });

            #endregion
        }

        #region 自动注入服务
        public void ConfigureContainer(ContainerBuilder builder)
        {
            var assemblysServices = Assembly.Load("Meiam.System.Interfaces");
            builder.RegisterAssemblyTypes(assemblysServices)
                .InstancePerDependency()//瞬时单例
               .AsImplementedInterfaces()////自动以其实现的所有接口类型暴露（包括IDisposable接口）
               .EnableInterfaceInterceptors(); //引用Autofac.Extras.DynamicProxy;
        }
        #endregion

        #region 同步任务初始化
        private void InitializeSyncJobs(IRecurringJobManager recurringJobManager, IHMDService syncService)
        {
            try
            {
                // 收货单同步 - 每30分钟执行一次
                recurringJobManager.AddOrUpdate("RC_Sync",
                    () => syncService.SyncRcDataAsync(syncService.GetLastSyncTime("INSPECT_IQC", "INSPECT_FPICREATEDATE")),
                    Configuration["SyncConfig:RC:Cron"] ?? "*/30 * * * *");

                // 报工单同步 - 每小时执行一次
                recurringJobManager.AddOrUpdate("WR_Sync",
                    () => syncService.SyncWrDataAsync(syncService.GetLastSyncTime("INSPECT_SI", "INSPECT_FPICREATEDATE")),
                    Configuration["SyncConfig:WR:Cron"] ?? "0 * * * *");

                // 物料同步 - 每天凌晨2点执行
                recurringJobManager.AddOrUpdate("ITEM_Sync",
                    () => syncService.SyncItemDataAsync(syncService.GetLastSyncTime("ITEM", "INSPECT_FPICREATEDATE")),
                    Configuration["SyncConfig:ITEM:Cron"] ?? "0 2 * * *");

                // 供应商同步 - 每天凌晨3点执行
                recurringJobManager.AddOrUpdate("VEND_Sync",
                    () => syncService.SyncVendDataAsync(syncService.GetLastSyncTime("SUPP", "INSPECT_FPICREATEDATE")),
                    Configuration["SyncConfig:VEND:Cron"] ?? "0 3 * * *");

                // 客户同步 - 每天凌晨4点执行
                recurringJobManager.AddOrUpdate("CUST_Sync",
                    () => syncService.SyncCustDataAsync(syncService.GetLastSyncTime("CUSTOM", "INSPECT_FPICREATEDATE")),
                    Configuration["SyncConfig:CUST:Cron"] ?? "0 4 * * *");

                Console.WriteLine("Sync jobs initialized successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing sync jobs: {ex.Message}");
                throw;
            }
        }
        #endregion
    }

    #region Hangfire授权过滤器
    public class HangfireAuthorizationFilter : IDashboardAuthorizationFilter
    {
        public bool Authorize(DashboardContext context)
        {
            // 生产环境应实现实际授权逻辑
            return true;
        }
    }
    #endregion
}