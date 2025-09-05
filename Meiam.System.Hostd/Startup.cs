using Autofac;
using Autofac.Extras.DynamicProxy;
using Hangfire;
using Hangfire.Dashboard;
using Hangfire.SqlServer;
using MapsterMapper;
using Meiam.System.Common;
using Meiam.System.Core;
using Meiam.System.Extensions;
using Meiam.System.Hostd.Extensions;
using Meiam.System.Hostd.Global;
using Meiam.System.Hostd.Middleware;
using Meiam.System.Hostd.Setup;
using Meiam.System.Interfaces;
using Meiam.System.Interfaces.IService;
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

            #region ����ע��
            Console.WriteLine("Enable database linking");
            //�������ݿ�����
            services.AddSqlsugarSetup();

            //Console.WriteLine("��������");
            ////��������
            //services.AddCorsSetup();

            Console.WriteLine("Swagger Setup");
            //˵���ĵ�
            services.AddSwaggerSetup();

            //ע��Cap (EventBus) ��������
            //services.AddCapSetup();

            Console.WriteLine("Inject cache");
            //ע�뻺��
            services.AddMemoryCache();

            Console.WriteLine("Inject HTTPCONTEXT");
            //ע�� HTTPCONTEXT
            services.AddHttpContextAccessor();

            //Console.WriteLine("Inject TokenManager");
            ////ע�� TokenManager
            //services.AddScoped<TokenManager>();

            Console.WriteLine("Inject entity mapping service");
            //ע��ʵ��ӳ�����
            services.AddScoped<IMapper, ServiceMapper>();

            //Console.WriteLine("ע����ŷ���");
            ////ע����ŷ���
            //services.AddSingleton<IAliyunSmsServices, AliyunSmsServices>();

            //Console.WriteLine("Inject email service");
            ////ע���ʼ�����
            //services.AddSingleton<IEmailSmtpServices, EmailSmtpServices>();

            //Console.WriteLine("Register for REDIS services");
            ////ע��REDIS ����
            //RedisServer.Initalize();

            //Console.WriteLine("Open the scheduled task");
            ////�����ƻ�����
            //services.AddTaskSchedulers();

            Console.WriteLine("Global Settings");
            #endregion

            #region Hangfire����
            Console.WriteLine("Configuring Hangfire...");
            services.AddHangfire(configuration => configuration
                .SetDataCompatibilityLevel(CompatibilityLevel.Version_170)
                .UseSimpleAssemblyNameTypeSerializer()
                .UseRecommendedSerializerSettings()
                // ���� Newtonsoft.Json ��Ϊ���л���
                .UseSerializerSettings(new JsonSerializerSettings {
                    TypeNameHandling = TypeNameHandling.All, // �ؼ�������������Ϣ
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                })
                .UseSqlServerStorage(Configuration["ConnectionStrings:HangfireConnection"], new SqlServerStorageOptions
                {
                    CommandBatchMaxTimeout = TimeSpan.FromMinutes(5),
                    SlidingInvisibilityTimeout = TimeSpan.FromMinutes(5),
                    QueuePollInterval = TimeSpan.Zero,
                    UseRecommendedIsolationLevel = true,
                    DisableGlobalLocks = true
                }));

            services.AddHangfireServer(options => {
                options.ServerName = "HMD_Sync_Server";
                options.WorkerCount = 50; // ���Ӳ����߳���
                options.Queues = new[] { "default", "sync" }; // �����Ķ���
                options.WorkerCount = 10; // ���������߳���
            });
            #endregion

            #region ȫ������
            //����Json��ʽ
            services.AddMvc().AddNewtonsoftJson(options =>
            {
                // ����ѭ������
                options.SerializerSettings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;
                // ��ʹ���շ�
                //options.SerializerSettings.ContractResolver = new DefaultContractResolver();
                // ����ʱ���ʽ
                options.SerializerSettings.DateFormatString = "yyyy-MM-dd HH:mm:ss";
                // ���ֶ�Ϊnullֵ�����ֶβ��᷵�ص�ǰ��
                //options.SerializerSettings.NullValueHandling = NullValueHandling.Ignore;
            });

            //��ȡ�ͻ��� IP
            services.Configure<ForwardedHeadersOptions>(options =>
            {
                options.ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto;
                options.KnownNetworks.Clear();
                options.KnownProxies.Clear();
            });

            //ע��ȫ���쳣����
            services.AddControllers(options =>
            {
                //ȫ���쳣����
                options.Filters.Add<GlobalExceptions>();
                //ȫ����־
                options.Filters.Add<GlobalActionMonitor>();

            })
            .ConfigureApiBehaviorOptions(options =>
            {
                //����ϵͳ�Դ�ģ����֤
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

            #region ͬ������ע��
            Console.WriteLine("Registering sync services...");
            services.AddScoped<IHMDService, HMDService>();
            #endregion

            Console.WriteLine("ע�����Ʒ���");
            // ע�����Ʒ���
            services.AddHttpClient<IYNKService, YNKService>();
            services.AddScoped<IYNKService, YNKService>();

        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env, ISysTasksQzService tasksQzService,
            IRecurringJobManager recurringJobManager,
            IHMDService syncService)
        {

            #region ����������ʾ

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            };
            #endregion

            // ������䣡
            app.UseCors("AllowAll");

            #region Hangfire�м��
            app.UseHangfireDashboard("/hangfire", new DashboardOptions
            {
                DashboardTitle = "����ͬ�����",
                Authorization = new[] { new HangfireAuthorizationFilter() }
            });
            #endregion

            #region ��ʼ��ͬ������
            Console.WriteLine("Initializing sync jobs...");
            string CustomID = AppSettings.Configuration["AppSettings:CustomID"];
            if (CustomID == "HMD") {
                InitializeSyncHMDJobs(recurringJobManager, syncService);
            }
            #endregion

            #region ����ע��
            //// ��������
            //app.UseCorsSetup();
            // ˵���ĵ�
            app.UseSwaggerSetup();
            // ������־���
            app.UseMiddleware<RequestMiddleware>();
            // ʹ�þ�̬�ļ�
            app.UseForwardedHeaders();
            // ʹ�þ�̬�ļ�
            app.UseStaticFiles();
            // ʹ��cookie
            app.UseCookiePolicy();
            // ʹ��Routing
            app.UseRouting();
            // ʹ����Ӧ�����м��
            app.UseResponseCaching();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
                endpoints.MapHangfireDashboard();
            });

            #endregion
        }

        #region �Զ�ע�����
        public void ConfigureContainer(ContainerBuilder builder)
        {
            var assemblysServices = Assembly.Load("Meiam.System.Interfaces");
            builder.RegisterAssemblyTypes(assemblysServices)
                .InstancePerDependency()//˲ʱ����
               .AsImplementedInterfaces()////�Զ�����ʵ�ֵ����нӿ����ͱ�¶������IDisposable�ӿڣ�
               //.EnableInterfaceInterceptors(); //����Autofac.Extras.DynamicProxy;
             .InstancePerLifetimeScope();
        }
        #endregion

        #region ͬ�������ʼ��
        private void InitializeSyncHMDJobs(IRecurringJobManager recurringJobManager, IHMDService syncService)
        {
            try
            {
                // �ջ���ͬ��
                recurringJobManager.AddOrUpdate("RC_Sync",
                    () => syncService.SyncRcDataAsync(),
                    Configuration["SyncConfig:RC"] ?? "* * * * *");

                // ������ͬ��
                recurringJobManager.AddOrUpdate("WR_Sync",
                    () => syncService.SyncWrDataAsync(),
                    Configuration["SyncConfig:WR"] ?? "* * * * *");

                // ����ͬ��
                recurringJobManager.AddOrUpdate("ITEM_Sync",
                    () => syncService.SyncItemDataAsync(),
                    Configuration["SyncConfig:ITEM"] ?? "* * * * *");

                // ��Ӧ��ͬ��
                recurringJobManager.AddOrUpdate("VEND_Sync",
                    () => syncService.SyncVendDataAsync(),
                    Configuration["SyncConfig:VEND"] ?? "* * * * *");

                // �ͻ�ͬ��
                recurringJobManager.AddOrUpdate("CUST_Sync",
                    () => syncService.SyncCustDataAsync(),
                    Configuration["SyncConfig:CUST"] ?? "* * * * *");

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

    #region Hangfire��Ȩ������
    public class HangfireAuthorizationFilter : IDashboardAuthorizationFilter
    {
        public bool Authorize(DashboardContext context)
        {
            // ��������Ӧʵ��ʵ����Ȩ�߼�
            return true;
        }
    }
    #endregion
}