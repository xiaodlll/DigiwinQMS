using Meiam.System.Common;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.OpenApi.Models;
using System;
using System.IO;

namespace Meiam.System.Hostd.Setup
{
    public static class SetupSwagger
    {
        public static void AddSwaggerSetup(this IServiceCollection services)
        {
            if (services == null) throw new ArgumentNullException(nameof(services));

            services.AddSwaggerGen(c =>
            {

                c.SwaggerDoc("v1", new OpenApiInfo
                {
                    Version = "v1",
                    Title = $"{AppSettings.Configuration["Startup:ApiName"]} 接口文档",
                    Description = $"{AppSettings.Configuration["Startup:ApiName"]} HTTP API "
                });

                try
                {
                    //就是这里
                    var xmlPath = Path.Combine(AppContext.BaseDirectory, "Meiam.System.Hostd.xml");//这个就是刚刚配置的xml文件名
                    c.IncludeXmlComments(xmlPath, true);//默认的第二个参数是false，这个是controller的注释，记得修改

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Xml 文件丢失，请检查并拷贝。\n{ex.Message}");
                }

                //// 开启加权小锁
                //c.OperationFilter<AppendAuthorizeFilter>();

            });
        }

        public static void UseSwaggerSetup(this IApplicationBuilder app) {// 获取环境信息
            string runtimeDirectory = AppContext.BaseDirectory;

            // 拼接 swagger-ui 目录路径（运行时目录下的 swagger-ui 文件夹）
            string swaggerUiPath = Path.Combine(runtimeDirectory, "swagger-ui");

            // 配置静态文件访问（仅当目录存在时）
            if (Directory.Exists(swaggerUiPath)) {
                app.UseStaticFiles(new StaticFileOptions {
                    FileProvider = new PhysicalFileProvider(swaggerUiPath),
                    RequestPath = "/swagger-ui" // 保持访问路径不变
                });
            }
            app.UseSwagger();
            app.UseSwaggerUI(c =>
            {
                var ApiName = AppSettings.Configuration["Startup:ApiName"];

                c.SwaggerEndpoint("/swagger/v1/swagger.json", "API v1");
                // 添加自定义CSS
                c.InjectStylesheet("/swagger-ui/custom.css");
                // 添加自定义JavaScript
                c.InjectJavascript("/swagger-ui/custom.js");
                c.RoutePrefix = string.Empty;
            });
        }
    }
}
