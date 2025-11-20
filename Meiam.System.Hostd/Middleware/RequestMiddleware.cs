using Microsoft.AspNetCore.Http;
using NLog;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Meiam.System.Hostd.Middleware {
    public class RequestMiddleware {
        private readonly RequestDelegate _next;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RequestMiddleware(RequestDelegate next) {
            _next = next;
        }

        public async Task InvokeAsync(HttpContext context) {
            // 只记录 API 请求
            if (!context.Request.Path.Value.Contains("api", StringComparison.OrdinalIgnoreCase)) {
                await _next(context);
                return;
            }

            // 1. 记录请求信息（Path + Body）
            string requestBody = await ReadRequestBody(context.Request);
            _logger.Info($"Request | Path: {context.Request.Path} | Body: {requestBody}");

            // 2. 捕获响应
            var originalResponseBody = context.Response.Body;
            using var responseStream = new MemoryStream();
            context.Response.Body = responseStream;

            // 3. 执行后续中间件
            await _next(context);

            // 4. 记录响应信息（Status + Body）
            string responseBody = await ReadResponseBody(context.Response);
            var logBody = responseBody?.Length > 200 ? responseBody.Substring(0, 200) + "..." : responseBody;
            _logger.Info($"Response | Status: {context.Response.StatusCode} | Body: {logBody}");

            // 5. 回写响应到客户端
            responseStream.Position = 0;
            await responseStream.CopyToAsync(originalResponseBody);
        }

        private async Task<string> ReadRequestBody(HttpRequest request) {
            request.EnableBuffering(); // 允许重复读取 Body
            using var reader = new StreamReader(request.Body, leaveOpen: true);
            string body = await reader.ReadToEndAsync();
            request.Body.Position = 0; // 重置流位置
            return body;
        }

        private async Task<string> ReadResponseBody(HttpResponse response) {
            response.Body.Seek(0, SeekOrigin.Begin);
            string body = await new StreamReader(response.Body).ReadToEndAsync();
            response.Body.Seek(0, SeekOrigin.Begin); // 重置流位置
            return body;
        }
    }
}