using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Linq;

namespace Meiam.System.Hostd.Middleware {
    public class ApiRequestLoggingMiddleware {
        private readonly RequestDelegate _next;
        private readonly ILogger<ApiRequestLoggingMiddleware> _logger;
        private readonly PathString _apiPathPrefix = new PathString("/api");
        private readonly List<string> _excludedPaths = new List<string>
        {
        "/swagger",
        "/favicon.ico",
        "/health",
        // 添加更多需要排除的路径前缀
    };

        public ApiRequestLoggingMiddleware(RequestDelegate next, ILogger<ApiRequestLoggingMiddleware> logger) {
            _next = next;
            _logger = logger;
        }

        public async Task InvokeAsync(HttpContext context) {
            // 快速判断：只处理API请求并排除特定路径
            if (!IsApiRequest(context.Request.Path) || IsExcludedPath(context.Request.Path)) {
                await _next(context);
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            string requestBody = null;
            string responseBody = null;
            int statusCode = 200;

            try {
                // 保存原始响应流
                var originalResponseBodyStream = context.Response.Body;
                using var responseBodyStream = new MemoryStream();
                context.Response.Body = responseBodyStream;

                // 读取请求体（可重用）
                context.Request.EnableBuffering();
                if (IsReadableContent(context.Request)) {
                    requestBody = await ReadRequestBody(context.Request);
                    // 重置请求流位置，确保后续处理不受影响
                    context.Request.Body.Position = 0;
                }

                // 执行后续中间件管道
                await _next(context);

                // 读取响应体
                statusCode = context.Response.StatusCode;
                if (IsReadableContent(context.Response)) {
                    responseBody = await ReadResponseBody(responseBodyStream);
                    // 将响应写回原始流
                    responseBodyStream.Position = 0;
                    await responseBodyStream.CopyToAsync(originalResponseBodyStream);
                }
            }
            finally {
                stopwatch.Stop();

                // 使用结构化日志记录关键信息
                _logger.LogInformation(
                    "API Request Completed - Method: {HttpMethod}, Path: {Path}, Status: {StatusCode}, Duration: {Duration}ms",
                    context.Request.Method,
                    context.Request.Path,
                    statusCode,
                    stopwatch.ElapsedMilliseconds);

                // 记录详细内容（可根据需要调整日志级别）
                if (_logger.IsEnabled(LogLevel.Debug)) {
                    _logger.LogDebug("Request Body: {RequestBody}", requestBody ?? "No content");
                    _logger.LogDebug("Response Body: {ResponseBody}", responseBody ?? "No content");
                }
            }
        }

        private bool IsApiRequest(PathString path)
            => path.StartsWithSegments(_apiPathPrefix, StringComparison.OrdinalIgnoreCase);

        private bool IsExcludedPath(PathString path)
            => _excludedPaths.Any(p => path.StartsWithSegments(p, StringComparison.OrdinalIgnoreCase));

        private bool IsReadableContent(HttpRequest request)
            => request.Body.CanRead && IsTextBasedContentType(request.ContentType);

        private bool IsReadableContent(HttpResponse response)
            => response.Body.CanRead && IsTextBasedContentType(response.ContentType);

        private bool IsTextBasedContentType(string contentType) {
            if (string.IsNullOrEmpty(contentType)) return false;

            return contentType.Contains("json", StringComparison.OrdinalIgnoreCase) ||
                   contentType.Contains("xml", StringComparison.OrdinalIgnoreCase) ||
                   contentType.Contains("text", StringComparison.OrdinalIgnoreCase) ||
                   contentType.Contains("html", StringComparison.OrdinalIgnoreCase);
        }

        private async Task<string> ReadRequestBody(HttpRequest request) {
            using var reader = new StreamReader(
                request.Body,
                encoding: Encoding.UTF8,
                detectEncodingFromByteOrderMarks: false,
                bufferSize: 8192,
                leaveOpen: true);

            return await reader.ReadToEndAsync();
        }

        private async Task<string> ReadResponseBody(Stream stream) {
            stream.Position = 0;
            using var reader = new StreamReader(stream, Encoding.UTF8);
            return await reader.ReadToEndAsync();
        }
    }
}