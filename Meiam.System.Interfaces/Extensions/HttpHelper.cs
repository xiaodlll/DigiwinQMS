using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Meiam.System.Interfaces.Extensions
{
    public static class HttpHelper
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        /// <summary>
        /// 发送 POST 请求，提交 JSON 数据
        /// </summary>
        public static async Task<string> PostJsonAsync(string url, object data, string? token = null)
        {
            try
            {
                string json = JsonConvert.SerializeObject(data);
                using var content = new StringContent(json, Encoding.UTF8, "application/json");

                if (!string.IsNullOrEmpty(token))
                {
                    _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                }

                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"POST 请求失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 发送 GET 请求，支持自动拼接参数
        /// </summary>
        public static async Task<string> GetAsync(string url, Dictionary<string, string>? parameters = null, string? token = null)
        {
            try
            {
                if (parameters != null && parameters.Count > 0)
                {
                    var query = new FormUrlEncodedContent(parameters).ReadAsStringAsync().Result;
                    url = $"{url}?{query}";
                }

                if (!string.IsNullOrEmpty(token))
                {
                    _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                }

                var response = await _httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"GET 请求失败: {ex.Message}", ex);
            }
        }
    }
}