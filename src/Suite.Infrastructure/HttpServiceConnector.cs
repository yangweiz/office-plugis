using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Suite.Infrastructure
{
    public sealed class HttpServiceConnector
    {
        private readonly HttpClient _client;
        public HttpServiceConnector(HttpClient client) { _client = client ?? throw new ArgumentNullException("client"); }

        public async Task<string> PostJsonAsync(Uri endpoint, string json, string secret, CancellationToken cancellationToken)
        {
            if (endpoint == null) throw new ArgumentNullException("endpoint");
            if (!endpoint.IsAbsoluteUri || endpoint.Scheme != Uri.UriSchemeHttps) throw new InvalidOperationException("在线服务必须使用有效的 HTTPS 地址。");
            if (json == null) throw new ArgumentNullException("json");
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
            {
                request.Content = content;
                if (!string.IsNullOrEmpty(secret)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", secret);
                using (HttpResponseMessage response = await _client.SendAsync(request, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new InvalidOperationException("服务返回 HTTP " + (int)response.StatusCode + "；未应用任何结果。");
                    return body;
                }
            }
        }
    }
}
