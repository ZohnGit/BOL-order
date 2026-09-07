using System.Globalization;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace BolOrderExporter;

public sealed class BolApiClient : IDisposable
{
    private const string ApiMediaType = "application/vnd.retailer.v10+json";
    private readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(60) };

    public async Task<string> LoginAsync(string user, string password, CancellationToken cancellationToken = default)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, "https://login.bol.com/token");
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Authorization = new AuthenticationHeaderValue(
            "Basic",
            Convert.ToBase64String(Encoding.UTF8.GetBytes($"{user}:{password}")));
        request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["grant_type"] = "client_credentials"
        });

        using var document = await SendJsonAsync(request, cancellationToken);
        if (!TryString(document.RootElement, "access_token", out var token) || string.IsNullOrWhiteSpace(token))
        {
            throw new InvalidOperationException("登录响应中没有 access_token。");
        }
        return token;
    }

    public async Task<JsonDocument> GetOrdersAsync(string token, int page, CancellationToken cancellationToken = default)
    {
        var url = "https://api.bol.com/retailer/orders" +
                  $"?fulfilment-method=FBB&status=ALL&page={page.ToString(CultureInfo.InvariantCulture)}";
        return await GetAsync(url, token, cancellationToken);
    }

    public Task<JsonDocument> GetOrderAsync(string token, string orderId, CancellationToken cancellationToken = default) =>
        GetAsync($"https://api.bol.com/retailer/orders/{Uri.EscapeDataString(orderId)}", token, cancellationToken);

    public Task<JsonDocument> GetShipmentsAsync(string token, string orderId, CancellationToken cancellationToken = default) =>
        GetAsync($"https://api.bol.com/retailer/shipments?order-id={Uri.EscapeDataString(orderId)}", token, cancellationToken);

    public Task<JsonDocument> GetShipmentAsync(string token, string shipmentId, CancellationToken cancellationToken = default) =>
        GetAsync($"https://api.bol.com/retailer/shipments/{Uri.EscapeDataString(shipmentId)}", token, cancellationToken);

    private async Task<JsonDocument> GetAsync(string url, string token, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(ApiMediaType));
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        return await SendJsonAsync(request, cancellationToken);
    }

    private async Task<JsonDocument> SendJsonAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.SendAsync(request, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var message = ExtractApiError(body);
            throw new HttpRequestException(
                $"bol.com API 返回 {(int)response.StatusCode} {response.ReasonPhrase}" +
                (string.IsNullOrWhiteSpace(message) ? "。" : $"：{message}"));
        }

        try
        {
            return JsonDocument.Parse(body);
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException("bol.com API 返回了无法解析的数据。", ex);
        }
    }

    private static string? ExtractApiError(string body)
    {
        try
        {
            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;
            foreach (var name in new[] { "detail", "title", "message", "error_description", "error" })
            {
                if (TryString(root, name, out var value) && !string.IsNullOrWhiteSpace(value)) return value;
            }
        }
        catch (JsonException)
        {
            // Do not include an arbitrary HTML response in the UI.
        }
        return null;
    }

    internal static bool TryString(JsonElement element, string name, out string? value)
    {
        value = null;
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var property)) return false;
        value = property.ValueKind switch
        {
            JsonValueKind.String => property.GetString(),
            JsonValueKind.Number => property.GetRawText(),
            _ => null
        };
        return value is not null;
    }

    public void Dispose() => _httpClient.Dispose();
}
