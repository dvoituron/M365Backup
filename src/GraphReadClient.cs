using System.Net;
using System.Runtime.CompilerServices;
using System.Text.Json;
using Azure.Core;

namespace M365Backup;

internal sealed class GraphReadClient(HttpClient httpClient, TokenCredential credential, string[] scopes)
{
    private const string GraphHost = "graph.microsoft.com";
    private static readonly Uri GraphBaseUri = new("https://graph.microsoft.com/v1.0/");

    public async IAsyncEnumerable<JsonElement> GetCollectionAsync(
        string url,
        [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        string? nextUrl = url;
        while (nextUrl is not null)
        {
            using var response = await SendGetAsync(nextUrl, HttpCompletionOption.ResponseContentRead, cancellationToken);
            await EnsureSuccessAsync(response, cancellationToken);

            using var document = await JsonDocument.ParseAsync(
                await response.Content.ReadAsStreamAsync(cancellationToken),
                cancellationToken: cancellationToken);
            var root = document.RootElement;
            if (!root.TryGetProperty("value", out var values) || values.ValueKind != JsonValueKind.Array)
            {
                throw new InvalidOperationException("Microsoft Graph returned an unexpected collection.");
            }

            foreach (var item in values.EnumerateArray())
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return item.Clone();
            }

            nextUrl = root.TryGetProperty("@odata.nextLink", out var nextLink)
                ? nextLink.GetString()
                : null;
        }
    }

    public async Task DownloadToFileAsync(string url, string destination, CancellationToken cancellationToken)
    {
        using var response = await SendGetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        await EnsureSuccessAsync(response, cancellationToken);

        await using var source = await response.Content.ReadAsStreamAsync(cancellationToken);
        await using var target = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        await source.CopyToAsync(target, cancellationToken);
    }

    private async Task<HttpResponseMessage> SendGetAsync(
        string url,
        HttpCompletionOption completionOption,
        CancellationToken cancellationToken)
    {
        var requestUri = Uri.TryCreate(url, UriKind.Absolute, out var absoluteUri)
            ? absoluteUri
            : new Uri(GraphBaseUri, url);
        if (requestUri.Scheme != Uri.UriSchemeHttps ||
            !string.Equals(requestUri.Host, GraphHost, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Microsoft Graph URL is not allowed.");
        }

        var token = await credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken);
        using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
        request.Headers.Accept.ParseAdd(url.EndsWith("/$value", StringComparison.Ordinal)
            ? "message/rfc822"
            : "application/json");
        if (!url.EndsWith("/$value", StringComparison.Ordinal))
        {
            request.Headers.TryAddWithoutValidation("Prefer", "outlook.timezone=\"UTC\"");
        }

        return await httpClient.SendAsync(request, completionOption, cancellationToken);
    }

    private static async Task EnsureSuccessAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        if (response.IsSuccessStatusCode)
        {
            return;
        }

        if (response.StatusCode == HttpStatusCode.TooManyRequests)
        {
            var retryAfter = response.Headers.RetryAfter?.ToString() ?? "unknown";
            throw new HttpRequestException(
                $"Microsoft Graph is throttling requests (429). Retry-After: {retryAfter}.");
        }

        var detail = await response.Content.ReadAsStringAsync(cancellationToken);
        throw new HttpRequestException(
            $"Microsoft Graph returned {(int)response.StatusCode} ({response.ReasonPhrase}): {detail}");
    }
}
