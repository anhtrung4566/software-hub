param(
    [int]$Port = 8080,
    [switch]$BackgroundTestMode
)

$ErrorActionPreference = "Stop"
$RootDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$csharpServer = @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

public class LocalHubResponse {
    public int StatusCode;
    public string StatusText;
    public string ContentType;
    public Dictionary<string, string> Headers;
    public byte[] BodyBytes;
    public string BodyText {
        get { return BodyBytes != null ? Encoding.UTF8.GetString(BodyBytes) : ""; }
    }
}

public class LocalHubServer {
    private TcpListener _listener;
    private bool _running;
    private string _rootDir;
    public int Port { get; private set; }

    public LocalHubServer(string rootDir, int port) {
        _rootDir = Path.GetFullPath(rootDir);
        try {
            _listener = new TcpListener(IPAddress.Loopback, port);
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            _running = true;
            ThreadPool.QueueUserWorkItem(_ => AcceptLoop());
        } catch {
            Port = port > 0 ? port : 8080;
            _running = false;
        }
    }

    public void Stop() {
        _running = false;
        try { if (_listener != null) _listener.Stop(); } catch {}
    }

    public LocalHubResponse DispatchRequest(string method, string rawUrl, string bodyText) {
        method = (method ?? "GET").ToUpperInvariant();
        rawUrl = rawUrl ?? "/";
        bodyText = bodyText ?? "";

        Dictionary<string, string> headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        headers["Access-Control-Allow-Origin"] = "*";
        headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS";
        headers["Access-Control-Allow-Headers"] = "Content-Type";
        headers["X-Content-Type-Options"] = "nosniff";
        headers["X-Frame-Options"] = "SAMEORIGIN";

        if (Encoding.UTF8.GetByteCount(bodyText) > 2 * 1024 * 1024) {
            return MakeResp(413, "Payload Too Large", "application/json; charset=utf-8", headers, Encoding.UTF8.GetBytes("{\"status\":\"error\",\"message\":\"Payload too large\"}"));
        }

        string pathOnly = rawUrl.Split('?')[0];
        string queryOnly = rawUrl.Contains("?") ? rawUrl.Substring(rawUrl.IndexOf('?') + 1) : "";
        string decodedPath = Uri.UnescapeDataString(pathOnly);

        if (method == "OPTIONS") {
            return MakeResp(204, "No Content", "text/plain", headers, new byte[0]);
        }

        if (decodedPath == "/api/health") {
            string json = "{\"status\":\"ok\",\"service\":\"SieuAppVanPhong-Backend\",\"timestamp\":\"" + DateTime.UtcNow.ToString("o") + "\"}";
            return MakeResp(200, "OK", "application/json; charset=utf-8", headers, Encoding.UTF8.GetBytes(json));
        }

        if (decodedPath == "/api/stats" || decodedPath == "/api/stats.php") {
            string statsFile = Path.Combine(_rootDir, "data", "stats.json");
            int visits = 6150;
            int downloads = 2180;
            if (File.Exists(statsFile)) {
                string s = File.ReadAllText(statsFile, Encoding.UTF8);
                visits = ExtractInt(s, "total_visits", 6150);
                downloads = ExtractInt(s, "total_downloads", 2180);
            }
            if (queryOnly.Contains("action=download") || method == "POST") {
                downloads++;
            } else {
                visits++;
            }
            string outJson = "{\n  \"status\": \"ok\",\n  \"date\": \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",\n  \"month\": \"2026-11\",\n  \"month_label\": \"Tháng 11\",\n  \"today_visits\": " + (visits - 5790) + ",\n  \"today_downloads\": " + (downloads - 2030) + ",\n  \"month_visits\": " + (visits - 1300) + ",\n  \"month_downloads\": " + (downloads - 500) + ",\n  \"total_visits\": " + visits + ",\n  \"total_downloads\": " + downloads + "\n}";
            try { File.WriteAllText(statsFile, outJson, new UTF8Encoding(false)); } catch {}
            return MakeResp(200, "OK", "application/json; charset=utf-8", headers, Encoding.UTF8.GetBytes(outJson));
        }

        if (decodedPath == "/api/products") {
            string prodFile = Path.Combine(_rootDir, "data", "products.json");
            if (method == "GET") {
                byte[] bytes = File.Exists(prodFile) ? File.ReadAllBytes(prodFile) : Encoding.UTF8.GetBytes("[]");
                return MakeResp(200, "OK", "application/json; charset=utf-8", headers, bytes);
            }
            if (method == "POST") {
                string trimmed = bodyText.Trim();
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]")) {
                    File.WriteAllText(prodFile, trimmed, new UTF8Encoding(false));
                    return MakeResp(200, "OK", "application/json; charset=utf-8", headers, Encoding.UTF8.GetBytes("{\"status\":\"ok\"}"));
                } else {
                    return MakeResp(400, "Bad Request", "application/json; charset=utf-8", headers, Encoding.UTF8.GetBytes("{\"status\":\"error\",\"message\":\"Invalid JSON array\"}"));
                }
            }
        }

        if (pathOnly.Contains("..") || decodedPath.Contains("..") || decodedPath.Contains("\0")) {
            return MakeResp(403, "Forbidden", "text/plain; charset=utf-8", headers, Encoding.UTF8.GetBytes("403 Forbidden"));
        }

        string rel = decodedPath == "/" ? "index.html" : decodedPath.TrimStart('/');
        string fullPath = Path.GetFullPath(Path.Combine(_rootDir, rel.Replace('/', Path.DirectorySeparatorChar)));
        if (!fullPath.StartsWith(_rootDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) && !string.Equals(fullPath, _rootDir, StringComparison.OrdinalIgnoreCase)) {
            return MakeResp(403, "Forbidden", "text/plain; charset=utf-8", headers, Encoding.UTF8.GetBytes("403 Forbidden"));
        }
        if (Directory.Exists(fullPath)) {
            fullPath = Path.Combine(fullPath, "index.html");
        }
        if (!File.Exists(fullPath)) {
            return MakeResp(404, "Not Found", "text/plain; charset=utf-8", headers, Encoding.UTF8.GetBytes("404 Not Found"));
        }

        string ext = Path.GetExtension(fullPath).ToLowerInvariant();
        string mime = "application/octet-stream";
        if (ext == ".html") mime = "text/html; charset=utf-8";
        else if (ext == ".css") mime = "text/css; charset=utf-8";
        else if (ext == ".js") mime = "application/javascript; charset=utf-8";
        else if (ext == ".json") mime = "application/json; charset=utf-8";
        else if (ext == ".svg") mime = "image/svg+xml";
        else if (ext == ".png") mime = "image/png";
        else if (ext == ".md") mime = "text/markdown; charset=utf-8";

        byte[] fileBytes = File.ReadAllBytes(fullPath);
        return MakeResp(200, "OK", mime, headers, fileBytes);
    }

    private LocalHubResponse MakeResp(int code, string status, string contentType, Dictionary<string, string> headers, byte[] body) {
        LocalHubResponse r = new LocalHubResponse();
        r.StatusCode = code;
        r.StatusText = status;
        r.ContentType = contentType;
        r.Headers = headers;
        r.BodyBytes = body;
        return r;
    }

    private void AcceptLoop() {
        while (_running) {
            try {
                TcpClient client = _listener.AcceptTcpClient();
                ThreadPool.QueueUserWorkItem(_ => HandleClient(client));
            } catch {
                if (!_running) break;
            }
        }
    }

    private void HandleClient(TcpClient client) {
        using (client)
        using (NetworkStream stream = client.GetStream()) {
            stream.ReadTimeout = 5000;
            stream.WriteTimeout = 5000;
            byte[] buffer = new byte[65536];
            int bytesRead = 0;
            try {
                bytesRead = stream.Read(buffer, 0, buffer.Length);
            } catch { return; }
            if (bytesRead <= 0) return;

            string rawReq = Encoding.UTF8.GetString(buffer, 0, bytesRead);
            int headerEnd = rawReq.IndexOf("\r\n\r\n");
            if (headerEnd < 0) return;

            string headerPart = rawReq.Substring(0, headerEnd);
            string[] lines = headerPart.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            if (lines.Length == 0) return;

            string[] reqLine = lines[0].Split(' ');
            if (reqLine.Length < 2) return;
            string method = reqLine[0].ToUpperInvariant();
            string rawUrl = reqLine[1];

            int contentLength = 0;
            foreach (string l in lines) {
                if (l.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)) {
                    int.TryParse(l.Substring(15).Trim(), out contentLength);
                }
            }

            int bodyStartByte = Encoding.UTF8.GetByteCount(headerPart + "\r\n\r\n");
            int bodyBytesReceived = bytesRead - bodyStartByte;
            MemoryStream bodyMs = new MemoryStream();
            if (bodyBytesReceived > 0) {
                bodyMs.Write(buffer, bodyStartByte, bodyBytesReceived);
            }
            while (bodyMs.Length < contentLength) {
                int r = stream.Read(buffer, 0, buffer.Length);
                if (r <= 0) break;
                bodyMs.Write(buffer, 0, r);
            }
            string bodyText = Encoding.UTF8.GetString(bodyMs.ToArray());

            LocalHubResponse resp = DispatchRequest(method, rawUrl, bodyText);
            SendResponse(stream, resp);
        }
    }

    private int ExtractInt(string json, string key, int defVal) {
        int idx = json.IndexOf("\"" + key + "\"");
        if (idx < 0) return defVal;
        int colon = json.IndexOf(':', idx);
        if (colon < 0) return defVal;
        int end = colon + 1;
        while (end < json.Length && (char.IsWhiteSpace(json[end]))) end++;
        int start = end;
        while (end < json.Length && char.IsDigit(json[end])) end++;
        int val;
        if (int.TryParse(json.Substring(start, end - start), out val)) return val;
        return defVal;
    }

    private void SendResponse(NetworkStream stream, LocalHubResponse resp) {
        StringBuilder sb = new StringBuilder();
        sb.Append("HTTP/1.1 ").Append(resp.StatusCode).Append(" ").Append(resp.StatusText).Append("\r\n");
        sb.Append("Content-Type: ").Append(resp.ContentType).Append("\r\n");
        sb.Append("Content-Length: ").Append(resp.BodyBytes.Length).Append("\r\n");
        foreach (KeyValuePair<string, string> kv in resp.Headers) {
            sb.Append(kv.Key).Append(": ").Append(kv.Value).Append("\r\n");
        }
        sb.Append("Connection: close\r\n\r\n");
        byte[] headerBytes = Encoding.UTF8.GetBytes(sb.ToString());
        stream.Write(headerBytes, 0, headerBytes.Length);
        if (resp.BodyBytes.Length > 0) {
            stream.Write(resp.BodyBytes, 0, resp.BodyBytes.Length);
        }
        stream.Flush();
    }
}
"@

if (-not ("LocalHubServer" -as [type])) {
    Add-Type -TypeDefinition $csharpServer -Language CSharp
}

if (-not $BackgroundTestMode) {
    $srv = New-Object LocalHubServer($RootDir, $Port)
    Write-Host "=================================================================" -ForegroundColor Yellow
    Write-Host " SIÊU APP VĂN PHÒNG (ĐẢNG ỦY & UBND PHƯỜNG) - LOCALHOST SERVER" -ForegroundColor Red
    Write-Host " Đang chạy tại: http://localhost:$($srv.Port)/" -ForegroundColor Green
    Write-Host " Trang Quản Trị: http://localhost:$($srv.Port)/admin.html" -ForegroundColor Cyan
    Write-Host " Nhấn Ctrl+C để dừng máy chủ." -ForegroundColor Gray
    Write-Host "=================================================================" -ForegroundColor Yellow
    try {
        while ($true) { Start-Sleep -Seconds 1 }
    } finally {
        $srv.Stop()
    }
}
