using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace STB2026.Shared.Protocol
{
    // ═══════════════════════════════════════════════════════════════
    //  Запрос MCP → Revit Bridge
    // ═══════════════════════════════════════════════════════════════
    
    /// <summary>
    /// Запрос от McpServer к RevitBridge через Named Pipe.
    /// ВАЖНО: имена свойств PascalCase — обе стороны должны 
    /// сериализовать без переименования (camelCase запрещён!).
    /// </summary>
    public class BridgeRequest
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Method { get; set; }
        public Dictionary<string, object> Params { get; set; } = new Dictionary<string, object>();
        public long Timestamp { get; set; } = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
    }

    // ═══════════════════════════════════════════════════════════════
    //  Ответ Revit Bridge → MCP
    // ═══════════════════════════════════════════════════════════════

    public class BridgeResponse
    {
        public string Id { get; set; }
        public bool Success { get; set; }
        public object Data { get; set; }
        public string Error { get; set; }
        public long DurationMs { get; set; }

        public static BridgeResponse Ok(string id, object data, long ms = 0)
            => new BridgeResponse { Id = id, Success = true, Data = data, DurationMs = ms };

        public static BridgeResponse Fail(string id, string error)
            => new BridgeResponse { Id = id, Success = false, Error = error };
    }

    // ═══════════════════════════════════════════════════════════════
    //  Length-prefixed JSON framing для Named Pipe
    //  Формат: [4 байта uint32 LE — длина][UTF-8 JSON payload]
    // ═══════════════════════════════════════════════════════════════

    public static class PipeIO
    {
        public const int MaxSize = 16 * 1024 * 1024; // 16 МБ

        public static async Task WriteAsync(Stream s, string json, CancellationToken ct = default)
        {
            byte[] payload = Encoding.UTF8.GetBytes(json);
            if (payload.Length > MaxSize)
                throw new InvalidOperationException($"Фрейм {payload.Length} байт превышает {MaxSize}");

            byte[] header = BitConverter.GetBytes((uint)payload.Length);
            await s.WriteAsync(header, 0, 4, ct);
            await s.WriteAsync(payload, 0, payload.Length, ct);
            await s.FlushAsync(ct);
        }

        public static async Task<string> ReadAsync(Stream s, CancellationToken ct = default)
        {
            byte[] header = new byte[4];
            if (await ReadExact(s, header, 4, ct) < 4) return null;

            uint len = BitConverter.ToUInt32(header, 0);
            if (len == 0) return "";
            if (len > MaxSize) throw new InvalidOperationException($"Фрейм {len} байт превышает {MaxSize}");

            byte[] buf = new byte[len];
            if (await ReadExact(s, buf, (int)len, ct) < (int)len) return null;

            return Encoding.UTF8.GetString(buf);
        }

        private static async Task<int> ReadExact(Stream s, byte[] buf, int count, CancellationToken ct)
        {
            int total = 0;
            while (total < count)
            {
                int n = await s.ReadAsync(buf, total, count - total, ct);
                if (n == 0) break;
                total += n;
            }
            return total;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  Константы
    // ═══════════════════════════════════════════════════════════════

    public static class Constants
    {
        /// <summary>Pipe включает PID Revit для поддержки нескольких экземпляров.</summary>
        public static string PipeName(int revitPid) => $"STB2026_Bridge_{revitPid}";

        /// <summary>Discovery pipe — McpServer подключается сюда чтобы узнать PID Revit.</summary>
        public const string PipeDiscovery = "STB2026_Bridge_Discovery";
    }
}
