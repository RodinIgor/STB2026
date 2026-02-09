using System;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using STB2026.Shared.Protocol;

namespace STB2026.McpServer.Pipe
{
    /// <summary>
    /// Клиент Named Pipe — отправляет запросы в Revit Bridge.
    /// Автоматически находит запущенный Revit по discovery pipe.
    /// </summary>
    public sealed class RevitPipeClient : IDisposable
    {
        private NamedPipeClientStream _pipe;
        private readonly SemaphoreSlim _lock = new(1, 1);

        // КРИТИЧНО: PropertyNamingPolicy = null = PascalCase.
        // На стороне Revit — Newtonsoft.Json тоже PascalCase по умолчанию.
        // Если поставить CamelCase — pipe-сообщения не распарсятся!
        private readonly JsonSerializerOptions _json = new()
        {
            PropertyNamingPolicy = null, // PascalCase — как Newtonsoft.Json
            PropertyNameCaseInsensitive = true, // Для надёжности при чтении
            WriteIndented = false
        };

        /// <summary>Подключение к Revit Bridge.</summary>
        public async Task<bool> ConnectAsync(int timeoutMs = 10000, CancellationToken ct = default)
        {
            try
            {
                int revitPid = await DiscoverRevitPidAsync(timeoutMs, ct);
                if (revitPid <= 0)
                {
                    Console.Error.WriteLine("[STB2026] Revit не обнаружен через discovery pipe");
                    return false;
                }

                string pipeName = Constants.PipeName(revitPid);
                _pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                await _pipe.ConnectAsync(timeoutMs, ct);

                Console.Error.WriteLine($"[STB2026] Подключён к Revit (PID {revitPid}) через '{pipeName}'");
                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[STB2026] Ошибка подключения: {ex.Message}");
                return false;
            }
        }

        /// <summary>Отправить запрос и дождаться ответа.</summary>
        public async Task<BridgeResponse> SendAsync(BridgeRequest request, CancellationToken ct = default)
        {
            await _lock.WaitAsync(ct);
            try
            {
                if (_pipe == null || !_pipe.IsConnected)
                    return BridgeResponse.Fail(request.Id,
                        "Нет подключения к Revit. Откройте Revit и убедитесь что STB2026 Bridge активен.");

                string reqJson = JsonSerializer.Serialize(request, _json);
                await PipeIO.WriteAsync(_pipe, reqJson, ct);

                string respJson = await PipeIO.ReadAsync(_pipe, ct);
                if (respJson == null)
                    return BridgeResponse.Fail(request.Id, "Revit разорвал соединение");

                return JsonSerializer.Deserialize<BridgeResponse>(respJson, _json)
                    ?? BridgeResponse.Fail(request.Id, "Пустой ответ от Revit");
            }
            catch (Exception ex)
            {
                return BridgeResponse.Fail(request.Id, $"Ошибка IPC: {ex.Message}");
            }
            finally
            {
                _lock.Release();
            }
        }

        /// <summary>Находит PID запущенного Revit через discovery pipe.</summary>
        private async Task<int> DiscoverRevitPidAsync(int timeoutMs, CancellationToken ct)
        {
            try
            {
                using var disc = new NamedPipeClientStream(".", Constants.PipeDiscovery,
                    PipeDirection.InOut, PipeOptions.Asynchronous);
                await disc.ConnectAsync(timeoutMs, ct);

                string resp = await PipeIO.ReadAsync(disc, ct);
                if (int.TryParse(resp, out int pid))
                    return pid;
            }
            catch
            {
                // Discovery pipe не найден — fallback
            }

            // Fallback: ищем процесс Revit
            foreach (var proc in Process.GetProcessesByName("Revit"))
            {
                string pipeName = Constants.PipeName(proc.Id);
                try
                {
                    using var test = new NamedPipeClientStream(".", pipeName,
                        PipeDirection.InOut, PipeOptions.Asynchronous);
                    await test.ConnectAsync(1000, ct);
                    return proc.Id;
                }
                catch { /* Этот Revit без нашего bridge */ }
            }

            return -1;
        }

        public void Dispose()
        {
            _pipe?.Dispose();
            _lock?.Dispose();
        }
    }
}
