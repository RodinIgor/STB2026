using System;
using System.Diagnostics;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using STB2026.Shared.Protocol;

namespace STB2026.RevitBridge.Infrastructure
{
    /// <summary>
    /// Named Pipe сервер — принимает подключения от STB2026.McpServer.exe.
    /// Работает в фоновом потоке, переправляет запросы через EventBridge.
    /// 
    /// Безопасность: ACL ограничен текущим пользователем Windows.
    /// </summary>
    public sealed class PipeServer : IDisposable
    {
        private readonly CommandRouter _router;
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private readonly int _revitPid;
        private Task _listenerTask;
        private Task _discoveryTask;

        public PipeServer(CommandRouter router)
        {
            _router = router;
            _revitPid = Process.GetCurrentProcess().Id;
        }

        /// <summary>Запуск в фоне.</summary>
        public void Start()
        {
            _listenerTask = Task.Run(() => ListenLoop(_cts.Token));
            _discoveryTask = Task.Run(() => DiscoveryLoop(_cts.Token));
            Debug.WriteLine($"[STB2026 Bridge] Pipe сервер запущен (PID {_revitPid})");
        }

        /// <summary>Основной цикл — принимаем подключения одно за другим.</summary>
        private async Task ListenLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    string pipeName = Constants.PipeName(_revitPid);
                    using var server = CreateSecurePipe(pipeName);

                    Debug.WriteLine($"[STB2026 Bridge] Ожидание подключения на '{pipeName}'...");
                    await server.WaitForConnectionAsync(ct);
                    Debug.WriteLine("[STB2026 Bridge] MCP сервер подключился");

                    await HandleConnection(server, ct);
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[STB2026 Bridge] Ошибка listener: {ex.Message}");
                    await Task.Delay(1000, ct); // Пауза перед переподключением
                }
            }
        }

        /// <summary>Discovery pipe — возвращает PID Revit при подключении.</summary>
        private async Task DiscoveryLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    using var server = CreateSecurePipe(Constants.PipeDiscovery);
                    await server.WaitForConnectionAsync(ct);
                    await PipeIO.WriteAsync(server, _revitPid.ToString(), ct);
                    server.Disconnect();
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[STB2026 Bridge] Discovery error: {ex.Message}");
                    await Task.Delay(1000, ct);
                }
            }
        }

        /// <summary>Обработка одного подключения — цикл запрос/ответ.</summary>
        private async Task HandleConnection(NamedPipeServerStream pipe, CancellationToken ct)
        {
            while (pipe.IsConnected && !ct.IsCancellationRequested)
            {
                try
                {
                    string reqJson = await PipeIO.ReadAsync(pipe, ct);
                    if (reqJson == null) break; // Клиент отключился

                    var request = JsonConvert.DeserializeObject<BridgeRequest>(reqJson);
                    if (request == null) continue;

                    var sw = Stopwatch.StartNew();

                    // Маршрутизируем запрос к нужному обработчику
                    BridgeResponse response;
                    try
                    {
                        response = await _router.HandleAsync(request);
                        response.DurationMs = sw.ElapsedMilliseconds;
                    }
                    catch (Exception ex)
                    {
                        response = BridgeResponse.Fail(request.Id, $"Ошибка обработки: {ex.Message}");
                        response.DurationMs = sw.ElapsedMilliseconds;
                    }

                    string respJson = JsonConvert.SerializeObject(response);
                    await PipeIO.WriteAsync(pipe, respJson, ct);
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[STB2026 Bridge] Ошибка обработки запроса: {ex.Message}");
                }
            }
        }

        /// <summary>Создание Named Pipe с ACL только для текущего пользователя.</summary>
        private static NamedPipeServerStream CreateSecurePipe(string name)
        {
            var security = new PipeSecurity();
            var currentUser = WindowsIdentity.GetCurrent().User;
            security.AddAccessRule(new PipeAccessRule(
                currentUser,
                PipeAccessRights.ReadWrite,
                AccessControlType.Allow));

            return NamedPipeServerStreamAcl.Create(
                name,
                PipeDirection.InOut,
                1, // Один клиент за раз
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous,
                inBufferSize: 64 * 1024,
                outBufferSize: 64 * 1024,
                pipeSecurity: security);
        }

        public void Dispose()
        {
            _cts.Cancel();
            _cts.Dispose();
        }
    }
}
