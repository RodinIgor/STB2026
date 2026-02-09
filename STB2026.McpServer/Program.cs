using System;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using STB2026.McpServer.Pipe;

namespace STB2026.McpServer
{
    /// <summary>
    /// STB2026 MCP Server — транспорт между Claude и Revit.
    /// КРИТИЧНО: stdout используется ТОЛЬКО для MCP JSON-RPC.
    /// Все логи → stderr (Console.Error).
    /// </summary>
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.Error.WriteLine("[STB2026] MCP Server запущен");
            Console.Error.WriteLine("[STB2026] Подключение к Revit...");

            var pipeClient = new RevitPipeClient();
            bool connected = await pipeClient.ConnectAsync(timeoutMs: 15000);

            if (!connected)
            {
                Console.Error.WriteLine("[STB2026] ВНИМАНИЕ: Revit не подключён. Tools будут возвращать ошибку.");
                Console.Error.WriteLine("[STB2026] Откройте Revit и убедитесь что STB2026 Bridge активен.");
            }

            var builder = Host.CreateApplicationBuilder(args);
            
            // КРИТИЧНО: убрать ВСЕ логи из stdout — иначе ломает MCP JSON-RPC
            builder.Logging.ClearProviders();
            builder.Logging.AddConsole(options =>
            {
                options.LogToStandardErrorThreshold = LogLevel.Trace; // ВСЕ логи → stderr
            });
            
            builder.Services.AddSingleton(pipeClient);
            builder.Services
                .AddMcpServer()
                .WithStdioServerTransport()
                .WithToolsFromAssembly();

            var app = builder.Build();
            await app.RunAsync();
        }
    }
}
