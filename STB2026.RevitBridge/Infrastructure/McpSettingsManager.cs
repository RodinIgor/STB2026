using System;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Infrastructure
{
    /// <summary>
    /// Управление настройками MCP-подключения.
    /// 
    /// Функционал (аналог Nonica AI Connector Settings):
    /// 1. Auto-setup: находит claude_desktop_config.json и добавляет STB2026
    /// 2. Manual: показывает JSON для копирования
    /// 3. Детектирует запущенные AI-приложения (Claude Desktop, Cursor)
    /// 4. Тогл read-only / read-write доступа
    /// </summary>
    public sealed class McpSettingsManager
    {
        // ═══ Пути к конфигам AI-приложений ═══

        private static readonly string ClaudeConfigDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Claude");

        private static readonly string ClaudeConfigPath =
            Path.Combine(ClaudeConfigDir, "claude_desktop_config.json");

        private static readonly string CursorConfigDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Cursor");

        private static readonly string CursorConfigPath =
            Path.Combine(CursorConfigDir, "mcp.json");

        // ═══ Настройки STB2026 ═══

        private static readonly string Stb2026ConfigDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "STB2026");

        private static readonly string Stb2026ConfigPath =
            Path.Combine(Stb2026ConfigDir, "settings.json");

        /// <summary>Имя MCP-сервера в конфигах AI-приложений.</summary>
        public const string McpServerName = "Revit_STB2026";

        /// <summary>Разрешено ли AI редактировать модель (modify_model).</summary>
        public bool EnableEditTools { get; set; } = true;

        /// <summary>Включено ли MCP-подключение.</summary>
        public bool EnableConnection { get; set; } = true;

        // ═══════════════════════════════════════════════════════
        //  Детекция AI-приложений
        // ═══════════════════════════════════════════════════════

        /// <summary>Статус AI-приложения.</summary>
        public class AiAppStatus
        {
            public string Name { get; set; }
            public bool Installed { get; set; }
            public bool Running { get; set; }
            public bool Configured { get; set; }
            public bool NeedsRestart { get; set; }
            public string ConfigPath { get; set; }
            public string Message { get; set; }
        }

        /// <summary>Проверить статус Claude Desktop.</summary>
        public AiAppStatus DetectClaudeDesktop()
        {
            var status = new AiAppStatus
            {
                Name = "Claude Desktop App",
                ConfigPath = ClaudeConfigPath
            };

            // Установлен?
            status.Installed = Directory.Exists(ClaudeConfigDir);

            // Запущен?
            status.Running = IsProcessRunning("Claude");

            // Настроен?
            if (status.Installed && File.Exists(ClaudeConfigPath))
            {
                try
                {
                    var json = JObject.Parse(File.ReadAllText(ClaudeConfigPath));
                    var servers = json["mcpServers"] as JObject;
                    status.Configured = servers != null && servers.ContainsKey(McpServerName);
                }
                catch { status.Configured = false; }
            }

            // Сообщение
            if (!status.Installed)
                status.Message = "Claude Desktop App не обнаружен.";
            else if (!status.Configured)
                status.Message = "Требуется настройка. Нажмите 'Подключить'.";
            else if (status.NeedsRestart)
                status.Message = "Настроен. Перезапустите Claude Desktop (трей, рядом с часами).";
            else
                status.Message = "✓ Готов к работе.";

            return status;
        }

        /// <summary>Проверить статус Cursor.</summary>
        public AiAppStatus DetectCursor()
        {
            var status = new AiAppStatus
            {
                Name = "Cursor",
                ConfigPath = CursorConfigPath
            };

            status.Installed = Directory.Exists(CursorConfigDir);
            status.Running = IsProcessRunning("Cursor");

            if (status.Installed && File.Exists(CursorConfigPath))
            {
                try
                {
                    var json = JObject.Parse(File.ReadAllText(CursorConfigPath));
                    var servers = json["mcpServers"] as JObject;
                    status.Configured = servers != null && servers.ContainsKey(McpServerName);
                }
                catch { status.Configured = false; }
            }

            if (!status.Installed)
                status.Message = "Cursor не обнаружен.";
            else if (!status.Configured)
                status.Message = "Требуется настройка.";
            else
                status.Message = "✓ Готов к работе.";

            return status;
        }

        /// <summary>Проверить статус Claude Code (VS Code + CLI).</summary>
        public AiAppStatus DetectClaudeCode()
        {
            var status = new AiAppStatus { Name = "Claude Code (Multi Agent)" };
            // Claude Code обнаруживается через наличие CLI-команды
            status.Installed = FindInPath("claude.exe") || FindInPath("claude");
            status.Running = IsProcessRunning("claude");
            status.Configured = status.Installed; // Claude Code использует свой MCP discovery
            status.Message = status.Installed
                ? "✓ Готов к использованию через Multi Agent."
                : "Claude Code не обнаружен.";
            return status;
        }

        // ═══════════════════════════════════════════════════════
        //  Auto-setup: запись конфига
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Автоматически добавить STB2026 в claude_desktop_config.json.
        /// Возвращает true если записано, false если ошибка.
        /// </summary>
        public SetupResult AutoSetupClaude()
        {
            return AutoSetupConfig(ClaudeConfigDir, ClaudeConfigPath, "Claude Desktop");
        }

        /// <summary>Автоматически добавить STB2026 в Cursor mcp.json.</summary>
        public SetupResult AutoSetupCursor()
        {
            return AutoSetupConfig(CursorConfigDir, CursorConfigPath, "Cursor");
        }

        /// <summary>Удалить STB2026 из claude_desktop_config.json.</summary>
        public SetupResult RemoveFromClaude()
        {
            return RemoveFromConfig(ClaudeConfigPath, "Claude Desktop");
        }

        /// <summary>Удалить STB2026 из Cursor mcp.json.</summary>
        public SetupResult RemoveFromCursor()
        {
            return RemoveFromConfig(CursorConfigPath, "Cursor");
        }

        private SetupResult AutoSetupConfig(string configDir, string configPath, string appName)
        {
            try
            {
                // Создаём директорию если нет
                if (!Directory.Exists(configDir))
                    Directory.CreateDirectory(configDir);

                // Читаем существующий конфиг или создаём новый
                JObject config;
                if (File.Exists(configPath))
                {
                    string existing = File.ReadAllText(configPath);
                    config = string.IsNullOrWhiteSpace(existing)
                        ? new JObject()
                        : JObject.Parse(existing);
                }
                else
                {
                    config = new JObject();
                }

                // Добавляем/обновляем mcpServers
                if (config["mcpServers"] == null)
                    config["mcpServers"] = new JObject();

                var servers = (JObject)config["mcpServers"];
                servers[McpServerName] = JObject.FromObject(GetMcpServerConfig());

                // Записываем с красивым форматированием
                string json = config.ToString(Formatting.Indented);

                // Бэкап перед записью
                if (File.Exists(configPath))
                {
                    string backup = configPath + ".bak";
                    File.Copy(configPath, backup, true);
                }

                File.WriteAllText(configPath, json);

                return new SetupResult
                {
                    Success = true,
                    Message = $"✓ STB2026 добавлен в {appName}.\nПерезапустите {appName} для применения.",
                    NeedsRestart = true
                };
            }
            catch (UnauthorizedAccessException)
            {
                return new SetupResult
                {
                    Success = false,
                    Message = $"Нет доступа к файлу конфигурации {appName}.\n" +
                              $"Возможно, {appName} установлен для другого пользователя.\n" +
                              $"Путь: {configPath}"
                };
            }
            catch (Exception ex)
            {
                return new SetupResult
                {
                    Success = false,
                    Message = $"Ошибка настройки {appName}: {ex.Message}"
                };
            }
        }

        private SetupResult RemoveFromConfig(string configPath, string appName)
        {
            try
            {
                if (!File.Exists(configPath))
                    return new SetupResult { Success = true, Message = $"Конфиг {appName} не найден." };

                var config = JObject.Parse(File.ReadAllText(configPath));
                var servers = config["mcpServers"] as JObject;
                if (servers != null && servers.ContainsKey(McpServerName))
                {
                    servers.Remove(McpServerName);
                    File.WriteAllText(configPath, config.ToString(Formatting.Indented));
                    return new SetupResult
                    {
                        Success = true,
                        Message = $"✓ STB2026 удалён из {appName}. Перезапустите {appName}.",
                        NeedsRestart = true
                    };
                }

                return new SetupResult { Success = true, Message = $"STB2026 не найден в конфиге {appName}." };
            }
            catch (Exception ex)
            {
                return new SetupResult { Success = false, Message = $"Ошибка: {ex.Message}" };
            }
        }

        // ═══════════════════════════════════════════════════════
        //  Manual mode: JSON для копирования
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// JSON для ручной вставки в claude_desktop_config.json или mcp.json.
        /// Как на скриншоте 3 Nonica (Manual tab).
        /// </summary>
        public string GetManualConfigJson()
        {
            var config = new JObject
            {
                ["mcpServers"] = new JObject
                {
                    [McpServerName] = JObject.FromObject(GetMcpServerConfig())
                }
            };

            return config.ToString(Formatting.Indented);
        }

        // ═══════════════════════════════════════════════════════
        //  Конфиг MCP-сервера
        // ═══════════════════════════════════════════════════════

        /// <summary>Путь к STB2026.McpServer.exe (рядом с RevitBridge.dll).</summary>
        public string GetMcpServerExePath()
        {
            // McpServer.exe лежит рядом с RevitBridge.dll
            string bridgeDir = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);

            return Path.Combine(bridgeDir, "STB2026.McpServer.exe");
        }

        private object GetMcpServerConfig()
        {
            return new
            {
                type = "stdio",
                command = GetMcpServerExePath().Replace("\\", "\\\\"),
                args = new string[0],
                timeout = 15000
            };
        }

        // ═══════════════════════════════════════════════════════
        //  Сохранение/загрузка настроек STB2026
        // ═══════════════════════════════════════════════════════

        /// <summary>Сохранить настройки STB2026 (тоглы).</summary>
        public void SaveSettings()
        {
            try
            {
                if (!Directory.Exists(Stb2026ConfigDir))
                    Directory.CreateDirectory(Stb2026ConfigDir);

                var settings = new JObject
                {
                    ["enableConnection"] = EnableConnection,
                    ["enableEditTools"] = EnableEditTools
                };

                File.WriteAllText(Stb2026ConfigPath, settings.ToString(Formatting.Indented));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[STB2026] Ошибка сохранения настроек: {ex.Message}");
            }
        }

        /// <summary>Загрузить настройки STB2026.</summary>
        public void LoadSettings()
        {
            try
            {
                if (!File.Exists(Stb2026ConfigPath)) return;

                var settings = JObject.Parse(File.ReadAllText(Stb2026ConfigPath));
                EnableConnection = settings.Value<bool?>("enableConnection") ?? true;
                EnableEditTools = settings.Value<bool?>("enableEditTools") ?? true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[STB2026] Ошибка загрузки настроек: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════
        //  Утилиты
        // ═══════════════════════════════════════════════════════

        private static bool IsProcessRunning(string name)
        {
            try { return Process.GetProcessesByName(name).Length > 0; }
            catch { return false; }
        }

        private static bool FindInPath(string exe)
        {
            var pathDirs = Environment.GetEnvironmentVariable("PATH")?.Split(';') ?? new string[0];
            foreach (var dir in pathDirs)
            {
                try
                {
                    if (File.Exists(Path.Combine(dir, exe)))
                        return true;
                }
                catch { }
            }
            return false;
        }

        public class SetupResult
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public bool NeedsRestart { get; set; }
        }
    }
}
