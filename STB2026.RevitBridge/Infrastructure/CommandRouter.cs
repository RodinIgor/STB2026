using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using STB2026.Shared.Protocol;

namespace STB2026.RevitBridge.Infrastructure
{
    /// <summary>
    /// Маршрутизатор команд MCP → обработчики Revit API.
    /// 
    /// Интеграция с McpSettingsManager:
    /// - EnableConnection = false → все запросы отклоняются
    /// - EnableEditTools = false  → write-методы отклоняются (read-only режим)
    /// </summary>
    public sealed class CommandRouter
    {
        private readonly EventBridge _bridge;
        private McpSettingsManager _settings;

        private readonly Dictionary<string, Func<UIApplication, Dictionary<string, object>, object>> _handlers
            = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Методы, требующие разрешения на редактирование.</summary>
        private static readonly HashSet<string> WriteMethods = new(StringComparer.OrdinalIgnoreCase)
        {
            "modify_model",
            "manage_sheets",
            "manage_views",
            "create_elements",
            "manage_families",
            "manage_project"
            // run_hvac_check — анализ с визуализацией, не прямая модификация
            // query_geometry — только чтение
        };

        public CommandRouter(EventBridge bridge) => _bridge = bridge;

        public void SetSettingsManager(McpSettingsManager settings) => _settings = settings;

        public void Register(string method, Func<UIApplication, Dictionary<string, object>, object> handler)
        {
            _handlers[method] = handler;
            Debug.WriteLine($"[STB2026 Router] Зарегистрирован: {method}");
        }

        public async Task<BridgeResponse> HandleAsync(BridgeRequest request)
        {
            // Проверка: подключение включено?
            if (_settings != null && !_settings.EnableConnection)
                return BridgeResponse.Fail(request.Id,
                    "Подключение AI к Revit отключено. Включите 'Enable Connection' в настройках.");

            // Проверка: метод существует?
            if (!_handlers.TryGetValue(request.Method, out var handler))
                return BridgeResponse.Fail(request.Id,
                    $"Неизвестный метод: '{request.Method}'. Доступные: {string.Join(", ", _handlers.Keys)}");

            // Проверка: write-операция при read-only?
            if (_settings != null && !_settings.EnableEditTools && WriteMethods.Contains(request.Method))
                return BridgeResponse.Fail(request.Id,
                    "Редактирование модели через AI отключено. " +
                    "Включите 'Enable A.I. Tools to Edit' в настройках. " +
                    "Доступны: get_model_info, get_elements, get_element_params, run_hvac_check, query_geometry.");

            // Выполнение на UI-потоке Revit
            try
            {
                var result = await _bridge.ExecuteOnRevitThread(uiApp => handler(uiApp, request.Params));
                return BridgeResponse.Ok(request.Id, result);
            }
            catch (TimeoutException)
            {
                return BridgeResponse.Fail(request.Id,
                    "Revit не ответил за 30 секунд. Возможно, открыт диалог или идёт долгая операция.");
            }
            catch (Exception ex)
            {
                return BridgeResponse.Fail(request.Id, $"{ex.GetType().Name}: {ex.Message}");
            }
        }
    }
}
