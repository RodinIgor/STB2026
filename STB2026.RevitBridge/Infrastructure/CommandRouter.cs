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
    /// - Если EnableConnection = false → все запросы отклоняются
    /// - Если EnableEditTools = false → modify_model отклоняется (read-only режим)
    /// </summary>
    public sealed class CommandRouter
    {
        private readonly EventBridge _bridge;
        private McpSettingsManager _settings;

        private readonly Dictionary<string, Func<UIApplication, Dictionary<string, object>, object>> _handlers
            = new Dictionary<string, Func<UIApplication, Dictionary<string, object>, object>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Методы, требующие разрешения на редактирование.</summary>
        private static readonly HashSet<string> WriteMethods = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "modify_model"
            // run_hvac_check тоже может менять модель (velocity окрашивает, tag создаёт марки),
            // но это считается "анализом с визуализацией", не прямой модификацией
        };

        public CommandRouter(EventBridge bridge)
        {
            _bridge = bridge;
        }

        /// <summary>Установить менеджер настроек (для тоглов).</summary>
        public void SetSettingsManager(McpSettingsManager settings)
        {
            _settings = settings;
        }

        /// <summary>Зарегистрировать обработчик метода.</summary>
        public void Register(string method, Func<UIApplication, Dictionary<string, object>, object> handler)
        {
            _handlers[method] = handler;
            Debug.WriteLine($"[STB2026 Router] Зарегистрирован: {method}");
        }

        /// <summary>Обработать входящий запрос.</summary>
        public async Task<BridgeResponse> HandleAsync(BridgeRequest request)
        {
            // ═══ Проверка: подключение включено? ═══
            if (_settings != null && !_settings.EnableConnection)
            {
                return BridgeResponse.Fail(request.Id,
                    "Подключение AI к Revit отключено. " +
                    "Включите тогл 'Enable Connection' в настройках STB2026 AI Connector.");
            }

            // ═══ Проверка: метод существует? ═══
            if (!_handlers.TryGetValue(request.Method, out var handler))
            {
                return BridgeResponse.Fail(request.Id,
                    $"Неизвестный метод: '{request.Method}'. " +
                    $"Доступные: {string.Join(", ", _handlers.Keys)}");
            }

            // ═══ Проверка: write-операция при выключенном тогле? ═══
            if (_settings != null && !_settings.EnableEditTools && WriteMethods.Contains(request.Method))
            {
                return BridgeResponse.Fail(request.Id,
                    "Редактирование модели через AI отключено. " +
                    "Включите тогл 'Enable A.I. Tools to Edit' в настройках STB2026 AI Connector. " +
                    "Доступны только операции чтения: get_model_info, get_elements, get_element_params.");
            }

            // ═══ Выполнение на UI-потоке Revit ═══
            try
            {
                var result = await _bridge.ExecuteOnRevitThread(uiApp =>
                    handler(uiApp, request.Params));

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
