using System;
using System.Diagnostics;
using System.Reflection;
using Autodesk.Revit.UI;
using STB2026.RevitBridge.Handlers;
using STB2026.RevitBridge.Infrastructure;

namespace STB2026.RevitBridge
{
    /// <summary>
    /// STB2026 MCP Bridge — единая точка входа (IExternalApplication).
    /// Только AI Connector панель. HVAC-плагины в отдельном проекте Application.
    /// </summary>
    public class BridgeApp : IExternalApplication
    {
        public static McpSettingsManager Settings { get; private set; }
        public static bool IsPipeActive { get; private set; }

        private EventBridge _eventBridge;
        private CommandRouter _router;
        private PipeServer _pipeServer;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                try { application.CreateRibbonTab("STB2026"); }
                catch { /* Уже создана другим плагином */ }

                // MCP Bridge инфраструктура
                Settings = new McpSettingsManager();
                Settings.LoadSettings();

                _eventBridge = new EventBridge();
                _eventBridge.Initialize(application);

                _router = new CommandRouter(_eventBridge);
                RegisterHandlers();
                _router.SetSettingsManager(Settings);

                if (Settings.EnableConnection)
                {
                    _pipeServer = new PipeServer(_router);
                    _pipeServer.Start();
                    IsPipeActive = true;
                }

                // Панель AI Connector
                try
                {
                    RibbonPanel panelAI = application.CreateRibbonPanel("STB2026", "AI Connector");
                    var btnMcp = new PushButtonData(
                        "cmdMcpSettings", "A.I.\nConnector", assemblyPath,
                        "STB2026.RevitBridge.McpSettingsCommand");
                    btnMcp.ToolTip = "Настройки подключения Claude / Cursor к Revit через MCP.";
                    panelAI.AddItem(btnMcp);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[STB2026 Bridge] Не удалось добавить кнопку AI: {ex.Message}");
                }

                Debug.WriteLine("[STB2026 Bridge] MCP Bridge запущен (11 обработчиков)");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("STB2026 Bridge",
                    $"Не удалось запустить MCP Bridge:\n{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            _pipeServer?.Dispose();
            IsPipeActive = false;
            return Result.Succeeded;
        }

        /// <summary>Регистрация всех 11 MCP обработчиков.</summary>
        private void RegisterHandlers()
        {
            // ── Чтение ──
            _router.Register("get_model_info",    ModelInfoHandler.Handle);
            _router.Register("get_elements",      ElementsHandler.Handle);
            _router.Register("get_element_params", ElementParamsHandler.Handle);

            // ── Модификация ──
            _router.Register("modify_model",      ModifyHandler.Handle);

            // ── HVAC ──
            _router.Register("run_hvac_check",    HvacCheckHandler.Handle);

            // ── Семейства ──
            _router.Register("manage_families",   FamilyHandler.Handle);

            // ── Листы и видовые экраны ──
            _router.Register("manage_sheets",     SheetHandler.Handle);

            // ── Виды ──
            _router.Register("manage_views",      ViewHandler.Handle);

            // ── Создание элементов ──
            _router.Register("create_elements",   CreationHandler.Handle);

            // ── Геометрия и измерения ──
            _router.Register("query_geometry",    GeometryHandler.Handle);

            // ── Проект ──
            _router.Register("manage_project",    ProjectHandler.Handle);
        }
    }
}
