using System;
using System.Diagnostics;
using System.Reflection;
using Autodesk.Revit.UI;
using STB2026.RevitBridge.Handlers;
using STB2026.RevitBridge.Infrastructure;

namespace STB2026.RevitBridge
{
    /// <summary>
    /// STB2026 — единая точка входа (IExternalApplication).
    /// 
    /// Объединяет:
    /// 1. MCP Bridge (Named Pipe → Claude/Cursor)
    /// 2. HVAC инструменты (кнопки на ribbon)
    /// 
    /// Все на одной вкладке "STB2026":
    /// ┌──────────────────────────────────────────────────────────────┐
    /// │  Оформление  │  Проверки  │  Расчёты  │  AI Connector      │
    /// │  Маркировка  │  Скорости  │  Пересеч. │  A.I. Connector    │
    /// │  воздухов.   │  Валидация │  со стен. │                    │
    /// │              │  Сбросить  │           │                    │
    /// │              │  цвета     │           │                    │
    /// └──────────────────────────────────────────────────────────────┘
    /// </summary>
    public class BridgeApp : IExternalApplication
    {
        // ═══ Статические ссылки для доступа из команд ═══
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

                // ═══════════════════════════════════════
                // 1. Создаём единую вкладку STB2026
                // ═══════════════════════════════════════
                try { application.CreateRibbonTab("STB2026"); }
                catch { /* Уже создана — нормально */ }

                // ═══════════════════════════════════════
                // 2. HVAC панели (кнопки)
                // ═══════════════════════════════════════
                CreateHvacPanels(application, assemblyPath);

                // ═══════════════════════════════════════
                // 3. MCP Bridge (AI Connector)
                // ═══════════════════════════════════════
                InitializeMcpBridge(application, assemblyPath);

                Debug.WriteLine("[STB2026] ✅ Плагин запущен (HVAC + MCP Bridge)");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("STB2026 — Ошибка",
                    $"Не удалось запустить плагин:\n{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            _pipeServer?.Dispose();
            IsPipeActive = false;
            Debug.WriteLine("[STB2026] Остановлен");
            return Result.Succeeded;
        }

        // ═══════════════════════════════════════════════════════════
        //  HVAC панели
        // ═══════════════════════════════════════════════════════════

        private void CreateHvacPanels(UIControlledApplication app, string assemblyPath)
        {
            // ─── Панель: Оформление ───
            RibbonPanel panelDecor = app.CreateRibbonPanel("STB2026", "Оформление");

            var btnTagDucts = new PushButtonData(
                "cmdTagDucts",
                "Маркировка\nвоздуховодов",
                assemblyPath,
                "STB2026.RevitBridge.Commands.TagDuctsCommand"
            );
            btnTagDucts.ToolTip = "Автоматическая маркировка воздуховодов:\nрасход, размер, скорость";
            panelDecor.AddItem(btnTagDucts);

            // ─── Панель: Проверки ───
            RibbonPanel panelChecks = app.CreateRibbonPanel("STB2026", "Проверки");

            var btnVelocity = new PushButtonData(
                "cmdVelocityCheck",
                "Проверка\nскоростей",
                assemblyPath,
                "STB2026.RevitBridge.Commands.VelocityCheckCommand"
            );
            btnVelocity.ToolTip = "Проверка скоростей воздуха по СП 60.13330.2020\nЦветовая карта: зелёный/жёлтый/оранжевый";
            panelChecks.AddItem(btnVelocity);

            var btnValidation = new PushButtonData(
                "cmdSystemValidation",
                "Валидация\nсистемы",
                assemblyPath,
                "STB2026.RevitBridge.Commands.SystemValidationCommand"
            );
            btnValidation.ToolTip = "Проверка: нулевые расходы, отключённые элементы,\nненазначенные системы";
            panelChecks.AddItem(btnValidation);

            var btnResetColors = new PushButtonData(
                "cmdResetColors",
                "Сбросить\nцвета",
                assemblyPath,
                "STB2026.RevitBridge.Commands.ResetColorsCommand"
            );
            btnResetColors.ToolTip = "Сбросить цветовые переопределения\nдля всех воздуховодов на текущем виде";
            panelChecks.AddItem(btnResetColors);

            // ─── Панель: Расчёты ───
            RibbonPanel panelCalc = app.CreateRibbonPanel("STB2026", "Расчёты");

            var btnWallInt = new PushButtonData(
                "cmdWallIntersections",
                "Пересечения\nсо стенами",
                assemblyPath,
                "STB2026.RevitBridge.Commands.WallIntersectionsCommand"
            );
            btnWallInt.ToolTip = "Поиск пересечений воздуховодов со стенами\n(включая связанные файлы) и координатный отчёт";
            panelCalc.AddItem(btnWallInt);
        }

        // ═══════════════════════════════════════════════════════════
        //  MCP Bridge
        // ═══════════════════════════════════════════════════════════

        private void InitializeMcpBridge(UIControlledApplication app, string assemblyPath)
        {
            // 1. Настройки
            Settings = new McpSettingsManager();
            Settings.LoadSettings();

            // 2. EventBridge — мост фоновый поток → UI поток Revit
            _eventBridge = new EventBridge();
            _eventBridge.Initialize(app);

            // 3. CommandRouter — маршрутизация MCP команд
            _router = new CommandRouter(_eventBridge);
            RegisterHandlers();
            _router.SetSettingsManager(Settings);

            // 4. PipeServer — если подключение включено
            if (Settings.EnableConnection)
            {
                _pipeServer = new PipeServer(_router);
                _pipeServer.Start();
                IsPipeActive = true;
            }

            // 5. Кнопка AI Connector
            try
            {
                RibbonPanel panelAI = app.CreateRibbonPanel("STB2026", "AI Connector");

                var btnMcp = new PushButtonData(
                    "cmdMcpSettings",
                    "A.I.\nConnector",
                    assemblyPath,
                    "STB2026.RevitBridge.McpSettingsCommand"
                );
                btnMcp.ToolTip = "Настройки подключения Claude / Cursor к Revit через MCP.\n" +
                                  "Auto-setup, Manual JSON, тоглы read/write доступа.";
                panelAI.AddItem(btnMcp);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[STB2026] Не удалось добавить кнопку AI: {ex.Message}");
            }
        }

        /// <summary>Регистрация всех 6 MCP обработчиков.</summary>
        private void RegisterHandlers()
        {
            _router.Register("get_model_info", ModelInfoHandler.Handle);
            _router.Register("get_elements", ElementsHandler.Handle);
            _router.Register("get_element_params", ElementParamsHandler.Handle);
            _router.Register("modify_model", ModifyHandler.Handle);
            _router.Register("run_hvac_check", HvacCheckHandler.Handle);
            _router.Register("manage_families", FamilyHandler.Handle);
        }
    }
}
