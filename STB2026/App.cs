using Autodesk.Revit.UI;
using System;
using System.Reflection;

namespace STB2026
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                // Создаём вкладку на ленте Revit
                application.CreateRibbonTab("STB2026");

                // ═══════════════════════════════════════
                // Панель 1: Оформление
                // ═══════════════════════════════════════
                RibbonPanel panelDecor = application.CreateRibbonPanel("STB2026", "Оформление");

                PushButtonData btnTagDucts = new PushButtonData(
                    "cmdTagDucts",
                    "Маркировка\nвоздуховодов",
                    assemblyPath,
                    "STB2026.Commands.TagDuctsCommand"
                );
                btnTagDucts.ToolTip = "Автоматическая маркировка воздуховодов:\nрасход, размер, скорость";
                btnTagDucts.LargeImage = null; // TODO: иконка 32x32
                panelDecor.AddItem(btnTagDucts);

                // ═══════════════════════════════════════
                // Панель 2: Проверки
                // ═══════════════════════════════════════
                RibbonPanel panelChecks = application.CreateRibbonPanel("STB2026", "Проверки");

                PushButtonData btnVelocity = new PushButtonData(
                    "cmdVelocityCheck",
                    "Проверка\nскоростей",
                    assemblyPath,
                    "STB2026.Commands.VelocityCheckCommand"
                );
                btnVelocity.ToolTip = "Проверка скоростей воздуха по СП 60.13330.2020\nЦветовая карта: зелёный/жёлтый/красный";
                panelChecks.AddItem(btnVelocity);

                PushButtonData btnValidation = new PushButtonData(
                    "cmdSystemValidation",
                    "Валидация\nсистемы",
                    assemblyPath,
                    "STB2026.Commands.SystemValidationCommand"
                );
                btnValidation.ToolTip = "Проверка: нулевые расходы, отключённые элементы,\nненазначенные системы";
                panelChecks.AddItem(btnValidation);

                // ═══════════════════════════════════════
                // Панель 3: Расчёты
                // ═══════════════════════════════════════
                RibbonPanel panelCalc = application.CreateRibbonPanel("STB2026", "Расчёты");

                PushButtonData btnWallInt = new PushButtonData(
                    "cmdWallIntersections",
                    "Пересечения\nсо стенами",
                    assemblyPath,
                    "STB2026.Commands.WallIntersectionsCommand"
                );
                btnWallInt.ToolTip = "Поиск пересечений воздуховодов со стенами\nи формирование координатного отчёта";
                panelCalc.AddItem(btnWallInt);

                // ═══════════════════════════════════════
                // Панель 4: Отчёты (зарезервирована)
                // ═══════════════════════════════════════
                RibbonPanel panelReports = application.CreateRibbonPanel("STB2026", "Отчёты");

                PushButtonData btnHello = new PushButtonData(
                    "cmdHelloWorld",
                    "О плагине",
                    assemblyPath,
                    "STB2026.HelloWorldCommand"
                );
                btnHello.ToolTip = "Информация о плагине STB2026";
                panelReports.AddItem(btnHello);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("STB2026 — Ошибка", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
