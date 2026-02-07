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
                application.CreateRibbonTab("STB2026");
                RibbonPanel panel = application.CreateRibbonPanel("STB2026", "Основные");

                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                PushButtonData buttonData = new PushButtonData(
                    "cmdHelloWorld",
                    "Привет\nмир",
                    assemblyPath,
                    "STB2026.HelloWorldCommand"
                );
                buttonData.ToolTip = "Показать приветственное окно STB2026";
                panel.AddItem(buttonData);

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