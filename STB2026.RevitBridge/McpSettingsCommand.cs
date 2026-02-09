using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using STB2026.RevitBridge.UI;

namespace STB2026.RevitBridge
{
    /// <summary>
    /// Команда "A.I. Connector" — открывает окно настроек MCP.
    /// Кнопка на панели STB2026 в Revit.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class McpSettingsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получаем ссылки из статических полей BridgeApp
            var settings = BridgeApp.Settings;
            bool pipeActive = BridgeApp.IsPipeActive;

            if (settings == null)
            {
                TaskDialog.Show("STB2026", "Bridge не инициализирован. Перезапустите Revit.");
                return Result.Failed;
            }

            var window = new McpSettingsWindow(settings, pipeActive);
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}
