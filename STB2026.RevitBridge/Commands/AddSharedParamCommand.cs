using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace STB2026.RevitBridge.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AddSharedParamCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                if (uiApp.ActiveUIDocument == null)
                {
                    TaskDialog.Show("STB2026", "Нет открытого проекта.");
                    return Result.Cancelled;
                }

                var window = new UI.AddSharedParamWindow(uiApp);
                window.ShowDialog();
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Ошибка", $"Не удалось открыть окно:\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
