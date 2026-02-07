using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace STB2026.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class TagDuctsCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                View view = doc.ActiveView;

                // Проверяем, что вид подходит для маркировки
                if (view is ViewSheet || view is ViewSchedule || view is View3D)
                {
                    TaskDialog.Show("STB2026",
                        "Маркировка доступна только на планах, разрезах и фасадах.");
                    return Result.Cancelled;
                }

                var service = new Services.DuctTaggerService(doc, view);
                var result = service.TagAllDucts();

                TaskDialog dlg = new TaskDialog("STB2026 — Маркировка воздуховодов");
                dlg.MainInstruction = "Маркировка завершена";
                dlg.MainContent =
                    $"Промаркировано: {result.Tagged}\n" +
                    $"Уже имели марки: {result.AlreadyTagged}\n" +
                    $"Пропущено (ошибки): {result.Skipped}";

                if (result.Skipped > 0 && result.Errors.Count > 0)
                {
                    dlg.ExpandedContent = string.Join("\n", result.Errors);
                }

                dlg.Show();
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
