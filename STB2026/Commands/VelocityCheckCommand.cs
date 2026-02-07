using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace STB2026.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class VelocityCheckCommand : IExternalCommand
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

                if (view is ViewSheet || view is ViewSchedule)
                {
                    TaskDialog.Show("STB2026",
                        "Проверка скоростей доступна только на видах модели.");
                    return Result.Cancelled;
                }

                var service = new Services.VelocityCheckerService(doc, view);
                var result = service.CheckAndColorize();

                TaskDialog dlg = new TaskDialog("STB2026 — Проверка скоростей");
                dlg.MainInstruction = "Проверка скоростей по СП 60.13330.2020";
                dlg.MainContent =
                    $"Всего воздуховодов: {result.Total}\n" +
                    $"🟢 В норме: {result.Normal}\n" +
                    $"🟡 Предупреждение: {result.Warning}\n" +
                    $"🔴 Превышение: {result.Exceeded}\n" +
                    $"⚪ Нет данных: {result.NoData}";

                if (result.Exceeded > 0)
                {
                    dlg.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                }

                dlg.FooterText = "Цветовая карта применена к текущему виду.\n" +
                                 "Для сброса: Вид → Графика → Сбросить переопределения.";
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
