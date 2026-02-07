using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace STB2026.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class SystemValidationCommand : IExternalCommand
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
                    TaskDialog.Show("STB2026", "Валидация доступна только на видах модели.");
                    return Result.Cancelled;
                }

                // Валидация с автоматической подсветкой
                var service = new Services.SystemValidatorService(doc, view);
                var result = service.ValidateAndColorize();

                TaskDialog dlg = new TaskDialog("STB2026 — Валидация системы");
                dlg.MainInstruction = result.HasErrors
                    ? "Обнаружены проблемы"
                    : "Система в порядке";

                string content =
                    $"Элементов на виде: {result.TotalElements}\n\n" +
                    $"🔴 Нулевой расход: {result.ZeroFlowCount}\n" +
                    $"🟡 Не подключены: {result.DisconnectedCount}\n" +
                    $"⚪ Без системы: {result.NoSystemCount}";

                if (result.HasErrors)
                    content += "\n\nПроблемные элементы подсвечены на виде.";

                dlg.MainContent = content;

                if (result.HasErrors)
                {
                    dlg.MainIcon = TaskDialogIcon.TaskDialogIconWarning;

                    string details = "";
                    if (result.ZeroFlowIds.Any())
                    {
                        details += $"🔴 Нулевой расход (ID): {string.Join(", ", result.ZeroFlowIds.Take(20))}";
                        if (result.ZeroFlowIds.Count > 20)
                            details += $"... и ещё {result.ZeroFlowIds.Count - 20}";
                        details += "\n\n";
                    }
                    if (result.DisconnectedIds.Any())
                    {
                        details += $"🟡 Не подключены (ID): {string.Join(", ", result.DisconnectedIds.Take(20))}";
                        if (result.DisconnectedIds.Count > 20)
                            details += $"... и ещё {result.DisconnectedIds.Count - 20}";
                        details += "\n\n";
                    }
                    if (result.NoSystemIds.Any())
                    {
                        details += $"⚪ Без системы (ID): {string.Join(", ", result.NoSystemIds.Take(20))}";
                        if (result.NoSystemIds.Count > 20)
                            details += $"... и ещё {result.NoSystemIds.Count - 20}";
                    }

                    dlg.ExpandedContent = details;

                    dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                        "Выделить проблемные элементы в модели");
                }

                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                    "Сбросить цвета на текущем виде");

                var dialogResult = dlg.Show();

                if (dialogResult == TaskDialogResult.CommandLink1 && result.HasErrors)
                {
                    // Только выделение (Selection) — подсветка уже сделана
                    var allProblemIds = result.ZeroFlowIds
                        .Concat(result.DisconnectedIds)
                        .Concat(result.NoSystemIds)
                        .Distinct()
                        .Select(id => new ElementId(id))
                        .ToList();

                    uidoc.Selection.SetElementIds(allProblemIds);
                }
                else if (dialogResult == TaskDialogResult.CommandLink2)
                {
                    ResetAllMepColors(doc, view);
                }

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

        private void ResetAllMepColors(Document doc, View view)
        {
            using (Transaction tx = new Transaction(doc, "STB2026: Сброс цветов"))
            {
                tx.Start();
                var emptyOverride = new OverrideGraphicSettings();

                var allMepElements = new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType()
                    .ToList();

                var fittings = new FilteredElementCollector(doc, view.Id)
                    .OfCategory(BuiltInCategory.OST_DuctFitting)
                    .WhereElementIsNotElementType()
                    .ToList();

                var terminals = new FilteredElementCollector(doc, view.Id)
                    .OfCategory(BuiltInCategory.OST_DuctTerminal)
                    .WhereElementIsNotElementType()
                    .ToList();

                foreach (var elem in allMepElements.Concat(fittings).Concat(terminals))
                {
                    try { view.SetElementOverrides(elem.Id, emptyOverride); } catch { }
                }

                tx.Commit();
            }
            TaskDialog.Show("STB2026", "Цвета сброшены.");
        }
    }
}
