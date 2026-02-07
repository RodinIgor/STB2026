using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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

                var service = new Services.SystemValidatorService(doc);
                var result = service.Validate();

                string statusIcon = result.HasErrors ? "⚠️" : "✅";

                TaskDialog dlg = new TaskDialog("STB2026 — Валидация системы");
                dlg.MainInstruction = result.HasErrors
                    ? "Обнаружены проблемы"
                    : "Система в порядке";

                dlg.MainContent =
                    $"Всего элементов: {result.TotalElements}\n\n" +
                    $"Нулевой расход: {result.ZeroFlowCount}\n" +
                    $"Не подключены: {result.DisconnectedCount}\n" +
                    $"Без системы: {result.NoSystemCount}";

                if (result.HasErrors)
                {
                    dlg.MainIcon = TaskDialogIcon.TaskDialogIconWarning;

                    string details = "";
                    if (result.ZeroFlowIds.Any())
                    {
                        details += $"Нулевой расход (ID): {string.Join(", ", result.ZeroFlowIds.Take(20))}";
                        if (result.ZeroFlowIds.Count > 20)
                            details += $"... и ещё {result.ZeroFlowIds.Count - 20}";
                        details += "\n\n";
                    }
                    if (result.DisconnectedIds.Any())
                    {
                        details += $"Не подключены (ID): {string.Join(", ", result.DisconnectedIds.Take(20))}";
                        if (result.DisconnectedIds.Count > 20)
                            details += $"... и ещё {result.DisconnectedIds.Count - 20}";
                        details += "\n\n";
                    }
                    if (result.NoSystemIds.Any())
                    {
                        details += $"Без системы (ID): {string.Join(", ", result.NoSystemIds.Take(20))}";
                        if (result.NoSystemIds.Count > 20)
                            details += $"... и ещё {result.NoSystemIds.Count - 20}";
                    }

                    dlg.ExpandedContent = details;

                    // Подсветка проблемных элементов
                    dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                        "Выделить проблемные элементы в модели");
                }

                var dialogResult = dlg.Show();

                if (dialogResult == TaskDialogResult.CommandLink1 && result.HasErrors)
                {
                    // Выделяем проблемные элементы
                    var allProblemIds = result.ZeroFlowIds
                        .Concat(result.DisconnectedIds)
                        .Concat(result.NoSystemIds)
                        .Distinct()
                        .Select(id => new ElementId(id))
                        .ToList();

                    uidoc.Selection.SetElementIds(allProblemIds);

                    // Цветовая подсветка
                    using (Transaction tx = new Transaction(doc, "STB2026: Подсветка проблем"))
                    {
                        tx.Start();
                        View view = doc.ActiveView;

                        var redOverride = new OverrideGraphicSettings();
                        redOverride.SetProjectionLineColor(new Color(255, 0, 0));

                        var yellowOverride = new OverrideGraphicSettings();
                        yellowOverride.SetProjectionLineColor(new Color(255, 200, 0));

                        var grayOverride = new OverrideGraphicSettings();
                        grayOverride.SetProjectionLineColor(new Color(128, 128, 128));

                        foreach (int id in result.ZeroFlowIds)
                        {
                            try { view.SetElementOverrides(new ElementId(id), redOverride); } catch { }
                        }
                        foreach (int id in result.DisconnectedIds)
                        {
                            try { view.SetElementOverrides(new ElementId(id), yellowOverride); } catch { }
                        }
                        foreach (int id in result.NoSystemIds)
                        {
                            try { view.SetElementOverrides(new ElementId(id), grayOverride); } catch { }
                        }

                        tx.Commit();
                    }
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
    }
}
