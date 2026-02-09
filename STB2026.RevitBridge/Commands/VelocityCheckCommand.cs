using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace STB2026.RevitBridge.Commands
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
                    TaskDialog.Show("STB2026", "Проверка скоростей доступна только на видах модели.");
                    return Result.Cancelled;
                }

                var service = new STB2026.RevitBridge.Services.VelocityCheckerService(doc, view);
                var result = service.CheckAndColorize();

                TaskDialog dlg = new TaskDialog("STB2026 — Проверка скоростей");
                dlg.MainInstruction = "Проверка скоростей по СП 60.13330.2020";
                dlg.MainContent =
                    $"Всего воздуховодов: {result.Total}\n" +
                    $"🟢 В норме: {result.Normal}\n" +
                    $"🟡 Предупреждение: {result.Warning}\n" +
                    $"🔴 Превышение: {result.Exceeded}\n" +
                    $"⚪ Нет данных: {result.NoData}";

                // Добавляем информацию о применённых диапазонах
                if (result.RangeUsage.Count > 0)
                {
                    string rangeInfo = "Применённые нормы СП 60.13330.2020:\n";
                    foreach (var kvp in result.RangeUsage.OrderByDescending(x => x.Value))
                    {
                        rangeInfo += $"  • {kvp.Key} — {kvp.Value} шт.\n";
                    }
                    dlg.ExpandedContent = rangeInfo;
                }

                if (result.Exceeded > 0)
                {
                    dlg.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                }

                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                    "Сбросить цвета на текущем виде");

                dlg.FooterText = "Цветовая карта применена к текущему виду.";

                var dialogResult = dlg.Show();

                // Сброс цветов по запросу
                if (dialogResult == TaskDialogResult.CommandLink1)
                {
                    ResetColors(doc, view);
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

        /// <summary>
        /// Сбрасывает цветовые переопределения для всех воздуховодов на виде
        /// </summary>
        private void ResetColors(Document doc, View view)
        {
            using (Transaction tx = new Transaction(doc, "STB2026: Сброс цветов"))
            {
                tx.Start();
                var ducts = new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType()
                    .ToList();

                var emptyOverride = new OverrideGraphicSettings();
                foreach (var duct in ducts)
                {
                    try { view.SetElementOverrides(duct.Id, emptyOverride); } catch { }
                }
                tx.Commit();
            }

            TaskDialog.Show("STB2026", $"Цвета сброшены.");
        }
    }
}
