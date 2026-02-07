using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Text;

namespace STB2026.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class WallIntersectionsCommand : IExternalCommand
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

                var service = new Services.WallIntersectorService(doc);
                var result = service.FindIntersections();

                if (result.Intersections.Count == 0)
                {
                    TaskDialog.Show("STB2026",
                        "Пересечения воздуховодов со стенами не обнаружены.");
                    return Result.Succeeded;
                }

                // Формируем отчёт
                var sb = new StringBuilder();
                sb.AppendLine("Воздуховод;Система;Размер;Стена;Тип стены;X (мм);Y (мм);Z (мм);Отметка (мм)");

                foreach (var item in result.Intersections)
                {
                    sb.AppendLine(
                        $"{item.DuctId};{item.SystemName};{item.DuctSize};" +
                        $"{item.WallId};{item.WallType};" +
                        $"{item.X:F0};{item.Y:F0};{item.Z:F0};{item.Elevation:F0}");
                }

                // Сохраняем CSV
                string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"STB2026_Пересечения_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                string filePath = Path.Combine(folder, fileName);
                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

                // Подсветка пересекающихся воздуховодов
                using (Transaction tx = new Transaction(doc, "STB2026: Подсветка пересечений"))
                {
                    tx.Start();
                    var overrideSettings = new OverrideGraphicSettings();
                    overrideSettings.SetProjectionLineColor(new Color(255, 140, 0)); // оранжевый

                    foreach (var item in result.Intersections)
                    {
                        try
                        {
                            doc.ActiveView.SetElementOverrides(
                                new ElementId(item.DuctId), overrideSettings);
                        }
                        catch { }
                    }
                    tx.Commit();
                }

                TaskDialog dlg = new TaskDialog("STB2026 — Пересечения со стенами");
                dlg.MainInstruction = $"Найдено пересечений: {result.Intersections.Count}";
                dlg.MainContent =
                    $"Воздуховодов: {result.UniqueDucts}\n" +
                    $"Стен: {result.UniqueWalls}\n\n" +
                    $"Отчёт сохранён:\n{filePath}";
                dlg.FooterText = "Пересекающиеся воздуховоды подсвечены оранжевым.";
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
