using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Text;

namespace STB2026.RevitBridge.Commands
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

                var service = new STB2026.RevitBridge.Services.WallIntersectorService(doc);
                var result = service.FindIntersections();

                if (result.Intersections.Count == 0)
                {
                    TaskDialog.Show("STB2026", "Пересечения воздуховодов со стенами не обнаружены.\n" +
                        "(Проверены стены текущего документа и связанных файлов)");
                    return Result.Succeeded;
                }

                // CSV с колонкой источника стены
                var sb = new StringBuilder();
                sb.AppendLine("Воздуховод;Система;Размер;Стена;Тип стены;Источник;X (мм);Y (мм);Z (мм);Отметка (мм)");

                foreach (var item in result.Intersections)
                {
                    string source = string.IsNullOrEmpty(item.WallSource) ? "Текущий" : item.WallSource;
                    sb.AppendLine(
                        $"{item.DuctId};{item.SystemName};{item.DuctSize};" +
                        $"{item.WallId};{item.WallType};{source};" +
                        $"{item.X:F0};{item.Y:F0};{item.Z:F0};{item.Elevation:F0}");
                }

                string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"STB2026_Пересечения_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                string filePath = Path.Combine(folder, fileName);
                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

                // Подсветка пересекающихся воздуховодов
                using (Transaction tx = new Transaction(doc, "STB2026: Подсветка пересечений"))
                {
                    tx.Start();
                    var overrideSettings = new OverrideGraphicSettings();
                    overrideSettings.SetProjectionLineColor(new Color(255, 140, 0));

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

                // Диалог результатов
                string linkedInfo = result.LinkedWallCount > 0
                    ? $"\n(в т.ч. из связанных файлов: {result.LinkedWallCount})"
                    : "";

                TaskDialog dlg = new TaskDialog("STB2026 — Пересечения со стенами");
                dlg.MainInstruction = $"Найдено проходок: {result.Intersections.Count}";
                dlg.MainContent =
                    $"Воздуховодов: {result.UniqueDucts}\n" +
                    $"Стен: {result.UniqueWalls}{linkedInfo}\n\n" +
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
