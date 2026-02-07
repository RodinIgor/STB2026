using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace STB2026.Commands
{
    /// <summary>
    /// Сбрасывает все цветовые переопределения для MEP-элементов на текущем виде.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ResetColorsCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                View view = doc.ActiveView;

                if (view is ViewSheet || view is ViewSchedule)
                {
                    TaskDialog.Show("STB2026", "Сброс цветов доступен только на видах модели.");
                    return Result.Cancelled;
                }

                int resetCount = 0;

                using (Transaction tx = new Transaction(doc, "STB2026: Сброс цветов"))
                {
                    tx.Start();
                    var emptyOverride = new OverrideGraphicSettings();

                    var ducts = new FilteredElementCollector(doc, view.Id)
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

                    var equipment = new FilteredElementCollector(doc, view.Id)
                        .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                        .WhereElementIsNotElementType()
                        .ToList();

                    foreach (var elem in ducts.Concat(fittings).Concat(terminals).Concat(equipment))
                    {
                        try
                        {
                            var currentOverride = view.GetElementOverrides(elem.Id);
                            if (currentOverride.ProjectionLineColor.IsValid ||
                                currentOverride.SurfaceForegroundPatternColor.IsValid)
                            {
                                view.SetElementOverrides(elem.Id, emptyOverride);
                                resetCount++;
                            }
                        }
                        catch { }
                    }

                    tx.Commit();
                }

                TaskDialog.Show("STB2026",
                    resetCount > 0
                        ? $"Сброшены цвета для {resetCount} элементов."
                        : "Переопределения цветов не обнаружены.");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}