using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace STB2026
{
    [Transaction(TransactionMode.Manual)]
    public class HelloWorldCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                Document doc = uiApp.ActiveUIDocument.Document;

                string info =
                    $"Проект: {doc.Title}\n" +
                    $"Пользователь: {uiApp.Application.Username}\n" +
                    $"Revit: {uiApp.Application.VersionNumber}\n" +
                    $"Дата: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";

                TaskDialog dlg = new TaskDialog("STB2026");
                dlg.MainInstruction = "Привет мир!";
                dlg.MainContent = info;
                dlg.CommonButtons = TaskDialogCommonButtons.Ok;

                dlg.Show();

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