using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// manage_sheets — операции с листами и видовыми экранами.
    /// 
    /// Действия:
    ///   create_sheet   — создать лист (sheet_name, sheet_number, title_block_name, title_block_type)
    ///   place_view     — разместить вид на листе (sheet_id, view_id, x_mm, y_mm)
    ///   list_viewports — список видовых экранов на листе (sheet_id)
    /// </summary>
    internal static class SheetHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null)
                return new { error = "Нет открытого документа" };

            var doc = uidoc.Document;
            string action = p.TryGetValue("action", out var a) ? a?.ToString() ?? "" : "";
            string dataStr = p.TryGetValue("data", out var d) ? d?.ToString() ?? "{}" : "{}";

            JObject data;
            try { data = JObject.Parse(dataStr); }
            catch { return new { error = $"Невалидный JSON в data: {dataStr}" }; }

            switch (action.ToLowerInvariant())
            {
                case "create_sheet":    return CreateSheet(doc, data);
                case "place_view":      return PlaceView(doc, data);
                case "list_viewports":  return ListViewports(doc, data);
                case "move_viewport":   return MoveViewport(doc, data);
                case "align_viewports": return AlignViewports(doc, data);
                default:
                    return new
                    {
                        error = $"Неизвестное действие: '{action}'",
                        available = new[] { "create_sheet", "place_view", "list_viewports",
                                            "move_viewport", "align_viewports" }
                    };
            }
        }

        // ═══ create_sheet ═══
        private static object CreateSheet(Document doc, JObject data)
        {
            string sheetName = data.Value<string>("sheet_name") ?? "";
            string sheetNumber = data.Value<string>("sheet_number") ?? "";
            string tbFamily = data.Value<string>("title_block_name") ?? "";
            string tbType = data.Value<string>("title_block_type") ?? "";

            if (string.IsNullOrWhiteSpace(sheetName))
                return new { error = "Укажите sheet_name (имя листа)" };
            if (string.IsNullOrWhiteSpace(sheetNumber))
                return new { error = "Укажите sheet_number (номер листа)" };

            // Проверка уникальности номера
            var existingSheet = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .FirstOrDefault(s => s.SheetNumber == sheetNumber);

            if (existingSheet != null)
                return new { error = $"Лист с номером '{sheetNumber}' уже существует (id: {existingSheet.Id.IntegerValue})" };

            // Поиск основной надписи
            var titleBlockId = FindTitleBlock(doc, tbFamily, tbType);
            if (titleBlockId == ElementId.InvalidElementId)
            {
                var available = GetAvailableTitleBlocks(doc);
                return new { error = $"Основная надпись не найдена: '{tbFamily}' / '{tbType}'", available_title_blocks = available };
            }

            ViewSheet sheet;
            using (var trans = new Transaction(doc, $"STB2026: Создание листа {sheetNumber}"))
            {
                trans.Start();
                sheet = ViewSheet.Create(doc, titleBlockId);
                sheet.SheetNumber = sheetNumber;
                sheet.Name = sheetName;
                trans.Commit();
            }

            return new
            {
                action = "create_sheet",
                sheet_id = sheet.Id.IntegerValue,
                sheet_number = sheet.SheetNumber,
                sheet_name = sheet.Name,
                title_block = doc.GetElement(titleBlockId)?.Name ?? ""
            };
        }

        // ═══ place_view ═══
        private static object PlaceView(Document doc, JObject data)
        {
            int sheetIdInt = data.Value<int?>("sheet_id") ?? -1;
            int viewIdInt = data.Value<int?>("view_id") ?? -1;
            double xMm = data.Value<double?>("x_mm") ?? 0;
            double yMm = data.Value<double?>("y_mm") ?? 0;

            if (sheetIdInt <= 0)
                return new { error = "Укажите sheet_id" };
            if (viewIdInt <= 0)
                return new { error = "Укажите view_id" };

            var sheet = doc.GetElement(new ElementId(sheetIdInt)) as ViewSheet;
            if (sheet == null)
                return new { error = $"Лист id={sheetIdInt} не найден" };

            var view = doc.GetElement(new ElementId(viewIdInt)) as View;
            if (view == null)
                return new { error = $"Вид id={viewIdInt} не найден" };

            // Проверка: вид уже размещён на другом листе?
            if (Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id) == false)
                return new { error = $"Вид '{view.Name}' нельзя разместить на листе (уже размещён или является листом/спецификацией)" };

            // Координата на листе (из мм в футы)
            var point = new XYZ(xMm / 304.8, yMm / 304.8, 0);

            // Если координаты нулевые — центр листа
            if (Math.Abs(xMm) < 1 && Math.Abs(yMm) < 1)
            {
                point = GetSheetCenter(doc, sheet);
            }

            Viewport viewport;
            using (var trans = new Transaction(doc, $"STB2026: Вид на лист {sheet.SheetNumber}"))
            {
                trans.Start();
                viewport = Viewport.Create(doc, sheet.Id, view.Id, point);
                trans.Commit();
            }

            return new
            {
                action = "place_view",
                viewport_id = viewport.Id.IntegerValue,
                sheet_id = sheetIdInt,
                sheet_number = sheet.SheetNumber,
                view_name = view.Name,
                position_mm = new { x = xMm, y = yMm }
            };
        }

        // ═══ list_viewports ═══
        private static object ListViewports(Document doc, JObject data)
        {
            int sheetIdInt = data.Value<int?>("sheet_id") ?? -1;

            if (sheetIdInt <= 0)
                return new { error = "Укажите sheet_id" };

            var sheet = doc.GetElement(new ElementId(sheetIdInt)) as ViewSheet;
            if (sheet == null)
                return new { error = $"Лист id={sheetIdInt} не найден" };

            var vpIds = sheet.GetAllViewports();
            var viewports = new List<object>();

            foreach (var vpId in vpIds)
            {
                var vp = doc.GetElement(vpId) as Viewport;
                if (vp == null) continue;

                var view = doc.GetElement(vp.ViewId) as View;
                var center = vp.GetBoxCenter();

                viewports.Add(new
                {
                    viewport_id = vp.Id.IntegerValue,
                    view_id = vp.ViewId.IntegerValue,
                    view_name = view?.Name ?? "",
                    view_type = view?.ViewType.ToString() ?? "",
                    center_x_mm = Math.Round(center.X * 304.8, 1),
                    center_y_mm = Math.Round(center.Y * 304.8, 1)
                });
            }

            return new
            {
                action = "list_viewports",
                sheet_id = sheetIdInt,
                sheet_number = sheet.SheetNumber,
                sheet_name = sheet.Name,
                viewports_count = viewports.Count,
                viewports
            };
        }

        // ═══ move_viewport ═══
        private static object MoveViewport(Document doc, JObject data)
        {
            int vpIdInt = data.Value<int?>("viewport_id") ?? -1;
            double xMm = data.Value<double?>("center_x_mm") ?? double.NaN;
            double yMm = data.Value<double?>("center_y_mm") ?? double.NaN;

            if (vpIdInt <= 0)
                return new { error = "Укажите viewport_id" };
            if (double.IsNaN(xMm) || double.IsNaN(yMm))
                return new { error = "Укажите center_x_mm и center_y_mm" };

            var vp = doc.GetElement(new ElementId(vpIdInt)) as Viewport;
            if (vp == null)
                return new { error = $"Viewport id={vpIdInt} не найден" };

            var newCenter = new XYZ(xMm / 304.8, yMm / 304.8, 0);

            using (var trans = new Transaction(doc, "STB2026: move viewport"))
            {
                trans.Start();
                vp.SetBoxCenter(newCenter);
                trans.Commit();
            }

            var actualCenter = vp.GetBoxCenter();
            return new
            {
                action = "move_viewport",
                viewport_id = vpIdInt,
                center_x_mm = Math.Round(actualCenter.X * 304.8, 1),
                center_y_mm = Math.Round(actualCenter.Y * 304.8, 1)
            };
        }

        // ═══ align_viewports ═══
        /// <summary>
        /// Выравнивание всех viewport'ов на листе.
        /// horizontal: "left", "center", "right"
        /// vertical: "top", "center", "bottom"
        /// margin_mm: отступ от рамки (по умолчанию 5)
        /// </summary>
        private static object AlignViewports(Document doc, JObject data)
        {
            int sheetIdInt = data.Value<int?>("sheet_id") ?? -1;
            string horizontal = data.Value<string>("horizontal") ?? "left";
            string vertical = data.Value<string>("vertical") ?? "top";
            double marginMm = data.Value<double?>("margin_mm") ?? 5;

            if (sheetIdInt <= 0)
                return new { error = "Укажите sheet_id" };

            var sheet = doc.GetElement(new ElementId(sheetIdInt)) as ViewSheet;
            if (sheet == null)
                return new { error = $"Лист id={sheetIdInt} не найден" };

            // Получаем рабочую область листа из основной надписи
            var workArea = GetWorkArea(doc, sheet);
            if (workArea == null)
                return new { error = "Не удалось определить рабочую область листа" };

            var vpIds = sheet.GetAllViewports();
            if (vpIds.Count == 0)
                return new { error = "На листе нет видовых экранов" };

            var results = new List<object>();

            using (var trans = new Transaction(doc, "STB2026: align viewports"))
            {
                trans.Start();

                foreach (var vpId in vpIds)
                {
                    var vp = doc.GetElement(vpId) as Viewport;
                    if (vp == null) continue;

                    var outline = vp.GetBoxOutline();
                    var vpCenter = vp.GetBoxCenter();

                    // Размеры viewport'а
                    double vpWidth = (outline.MaximumPoint.X - outline.MinimumPoint.X) * 304.8;
                    double vpHeight = (outline.MaximumPoint.Y - outline.MinimumPoint.Y) * 304.8;

                    // Рассчитываем половинки
                    double halfW = vpWidth / 2;
                    double halfH = vpHeight / 2;

                    double newCenterX, newCenterY;

                    // Горизонтальное выравнивание
                    switch (horizontal.ToLowerInvariant())
                    {
                        case "left":
                            newCenterX = workArea.Value.Left + marginMm + halfW;
                            break;
                        case "right":
                            newCenterX = workArea.Value.Right - marginMm - halfW;
                            break;
                        default: // center
                            newCenterX = (workArea.Value.Left + workArea.Value.Right) / 2;
                            break;
                    }

                    // Вертикальное выравнивание
                    switch (vertical.ToLowerInvariant())
                    {
                        case "top":
                            newCenterY = workArea.Value.Top - marginMm - halfH;
                            break;
                        case "bottom":
                            newCenterY = workArea.Value.Bottom + marginMm + halfH;
                            break;
                        default: // center
                            newCenterY = (workArea.Value.Top + workArea.Value.Bottom) / 2;
                            break;
                    }

                    var newCenter = new XYZ(newCenterX / 304.8, newCenterY / 304.8, 0);
                    vp.SetBoxCenter(newCenter);

                    var view = doc.GetElement(vp.ViewId) as View;
                    results.Add(new
                    {
                        viewport_id = vp.Id.IntegerValue,
                        view_name = view?.Name ?? "",
                        center_x_mm = Math.Round(newCenterX, 1),
                        center_y_mm = Math.Round(newCenterY, 1),
                        size_w_mm = Math.Round(vpWidth, 1),
                        size_h_mm = Math.Round(vpHeight, 1)
                    });
                }

                trans.Commit();
            }

            return new
            {
                action = "align_viewports",
                sheet_id = sheetIdInt,
                horizontal,
                vertical,
                margin_mm = marginMm,
                work_area = new
                {
                    left = workArea.Value.Left,
                    top = workArea.Value.Top,
                    right = workArea.Value.Right,
                    bottom = workArea.Value.Bottom
                },
                viewports = results
            };
        }

        // ═══ Helpers ═══

        /// <summary>
        /// Рабочая область листа (мм) по основной надписи.
        /// 
        /// ADSK_ОсновнаяНадпись: origin (0,0) = правый нижний угол основного штампа.
        /// Для A3 альбомная (420×297): лист тянется влево и вверх от origin.
        /// 
        /// Координаты листа (мм):
        ///   Правый край листа  = origin.X + marginRight = +5
        ///   Левый край листа   = origin.X - sheetWidth + marginRight = -420 + 5 = -415
        ///   Левый край рамки   = -sheetWidth + marginRight + marginLeft = -420 + 5 + 20 = -395
        ///   Верхний край рамки = origin.Y + sheetHeight - marginTop = 0 + 297 - 5 = 292
        ///   Нижний край рамки  = origin.Y + marginBottom = 5
        ///   Верх штампа        = origin.Y + stampHeight = 55
        /// </summary>
        private static WorkAreaRect? GetWorkArea(Document doc, ViewSheet sheet)
        {
            var titleBlocks = new FilteredElementCollector(doc, sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .ToList();

            if (titleBlocks.Count == 0) return null;

            var tb = titleBlocks[0];

            double sheetWidth  = GetParamDouble(tb, "Ширина_Реальная", 420);  // мм (A3=420)
            double sheetHeight = GetParamDouble(tb, "Высота_Реальная", 297);  // мм (A3=297)
            double marginLeft  = GetParamDouble(tb, "Поле_слева", 20);        // мм подшивка
            double marginTop   = GetParamDouble(tb, "Поле_сверху", 5);        // мм
            double marginRight = 5;
            double marginBottom = 5;
            double stampHeight = 55; // высота штампа Форма 3

            // Origin основной надписи
            var tbLoc = tb.Location as LocationPoint;
            double ox = tbLoc != null ? tbLoc.Point.X * 304.8 : 0;
            double oy = tbLoc != null ? tbLoc.Point.Y * 304.8 : 0;

            // Origin = правый нижний угол штампа → лист тянется влево и вверх
            return new WorkAreaRect
            {
                Left   = ox - sheetWidth + marginRight + marginLeft,  // -420+5+20 = -395
                Right  = ox - marginRight,                             // -5
                Top    = oy + sheetHeight - marginTop,                 // 292
                Bottom = oy + stampHeight                              // 55 (над штампом)
            };
        }

        private static double GetParamDouble(Element el, string paramName, double fallback)
        {
            foreach (Parameter p in el.Parameters)
            {
                if (p?.Definition?.Name == paramName && p.HasValue && p.StorageType == StorageType.Double)
                    return p.AsDouble() * 304.8; // футы → мм
            }
            return fallback;
        }

        internal struct WorkAreaRect
        {
            public double Left, Top, Right, Bottom;
        }

        /// <summary>Поиск типоразмера основной надписи.</summary>
        private static ElementId FindTitleBlock(Document doc, string familyName, string typeName)
        {
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>();

            // Если указаны оба — точный поиск
            if (!string.IsNullOrWhiteSpace(familyName) && !string.IsNullOrWhiteSpace(typeName))
            {
                var exact = collector.FirstOrDefault(fs =>
                    fs.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                    fs.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
                if (exact != null) return exact.Id;
            }

            // Если указано только семейство — первый типоразмер
            if (!string.IsNullOrWhiteSpace(familyName))
            {
                var byFamily = collector.FirstOrDefault(fs =>
                    fs.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                if (byFamily != null) return byFamily.Id;

                // Частичное совпадение
                var partial = collector.FirstOrDefault(fs =>
                    fs.Family.Name.IndexOf(familyName, StringComparison.OrdinalIgnoreCase) >= 0);
                if (partial != null) return partial.Id;
            }

            // Если указан только тип
            if (!string.IsNullOrWhiteSpace(typeName))
            {
                var byType = collector.FirstOrDefault(fs =>
                    fs.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
                if (byType != null) return byType.Id;
            }

            // Fallback: первая доступная основная надпись (не начальный вид)
            var fallback = collector
                .FirstOrDefault(fs => !fs.Family.Name.Contains("Начальный", StringComparison.OrdinalIgnoreCase));

            return fallback?.Id ?? ElementId.InvalidElementId;
        }

        /// <summary>Получить центр рабочей области листа.</summary>
        private static XYZ GetSheetCenter(Document doc, ViewSheet sheet)
        {
            // Через рабочую область — самый точный способ
            var workArea = GetWorkArea(doc, sheet);
            if (workArea != null)
            {
                double cx = (workArea.Value.Left + workArea.Value.Right) / 2;
                double cy = (workArea.Value.Top + workArea.Value.Bottom) / 2;
                return new XYZ(cx / 304.8, cy / 304.8, 0);
            }

            // Fallback: BoundingBox основной надписи
            var titleBlocks = new FilteredElementCollector(doc, sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .ToList();

            if (titleBlocks.Count > 0)
            {
                var bb = titleBlocks[0].get_BoundingBox(sheet);
                if (bb != null)
                    return (bb.Min + bb.Max) / 2;
            }

            // Fallback: ADSK origin(0,0) = правый нижний штампа → центр A3 альбомной
            return new XYZ(-420.0 / 2 / 304.8, 297.0 / 2 / 304.8, 0);
        }

        /// <summary>Список доступных основных надписей.</summary>
        private static List<object> GetAvailableTitleBlocks(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .Select(fs => (object)new
                {
                    family = fs.Family.Name,
                    type = fs.Name,
                    id = fs.Id.IntegerValue
                })
                .ToList();
        }
    }
}
