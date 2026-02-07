using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.Services
{
    /// <summary>
    /// Результат маркировки
    /// </summary>
    public class TaggingResult
    {
        public int Tagged { get; set; }
        public int AlreadyTagged { get; set; }
        public int Skipped { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    /// <summary>
    /// Сервис автоматической маркировки воздуховодов.
    /// Расставляет марки с расходом, размером и скоростью.
    /// Включает простую защиту от наложения.
    /// </summary>
    public class DuctTaggerService
    {
        private readonly Document _doc;
        private readonly View _view;

        // Смещение марки от центра воздуховода (футы)
        private const double TagOffsetFt = 1.5;

        public DuctTaggerService(Document doc, View view)
        {
            _doc = doc;
            _view = view;
        }

        public TaggingResult TagAllDucts()
        {
            var result = new TaggingResult();

            // Собираем все воздуховоды на текущем виде
            var ducts = new FilteredElementCollector(_doc, _view.Id)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .ToList();

            if (ducts.Count == 0)
            {
                result.Errors.Add("На текущем виде нет воздуховодов.");
                return result;
            }

            // Собираем существующие марки воздуховодов
            var existingTagElements = new FilteredElementCollector(_doc, _view.Id)
                .OfClass(typeof(IndependentTag))
                .WhereElementIsNotElementType()
                .Cast<IndependentTag>()
                .ToList();

            var existingTags = new HashSet<ElementId>();
            foreach (var tag in existingTagElements)
            {
                try
                {
                    // Revit 2025: используем GetTaggedLocalElementIds()
                    var taggedIds = tag.GetTaggedLocalElementIds();
                    foreach (var id in taggedIds)
                        existingTags.Add(id);
                }
                catch
                {
                    // Fallback для совместимости
                }
            }

            // Занятые зоны для защиты от наложения
            var occupiedZones = new List<BoundingBoxXYZ>();

            using (Transaction tx = new Transaction(_doc, "STB2026: Маркировка воздуховодов"))
            {
                tx.Start();

                foreach (var duct in ducts)
                {
                    try
                    {
                        // Пропускаем уже промаркированные
                        if (existingTags.Contains(duct.Id))
                        {
                            result.AlreadyTagged++;
                            continue;
                        }

                        // Получаем центральную точку воздуховода
                        var locationCurve = duct.Location as LocationCurve;
                        if (locationCurve == null)
                        {
                            result.Skipped++;
                            result.Errors.Add($"ID {duct.Id}: нет кривой расположения");
                            continue;
                        }

                        XYZ midPoint = locationCurve.Curve.Evaluate(0.5, true);

                        // Определяем направление воздуховода и смещение марки
                        XYZ direction = (locationCurve.Curve.GetEndPoint(1) -
                                        locationCurve.Curve.GetEndPoint(0)).Normalize();

                        // Перпендикуляр для смещения (в плоскости вида)
                        XYZ viewUp = _view.UpDirection;
                        XYZ viewRight = _view.RightDirection;

                        // Смещаем марку перпендикулярно оси воздуховода
                        XYZ offset = GetSmartOffset(direction, viewUp, viewRight);
                        XYZ tagPosition = midPoint + offset * TagOffsetFt;

                        // Проверяем наложение
                        tagPosition = AdjustForCollision(tagPosition, occupiedZones, offset);

                        // Создаём марку
                        var tagRef = new Reference(duct);
                        IndependentTag tag = IndependentTag.Create(
                            _doc,
                            _view.Id,
                            tagRef,
                            false,       // leader
                            TagMode.TM_ADDBY_CATEGORY,
                            TagOrientation.Horizontal,
                            tagPosition
                        );

                        if (tag != null)
                        {
                            result.Tagged++;

                            // Добавляем зону марки в занятые
                            var tagBB = tag.get_BoundingBox(_view);
                            if (tagBB != null)
                                occupiedZones.Add(tagBB);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Skipped++;
                        result.Errors.Add($"ID {duct.Id}: {ex.Message}");
                    }
                }

                tx.Commit();
            }

            return result;
        }

        /// <summary>
        /// Определяет оптимальное направление смещения марки
        /// </summary>
        private XYZ GetSmartOffset(XYZ ductDirection, XYZ viewUp, XYZ viewRight)
        {
            // Для горизонтальных воздуховодов — смещаем вверх
            // Для вертикальных (на плане) — смещаем вправо
            double dotUp = Math.Abs(ductDirection.DotProduct(viewUp));
            double dotRight = Math.Abs(ductDirection.DotProduct(viewRight));

            if (dotRight > dotUp)
            {
                // Воздуховод идёт горизонтально → смещаем вверх
                return viewUp;
            }
            else
            {
                // Воздуховод идёт вертикально → смещаем вправо
                return viewRight;
            }
        }

        /// <summary>
        /// Корректирует позицию марки при наложении на существующие
        /// </summary>
        private XYZ AdjustForCollision(XYZ position, List<BoundingBoxXYZ> occupied, XYZ offsetDir)
        {
            const double step = 1.0; // шаг сдвига в футах
            const int maxAttempts = 5;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                bool hasCollision = false;

                foreach (var bb in occupied)
                {
                    if (IsInsideBB(position, bb))
                    {
                        hasCollision = true;
                        break;
                    }
                }

                if (!hasCollision)
                    return position;

                // Сдвигаем дальше
                position = position + offsetDir * step;
            }

            return position;
        }

        /// <summary>
        /// Проверяет, попадает ли точка в BoundingBox (с запасом)
        /// </summary>
        private bool IsInsideBB(XYZ point, BoundingBoxXYZ bb)
        {
            const double margin = 0.5; // запас в футах

            return point.X >= bb.Min.X - margin && point.X <= bb.Max.X + margin &&
                   point.Y >= bb.Min.Y - margin && point.Y <= bb.Max.Y + margin;
        }
    }
}
