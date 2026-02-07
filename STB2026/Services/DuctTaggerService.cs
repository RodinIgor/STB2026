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
        public int GroupsMerged { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    /// <summary>
    /// Группа воздуховодов с одинаковым расходом и сечением —
    /// маркируется одной маркой в середине участка.
    /// </summary>
    internal class DuctSegmentGroup
    {
        public string Key { get; set; } = "";
        public List<Duct> Ducts { get; set; } = new List<Duct>();
    }

    /// <summary>
    /// Сервис автоматической маркировки воздуховодов.
    /// Расставляет марки с расходом, размером и скоростью.
    /// Группирует воздуховоды с одинаковым расходом+сечением —
    /// одна марка на участок, по центру группы.
    /// Предпочитает типоразмер «Размер / Расход_10», иначе — первый доступный.
    /// </summary>
    public class DuctTaggerService
    {
        private readonly Document _doc;
        private readonly View _view;

        // Предпочтительный типоразмер марки
        private const string PreferredTagTypeName = "Размер / Расход_10";

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
            var existingTags = GetExistingTaggedDuctIds();

            // Определяем типоразмер марки
            ElementId tagTypeId = FindPreferredTagType();

            // Группируем воздуховоды по расходу + сечению
            var groups = GroupDuctsByFlowAndSize(ducts);

            // Занятые зоны для защиты от наложения
            var occupiedZones = new List<BoundingBoxXYZ>();

            using (Transaction tx = new Transaction(_doc, "STB2026: Маркировка воздуховодов"))
            {
                tx.Start();

                foreach (var group in groups)
                {
                    try
                    {
                        // Фильтруем уже промаркированные
                        var untaggedDucts = group.Ducts
                            .Where(d => !existingTags.Contains(d.Id))
                            .ToList();

                        if (untaggedDucts.Count == 0)
                        {
                            result.AlreadyTagged += group.Ducts.Count;
                            continue;
                        }

                        // Если все воздуховоды в группе уже промаркированы кроме некоторых
                        int alreadyCount = group.Ducts.Count - untaggedDucts.Count;
                        result.AlreadyTagged += alreadyCount;

                        // Выбираем "представителя" группы — ближайший к центру
                        Duct representativeDuct = FindRepresentativeDuct(untaggedDucts);
                        if (representativeDuct == null)
                        {
                            result.Skipped += untaggedDucts.Count;
                            continue;
                        }

                        // Вычисляем центр группы
                        XYZ groupCenter = CalculateGroupCenter(untaggedDucts);
                        if (groupCenter == null)
                        {
                            result.Skipped += untaggedDucts.Count;
                            continue;
                        }

                        // Определяем направление и смещение
                        var locationCurve = representativeDuct.Location as LocationCurve;
                        if (locationCurve == null)
                        {
                            result.Skipped += untaggedDucts.Count;
                            continue;
                        }

                        XYZ direction = (locationCurve.Curve.GetEndPoint(1) -
                                        locationCurve.Curve.GetEndPoint(0)).Normalize();
                        XYZ offset = GetSmartOffset(direction, _view.UpDirection, _view.RightDirection);
                        XYZ tagPosition = groupCenter + offset * TagOffsetFt;

                        // Проверяем наложение
                        tagPosition = AdjustForCollision(tagPosition, occupiedZones, offset);

                        // Создаём марку на представителе группы
                        var tagRef = new Reference(representativeDuct);
                        IndependentTag tag;

                        if (tagTypeId != ElementId.InvalidElementId)
                        {
                            tag = IndependentTag.Create(
                                _doc, _view.Id, tagRef, false,
                                TagMode.TM_ADDBY_CATEGORY,
                                TagOrientation.Horizontal, tagPosition);

                            if (tag != null)
                                tag.ChangeTypeId(tagTypeId);
                        }
                        else
                        {
                            tag = IndependentTag.Create(
                                _doc, _view.Id, tagRef, false,
                                TagMode.TM_ADDBY_CATEGORY,
                                TagOrientation.Horizontal, tagPosition);
                        }

                        if (tag != null)
                        {
                            result.Tagged++;
                            if (untaggedDucts.Count > 1)
                                result.GroupsMerged += untaggedDucts.Count - 1;

                            var tagBB = tag.get_BoundingBox(_view);
                            if (tagBB != null)
                                occupiedZones.Add(tagBB);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Skipped += group.Ducts.Count;
                        result.Errors.Add($"Группа '{group.Key}': {ex.Message}");
                    }
                }

                tx.Commit();
            }

            return result;
        }

        /// <summary>
        /// Ищет типоразмер марки «Размер / Расход_10».
        /// Если не найден — возвращает InvalidElementId (будет использован типоразмер по умолчанию).
        /// </summary>
        private ElementId FindPreferredTagType()
        {
            var tagTypes = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_DuctTags)
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .ToList();

            // Ищем точное совпадение
            var preferred = tagTypes.FirstOrDefault(t => t.Name == PreferredTagTypeName);
            if (preferred != null)
                return preferred.Id;

            // Ищем частичное совпадение (содержит "Расход")
            var fallback = tagTypes.FirstOrDefault(t =>
                t.Name.Contains("Расход", StringComparison.OrdinalIgnoreCase));
            if (fallback != null)
                return fallback.Id;

            // Не нашли — вернём InvalidElementId, будет использован дефолтный
            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Группирует воздуховоды по расходу + размеру сечения.
        /// Воздуховоды с одинаковым расходом и сечением = один участок.
        /// </summary>
        private List<DuctSegmentGroup> GroupDuctsByFlowAndSize(List<Duct> ducts)
        {
            var groups = new Dictionary<string, DuctSegmentGroup>();

            foreach (var duct in ducts)
            {
                string flow = GetFlowString(duct);
                string size = GetSizeString(duct);
                string system = GetSystemName(duct);
                string key = $"{system}|{size}|{flow}";

                if (!groups.ContainsKey(key))
                    groups[key] = new DuctSegmentGroup { Key = key };

                groups[key].Ducts.Add(duct);
            }

            return groups.Values.ToList();
        }

        /// <summary>
        /// Находит центр группы воздуховодов (среднее всех midpoint).
        /// </summary>
        private XYZ CalculateGroupCenter(List<Duct> ducts)
        {
            var midPoints = new List<XYZ>();

            foreach (var duct in ducts)
            {
                var lc = duct.Location as LocationCurve;
                if (lc != null)
                    midPoints.Add(lc.Curve.Evaluate(0.5, true));
            }

            if (midPoints.Count == 0) return null;

            double avgX = midPoints.Average(p => p.X);
            double avgY = midPoints.Average(p => p.Y);
            double avgZ = midPoints.Average(p => p.Z);

            // Проецируем центр группы на ближайший воздуховод
            XYZ avgPoint = new XYZ(avgX, avgY, avgZ);
            double minDist = double.MaxValue;
            XYZ closestMid = midPoints[0];

            foreach (var mid in midPoints)
            {
                double dist = mid.DistanceTo(avgPoint);
                if (dist < minDist)
                {
                    minDist = dist;
                    closestMid = mid;
                }
            }

            return closestMid;
        }

        /// <summary>
        /// Находит представительный воздуховод (ближайший к центру группы).
        /// </summary>
        private Duct FindRepresentativeDuct(List<Duct> ducts)
        {
            if (ducts.Count == 1) return ducts[0];

            var midPoints = ducts
                .Select(d => new { Duct = d, Mid = (d.Location as LocationCurve)?.Curve.Evaluate(0.5, true) })
                .Where(x => x.Mid != null)
                .ToList();

            if (midPoints.Count == 0) return null;

            double avgX = midPoints.Average(p => p.Mid.X);
            double avgY = midPoints.Average(p => p.Mid.Y);
            double avgZ = midPoints.Average(p => p.Mid.Z);
            XYZ center = new XYZ(avgX, avgY, avgZ);

            return midPoints
                .OrderBy(p => p.Mid.DistanceTo(center))
                .First().Duct;
        }

        private HashSet<ElementId> GetExistingTaggedDuctIds()
        {
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
                    var taggedIds = tag.GetTaggedLocalElementIds();
                    foreach (var id in taggedIds)
                        existingTags.Add(id);
                }
                catch { }
            }
            return existingTags;
        }

        private string GetFlowString(Duct duct)
        {
            Parameter p = duct.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM);
            return p?.AsValueString() ?? "0";
        }

        private string GetSizeString(Duct duct)
        {
            Parameter p = duct.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE);
            return p?.AsString() ?? "N/A";
        }

        private string GetSystemName(Duct duct)
        {
            Parameter p = duct.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            return p?.AsString() ?? "";
        }

        private XYZ GetSmartOffset(XYZ ductDirection, XYZ viewUp, XYZ viewRight)
        {
            double dotUp = Math.Abs(ductDirection.DotProduct(viewUp));
            double dotRight = Math.Abs(ductDirection.DotProduct(viewRight));

            if (dotRight > dotUp)
                return viewUp;
            else
                return viewRight;
        }

        private XYZ AdjustForCollision(XYZ position, List<BoundingBoxXYZ> occupied, XYZ offsetDir)
        {
            const double step = 1.0;
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
                if (!hasCollision) return position;
                position = position + offsetDir * step;
            }
            return position;
        }

        private bool IsInsideBB(XYZ point, BoundingBoxXYZ bb)
        {
            const double margin = 0.5;
            return point.X >= bb.Min.X - margin && point.X <= bb.Max.X + margin &&
                   point.Y >= bb.Min.Y - margin && point.Y <= bb.Max.Y + margin;
        }
    }
}
