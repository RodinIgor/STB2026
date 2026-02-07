using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.Services
{
    /// <summary>
    /// Точка пересечения воздуховода со стеной
    /// </summary>
    public class IntersectionPoint
    {
        public int DuctId { get; set; }
        public string SystemName { get; set; } = "";
        public string DuctSize { get; set; } = "";
        public int WallId { get; set; }
        public string WallType { get; set; } = "";
        /// <summary>X координата, мм</summary>
        public double X { get; set; }
        /// <summary>Y координата, мм</summary>
        public double Y { get; set; }
        /// <summary>Z координата, мм</summary>
        public double Z { get; set; }
        /// <summary>Отметка от уровня, мм</summary>
        public double Elevation { get; set; }
    }

    /// <summary>
    /// Результат поиска пересечений
    /// </summary>
    public class WallIntersectionResult
    {
        public List<IntersectionPoint> Intersections { get; set; } = new List<IntersectionPoint>();
        public int UniqueDucts => Intersections.Select(i => i.DuctId).Distinct().Count();
        public int UniqueWalls => Intersections.Select(i => i.WallId).Distinct().Count();
    }

    /// <summary>
    /// Сервис поиска пересечений воздуховодов со стенами.
    /// Использует Revit Interference Check (ElementIntersectsElementFilter).
    /// </summary>
    public class WallIntersectorService
    {
        private readonly Document _doc;
        private const double FtToMm = 304.8;

        public WallIntersectorService(Document doc)
        {
            _doc = doc;
        }

        public WallIntersectionResult FindIntersections()
        {
            var result = new WallIntersectionResult();

            // Собираем все воздуховоды в модели
            var ducts = new FilteredElementCollector(_doc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .ToList();

            // Собираем все стены
            var walls = new FilteredElementCollector(_doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            if (ducts.Count == 0 || walls.Count == 0)
                return result;

            foreach (var duct in ducts)
            {
                try
                {
                    // Получаем геометрию (Solid) воздуховода для поиска пересечений
                    var ductSolid = GetElementSolid(duct);
                    if (ductSolid == null)
                        continue;

                    // Фильтр пересечений: стены, пересекающиеся с данным воздуховодом
                    var intersectionFilter = new ElementIntersectsElementFilter(duct);

                    var intersectingWalls = new FilteredElementCollector(_doc)
                        .OfClass(typeof(Wall))
                        .WhereElementIsNotElementType()
                        .WherePasses(intersectionFilter)
                        .Cast<Wall>()
                        .ToList();

                    foreach (var wall in intersectingWalls)
                    {
                        try
                        {
                            // Находим точку пересечения
                            XYZ? intersectionPoint = FindIntersectionPoint(duct, wall);

                            if (intersectionPoint != null)
                            {
                                var point = new IntersectionPoint
                                {
                                    DuctId = duct.Id.IntegerValue,
                                    SystemName = GetSystemName(duct),
                                    DuctSize = GetDuctSize(duct),
                                    WallId = wall.Id.IntegerValue,
                                    WallType = wall.WallType?.Name ?? "N/A",
                                    X = Math.Round(intersectionPoint.X * FtToMm, 0),
                                    Y = Math.Round(intersectionPoint.Y * FtToMm, 0),
                                    Z = Math.Round(intersectionPoint.Z * FtToMm, 0),
                                    Elevation = Math.Round(intersectionPoint.Z * FtToMm, 0)
                                };

                                result.Intersections.Add(point);
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return result;
        }

        /// <summary>
        /// Находит точку пересечения оси воздуховода с геометрией стены
        /// </summary>
        private XYZ? FindIntersectionPoint(Duct duct, Wall wall)
        {
            var locationCurve = duct.Location as LocationCurve;
            if (locationCurve == null)
                return null;

            Line ductLine = locationCurve.Curve as Line;
            if (ductLine == null)
                return null;

            // Получаем Solid стены
            Solid? wallSolid = GetElementSolid(wall);
            if (wallSolid == null)
                return null;

            // Находим пересечение линии воздуховода с гранями стены
            foreach (Face face in wallSolid.Faces)
            {
                if (face is PlanarFace planarFace)
                {
                    // Пересечение линии с плоскостью грани
                    XYZ? hit = LinePlaneIntersection(
                        ductLine.GetEndPoint(0),
                        ductLine.Direction,
                        planarFace.Origin,
                        planarFace.FaceNormal);

                    if (hit != null)
                    {
                        // Проверяем что точка на отрезке воздуховода
                        double param = ductLine.Project(hit).Parameter;
                        if (param >= 0 && param <= 1)
                        {
                            // Проверяем что точка на грани стены
                            var faceResult = face.Project(hit);
                            if (faceResult != null)
                            {
                                return hit;
                            }
                        }
                    }
                }
            }

            // Fallback: средняя точка между BB воздуховода и стены
            var ductBB = duct.get_BoundingBox(null);
            var wallBB = wall.get_BoundingBox(null);

            if (ductBB != null && wallBB != null)
            {
                // Приближённая точка пересечения по BoundingBox
                XYZ ductCenter = (ductBB.Min + ductBB.Max) / 2;
                XYZ wallCenter = (wallBB.Min + wallBB.Max) / 2;

                // Проецируем центр воздуховода на ось стены
                var wallLocation = wall.Location as LocationCurve;
                if (wallLocation?.Curve is Line wallLine)
                {
                    var proj = wallLine.Project(ductCenter);
                    if (proj != null)
                    {
                        return new XYZ(proj.XYZPoint.X, proj.XYZPoint.Y, ductCenter.Z);
                    }
                }

                return new XYZ(
                    (ductCenter.X + wallCenter.X) / 2,
                    (ductCenter.Y + wallCenter.Y) / 2,
                    ductCenter.Z);
            }

            return null;
        }

        /// <summary>
        /// Пересечение луча с плоскостью
        /// </summary>
        private XYZ? LinePlaneIntersection(XYZ linePoint, XYZ lineDir,
                                           XYZ planePoint, XYZ planeNormal)
        {
            double denom = lineDir.DotProduct(planeNormal);
            if (Math.Abs(denom) < 1e-10)
                return null; // параллельны

            double t = (planePoint - linePoint).DotProduct(planeNormal) / denom;
            return linePoint + lineDir * t;
        }

        /// <summary>
        /// Извлекает Solid из геометрии элемента
        /// </summary>
        private Solid? GetElementSolid(Element element)
        {
            var options = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Medium
            };

            GeometryElement geomElem = element.get_Geometry(options);
            if (geomElem == null) return null;

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                    return solid;

                if (geomObj is GeometryInstance instance)
                {
                    foreach (GeometryObject instObj in instance.GetInstanceGeometry())
                    {
                        if (instObj is Solid instSolid && instSolid.Volume > 0)
                            return instSolid;
                    }
                }
            }

            return null;
        }

        private string GetSystemName(Duct duct)
        {
            Parameter param = duct.get_Parameter(
                BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            return param?.AsString() ?? "N/A";
        }

        private string GetDuctSize(Duct duct)
        {
            Parameter param = duct.get_Parameter(
                BuiltInParameter.RBS_CALCULATED_SIZE);
            return param?.AsString() ?? "N/A";
        }
    }
}
