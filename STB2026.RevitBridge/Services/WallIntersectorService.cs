using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.RevitBridge.Services
{
    public class IntersectionPoint
    {
        public int DuctId { get; set; }
        public string SystemName { get; set; } = "";
        public string DuctSize { get; set; } = "";
        public int WallId { get; set; }
        public string WallType { get; set; } = "";
        /// <summary>Ð˜ÑÑ‚Ð¾Ñ‡Ð½Ð¸Ðº ÑÑ‚ÐµÐ½Ñ‹: "" Ð´Ð»Ñ Ñ‚ÐµÐºÑƒÑ‰ÐµÐ³Ð¾ Ð´Ð¾ÐºÑƒÐ¼ÐµÐ½Ñ‚Ð°, Ð¸Ð¼Ñ Ð»Ð¸Ð½ÐºÐ° Ð´Ð»Ñ ÑÐ²ÑÐ·Ð°Ð½Ð½Ð¾Ð³Ð¾</summary>
        public string WallSource { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public double Elevation { get; set; }
    }

    public class WallIntersectionResult
    {
        public List<IntersectionPoint> Intersections { get; set; } = new List<IntersectionPoint>();
        public int UniqueDucts => Intersections.Select(i => i.DuctId).Distinct().Count();
        public int UniqueWalls => Intersections.Select(i => i.WallId).Distinct().Count();
        public int LinkedWallCount { get; set; }
    }

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
            var rawIntersections = new List<IntersectionPoint>();

            var ducts = new FilteredElementCollector(_doc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .ToList();

            if (ducts.Count == 0) return result;

            // 1. Ð¡Ñ‚ÐµÐ½Ñ‹ Ñ‚ÐµÐºÑƒÑ‰ÐµÐ³Ð¾ Ð´Ð¾ÐºÑƒÐ¼ÐµÐ½Ñ‚Ð°
            FindIntersectionsWithLocalWalls(ducts, rawIntersections);

            // 2. Ð¡Ñ‚ÐµÐ½Ñ‹ Ð¸Ð· ÑÐ²ÑÐ·Ð°Ð½Ð½Ñ‹Ñ… Ñ„Ð°Ð¹Ð»Ð¾Ð² (RevitLinkInstance)
            var linkInstances = new FilteredElementCollector(_doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            int linkedWallCount = 0;
            foreach (var linkInstance in linkInstances)
            {
                try
                {
                    Document linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc == null) continue;

                    Transform linkTransform = linkInstance.GetTotalTransform();
                    string linkName = linkDoc.Title ?? linkInstance.Name;

                    int count = FindIntersectionsWithLinkedWalls(
                        ducts, linkDoc, linkTransform, rawIntersections, linkName);
                    linkedWallCount += count;
                }
                catch { }
            }
            result.LinkedWallCount = linkedWallCount;

            // 3. Ð“Ñ€ÑƒÐ¿Ð¿Ð¸Ñ€Ð¾Ð²ÐºÐ°: DuctId+WallId = Ð¾Ð´Ð½Ð° Ð¿Ñ€Ð¾Ñ…Ð¾Ð´ÐºÐ°
            result.Intersections = GroupIntersections(rawIntersections);

            return result;
        }

        private void FindIntersectionsWithLocalWalls(
            List<Duct> ducts, List<IntersectionPoint> results)
        {
            foreach (var duct in ducts)
            {
                try
                {
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
                            XYZ point = FindLocalIntersectionPoint(duct, wall);
                            if (point != null)
                            {
                                results.Add(new IntersectionPoint
                                {
                                    DuctId = duct.Id.IntegerValue,
                                    SystemName = GetSystemName(duct),
                                    DuctSize = GetDuctSize(duct),
                                    WallId = wall.Id.IntegerValue,
                                    WallType = wall.WallType?.Name ?? "N/A",
                                    WallSource = "",
                                    X = Math.Round(point.X * FtToMm, 0),
                                    Y = Math.Round(point.Y * FtToMm, 0),
                                    Z = Math.Round(point.Z * FtToMm, 0),
                                    Elevation = Math.Round(point.Z * FtToMm, 0)
                                });
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }
        }

        private int FindIntersectionsWithLinkedWalls(
            List<Duct> ducts, Document linkDoc, Transform linkTransform,
            List<IntersectionPoint> results, string linkName)
        {
            int foundCount = 0;

            var linkedWalls = new FilteredElementCollector(linkDoc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            if (linkedWalls.Count == 0) return 0;

            foreach (var duct in ducts)
            {
                try
                {
                    var ductLocation = duct.Location as LocationCurve;
                    if (ductLocation == null) continue;

                    Line ductLine = ductLocation.Curve as Line;
                    if (ductLine == null) continue;

                    BoundingBoxXYZ ductBB = duct.get_BoundingBox(null);
                    if (ductBB == null) continue;

                    foreach (var wall in linkedWalls)
                    {
                        try
                        {
                            BoundingBoxXYZ wallBB = wall.get_BoundingBox(null);
                            if (wallBB == null) continue;

                            XYZ wallMin = linkTransform.OfPoint(wallBB.Min);
                            XYZ wallMax = linkTransform.OfPoint(wallBB.Max);

                            if (!BoundingBoxesOverlap(ductBB.Min, ductBB.Max, wallMin, wallMax, 1.0))
                                continue;

                            XYZ hit = FindLinkedIntersectionPoint(ductLine, wall, linkTransform);
                            if (hit != null)
                            {
                                results.Add(new IntersectionPoint
                                {
                                    DuctId = duct.Id.IntegerValue,
                                    SystemName = GetSystemName(duct),
                                    DuctSize = GetDuctSize(duct),
                                    WallId = wall.Id.IntegerValue,
                                    WallType = wall.WallType?.Name ?? "N/A",
                                    WallSource = linkName,
                                    X = Math.Round(hit.X * FtToMm, 0),
                                    Y = Math.Round(hit.Y * FtToMm, 0),
                                    Z = Math.Round(hit.Z * FtToMm, 0),
                                    Elevation = Math.Round(hit.Z * FtToMm, 0)
                                });
                                foundCount++;
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }
            return foundCount;
        }

        private XYZ FindLinkedIntersectionPoint(Line ductLine, Wall wall, Transform linkTransform)
        {
            Solid wallSolid = GetElementSolid(wall);
            if (wallSolid == null) return null;

            foreach (Face face in wallSolid.Faces)
            {
                if (face is PlanarFace planarFace)
                {
                    XYZ faceOrigin = linkTransform.OfPoint(planarFace.Origin);
                    XYZ faceNormal = linkTransform.OfVector(planarFace.FaceNormal);

                    XYZ hit = LinePlaneIntersection(
                        ductLine.GetEndPoint(0), ductLine.Direction,
                        faceOrigin, faceNormal);

                    if (hit != null)
                    {
                        IntersectionResult proj = ductLine.Project(hit);
                        if (proj != null && proj.Parameter >= 0 && proj.Parameter <= ductLine.Length)
                        {
                            XYZ hitInLink = linkTransform.Inverse.OfPoint(hit);
                            IntersectionResult faceProj = face.Project(hitInLink);
                            if (faceProj != null)
                                return hit;
                        }
                    }
                }
            }

            // Fallback
            XYZ ductMid = (ductLine.GetEndPoint(0) + ductLine.GetEndPoint(1)) / 2;
            BoundingBoxXYZ wallBB = wall.get_BoundingBox(null);
            if (wallBB != null)
            {
                XYZ wallCenter = linkTransform.OfPoint((wallBB.Min + wallBB.Max) / 2);
                return new XYZ((ductMid.X + wallCenter.X) / 2, (ductMid.Y + wallCenter.Y) / 2, ductMid.Z);
            }
            return null;
        }

        private List<IntersectionPoint> GroupIntersections(List<IntersectionPoint> raw)
        {
            return raw
                .GroupBy(p => new { p.DuctId, p.WallId })
                .Select(g =>
                {
                    var first = g.First();
                    return new IntersectionPoint
                    {
                        DuctId = first.DuctId,
                        SystemName = first.SystemName,
                        DuctSize = first.DuctSize,
                        WallId = first.WallId,
                        WallType = first.WallType,
                        WallSource = first.WallSource,
                        X = Math.Round(g.Average(p => p.X), 0),
                        Y = Math.Round(g.Average(p => p.Y), 0),
                        Z = Math.Round(g.Average(p => p.Z), 0),
                        Elevation = Math.Round(g.Average(p => p.Elevation), 0)
                    };
                })
                .ToList();
        }

        private bool BoundingBoxesOverlap(XYZ min1, XYZ max1, XYZ min2, XYZ max2, double tolerance)
        {
            double xMin2 = Math.Min(min2.X, max2.X), xMax2 = Math.Max(min2.X, max2.X);
            double yMin2 = Math.Min(min2.Y, max2.Y), yMax2 = Math.Max(min2.Y, max2.Y);
            double zMin2 = Math.Min(min2.Z, max2.Z), zMax2 = Math.Max(min2.Z, max2.Z);

            return !(max1.X + tolerance < xMin2 || min1.X - tolerance > xMax2 ||
                     max1.Y + tolerance < yMin2 || min1.Y - tolerance > yMax2 ||
                     max1.Z + tolerance < zMin2 || min1.Z - tolerance > zMax2);
        }

        private XYZ FindLocalIntersectionPoint(Duct duct, Wall wall)
        {
            var locationCurve = duct.Location as LocationCurve;
            if (locationCurve == null) return null;
            Line ductLine = locationCurve.Curve as Line;
            if (ductLine == null) return null;
            Solid wallSolid = GetElementSolid(wall);
            if (wallSolid == null) return null;

            foreach (Face face in wallSolid.Faces)
            {
                if (face is PlanarFace planarFace)
                {
                    XYZ hit = LinePlaneIntersection(
                        ductLine.GetEndPoint(0), ductLine.Direction,
                        planarFace.Origin, planarFace.FaceNormal);
                    if (hit != null)
                    {
                        IntersectionResult proj = ductLine.Project(hit);
                        if (proj != null && proj.Parameter >= 0 && proj.Parameter <= ductLine.Length)
                        {
                            var faceResult = face.Project(hit);
                            if (faceResult != null) return hit;
                        }
                    }
                }
            }

            var ductBB = duct.get_BoundingBox(null);
            var wallBB = wall.get_BoundingBox(null);
            if (ductBB != null && wallBB != null)
            {
                XYZ ductCenter = (ductBB.Min + ductBB.Max) / 2;
                XYZ wallCenter = (wallBB.Min + wallBB.Max) / 2;
                var wallLocation = wall.Location as LocationCurve;
                if (wallLocation?.Curve is Line wallLine)
                {
                    var proj = wallLine.Project(ductCenter);
                    if (proj != null)
                        return new XYZ(proj.XYZPoint.X, proj.XYZPoint.Y, ductCenter.Z);
                }
                return new XYZ((ductCenter.X + wallCenter.X) / 2, (ductCenter.Y + wallCenter.Y) / 2, ductCenter.Z);
            }
            return null;
        }

        private XYZ LinePlaneIntersection(XYZ linePoint, XYZ lineDir, XYZ planePoint, XYZ planeNormal)
        {
            double denom = lineDir.DotProduct(planeNormal);
            if (Math.Abs(denom) < 1e-10) return null;
            double t = (planePoint - linePoint).DotProduct(planeNormal) / denom;
            return linePoint + lineDir * t;
        }

        private Solid GetElementSolid(Element element)
        {
            var options = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Medium };
            GeometryElement geomElem = element.get_Geometry(options);
            if (geomElem == null) return null;
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid solid && solid.Volume > 0) return solid;
                if (geomObj is GeometryInstance instance)
                    foreach (GeometryObject instObj in instance.GetInstanceGeometry())
                        if (instObj is Solid instSolid && instSolid.Volume > 0) return instSolid;
            }
            return null;
        }

        private string GetSystemName(Duct duct)
        {
            Parameter param = duct.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            return param?.AsString() ?? "N/A";
        }

        private string GetDuctSize(Duct duct)
        {
            Parameter param = duct.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE);
            return param?.AsString() ?? "N/A";
        }
    }
}
