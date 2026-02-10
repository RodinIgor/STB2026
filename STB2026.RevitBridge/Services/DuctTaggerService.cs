using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.RevitBridge.Services
{
    public class TaggingResult
    {
        public int Tagged { get; set; }
        public int AlreadyTagged { get; set; }
        public int Skipped { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    public class DuctTaggerService
    {
        private readonly Document _doc;
        private readonly View _view;
        private const double TagOffsetFt = 1.5;

        public DuctTaggerService(Document doc, View view)
        {
            _doc = doc;
            _view = view;
        }

        public TaggingResult TagAllDucts()
        {
            var result = new TaggingResult();

            var ducts = new FilteredElementCollector(_doc, _view.Id)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .ToList();

            if (ducts.Count == 0)
            {
                result.Errors.Add("ÐÐ° Ñ‚ÐµÐºÑƒÑ‰ÐµÐ¼ Ð²Ð¸Ð´Ðµ Ð½ÐµÑ‚ Ð²Ð¾Ð·Ð´ÑƒÑ…Ð¾Ð²Ð¾Ð´Ð¾Ð².");
                return result;
            }

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

            var occupiedZones = new List<BoundingBoxXYZ>();

            using (Transaction tx = new Transaction(_doc, "STB2026: ÐœÐ°Ñ€ÐºÐ¸Ñ€Ð¾Ð²ÐºÐ° Ð²Ð¾Ð·Ð´ÑƒÑ…Ð¾Ð²Ð¾Ð´Ð¾Ð²"))
            {
                tx.Start();

                foreach (var duct in ducts)
                {
                    try
                    {
                        if (existingTags.Contains(duct.Id))
                        {
                            result.AlreadyTagged++;
                            continue;
                        }

                        var locationCurve = duct.Location as LocationCurve;
                        if (locationCurve == null)
                        {
                            result.Skipped++;
                            result.Errors.Add($"ID {duct.Id}: Ð½ÐµÑ‚ ÐºÑ€Ð¸Ð²Ð¾Ð¹ Ñ€Ð°ÑÐ¿Ð¾Ð»Ð¾Ð¶ÐµÐ½Ð¸Ñ");
                            continue;
                        }

                        XYZ midPoint = locationCurve.Curve.Evaluate(0.5, true);
                        XYZ direction = (locationCurve.Curve.GetEndPoint(1) - locationCurve.Curve.GetEndPoint(0)).Normalize();
                        XYZ viewUp = _view.UpDirection;
                        XYZ viewRight = _view.RightDirection;
                        XYZ offset = GetSmartOffset(direction, viewUp, viewRight);
                        XYZ tagPosition = midPoint + offset * TagOffsetFt;
                        tagPosition = AdjustForCollision(tagPosition, occupiedZones, offset);

                        var tagRef = new Reference(duct);
                        IndependentTag tag = IndependentTag.Create(
                            _doc, _view.Id, tagRef, false,
                            TagMode.TM_ADDBY_CATEGORY,
                            TagOrientation.Horizontal,
                            tagPosition);

                        if (tag != null)
                        {
                            result.Tagged++;
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
