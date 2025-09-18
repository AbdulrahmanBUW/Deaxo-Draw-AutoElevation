using System;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Extract element geometry properties: origin, vector, width, height, depth, type name
    /// Mirrors the logic of your Python ElementProperties.
    /// </summary>
    public class ElementProperties
    {
        public Document Doc { get; private set; }
        public Element Element { get; private set; }

        public XYZ Origin { get; private set; }
        public XYZ Vector { get; private set; }
        public double Width { get; private set; }
        public double Height { get; private set; }
        public double Depth { get; private set; }

        public bool IsValid { get; private set; } = false;
        public string TypeName { get; private set; }

        public ElementProperties(Document doc, Element el)
        {
            Doc = doc;
            Element = el;
            try
            {
                TypeName = GetElementTypeName(el);
                if (el is Wall) GetWallProperties(el as Wall);
                else GetGenericProperties(el);

                IsValid = (Origin != null && Vector != null && Width > 0 && Height > 0);
            }
            catch
            {
                IsValid = false;
            }
        }

        private string GetElementTypeName(Element el)
        {
            try
            {
                var elType = Doc.GetElement(el.GetTypeId());
                if (elType != null)
                {
                    Parameter p = elType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    if (p != null && p.HasValue) return p.AsString();
                }
            }
            catch { }
            return el.Name;
        }

        private void GetWallProperties(Wall w)
        {
            var locCurve = w.Location as LocationCurve;
            if (locCurve == null) return;
            var curve = locCurve.Curve;
            var p0 = curve.GetEndPoint(0);
            var p1 = curve.GetEndPoint(1);
            Vector = p1 - p0;
            Width = Vector.GetLength();

            // height param
            try
            {
                var p = w.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (p != null && p.StorageType == StorageType.Double)
                    Height = p.AsDouble();
            }
            catch { }

            var bb = w.get_BoundingBox(null);
            Origin = (bb.Min + bb.Max) / 2;
            Depth = bb.Max.Y - bb.Min.Y;
        }

        private void GetGenericProperties(Element el)
        {
            var elType = Doc.GetElement(el.GetTypeId());
            BoundingBoxXYZ bb = el.get_BoundingBox(null);
            BoundingBoxXYZ bbType = null;
            if (elType != null)
                bbType = elType.get_BoundingBox(null);

            if (el is FamilyInstance fi)
            {
                var famPlace = fi.Symbol.Family.FamilyPlacementType;
                if (famPlace == FamilyPlacementType.OneLevelBased || famPlace == FamilyPlacementType.TwoLevelsBased || famPlace == FamilyPlacementType.WorkPlaneBased)
                {
                    Origin = (bb.Min + bb.Max) / 2;
                    if (bbType != null)
                    {
                        Width = Math.Abs(bbType.Max.X - bbType.Min.X);
                        Height = Math.Abs(bbType.Max.Z - bbType.Min.Z);
                        Depth = Math.Abs(bbType.Max.Y - bbType.Min.Y);
                        // define local vector along X of family type and rotate by instance rotation if present
                        var ptStart = new XYZ(bbType.Min.X, (bbType.Min.Y + bbType.Max.Y) / 2, bbType.Min.Z);
                        var ptEnd = new XYZ(bbType.Max.X, (bbType.Min.Y + bbType.Max.Y) / 2, bbType.Min.Z);
                        var vec = ptEnd - ptStart;

                        try
                        {
                            var locPoint = fi.Location as LocationPoint;
                            if (locPoint != null)
                            {
                                var rot = locPoint.Rotation;
                                Vector = GeometryHelpers.RotateVector(vec, rot);
                            }
                            else
                            {
                                Vector = vec;
                            }
                        }
                        catch { Vector = vec; }

                        return;
                    }
                }
                else if (famPlace == FamilyPlacementType.CurveBased || famPlace == FamilyPlacementType.CurveDrivenStructural)
                {
                    var lc = fi.Location as LocationCurve;
                    if (lc != null)
                    {
                        var curve = lc.Curve;
                        var p0 = curve.GetEndPoint(0);
                        var p1 = curve.GetEndPoint(1);
                        if (p0.Z != p1.Z)
                        {
                            p1 = new XYZ(p1.X, p1.Y, p0.Z);
                        }
                        Vector = p1 - p0;
                        Width = Vector.GetLength();
                        var bbEl = fi.get_BoundingBox(null);
                        Height = Math.Abs(bbEl.Max.Z - bbEl.Min.Z);
                        Origin = (bbEl.Min + bbEl.Max) / 2;
                        return;
                    }
                }
                else if (famPlace == FamilyPlacementType.OneLevelBasedHosted)
                {
                    var host = fi.Host;
                    if (host is Wall wh)
                    {
                        var loc = wh.Location as LocationCurve;
                        var curve = loc?.Curve;
                        if (curve != null)
                        {
                            var p0 = curve.GetEndPoint(0);
                            var p1 = curve.GetEndPoint(1);
                            Vector = p1 - p0;
                            // respect facing flipped
                            try
                            {
                                if (fi.FacingFlipped) Vector = -Vector;
                            }
                            catch { }
                            Origin = (bb.Min + bb.Max) / 2;
                            if (bbType != null)
                            {
                                Width = Math.Abs(bbType.Max.X - bbType.Min.X);
                                Height = Math.Abs(bbType.Max.Z - bbType.Min.Z);
                                return;
                            }
                        }
                    }
                }
            }

            // fallback for non-family instances
            if (bbType != null)
            {
                Origin = (bb.Min + bb.Max) / 2;
                Width = Math.Abs(bbType.Max.X - bbType.Min.X);
                Height = Math.Abs(bbType.Max.Z - bbType.Min.Z);
                Depth = Math.Abs(bbType.Max.Y - bbType.Min.Y);
                var ptS = new XYZ(bbType.Min.X, (bbType.Min.Y + bbType.Max.Y) / 2, bbType.Min.Z);
                var ptE = new XYZ(bbType.Max.X, (bbType.Min.Y + bbType.Max.Y) / 2, bbType.Min.Z);
                Vector = ptE - ptS;
            }
        }
    }
}
