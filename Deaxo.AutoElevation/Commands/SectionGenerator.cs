using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    public class ElevationResult
    {
        public ViewSection elevation;
    }

    public class SectionResult
    {
        public ViewSection elevation;
        public ViewSection cross;
        public ViewSection plan;
    }

    public static class SectionGenerator
    {
        public static ElevationResult CreateElevationOnly(Document doc, ElementProperties props)
        {
            if (props == null || !props.IsValid) return null;

            try
            {
                var boxElev = CreateSectionBox(props, "elevation");
                var elev = CreateViewSection(doc, boxElev);
                if (elev == null) return null;

                RenameViewSafe(elev, $"{props.TypeName}_{props.Element.Id}_Elevation");
                return new ElevationResult { elevation = elev };
            }
            catch
            {
                return null;
            }
        }

        public static ElevationResult CreateInternalElevation(Document doc, ElementProperties props)
        {
            if (props == null || !props.IsValid) return null;

            try
            {
                // Try internal mode first
                var boxElev = CreateSectionBox(props, "internal");
                var elev = CreateViewSection(doc, boxElev);
                if (elev == null) return null;

                RenameViewSafe(elev, $"{props.TypeName}_{props.Element.Id}_Internal_Elevation");
                return new ElevationResult { elevation = elev };
            }
            catch
            {
                // Fallback to standard elevation if internal fails
                try
                {
                    var boxElev = CreateSectionBox(props, "elevation");
                    var elev = CreateViewSection(doc, boxElev);
                    if (elev == null) return null;

                    RenameViewSafe(elev, $"{props.TypeName}_{props.Element.Id}_Internal_Elevation");
                    return new ElevationResult { elevation = elev };
                }
                catch
                {
                    return null;
                }
            }
        }

        public static Dictionary<string, ViewSection> CreateGroupElevations(Document doc, BoundingBoxXYZ overallBB, View viewTemplate)
        {
            if (overallBB == null) return new Dictionary<string, ViewSection>();

            var elevations = new Dictionary<string, ViewSection>();
            string[] directions = { "North", "South", "East", "West" };

            try
            {
                XYZ center = (overallBB.Min + overallBB.Max) / 2;
                double width = overallBB.Max.X - overallBB.Min.X;
                double height = overallBB.Max.Z - overallBB.Min.Z;
                double depth = overallBB.Max.Y - overallBB.Min.Y;
                double padding = Math.Max(Math.Max(width, height), depth) * 0.3;

                foreach (string direction in directions)
                {
                    try
                    {
                        var elevation = CreateDirectionalElevation(doc, center, width, height, depth, padding, direction, viewTemplate);
                        if (elevation != null)
                            elevations[direction] = elevation;
                    }
                    catch
                    {
                        // Continue with other elevations
                    }
                }
            }
            catch
            {
                // Return whatever we managed to create
            }

            return elevations;
        }

        private static ViewSection CreateDirectionalElevation(Document doc, XYZ center, double width, double height,
            double depth, double padding, string direction, View viewTemplate)
        {
            try
            {
                var sectionBox = new BoundingBoxXYZ();
                var transform = CreateDirectionalTransform(center, width, height, depth, padding, direction);

                double halfWidth = (direction == "North" || direction == "South") ? width / 2 + padding : depth / 2 + padding;
                double halfHeight = height / 2 + padding;
                double sectionDepth = (direction == "North" || direction == "South") ? depth + padding * 2 : width + padding * 2;

                sectionBox.Min = new XYZ(-halfWidth, -halfHeight, 0);
                sectionBox.Max = new XYZ(halfWidth, halfHeight, sectionDepth);
                sectionBox.Transform = transform;

                var elevation = CreateViewSection(doc, sectionBox);
                if (elevation == null) return null;

                if (viewTemplate != null)
                    elevation.ViewTemplateId = viewTemplate.Id;

                RenameViewSafe(elevation, $"Group_Elevation_{direction}_{DateTime.Now:HHmmss}");
                return elevation;
            }
            catch
            {
                return null;
            }
        }

        private static Transform CreateDirectionalTransform(XYZ center, double width, double height, double depth,
            double padding, string direction)
        {
            var transform = Transform.Identity;
            transform.Origin = center;

            switch (direction)
            {
                case "North": // Looking south
                    transform.BasisX = XYZ.BasisX;
                    transform.BasisY = XYZ.BasisZ;
                    transform.BasisZ = -XYZ.BasisY;
                    transform.Origin = center + new XYZ(0, depth / 2 + padding, 0);
                    break;
                case "South": // Looking north
                    transform.BasisX = -XYZ.BasisX;
                    transform.BasisY = XYZ.BasisZ;
                    transform.BasisZ = XYZ.BasisY;
                    transform.Origin = center + new XYZ(0, -depth / 2 - padding, 0);
                    break;
                case "East": // Looking west
                    transform.BasisX = XYZ.BasisY;
                    transform.BasisY = XYZ.BasisZ;
                    transform.BasisZ = -XYZ.BasisX;
                    transform.Origin = center + new XYZ(width / 2 + padding, 0, 0);
                    break;
                case "West": // Looking east
                    transform.BasisX = -XYZ.BasisY;
                    transform.BasisY = XYZ.BasisZ;
                    transform.BasisZ = XYZ.BasisX;
                    transform.Origin = center + new XYZ(-width / 2 - padding, 0, 0);
                    break;
            }

            return transform;
        }

        private static ViewSection CreateViewSection(Document doc, BoundingBoxXYZ sectionBox)
        {
            try
            {
                ElementId sectionTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection);
                if (sectionTypeId == null || sectionTypeId == ElementId.InvalidElementId)
                    return null;

                return ViewSection.CreateSection(doc, sectionTypeId, sectionBox);
            }
            catch
            {
                return null;
            }
        }

        private static BoundingBoxXYZ CreateSectionBox(ElementProperties props, string mode)
        {
            var sectionBox = new BoundingBoxXYZ();
            var transform = CreateTransform(props.Origin, props.Vector, mode);

            double W_half = Math.Max(props.Width / 2.0, 0.5);
            double H_half = Math.Max(props.Height / 2.0, 1.0);
            double D_half = Math.Max(props.Depth / 2.0, 0.5);
            double offset = 1.0; // Fixed offset instead of extension method

            if (mode == "internal")
            {
                // For internal elevations: section extends from interior space through wall
                sectionBox.Min = new XYZ(-W_half - offset, -H_half - offset, -D_half - offset);
                sectionBox.Max = new XYZ(W_half + offset, H_half + offset, D_half + offset);
            }
            else
            {
                // Standard elevation
                sectionBox.Min = new XYZ(-W_half - offset, -H_half - offset, 0);
                sectionBox.Max = new XYZ(W_half + offset, H_half + offset, D_half + offset);
            }

            sectionBox.Transform = transform;
            return sectionBox;
        }

        private static Transform CreateTransform(XYZ origin, XYZ vector, string mode)
        {
            var transform = Transform.Identity;

            // Validate and normalize vector
            XYZ normalizedVector;
            try
            {
                if (vector.GetLength() < 0.001)
                    normalizedVector = XYZ.BasisX;
                else
                    normalizedVector = vector.Normalize();
            }
            catch
            {
                normalizedVector = XYZ.BasisX;
            }

            // Calculate perpendicular direction
            XYZ perpendicular;
            if (Math.Abs(normalizedVector.X) > Math.Abs(normalizedVector.Y))
            {
                perpendicular = new XYZ(-normalizedVector.Y, normalizedVector.X, 0);
            }
            else
            {
                perpendicular = new XYZ(normalizedVector.Y, -normalizedVector.X, 0);
            }

            try
            {
                perpendicular = perpendicular.Normalize();
            }
            catch
            {
                perpendicular = XYZ.BasisY;
            }

            if (mode == "internal")
            {
                // Position at interior side, looking toward wall
                double wallThickness = 2.0; // Assume 2 feet for positioning
                transform.Origin = origin + perpendicular * wallThickness;
                transform.BasisX = normalizedVector;
                transform.BasisY = XYZ.BasisZ;
                transform.BasisZ = -perpendicular; // Look toward wall
            }
            else
            {
                // Standard external elevation
                transform.Origin = origin;
                transform.BasisX = normalizedVector;
                transform.BasisY = XYZ.BasisZ;
                transform.BasisZ = perpendicular;
            }

            return transform;
        }

        private static void RenameViewSafe(View view, string newName)
        {
            if (view == null) return;

            for (int i = 0; i < 10; i++)
            {
                try
                {
                    view.Name = newName;
                    break;
                }
                catch
                {
                    newName += "*";
                }
            }
        }

        // Backwards compatibility
        public static SectionResult CreateSections(Document doc, ElementProperties props)
        {
            var elevationResult = CreateElevationOnly(doc, props);
            return new SectionResult { elevation = elevationResult?.elevation };
        }
    }

    public static class ElementPropertiesExtensions
    {
        public static double offset(this ElementProperties props)
        {
            return 1.0;
        }
    }
}