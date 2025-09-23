using System;
using System.Collections.Generic;
using System.Linq;
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
                var boxElev = CreateSectionBox(doc, props, "elevation");
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
                // Internal side elevation
                var boxElev = CreateSectionBox(doc, props, "internal");
                var elev = CreateViewSection(doc, boxElev);
                if (elev == null) return null;

                RenameViewSafe(elev, $"{props.TypeName}_{props.Element.Id}_Internal_Elevation");
                return new ElevationResult { elevation = elev };
            }
            catch
            {
                // fallback to standard
                try
                {
                    var boxElev = CreateSectionBox(doc, props, "elevation");
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

            try
            {
                XYZ center = (overallBB.Min + overallBB.Max) / 2;
                double width = Math.Abs(overallBB.Max.X - overallBB.Min.X);
                double height = Math.Abs(overallBB.Max.Z - overallBB.Min.Z);
                double depth = Math.Abs(overallBB.Max.Y - overallBB.Min.Y);

                // Ensure minimum dimensions
                width = Math.Max(width, 2.0);
                height = Math.Max(height, 8.0);
                depth = Math.Max(depth, 2.0);

                double padding = Math.Max(Math.Max(width, height), depth) * 0.5;

                System.Diagnostics.Debug.WriteLine($"Group elevation - Center: {center}, W: {width:F1}, H: {height:F1}, D: {depth:F1}, Padding: {padding:F1}");

                // Create North elevation (looking South)
                try
                {
                    var northElev = CreateSimpleElevation(doc, center, width, height, depth, padding, "North", viewTemplate);
                    if (northElev != null) elevations["North"] = northElev;
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"North failed: {ex.Message}"); }

                // Create South elevation (looking North)  
                try
                {
                    var southElev = CreateSimpleElevation(doc, center, width, height, depth, padding, "South", viewTemplate);
                    if (southElev != null) elevations["South"] = southElev;
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"South failed: {ex.Message}"); }

                // Create East elevation (looking West)
                try
                {
                    var eastElev = CreateSimpleElevation(doc, center, width, height, depth, padding, "East", viewTemplate);
                    if (eastElev != null) elevations["East"] = eastElev;
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"East failed: {ex.Message}"); }

                // Create West elevation (looking East)
                try
                {
                    var westElev = CreateSimpleElevation(doc, center, width, height, depth, padding, "West", viewTemplate);
                    if (westElev != null) elevations["West"] = westElev;
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"West failed: {ex.Message}"); }

                System.Diagnostics.Debug.WriteLine($"Total elevations created: {elevations.Count}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in CreateGroupElevations: {ex.Message}");
            }

            return elevations;
        }

        private static ViewSection CreateSimpleElevation(Document doc, XYZ center, double width, double height,
            double depth, double padding, string direction, View viewTemplate)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== Creating {direction} elevation ===");

                // Create a simple, robust section box
                var sectionBox = new BoundingBoxXYZ();

                // Use generous, consistent dimensions for all directions
                double boxWidth = Math.Max(width, depth) + padding * 2;
                double boxHeight = height + padding * 2;
                double boxDepth = Math.Max(width, depth) + padding * 2;

                sectionBox.Min = new XYZ(-boxWidth / 2, -boxHeight / 2, 0);
                sectionBox.Max = new XYZ(boxWidth / 2, boxHeight / 2, boxDepth);

                // Create simple, reliable transforms
                var transform = Transform.Identity;

                switch (direction)
                {
                    case "North": // Looking south (towards negative Y)
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, center.Y + depth / 2 + padding, center.Z);
                        break;

                    case "South": // Looking north (towards positive Y)
                        transform.BasisX = -XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, center.Y - depth / 2 - padding, center.Z);
                        break;

                    case "East": // Looking west (towards negative X)
                        transform.BasisX = XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisX;
                        transform.Origin = new XYZ(center.X + width / 2 + padding, center.Y, center.Z);
                        break;

                    case "West": // Looking east (towards positive X)
                        transform.BasisX = -XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisX;
                        transform.Origin = new XYZ(center.X - width / 2 - padding, center.Y, center.Z);
                        break;
                }

                sectionBox.Transform = transform;

                System.Diagnostics.Debug.WriteLine($"{direction}: Box size W={boxWidth:F1} H={boxHeight:F1} D={boxDepth:F1}");
                System.Diagnostics.Debug.WriteLine($"{direction}: Origin {transform.Origin}");

                // Create the view section
                var elevation = CreateViewSection(doc, sectionBox);
                if (elevation == null)
                {
                    System.Diagnostics.Debug.WriteLine($"*** {direction} elevation creation FAILED ***");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"*** {direction} elevation creation SUCCEEDED: ID {elevation.Id} ***");

                // Apply template
                if (viewTemplate != null)
                {
                    try
                    {
                        elevation.ViewTemplateId = viewTemplate.Id;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Template application failed for {direction}: {ex.Message}");
                    }
                }

                // Set name
                string viewName = $"Group_Elevation_{direction}_{DateTime.Now.Ticks % 1000000}";
                RenameViewSafe(elevation, viewName);

                return elevation;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"*** Exception creating {direction} elevation: {ex.Message} ***");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }



        private static ViewSection CreateViewSection(Document doc, BoundingBoxXYZ sectionBox)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("=== CreateViewSection called ===");

                ElementId sectionTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection);
                System.Diagnostics.Debug.WriteLine($"Default section type ID: {sectionTypeId}");

                if (sectionTypeId == null || sectionTypeId == ElementId.InvalidElementId)
                {
                    System.Diagnostics.Debug.WriteLine("No default section type found, searching for alternative...");

                    // Try to find any section view type
                    var sectionTypes = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .Where(vt => vt.ViewFamily == ViewFamily.Section)
                        .ToList();

                    System.Diagnostics.Debug.WriteLine($"Found {sectionTypes.Count} section view types");

                    if (sectionTypes.Count > 0)
                    {
                        sectionTypeId = sectionTypes.First().Id;
                        System.Diagnostics.Debug.WriteLine($"Using section type: {sectionTypeId}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("*** No section view types found in document ***");
                        return null;
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Creating section with type {sectionTypeId}");
                System.Diagnostics.Debug.WriteLine($"SectionBox Min: {sectionBox.Min}, Max: {sectionBox.Max}");
                System.Diagnostics.Debug.WriteLine($"SectionBox Origin: {sectionBox.Transform.Origin}");

                var section = ViewSection.CreateSection(doc, sectionTypeId, sectionBox);

                if (section != null)
                {
                    System.Diagnostics.Debug.WriteLine($"*** SUCCESS: Created section with ID: {section.Id} ***");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("*** ViewSection.CreateSection returned NULL ***");
                }

                return section;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"*** EXCEPTION in CreateViewSection: {ex.Message} ***");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        private static BoundingBoxXYZ CreateSectionBox(Document doc, ElementProperties props, string mode)
        {
            var sectionBox = new BoundingBoxXYZ();
            var transform = CreateTransform(doc, props, mode);

            double W_half = Math.Max(props.Width / 2.0, 0.5);
            double H_half = Math.Max(props.Height / 2.0, 1.0);
            double D_half = Math.Max(props.Depth / 2.0, 0.5);
            double offset = 1.0;

            // Section aligned parallel to wall interior
            sectionBox.Min = new XYZ(-W_half - offset, -H_half - offset, 0);
            sectionBox.Max = new XYZ(W_half + offset, H_half + offset, D_half + offset);
            sectionBox.Transform = transform;

            return sectionBox;
        }

        private static Transform CreateTransform(Document doc, ElementProperties props, string mode)
        {
            var transform = Transform.Identity;

            if (props.Element is Wall wall)
            {
                LocationCurve locCurve = wall.Location as LocationCurve;
                XYZ wallDir = (locCurve.Curve.GetEndPoint(1) - locCurve.Curve.GetEndPoint(0)).Normalize();
                XYZ extNormal = wall.Orientation;
                if (wall.Flipped) extNormal = -extNormal;
                XYZ intNormal = -extNormal;
                XYZ viewNormal = (mode == "internal") ? intNormal : extNormal;

                // Ensure orthogonal axes
                XYZ basisZ = viewNormal.Normalize();
                XYZ basisX = wallDir.Normalize();
                XYZ basisY = basisZ.CrossProduct(basisX).Normalize();

                transform.BasisX = basisX;
                transform.BasisY = basisY;
                transform.BasisZ = basisZ;

                double halfThickness = Math.Max(wall.Width / 2.0, 0.1);
                transform.Origin = props.Origin + ((mode == "internal") ? intNormal * halfThickness : extNormal * halfThickness);
            }
            else
            {
                XYZ forward = props.Vector.Normalize();
                XYZ perp = new XYZ(-forward.Y, forward.X, 0).Normalize();
                transform.BasisX = forward;
                transform.BasisZ = perp;
                transform.BasisY = transform.BasisZ.CrossProduct(transform.BasisX).Normalize();
                transform.Origin = props.Origin;
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
        public static double offset(this ElementProperties props) => 1.0;
    }
}