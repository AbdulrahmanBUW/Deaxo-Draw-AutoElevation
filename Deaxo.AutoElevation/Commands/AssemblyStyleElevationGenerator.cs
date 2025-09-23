// AssemblyStyleElevationGenerator.cs
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Creates assembly-style orthographic views without creating actual assemblies
    /// Replicates the "Create Views" functionality from Assembly Tab
    /// Compatible with C# 7.3 and Revit 2023
    /// </summary>
    public static class AssemblyStyleElevationGenerator
    {
        public class AssemblyViewResult
        {
            public ViewSection TopView { get; set; }
            public ViewSection BottomView { get; set; }
            public ViewSection LeftView { get; set; }
            public ViewSection RightView { get; set; }
            public ViewSection FrontView { get; set; }
            public ViewSection BackView { get; set; }
            public View3D ThreeDView { get; set; }
            public List<ViewSection> AllSectionViews
            {
                get
                {
                    var views = new List<ViewSection>();
                    if (TopView != null) views.Add(TopView);
                    if (BottomView != null) views.Add(BottomView);
                    if (LeftView != null) views.Add(LeftView);
                    if (RightView != null) views.Add(RightView);
                    if (FrontView != null) views.Add(FrontView);
                    if (BackView != null) views.Add(BackView);
                    return views;
                }
            }
        }

        /// <summary>
        /// Creates assembly-style orthographic views for selected elements
        /// </summary>
        public static AssemblyViewResult CreateAssemblyStyleViews(Document doc, List<Element> elements, View viewTemplate = null)
        {
            if (elements == null || elements.Count == 0)
                return null;

            try
            {
                // Step 1: Analyze geometry to determine optimal coordinate system
                var geometryAnalysis = AnalyzeElementGeometry(elements);
                if (geometryAnalysis == null) return null;

                // Step 2: Create all orthographic views
                var result = new AssemblyViewResult();

                // Create 6 orthographic section views
                result.TopView = CreateOrthographicView(doc, geometryAnalysis, OrthographicDirection.Top, viewTemplate);
                result.BottomView = CreateOrthographicView(doc, geometryAnalysis, OrthographicDirection.Bottom, viewTemplate);
                result.FrontView = CreateOrthographicView(doc, geometryAnalysis, OrthographicDirection.Front, viewTemplate);
                result.BackView = CreateOrthographicView(doc, geometryAnalysis, OrthographicDirection.Back, viewTemplate);
                result.LeftView = CreateOrthographicView(doc, geometryAnalysis, OrthographicDirection.Left, viewTemplate);
                result.RightView = CreateOrthographicView(doc, geometryAnalysis, OrthographicDirection.Right, viewTemplate);

                // Create 3D orthographic view
                result.ThreeDView = Create3DOrthographicView(doc, geometryAnalysis, viewTemplate);

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating assembly-style views: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Analyzes element geometry to determine optimal coordinate system and bounds
        /// This mimics what Revit does when creating assemblies
        /// </summary>
        private static GeometryAnalysisResult AnalyzeElementGeometry(List<Element> elements)
        {
            try
            {
                var analysis = new GeometryAnalysisResult();

                // Get all bounding boxes
                var boundingBoxes = elements.Select(el => el.get_BoundingBox(null))
                    .Where(bb => bb != null).ToList();

                if (boundingBoxes.Count == 0) return null;

                // Calculate overall bounding box
                analysis.OverallBounds = CalculateOverallBoundingBox(boundingBoxes);
                analysis.Center = (analysis.OverallBounds.Min + analysis.OverallBounds.Max) / 2;

                // Analyze primary directions (this is key to assembly-like behavior)
                analysis.CoordinateSystem = DetermineOptimalCoordinateSystem(elements, analysis.OverallBounds);

                // Calculate dimensions in local coordinate system
                analysis.LocalDimensions = CalculateLocalDimensions(analysis.OverallBounds, analysis.CoordinateSystem);

                // Calculate optimal view parameters
                analysis.ViewParameters = CalculateOptimalViewParameters(analysis);

                return analysis;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error analyzing geometry: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Determines the optimal coordinate system based on element geometry
        /// This is the key to creating assembly-like views that are properly oriented
        /// </summary>
        private static Transform DetermineOptimalCoordinateSystem(List<Element> elements, BoundingBoxXYZ bounds)
        {
            try
            {
                // Strategy 1: Use dominant linear elements (walls, beams) to determine orientation
                var linearElements = elements.Where(el => el is Wall ||
                    (el.Category != null && el.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)).ToList();

                if (linearElements.Any())
                {
                    var result = DetermineCoordinateFromLinearElements(linearElements, bounds);
                    if (result != null) return result;
                }

                // Strategy 2: Use building grid alignment if available
                var gridAlignment = TryDetermineGridAlignment(elements[0].Document, bounds);
                if (gridAlignment != null)
                {
                    return gridAlignment;
                }

                // Strategy 3: Analyze geometry distribution to find natural axes
                var geometryAxes = AnalyzeGeometryDistribution(elements, bounds);
                if (geometryAxes != null)
                {
                    return geometryAxes;
                }

                // Fallback: Use world coordinate system
                var transform = Transform.Identity;
                transform.Origin = (bounds.Min + bounds.Max) / 2;
                return transform;
            }
            catch
            {
                // Fallback to world coordinates
                var transform = Transform.Identity;
                transform.Origin = (bounds.Min + bounds.Max) / 2;
                return transform;
            }
        }

        /// <summary>
        /// Determines coordinate system based on linear elements like walls and beams
        /// </summary>
        private static Transform DetermineCoordinateFromLinearElements(List<Element> linearElements, BoundingBoxXYZ bounds)
        {
            var directions = new List<XYZ>();

            foreach (var element in linearElements)
            {
                if (element is Wall wall)
                {
                    var locationCurve = wall.Location as LocationCurve;
                    if (locationCurve?.Curve != null)
                    {
                        var dir = (locationCurve.Curve.GetEndPoint(1) - locationCurve.Curve.GetEndPoint(0)).Normalize();
                        directions.Add(dir);
                    }
                }
                else if (element.Location is LocationCurve locCurve)
                {
                    var dir = (locCurve.Curve.GetEndPoint(1) - locCurve.Curve.GetEndPoint(0)).Normalize();
                    directions.Add(dir);
                }
            }

            if (directions.Count > 0)
            {
                // Find dominant direction
                var primaryDir = FindDominantDirection(directions);
                var secondaryDir = FindPerpendicularDirection(directions, primaryDir);

                var transform = Transform.Identity;
                transform.BasisX = primaryDir;
                transform.BasisY = secondaryDir;
                transform.BasisZ = primaryDir.CrossProduct(secondaryDir).Normalize();
                transform.Origin = (bounds.Min + bounds.Max) / 2;

                return transform;
            }

            return null;
        }

        /// <summary>
        /// Tries to align with building grids for better orientation
        /// </summary>
        private static Transform TryDetermineGridAlignment(Document doc, BoundingBoxXYZ bounds)
        {
            try
            {
                var grids = new FilteredElementCollector(doc)
                    .OfClass(typeof(Grid))
                    .Cast<Grid>()
                    .Where(g => g.Curve != null)
                    .ToList();

                if (grids.Count >= 2)
                {
                    // Find grids that intersect our bounds
                    var relevantGrids = grids.Where(g => DoesGridIntersectBounds(g, bounds)).ToList();

                    if (relevantGrids.Count >= 2)
                    {
                        var directions = relevantGrids.Select(g => GetGridDirection(g)).ToList();
                        var primaryDir = FindDominantDirection(directions);
                        var secondaryDir = FindPerpendicularDirection(directions, primaryDir);

                        var transform = Transform.Identity;
                        transform.BasisX = primaryDir;
                        transform.BasisY = secondaryDir;
                        transform.BasisZ = XYZ.BasisZ;
                        transform.Origin = (bounds.Min + bounds.Max) / 2;

                        return transform;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Grid alignment failed: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Analyzes geometry distribution to find natural coordinate axes
        /// </summary>
        private static Transform AnalyzeGeometryDistribution(List<Element> elements, BoundingBoxXYZ bounds)
        {
            try
            {
                // This is a simplified approach - in a full implementation, you might use
                // principal component analysis to find the natural axes of the geometry

                var center = (bounds.Min + bounds.Max) / 2;
                var size = bounds.Max - bounds.Min;

                // If one dimension is significantly larger, align to that
                if (size.X > size.Y * 2 && size.X > size.Z * 2)
                {
                    // Elongated in X direction
                    var transform = Transform.Identity;
                    transform.Origin = center;
                    return transform;
                }
                else if (size.Y > size.X * 2 && size.Y > size.Z * 2)
                {
                    // Elongated in Y direction - rotate coordinate system
                    var transform = Transform.Identity;
                    transform.BasisX = XYZ.BasisY;
                    transform.BasisY = -XYZ.BasisX;
                    transform.BasisZ = XYZ.BasisZ;
                    transform.Origin = center;
                    return transform;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Geometry distribution analysis failed: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Creates an orthographic section view in the specified direction
        /// </summary>
        private static ViewSection CreateOrthographicView(Document doc, GeometryAnalysisResult analysis,
            OrthographicDirection direction, View viewTemplate)
        {
            try
            {
                var sectionBox = CreateSectionBox(analysis, direction);
                var section = CreateViewSection(doc, sectionBox);

                if (section != null)
                {
                    // Apply template
                    if (viewTemplate != null)
                    {
                        try { section.ViewTemplateId = viewTemplate.Id; } catch { }
                    }

                    // Set appropriate name
                    var baseName = $"Assembly_Style_{direction}_{DateTime.Now.Ticks % 1000000}";
                    SetUniqueViewName(section, baseName);

                    // Set detail level and visual style for better display
                    try
                    {
                        section.DetailLevel = ViewDetailLevel.Fine;
                    }
                    catch { }
                }

                return section;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to create {direction} view: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Creates a 3D orthographic view
        /// </summary>
        private static View3D Create3DOrthographicView(Document doc, GeometryAnalysisResult analysis, View viewTemplate)
        {
            try
            {
                // Get 3D view type
                var view3DType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vt => vt.ViewFamily == ViewFamily.ThreeDimensional);

                if (view3DType == null) return null;

                var view3D = View3D.CreateIsometric(doc, view3DType.Id);

                if (view3D != null)
                {
                    // Create perspective view first, then convert to orthographic if possible
                    try
                    {
                        // Try to set orthographic mode - this may not work in all Revit versions
                        var perspectiveParam = view3D.get_Parameter(BuiltInParameter.VIEWER_PERSPECTIVE);
                        if (perspectiveParam != null && !perspectiveParam.IsReadOnly)
                        {
                            perspectiveParam.Set(0); // 0 = orthographic, 1 = perspective
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not set orthographic mode: {ex.Message}");
                        // Continue with perspective view if orthographic setting fails
                    }

                    // Set section box for tight cropping
                    view3D.IsSectionBoxActive = true;
                    view3D.SetSectionBox(CreateSectionBox(analysis, OrthographicDirection.ThreeD));

                    // Apply template if provided
                    if (viewTemplate != null && viewTemplate is View3D)
                    {
                        try { view3D.ViewTemplateId = viewTemplate.Id; } catch { }
                    }

                    // Set name
                    var baseName = $"Assembly_Style_3D_{DateTime.Now.Ticks % 1000000}";
                    SetUniqueViewName(view3D, baseName);

                    // Set orientation for better default view
                    SetOptimal3DOrientation(view3D, analysis);
                }

                return view3D;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to create 3D view: {ex.Message}");
                return null;
            }
        }

        #region Helper Methods and Classes

        private enum OrthographicDirection
        {
            Top, Bottom, Front, Back, Left, Right, ThreeD
        }

        private class GeometryAnalysisResult
        {
            public BoundingBoxXYZ OverallBounds { get; set; }
            public XYZ Center { get; set; }
            public Transform CoordinateSystem { get; set; }
            public XYZ LocalDimensions { get; set; }
            public ViewParameters ViewParameters { get; set; }
        }

        private class ViewParameters
        {
            public double ViewDepth { get; set; }
            public double Padding { get; set; }
            public double OptimalScale { get; set; }
        }

        private static BoundingBoxXYZ CalculateOverallBoundingBox(List<BoundingBoxXYZ> boundingBoxes)
        {
            var minX = boundingBoxes.Min(bb => bb.Min.X);
            var minY = boundingBoxes.Min(bb => bb.Min.Y);
            var minZ = boundingBoxes.Min(bb => bb.Min.Z);
            var maxX = boundingBoxes.Max(bb => bb.Max.X);
            var maxY = boundingBoxes.Max(bb => bb.Max.Y);
            var maxZ = boundingBoxes.Max(bb => bb.Max.Z);

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }

        private static XYZ CalculateLocalDimensions(BoundingBoxXYZ bounds, Transform coordSystem)
        {
            var size = bounds.Max - bounds.Min;
            return new XYZ(size.X, size.Y, size.Z);
        }

        private static ViewParameters CalculateOptimalViewParameters(GeometryAnalysisResult analysis)
        {
            var dims = analysis.LocalDimensions;
            var maxDim = Math.Max(Math.Max(dims.X, dims.Y), dims.Z);

            return new ViewParameters
            {
                ViewDepth = maxDim * 1.5,
                Padding = maxDim * 0.2,
                OptimalScale = 100 // This could be calculated based on size
            };
        }

        private static BoundingBoxXYZ CreateSectionBox(GeometryAnalysisResult analysis, OrthographicDirection direction)
        {
            var bounds = analysis.OverallBounds;
            var coordSystem = analysis.CoordinateSystem;
            var params_ = analysis.ViewParameters;

            var sectionBox = new BoundingBoxXYZ();
            var size = bounds.Max - bounds.Min;
            var padding = params_.Padding;

            // Create section box based on direction
            switch (direction)
            {
                case OrthographicDirection.Top:
                    sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Y / 2 - padding, 0);
                    sectionBox.Max = new XYZ(size.X / 2 + padding, size.Y / 2 + padding, params_.ViewDepth);
                    sectionBox.Transform = CreateTopViewTransform(coordSystem, bounds);
                    break;

                case OrthographicDirection.Bottom:
                    sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Y / 2 - padding, 0);
                    sectionBox.Max = new XYZ(size.X / 2 + padding, size.Y / 2 + padding, params_.ViewDepth);
                    sectionBox.Transform = CreateBottomViewTransform(coordSystem, bounds);
                    break;

                case OrthographicDirection.Front:
                    sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Z / 2 - padding, 0);
                    sectionBox.Max = new XYZ(size.X / 2 + padding, size.Z / 2 + padding, params_.ViewDepth);
                    sectionBox.Transform = CreateFrontViewTransform(coordSystem, bounds);
                    break;

                case OrthographicDirection.Back:
                    sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Z / 2 - padding, 0);
                    sectionBox.Max = new XYZ(size.X / 2 + padding, size.Z / 2 + padding, params_.ViewDepth);
                    sectionBox.Transform = CreateBackViewTransform(coordSystem, bounds);
                    break;

                case OrthographicDirection.Left:
                    sectionBox.Min = new XYZ(-size.Y / 2 - padding, -size.Z / 2 - padding, 0);
                    sectionBox.Max = new XYZ(size.Y / 2 + padding, size.Z / 2 + padding, params_.ViewDepth);
                    sectionBox.Transform = CreateLeftViewTransform(coordSystem, bounds);
                    break;

                case OrthographicDirection.Right:
                    sectionBox.Min = new XYZ(-size.Y / 2 - padding, -size.Z / 2 - padding, 0);
                    sectionBox.Max = new XYZ(size.Y / 2 + padding, size.Z / 2 + padding, params_.ViewDepth);
                    sectionBox.Transform = CreateRightViewTransform(coordSystem, bounds);
                    break;

                case OrthographicDirection.ThreeD:
                    sectionBox.Min = bounds.Min - new XYZ(padding, padding, padding);
                    sectionBox.Max = bounds.Max + new XYZ(padding, padding, padding);
                    sectionBox.Transform = Transform.Identity;
                    break;
            }

            return sectionBox;
        }

        // Transform creation methods for each view direction
        private static Transform CreateTopViewTransform(Transform coordSystem, BoundingBoxXYZ bounds)
        {
            var transform = Transform.Identity;
            transform.BasisX = coordSystem.BasisX;
            transform.BasisY = coordSystem.BasisY;
            transform.BasisZ = -coordSystem.BasisZ; // Looking down
            transform.Origin = new XYZ(coordSystem.Origin.X, coordSystem.Origin.Y, bounds.Max.Z + 1);
            return transform;
        }

        private static Transform CreateBottomViewTransform(Transform coordSystem, BoundingBoxXYZ bounds)
        {
            var transform = Transform.Identity;
            transform.BasisX = coordSystem.BasisX;
            transform.BasisY = -coordSystem.BasisY;
            transform.BasisZ = coordSystem.BasisZ; // Looking up
            transform.Origin = new XYZ(coordSystem.Origin.X, coordSystem.Origin.Y, bounds.Min.Z - 1);
            return transform;
        }

        private static Transform CreateFrontViewTransform(Transform coordSystem, BoundingBoxXYZ bounds)
        {
            var transform = Transform.Identity;
            transform.BasisX = coordSystem.BasisX;
            transform.BasisY = coordSystem.BasisZ;
            transform.BasisZ = -coordSystem.BasisY; // Looking in negative Y direction
            transform.Origin = new XYZ(coordSystem.Origin.X, bounds.Max.Y + 1, coordSystem.Origin.Z);
            return transform;
        }

        private static Transform CreateBackViewTransform(Transform coordSystem, BoundingBoxXYZ bounds)
        {
            var transform = Transform.Identity;
            transform.BasisX = -coordSystem.BasisX;
            transform.BasisY = coordSystem.BasisZ;
            transform.BasisZ = coordSystem.BasisY; // Looking in positive Y direction
            transform.Origin = new XYZ(coordSystem.Origin.X, bounds.Min.Y - 1, coordSystem.Origin.Z);
            return transform;
        }

        private static Transform CreateLeftViewTransform(Transform coordSystem, BoundingBoxXYZ bounds)
        {
            var transform = Transform.Identity;
            transform.BasisX = coordSystem.BasisY;
            transform.BasisY = coordSystem.BasisZ;
            transform.BasisZ = -coordSystem.BasisX; // Looking in negative X direction
            transform.Origin = new XYZ(bounds.Max.X + 1, coordSystem.Origin.Y, coordSystem.Origin.Z);
            return transform;
        }

        private static Transform CreateRightViewTransform(Transform coordSystem, BoundingBoxXYZ bounds)
        {
            var transform = Transform.Identity;
            transform.BasisX = -coordSystem.BasisY;
            transform.BasisY = coordSystem.BasisZ;
            transform.BasisZ = coordSystem.BasisX; // Looking in positive X direction
            transform.Origin = new XYZ(bounds.Min.X - 1, coordSystem.Origin.Y, coordSystem.Origin.Z);
            return transform;
        }

        private static ViewSection CreateViewSection(Document doc, BoundingBoxXYZ sectionBox)
        {
            try
            {
                ElementId sectionTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection);

                if (sectionTypeId == null || sectionTypeId == ElementId.InvalidElementId)
                {
                    var sectionTypes = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .Where(vt => vt.ViewFamily == ViewFamily.Section)
                        .ToList();

                    if (sectionTypes.Count > 0)
                        sectionTypeId = sectionTypes.First().Id;
                    else
                        return null;
                }

                return ViewSection.CreateSection(doc, sectionTypeId, sectionBox);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CreateViewSection failed: {ex.Message}");
                return null;
            }
        }

        private static void SetOptimal3DOrientation(View3D view3D, GeometryAnalysisResult analysis)
        {
            try
            {
                // Set an isometric-style view direction that shows the elements well
                var viewDirection = new XYZ(-1, -1, -0.5).Normalize();
                var upDirection = XYZ.BasisZ;

                view3D.SetOrientation(new ViewOrientation3D(analysis.Center, upDirection, viewDirection));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to set 3D orientation: {ex.Message}");
            }
        }

        private static void SetUniqueViewName(View view, string baseName)
        {
            string viewName = baseName;
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    view.Name = viewName;
                    break;
                }
                catch
                {
                    viewName = baseName + $"_{i + 1}";
                }
            }
        }

        // Utility methods for direction analysis
        private static XYZ FindDominantDirection(List<XYZ> directions)
        {
            if (directions.Count == 0) return XYZ.BasisX;

            // Group similar directions and find the most common
            var groups = new List<List<XYZ>>();
            const double tolerance = 0.1; // ~6 degrees

            foreach (var dir in directions)
            {
                var found = false;
                foreach (var group in groups)
                {
                    if (group[0].IsAlmostEqualTo(dir, tolerance) ||
                        group[0].IsAlmostEqualTo(-dir, tolerance))
                    {
                        group.Add(dir);
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    groups.Add(new List<XYZ> { dir });
                }
            }

            return groups.OrderByDescending(g => g.Count).First()[0];
        }

        private static XYZ FindPerpendicularDirection(List<XYZ> directions, XYZ primaryDir)
        {
            foreach (var dir in directions)
            {
                var dot = Math.Abs(dir.DotProduct(primaryDir));
                if (dot < 0.3) // Close to perpendicular
                {
                    return dir.Normalize();
                }
            }

            // Fallback: create perpendicular direction
            if (Math.Abs(primaryDir.DotProduct(XYZ.BasisZ)) < 0.9)
            {
                return primaryDir.CrossProduct(XYZ.BasisZ).Normalize();
            }
            else
            {
                return primaryDir.CrossProduct(XYZ.BasisX).Normalize();
            }
        }

        private static bool DoesGridIntersectBounds(Grid grid, BoundingBoxXYZ bounds)
        {
            // Simplified intersection test
            try
            {
                var curve = grid.Curve;
                var start = curve.GetEndPoint(0);
                var end = curve.GetEndPoint(1);

                return (start.X <= bounds.Max.X && start.X >= bounds.Min.X &&
                        start.Y <= bounds.Max.Y && start.Y >= bounds.Min.Y) ||
                       (end.X <= bounds.Max.X && end.X >= bounds.Min.X &&
                        end.Y <= bounds.Max.Y && end.Y >= bounds.Min.Y);
            }
            catch
            {
                return false;
            }
        }

        private static XYZ GetGridDirection(Grid grid)
        {
            var curve = grid.Curve;
            return (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();
        }

        #endregion
    }
}