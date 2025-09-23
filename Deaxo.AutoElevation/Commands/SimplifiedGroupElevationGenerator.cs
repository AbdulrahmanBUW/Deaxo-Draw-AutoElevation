// SimplifiedGroupElevationGenerator.cs
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Simplified group elevation generator that creates assembly-style views
    /// without scope boxes - uses direct section creation with proper transforms
    /// </summary>
    public static class SimplifiedGroupElevationGenerator
    {
        public class GroupViewResult
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
        /// Creates 6 orthographic section views + 1 3D view for selected elements
        /// </summary>
        public static GroupViewResult CreateGroupViews(Document doc, List<Element> elements, View viewTemplate = null)
        {
            if (elements == null || elements.Count == 0)
                return null;

            try
            {
                // Calculate the overall bounding box of all elements
                var overallBounds = CalculateElementsBounds(elements);
                if (overallBounds == null)
                {
                    System.Diagnostics.Debug.WriteLine("Could not calculate bounds for elements");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"Overall bounds: Min({overallBounds.Min.X:F1}, {overallBounds.Min.Y:F1}, {overallBounds.Min.Z:F1}) Max({overallBounds.Max.X:F1}, {overallBounds.Max.Y:F1}, {overallBounds.Max.Z:F1})");

                var result = new GroupViewResult();

                // Create all 6 section views
                result.TopView = CreateSectionView(doc, overallBounds, ViewDirection.Top, viewTemplate);
                result.BottomView = CreateSectionView(doc, overallBounds, ViewDirection.Bottom, viewTemplate);
                result.FrontView = CreateSectionView(doc, overallBounds, ViewDirection.Front, viewTemplate);
                result.BackView = CreateSectionView(doc, overallBounds, ViewDirection.Back, viewTemplate);
                result.LeftView = CreateSectionView(doc, overallBounds, ViewDirection.Left, viewTemplate);
                result.RightView = CreateSectionView(doc, overallBounds, ViewDirection.Right, viewTemplate);

                // Create 3D view
                result.ThreeDView = Create3DView(doc, overallBounds, viewTemplate);

                // Count successful creations
                int sectionCount = result.AllSectionViews.Count;
                int totalCount = sectionCount + (result.ThreeDView != null ? 1 : 0);

                System.Diagnostics.Debug.WriteLine($"Successfully created {totalCount} views: {sectionCount} sections + {(result.ThreeDView != null ? 1 : 0)} 3D");

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in CreateGroupViews: {ex.Message}");
                return null;
            }
        }

        private enum ViewDirection
        {
            Top, Bottom, Front, Back, Left, Right
        }

        /// <summary>
        /// Calculate overall bounding box from elements - simplified and reliable
        /// </summary>
        private static BoundingBoxXYZ CalculateElementsBounds(List<Element> elements)
        {
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool hasValidBounds = false;

            foreach (var element in elements)
            {
                try
                {
                    var bb = element.get_BoundingBox(null);
                    if (bb != null)
                    {
                        minX = Math.Min(minX, bb.Min.X);
                        minY = Math.Min(minY, bb.Min.Y);
                        minZ = Math.Min(minZ, bb.Min.Z);
                        maxX = Math.Max(maxX, bb.Max.X);
                        maxY = Math.Max(maxY, bb.Max.Y);
                        maxZ = Math.Max(maxZ, bb.Max.Z);
                        hasValidBounds = true;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to get bounds for element {element.Id}: {ex.Message}");
                }
            }

            if (!hasValidBounds)
                return null;

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }

        /// <summary>
        /// Create a section view in the specified direction
        /// </summary>
        private static ViewSection CreateSectionView(Document doc, BoundingBoxXYZ bounds, ViewDirection direction, View viewTemplate)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== Creating {direction} section view ===");

                // Get section view type
                ElementId sectionTypeId = GetSectionViewType(doc);
                if (sectionTypeId == null)
                {
                    System.Diagnostics.Debug.WriteLine($"No section view type found for {direction}");
                    return null;
                }

                // Create the section box with proper transform
                var sectionBox = CreateSectionBoxForDirection(bounds, direction);
                if (sectionBox == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create section box for {direction}");
                    return null;
                }

                // Validate section box dimensions
                var boxSize = sectionBox.Max - sectionBox.Min;
                if (boxSize.X <= 0 || boxSize.Y <= 0 || boxSize.Z <= 0)
                {
                    System.Diagnostics.Debug.WriteLine($"Invalid section box size for {direction}: {boxSize.X:F1} x {boxSize.Y:F1} x {boxSize.Z:F1}");
                    return null;
                }

                // Validate transform
                var transform = sectionBox.Transform;
                if (transform == null || transform.Origin == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Invalid transform for {direction}");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"{direction}: Creating section with valid box and transform");

                // Create the section view
                ViewSection section = null;
                try
                {
                    section = ViewSection.CreateSection(doc, sectionTypeId, sectionBox);
                }
                catch (Exception createEx)
                {
                    System.Diagnostics.Debug.WriteLine($"ViewSection.CreateSection failed for {direction}: {createEx.Message}");

                    // Try with adjusted parameters
                    try
                    {
                        System.Diagnostics.Debug.WriteLine($"Retrying {direction} with adjusted parameters...");
                        var adjustedBox = CreateAdjustedSectionBox(bounds, direction);
                        if (adjustedBox != null)
                        {
                            section = ViewSection.CreateSection(doc, sectionTypeId, adjustedBox);
                            System.Diagnostics.Debug.WriteLine($"Retry successful for {direction}");
                        }
                    }
                    catch (Exception retryEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Retry also failed for {direction}: {retryEx.Message}");
                        return null;
                    }
                }

                if (section == null)
                {
                    System.Diagnostics.Debug.WriteLine($"ViewSection.CreateSection returned null for {direction}");
                    return null;
                }

                // Apply template
                if (viewTemplate != null)
                {
                    try
                    {
                        section.ViewTemplateId = viewTemplate.Id;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to apply template to {direction}: {ex.Message}");
                    }
                }

                // Set proper name using new naming convention
                string viewName = $"group_elements_view - {direction}";
                SetUniqueViewName(section, viewName);

                System.Diagnostics.Debug.WriteLine($"Successfully created {direction} section: ID {section.Id}");
                return section;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception creating {direction} section: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Create adjusted section box with simpler parameters as fallback
        /// </summary>
        private static BoundingBoxXYZ CreateAdjustedSectionBox(BoundingBoxXYZ bounds, ViewDirection direction)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Creating FALLBACK section box for {direction}");

                var center = (bounds.Min + bounds.Max) / 2;
                var size = bounds.Max - bounds.Min;

                // Use the exact same approach as Front/Back (which work) for all directions
                double boxSize = Math.Max(Math.Max(size.X, size.Y), size.Z) + 20.0;

                var sectionBox = new BoundingBoxXYZ();
                sectionBox.Min = new XYZ(-boxSize / 2, -boxSize / 2, 0);
                sectionBox.Max = new XYZ(boxSize / 2, boxSize / 2, boxSize);

                var transform = Transform.Identity;

                // Use very simple transforms - copy exactly what works for Front/Back
                switch (direction)
                {
                    case ViewDirection.Top:
                        // Use Front pattern but looking down
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisY;
                        transform.BasisZ = -XYZ.BasisZ;
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Max.Z + 10);
                        break;

                    case ViewDirection.Bottom:
                        // Use Front pattern but looking up
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = -XYZ.BasisY;
                        transform.BasisZ = XYZ.BasisZ;
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Min.Z - 10);
                        break;

                    case ViewDirection.Front:
                        // Copy working Front exactly
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Max.Y + 10, center.Z);
                        break;

                    case ViewDirection.Back:
                        // Copy working Back exactly
                        transform.BasisX = -XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Min.Y - 10, center.Z);
                        break;

                    case ViewDirection.Left:
                        // Use Back pattern but rotate to X direction
                        transform.BasisX = -XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Min.X - 10, center.Y, center.Z);
                        break;

                    case ViewDirection.Right:
                        // Use Front pattern but rotate to X direction
                        transform.BasisX = XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Max.X + 10, center.Y, center.Z);
                        break;
                }

                sectionBox.Transform = transform;

                System.Diagnostics.Debug.WriteLine($"FALLBACK {direction}: Created with box size {boxSize:F1} and origin {transform.Origin}");
                return sectionBox;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to create fallback section box for {direction}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create section box with proper transform for each direction
        /// </summary>
        private static BoundingBoxXYZ CreateSectionBoxForDirection(BoundingBoxXYZ bounds, ViewDirection direction)
        {
            try
            {
                var center = (bounds.Min + bounds.Max) / 2;
                var size = bounds.Max - bounds.Min;

                // Use consistent sizing approach that works for Front/Back
                double width = Math.Max(size.X, 5.0);
                double depth = Math.Max(size.Y, 5.0);
                double height = Math.Max(size.Z, 10.0);
                double padding = 10.0; // Fixed padding that works

                var sectionBox = new BoundingBoxXYZ();
                var transform = Transform.Identity;

                System.Diagnostics.Debug.WriteLine($"Creating {direction}: W={width:F1}, D={depth:F1}, H={height:F1}, Padding={padding:F1}");

                switch (direction)
                {
                    case ViewDirection.Top:
                        // Top view - like a plan view looking down
                        sectionBox.Min = new XYZ(-width / 2 - padding, -depth / 2 - padding, 0);
                        sectionBox.Max = new XYZ(width / 2 + padding, depth / 2 + padding, height + padding * 2);

                        // Same transform pattern as Front/Back that work
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisY;
                        transform.BasisZ = -XYZ.BasisZ; // Looking down
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Max.Z + padding);
                        break;

                    case ViewDirection.Bottom:
                        // Bottom view - looking up
                        sectionBox.Min = new XYZ(-width / 2 - padding, -depth / 2 - padding, 0);
                        sectionBox.Max = new XYZ(width / 2 + padding, depth / 2 + padding, height + padding * 2);

                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = -XYZ.BasisY;
                        transform.BasisZ = XYZ.BasisZ;  // Looking up
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Min.Z - padding);
                        break;

                    case ViewDirection.Front:
                        // Front view - WORKING - keep same as before
                        sectionBox.Min = new XYZ(-width / 2 - padding, -height / 2 - padding, 0);
                        sectionBox.Max = new XYZ(width / 2 + padding, height / 2 + padding, depth + padding * 2);

                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Max.Y + padding, center.Z);
                        break;

                    case ViewDirection.Back:
                        // Back view - WORKING - keep same as before
                        sectionBox.Min = new XYZ(-width / 2 - padding, -height / 2 - padding, 0);
                        sectionBox.Max = new XYZ(width / 2 + padding, height / 2 + padding, depth + padding * 2);

                        transform.BasisX = -XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Min.Y - padding, center.Z);
                        break;

                    case ViewDirection.Left:
                        // Left view - copy Front/Back pattern but rotate
                        sectionBox.Min = new XYZ(-depth / 2 - padding, -height / 2 - padding, 0);
                        sectionBox.Max = new XYZ(depth / 2 + padding, height / 2 + padding, width + padding * 2);

                        transform.BasisX = -XYZ.BasisY;  // Rotate 90 degrees from Front
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Min.X - padding, center.Y, center.Z);
                        break;

                    case ViewDirection.Right:
                        // Right view - copy Front/Back pattern but rotate opposite
                        sectionBox.Min = new XYZ(-depth / 2 - padding, -height / 2 - padding, 0);
                        sectionBox.Max = new XYZ(depth / 2 + padding, height / 2 + padding, width + padding * 2);

                        transform.BasisX = XYZ.BasisY;   // Rotate -90 degrees from Front
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Max.X + padding, center.Y, center.Z);
                        break;
                }

                sectionBox.Transform = transform;

                // Log the exact same way for all views
                System.Diagnostics.Debug.WriteLine($"{direction}: Box Min({sectionBox.Min.X:F1}, {sectionBox.Min.Y:F1}, {sectionBox.Min.Z:F1}) Max({sectionBox.Max.X:F1}, {sectionBox.Max.Y:F1}, {sectionBox.Max.Z:F1})");
                System.Diagnostics.Debug.WriteLine($"{direction}: Origin({transform.Origin.X:F1}, {transform.Origin.Y:F1}, {transform.Origin.Z:F1})");
                System.Diagnostics.Debug.WriteLine($"{direction}: BasisX({transform.BasisX.X:F1}, {transform.BasisX.Y:F1}, {transform.BasisX.Z:F1})");
                System.Diagnostics.Debug.WriteLine($"{direction}: BasisY({transform.BasisY.X:F1}, {transform.BasisY.Y:F1}, {transform.BasisY.Z:F1})");
                System.Diagnostics.Debug.WriteLine($"{direction}: BasisZ({transform.BasisZ.X:F1}, {transform.BasisZ.Y:F1}, {transform.BasisZ.Z:F1})");

                return sectionBox;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating section box for {direction}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create 3D view with proper section box
        /// </summary>
        private static View3D Create3DView(Document doc, BoundingBoxXYZ bounds, View viewTemplate)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("=== Creating 3D view ===");

                // Get 3D view type
                var view3DType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vt => vt.ViewFamily == ViewFamily.ThreeDimensional);

                if (view3DType == null)
                {
                    System.Diagnostics.Debug.WriteLine("No 3D view type found");
                    return null;
                }

                // Create isometric 3D view
                var view3D = View3D.CreateIsometric(doc, view3DType.Id);
                if (view3D == null)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create 3D view");
                    return null;
                }

                // Set section box to crop tightly around elements
                var size = bounds.Max - bounds.Min;
                double padding = Math.Max(Math.Max(size.X, size.Y), size.Z) * 0.2;

                var sectionBox = new BoundingBoxXYZ
                {
                    Min = bounds.Min - new XYZ(padding, padding, padding),
                    Max = bounds.Max + new XYZ(padding, padding, padding),
                    Transform = Transform.Identity
                };

                view3D.IsSectionBoxActive = true;
                view3D.SetSectionBox(sectionBox);

                // Try to set orthographic mode
                try
                {
                    var perspectiveParam = view3D.get_Parameter(BuiltInParameter.VIEWER_PERSPECTIVE);
                    if (perspectiveParam != null && !perspectiveParam.IsReadOnly)
                    {
                        perspectiveParam.Set(0); // 0 = orthographic
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not set orthographic mode: {ex.Message}");
                }

                // Apply template if provided and compatible
                if (viewTemplate != null && viewTemplate is View3D)
                {
                    try
                    {
                        view3D.ViewTemplateId = viewTemplate.Id;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to apply 3D template: {ex.Message}");
                    }
                }

                // Set proper name
                string viewName = "group_elements_view - 3D";
                SetUniqueViewName(view3D, viewName);

                // Set a good default orientation
                try
                {
                    var center = (bounds.Min + bounds.Max) / 2;
                    var viewDirection = new XYZ(-1, -1, -0.5).Normalize();
                    var upDirection = XYZ.BasisZ;
                    view3D.SetOrientation(new ViewOrientation3D(center, upDirection, viewDirection));
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to set 3D orientation: {ex.Message}");
                }

                System.Diagnostics.Debug.WriteLine($"Successfully created 3D view: ID {view3D.Id}");
                return view3D;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception creating 3D view: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get section view type from document
        /// </summary>
        private static ElementId GetSectionViewType(Document doc)
        {
            try
            {
                // Try default first
                ElementId sectionTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection);
                if (sectionTypeId != null && sectionTypeId != ElementId.InvalidElementId)
                {
                    return sectionTypeId;
                }

                // Fallback: find any section view type
                var sectionTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .Where(vt => vt.ViewFamily == ViewFamily.Section)
                    .ToList();

                return sectionTypes.Count > 0 ? sectionTypes.First().Id : null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting section view type: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Set unique view name with fallback
        /// </summary>
        private static void SetUniqueViewName(View view, string baseName)
        {
            if (view == null) return;

            string viewName = baseName;
            for (int i = 0; i < 20; i++)
            {
                try
                {
                    view.Name = viewName;
                    break;
                }
                catch
                {
                    // Name already exists, try with suffix
                    viewName = $"{baseName} ({i + 1})";
                }
            }
        }
    }
}