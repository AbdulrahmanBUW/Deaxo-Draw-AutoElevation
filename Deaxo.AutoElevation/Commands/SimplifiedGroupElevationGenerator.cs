// SimplifiedGroupElevationGenerator.cs - UPDATED WITH IMPROVED TEMPLATE HANDLING
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Simplified group elevation generator that creates assembly-style views
    /// UPDATED VERSION: Enhanced template handling with compatibility checking
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

            public int TotalViewCount => AllSectionViews.Count + (ThreeDView != null ? 1 : 0);
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
                System.Diagnostics.Debug.WriteLine("=== STARTING GROUP VIEW CREATION (ENHANCED TEMPLATE VERSION) ===");
                System.Diagnostics.Debug.WriteLine($"Input: {elements.Count} elements");

                if (viewTemplate != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Template provided: '{viewTemplate.Name}' (Type: {viewTemplate.ViewType})");
                }

                // Calculate the overall bounding box of all elements
                var overallBounds = CalculateElementsBounds(elements);
                if (overallBounds == null)
                {
                    System.Diagnostics.Debug.WriteLine("ERROR: Could not calculate bounds for elements");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"Overall bounds:");
                System.Diagnostics.Debug.WriteLine($"  Min: X={overallBounds.Min.X:F1}, Y={overallBounds.Min.Y:F1}, Z={overallBounds.Min.Z:F1}");
                System.Diagnostics.Debug.WriteLine($"  Max: X={overallBounds.Max.X:F1}, Y={overallBounds.Max.Y:F1}, Z={overallBounds.Max.Z:F1}");

                var result = new GroupViewResult();

                // Get section view type
                ElementId sectionTypeId = GetSectionViewType(doc);
                if (sectionTypeId == null)
                {
                    System.Diagnostics.Debug.WriteLine("ERROR: No section view type found");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"Using section view type: {sectionTypeId}");

                // Create all 6 section views using ENHANCED template handling
                System.Diagnostics.Debug.WriteLine("\n--- CREATING SECTION VIEWS WITH ENHANCED TEMPLATE HANDLING ---");

                result.TopView = CreateEnhancedSectionView(doc, overallBounds, ViewDirection.Top, viewTemplate, sectionTypeId);
                result.BottomView = CreateEnhancedSectionView(doc, overallBounds, ViewDirection.Bottom, viewTemplate, sectionTypeId);
                result.FrontView = CreateEnhancedSectionView(doc, overallBounds, ViewDirection.Front, viewTemplate, sectionTypeId);
                result.BackView = CreateEnhancedSectionView(doc, overallBounds, ViewDirection.Back, viewTemplate, sectionTypeId);
                result.LeftView = CreateEnhancedSectionView(doc, overallBounds, ViewDirection.Left, viewTemplate, sectionTypeId);
                result.RightView = CreateEnhancedSectionView(doc, overallBounds, ViewDirection.Right, viewTemplate, sectionTypeId);

                // Create 3D view with enhanced template handling
                System.Diagnostics.Debug.WriteLine("\n--- CREATING 3D VIEW WITH ENHANCED TEMPLATE HANDLING ---");
                result.ThreeDView = CreateEnhanced3DView(doc, overallBounds, viewTemplate);

                // Final summary
                int sectionCount = result.AllSectionViews.Count;
                int totalCount = result.TotalViewCount;

                System.Diagnostics.Debug.WriteLine($"\n=== ENHANCED TEMPLATE VERSION RESULTS ===");
                System.Diagnostics.Debug.WriteLine($"Section views: {sectionCount}/6");
                System.Diagnostics.Debug.WriteLine($"3D views: {(result.ThreeDView != null ? 1 : 0)}/1");
                System.Diagnostics.Debug.WriteLine($"Total: {totalCount}/7");

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in CreateGroupViews: {ex.Message}");
                return null;
            }
        }

        private enum ViewDirection
        {
            Top, Bottom, Front, Back, Left, Right
        }

        #region Template Application Methods

        /// <summary>
        /// Applies view template with proper error handling and compatibility checking
        /// </summary>
        private static bool ApplyViewTemplate(View view, View viewTemplate, string viewDescription = "")
        {
            if (viewTemplate == null)
            {
                System.Diagnostics.Debug.WriteLine($"No template selected for {viewDescription}");
                return false;
            }

            if (!viewTemplate.IsTemplate)
            {
                System.Diagnostics.Debug.WriteLine($"Selected view '{viewTemplate.Name}' is not a template");
                return false;
            }

            try
            {
                // Check compatibility first
                if (!IsTemplateCompatible(view, viewTemplate))
                {
                    System.Diagnostics.Debug.WriteLine($"⚠ Template '{viewTemplate.Name}' (Type: {viewTemplate.ViewType}) not compatible with {viewDescription} (Type: {view.ViewType})");
                    return false;
                }

                // Apply the template
                view.ViewTemplateId = viewTemplate.Id;

                // Verify it was applied
                if (view.ViewTemplateId == viewTemplate.Id)
                {
                    System.Diagnostics.Debug.WriteLine($"✓ Successfully applied template '{viewTemplate.Name}' to {viewDescription}");
                    LogTemplateProperties(view, viewTemplate);
                    return true;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"✗ Template application verification failed for {viewDescription}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"✗ Exception applying template '{viewTemplate.Name}' to {viewDescription}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if a template is compatible with a view type
        /// </summary>
        private static bool IsTemplateCompatible(View view, View template)
        {
            if (view == null || template == null || !template.IsTemplate)
                return false;

            // Section views can use Section templates
            if (view is ViewSection && template.ViewType == ViewType.Section)
                return true;

            // 3D views can use 3D templates  
            if (view is View3D && template.ViewType == ViewType.ThreeD)
                return true;

            // Plan views can use Plan templates
            if (view is ViewPlan && (template.ViewType == ViewType.FloorPlan ||
                                    template.ViewType == ViewType.CeilingPlan ||
                                    template.ViewType == ViewType.AreaPlan))
                return true;

            // Elevation views can use Elevation templates
            if (view.ViewType == ViewType.Elevation && template.ViewType == ViewType.Elevation)
                return true;

            return false;
        }

        /// <summary>
        /// Logs template properties for debugging
        /// </summary>
        private static void LogTemplateProperties(View view, View template)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"  Applied template properties:");
                System.Diagnostics.Debug.WriteLine($"    - Scale: {view.Scale}");
                System.Diagnostics.Debug.WriteLine($"    - Detail Level: {view.DetailLevel}");

                if (view is ViewSection section)
                {
                    System.Diagnostics.Debug.WriteLine($"    - Crop Box Active: {section.CropBoxActive}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"    Could not log template properties: {ex.Message}");
            }
        }

        #endregion

        /// <summary>
        /// Create section view using ENHANCED template logic
        /// </summary>
        private static ViewSection CreateEnhancedSectionView(Document doc, BoundingBoxXYZ bounds,
            ViewDirection direction, View viewTemplate, ElementId sectionTypeId)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"\n--- Creating {direction} Section (ENHANCED) ---");

                // Create section box using the proven working pattern
                var sectionBox = CreateWorkingSectionBox(bounds, direction);
                if (sectionBox == null)
                {
                    System.Diagnostics.Debug.WriteLine($"FAILED: Could not create section box for {direction}");
                    return null;
                }

                LogSectionBoxDetails(direction, sectionBox);

                // Create the section view
                ViewSection section = null;
                try
                {
                    section = ViewSection.CreateSection(doc, sectionTypeId, sectionBox);
                    System.Diagnostics.Debug.WriteLine($"SUCCESS: Created {direction} section view");
                }
                catch (Exception createEx)
                {
                    System.Diagnostics.Debug.WriteLine($"FAILED: {direction} creation error: {createEx.Message}");
                    return null;
                }

                if (section == null)
                {
                    System.Diagnostics.Debug.WriteLine($"FAILED: {direction} returned null");
                    return null;
                }

                // ENHANCED TEMPLATE APPLICATION:
                bool templateApplied = ApplyViewTemplate(section, viewTemplate, $"{direction} Section");

                // Set name
                string viewName = $"group_elements_view - {direction}";
                SetUniqueViewName(section, viewName);

                // Final status with template information
                System.Diagnostics.Debug.WriteLine($"SUCCESS: {direction} view created with ID {section.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");

                return section;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"EXCEPTION creating {direction}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create 3D view with ENHANCED template handling
        /// </summary>
        private static View3D CreateEnhanced3DView(Document doc, BoundingBoxXYZ bounds, View viewTemplate)
        {
            try
            {
                // Get 3D view type
                var view3DType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vt => vt.ViewFamily == ViewFamily.ThreeDimensional);

                if (view3DType == null) return null;

                // Create isometric 3D view
                var view3D = View3D.CreateIsometric(doc, view3DType.Id);
                if (view3D == null) return null;

                // Set section box
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

                // Try orthographic mode
                try
                {
                    var perspectiveParam = view3D.get_Parameter(BuiltInParameter.VIEWER_PERSPECTIVE);
                    if (perspectiveParam != null && !perspectiveParam.IsReadOnly)
                        perspectiveParam.Set(0);
                }
                catch { }

                // ENHANCED TEMPLATE APPLICATION FOR 3D:
                bool templateApplied = ApplyViewTemplate(view3D, viewTemplate, "3D View");

                // Set name and orientation
                SetUniqueViewName(view3D, "group_elements_view - 3D");

                try
                {
                    var center = (bounds.Min + bounds.Max) / 2;
                    var viewDirection = new XYZ(-1, -1, -0.5).Normalize();
                    var upDirection = XYZ.BasisZ;
                    view3D.SetOrientation(new ViewOrientation3D(center, upDirection, viewDirection));
                }
                catch { }

                System.Diagnostics.Debug.WriteLine($"SUCCESS: 3D view created (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                return view3D;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR creating 3D view: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create section box using the WORKING Front/Back transforms as templates for all directions
        /// This is the key fix - use proven working transforms
        /// </summary>
        private static BoundingBoxXYZ CreateWorkingSectionBox(BoundingBoxXYZ bounds, ViewDirection direction)
        {
            try
            {
                var center = (bounds.Min + bounds.Max) / 2;
                var size = bounds.Max - bounds.Min;

                // Base sizes (minimums to avoid degenerate boxes)
                double width = Math.Max(size.X, 10.0);
                double depth = Math.Max(size.Y, 10.0);
                double height = Math.Max(size.Z, 10.0);

                // Smart padding:
                // - small fraction of the model size so padding scales with object
                // - clamp between 5 and 40 (you can reduce 40 if you want even tighter)
                double maxModelDim = Math.Max(Math.Max(size.X, size.Y), size.Z);
                double padding = Math.Max(5.0, Math.Min(40.0, maxModelDim * 0.05)); // 5% of largest dimension, clamped

                // Also cap absolute extents to avoid extremely large boxes for huge models
                double maxExtentMultiplier = 1.5; // relative to model largest dimension
                double maxExtent = maxModelDim * maxExtentMultiplier;

                var sectionBox = new BoundingBoxXYZ();
                var transform = Transform.Identity;

                System.Diagnostics.Debug.WriteLine($"Creating {direction} with W={width:F1}, D={depth:F1}, H={height:F1}, pad={padding:F1}");

                switch (direction)
                {
                    case ViewDirection.Front:
                        // local extents: X = width, Y = height, Z = depth
                        sectionBox.Min = new XYZ(-Math.Min(width / 2 + padding, maxExtent), -Math.Min(height / 2 + padding, maxExtent), 0);
                        sectionBox.Max = new XYZ(Math.Min(width / 2 + padding, maxExtent), Math.Min(height / 2 + padding, maxExtent), Math.Min(depth + padding, maxExtent));
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Max.Y + padding, center.Z);
                        break;

                    case ViewDirection.Back:
                        sectionBox.Min = new XYZ(-Math.Min(width / 2 + padding, maxExtent), -Math.Min(height / 2 + padding, maxExtent), 0);
                        sectionBox.Max = new XYZ(Math.Min(width / 2 + padding, maxExtent), Math.Min(height / 2 + padding, maxExtent), Math.Min(depth + padding, maxExtent));
                        transform.BasisX = -XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Min.Y - padding, center.Z);
                        break;

                    case ViewDirection.Left:
                        // local extents: X = depth, Y = height, Z = width (rotated)
                        sectionBox.Min = new XYZ(-Math.Min(depth / 2 + padding, maxExtent), -Math.Min(height / 2 + padding, maxExtent), 0);
                        sectionBox.Max = new XYZ(Math.Min(depth / 2 + padding, maxExtent), Math.Min(height / 2 + padding, maxExtent), Math.Min(width + padding, maxExtent));
                        transform.BasisX = -XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Min.X - padding, center.Y, center.Z);
                        break;

                    case ViewDirection.Right:
                        sectionBox.Min = new XYZ(-Math.Min(depth / 2 + padding, maxExtent), -Math.Min(height / 2 + padding, maxExtent), 0);
                        sectionBox.Max = new XYZ(Math.Min(depth / 2 + padding, maxExtent), Math.Min(height / 2 + padding, maxExtent), Math.Min(width + padding, maxExtent));
                        // Look into the model from +X side
                        transform.BasisX = -XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Max.X + padding, center.Y, center.Z);
                        break;

                    case ViewDirection.Top:
                        sectionBox.Min = new XYZ(-Math.Min(width / 2 + padding, maxExtent), -Math.Min(depth / 2 + padding, maxExtent), 0);
                        sectionBox.Max = new XYZ(Math.Min(width / 2 + padding, maxExtent), Math.Min(depth / 2 + padding, maxExtent), Math.Min(height + padding, maxExtent));
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = -XYZ.BasisY;
                        transform.BasisZ = -XYZ.BasisZ;
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Max.Z + padding);
                        break;

                    case ViewDirection.Bottom:
                        sectionBox.Min = new XYZ(-Math.Min(width / 2 + padding, maxExtent), -Math.Min(depth / 2 + padding, maxExtent), 0);
                        sectionBox.Max = new XYZ(Math.Min(width / 2 + padding, maxExtent), Math.Min(depth / 2 + padding, maxExtent), Math.Min(height + padding, maxExtent));
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisY;
                        transform.BasisZ = XYZ.BasisZ;
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Min.Z - padding);
                        break;
                }

                // assign transform
                sectionBox.Transform = transform;

                // Validate local box size
                var boxSize = sectionBox.Max - sectionBox.Min;
                if (boxSize.X <= 0 || boxSize.Y <= 0 || boxSize.Z <= 0)
                {
                    System.Diagnostics.Debug.WriteLine($"ERROR: Invalid box size for {direction}: {boxSize.X:F1}x{boxSize.Y:F1}x{boxSize.Z:F1}");
                    return null;
                }

                // Orthonormalize the basis (safe-guard) — use BasisX/BasisY, compute BasisZ = BasisX x BasisY
                try
                {
                    var bx = transform.BasisX.Normalize();
                    var by = transform.BasisY.Normalize();
                    var bz = bx.CrossProduct(by);

                    double bzLen = bz.GetLength();
                    if (bzLen < 1e-6)
                    {
                        System.Diagnostics.Debug.WriteLine($"ERROR: Invalid basis vectors for {direction} (degenerate cross product).");
                        return null;
                    }

                    bz = bz.Divide(bzLen);
                    transform.BasisX = bx;
                    transform.BasisY = by;
                    transform.BasisZ = bz;
                    sectionBox.Transform = transform;

                    System.Diagnostics.Debug.WriteLine($"  {direction} Orthonormalized BasisX: ({bx.X:F2},{bx.Y:F2},{bx.Z:F2})");
                    System.Diagnostics.Debug.WriteLine($"  {direction} Orthonormalized BasisY: ({by.X:F2},{by.Y:F2},{by.Z:F2})");
                    System.Diagnostics.Debug.WriteLine($"  {direction} Orthonormalized BasisZ: ({bz.X:F2},{bz.Y:F2},{bz.Z:F2})");
                    System.Diagnostics.Debug.WriteLine($"  {direction} Box size (local): {boxSize.X:F1} x {boxSize.Y:F1} x {boxSize.Z:F1}");
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine($"ERROR orthonormalizing transform for {direction}: {e.Message}");
                    return null;
                }

                return sectionBox;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR creating section box for {direction}: {ex.Message}");
                return null;
            }
        }

        #region Helper Methods

        /// <summary>
        /// Calculate overall bounding box from elements
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
                catch { }
            }

            if (!hasValidBounds) return null;

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
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
                    return sectionTypeId;

                // Fallback: find any section view type
                var sectionTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .Where(vt => vt.ViewFamily == ViewFamily.Section)
                    .ToList();

                return sectionTypes.Count > 0 ? sectionTypes.First().Id : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Log section box details for debugging
        /// </summary>
        private static void LogSectionBoxDetails(ViewDirection direction, BoundingBoxXYZ sectionBox)
        {
            var transform = sectionBox.Transform;
            System.Diagnostics.Debug.WriteLine($"  {direction} Origin: ({transform.Origin.X:F1}, {transform.Origin.Y:F1}, {transform.Origin.Z:F1})");
            System.Diagnostics.Debug.WriteLine($"  {direction} BasisZ: ({transform.BasisZ.X:F2}, {transform.BasisZ.Y:F2}, {transform.BasisZ.Z:F2})");
        }

        /// <summary>
        /// Set unique view name
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
                    viewName = $"{baseName} ({i + 1})";
                }
            }
        }

        #endregion
    }
}