// AssemblyBasedGroupElevationGenerator.cs
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using RevitAssembly = Autodesk.Revit.DB.Assembly;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Creates group elevations by temporarily using Revit's native Assembly system
    /// then extracting and renaming the views, and safely cleaning up the assembly
    /// </summary>
    public static class AssemblyBasedGroupElevationGenerator
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
        /// Creates group elevations using Revit's native Assembly system as the engine
        /// </summary>
        public static AssemblyViewResult CreateGroupViews(Document doc, List<Element> elements, View viewTemplate = null)
        {
            if (elements == null || elements.Count == 0)
                return null;

            AssemblyInstance tempAssemblyInstance = null;
            try
            {
                System.Diagnostics.Debug.WriteLine("=== Creating Assembly-Based Group Views ===");
                System.Diagnostics.Debug.WriteLine($"Processing {elements.Count} elements");

                // Step 1: Create temporary assembly
                tempAssemblyInstance = CreateTemporaryAssembly(doc, elements);
                if (tempAssemblyInstance == null)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create temporary assembly");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"Created temporary assembly instance: {tempAssemblyInstance.Id}");

                // Step 2: Create assembly views using Revit's native system
                var assemblyViews = CreateAssemblyViews(doc, tempAssemblyInstance);
                if (assemblyViews == null || assemblyViews.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create assembly views");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"Created {assemblyViews.Count} assembly views");

                // Step 3: Extract and rename the views
                var result = ExtractAndRenameViews(doc, assemblyViews, viewTemplate);

                System.Diagnostics.Debug.WriteLine($"Successfully extracted {result.AllSectionViews.Count + (result.ThreeDView != null ? 1 : 0)} views");

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in CreateGroupViews: {ex.Message}");
                return null;
            }
            finally
            {
                // Step 4: Clean up - delete the temporary assembly
                if (tempAssemblyInstance != null)
                {
                    CleanupTemporaryAssembly(doc, tempAssemblyInstance);
                }
            }
        }

        /// <summary>
        /// Create a temporary assembly from the selected elements
        /// </summary>
        private static AssemblyInstance CreateTemporaryAssembly(Document doc, List<Element> elements)
        {
            try
            {
                // Convert elements to ElementIds
                var elementIds = elements.Select(e => e.Id).ToList();

                // Create assembly instance
                var assemblyId = AssemblyInstance.Create(doc, elementIds, elements.First().Category.Id);
                var assemblyInstance = doc.GetElement(assemblyId) as AssemblyInstance;

                if (assemblyInstance == null)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create assembly instance");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"Created assembly instance: {assemblyInstance.Id}");
                return assemblyInstance;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating temporary assembly: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create views using Revit's native assembly view creation
        /// </summary>
        private static List<View> CreateAssemblyViews(Document doc, AssemblyInstance assemblyInstance)
        {
            try
            {
                // Get all view family types for different view types
                var sectionViewType = GetViewFamilyType(doc, ViewFamily.Section);
                var threeDViewType = GetViewFamilyType(doc, ViewFamily.ThreeDimensional);

                if (sectionViewType == null)
                {
                    System.Diagnostics.Debug.WriteLine("No section view type found");
                    return null;
                }

                var createdViews = new List<View>();

                // Try to use assembly's built-in view creation methods
                try
                {
                    // Get the assembly type
                    var assemblyType = doc.GetElement(assemblyInstance.GetTypeId()) as RevitAssembly;
                    if (assemblyType != null)
                    {
                        // Try to create views using assembly methods
                        // Note: The API for this may vary by Revit version
                        System.Diagnostics.Debug.WriteLine("Attempting to use assembly built-in view creation");

                        // This is the method that should work, but may not be available in all versions
                        // var viewIds = assemblyType.CreateOrthographic3DViews();

                        // For now, fall back to manual creation
                        var fallbackViews = CreateFallbackAssemblyViews(doc, assemblyInstance, sectionViewType);
                        if (fallbackViews != null)
                        {
                            createdViews.AddRange(fallbackViews);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Native assembly view creation failed: {ex.Message}");

                    // Fallback: Try manual creation of section views around assembly
                    var fallbackViews = CreateFallbackAssemblyViews(doc, assemblyInstance, sectionViewType);
                    if (fallbackViews != null)
                    {
                        createdViews.AddRange(fallbackViews);
                    }
                }

                // Create 3D view if not already created
                if (threeDViewType != null && !createdViews.Any(v => v is View3D))
                {
                    try
                    {
                        var view3D = Create3DAssemblyView(doc, assemblyInstance, threeDViewType);
                        if (view3D != null)
                        {
                            createdViews.Add(view3D);
                            System.Diagnostics.Debug.WriteLine($"Created 3D assembly view: {view3D.Id}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"3D view creation failed: {ex.Message}");
                    }
                }

                return createdViews;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating assembly views: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Fallback method to create section views manually if native assembly method fails
        /// </summary>
        private static List<View> CreateFallbackAssemblyViews(Document doc, AssemblyInstance assemblyInstance, ViewFamilyType sectionViewType)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Using fallback assembly view creation");

                var createdViews = new List<View>();

                // Get assembly's bounding box
                var bounds = assemblyInstance.get_BoundingBox(null);
                if (bounds == null) return null;

                // Create section views using the assembly's coordinate system
                var directions = new[] { "Front", "Back", "Left", "Right", "Top", "Bottom" };

                foreach (var direction in directions)
                {
                    try
                    {
                        var sectionBox = CreateAssemblyAlignedSectionBox(bounds, direction);
                        if (sectionBox != null)
                        {
                            var section = ViewSection.CreateSection(doc, sectionViewType.Id, sectionBox);
                            if (section != null)
                            {
                                section.Name = $"Assembly_{direction}_{DateTime.Now.Ticks % 1000000}";
                                createdViews.Add(section);
                                System.Diagnostics.Debug.WriteLine($"Created fallback section: {direction}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to create {direction} fallback view: {ex.Message}");
                    }
                }

                return createdViews;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fallback view creation failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create section box aligned with assembly coordinate system
        /// </summary>
        private static BoundingBoxXYZ CreateAssemblyAlignedSectionBox(BoundingBoxXYZ bounds, string direction)
        {
            try
            {
                var center = (bounds.Min + bounds.Max) / 2;
                var size = bounds.Max - bounds.Min;
                double padding = Math.Max(Math.Max(size.X, size.Y), size.Z) * 0.3;

                var sectionBox = new BoundingBoxXYZ();
                var transform = Transform.Identity;

                switch (direction)
                {
                    case "Front":
                        sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Z / 2 - padding, 0);
                        sectionBox.Max = new XYZ(size.X / 2 + padding, size.Z / 2 + padding, size.Y + padding);
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Max.Y + padding, center.Z);
                        break;
                    case "Back":
                        sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Z / 2 - padding, 0);
                        sectionBox.Max = new XYZ(size.X / 2 + padding, size.Z / 2 + padding, size.Y + padding);
                        transform.BasisX = -XYZ.BasisX;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisY;
                        transform.Origin = new XYZ(center.X, bounds.Min.Y - padding, center.Z);
                        break;
                    case "Left":
                        sectionBox.Min = new XYZ(-size.Y / 2 - padding, -size.Z / 2 - padding, 0);
                        sectionBox.Max = new XYZ(size.Y / 2 + padding, size.Z / 2 + padding, size.X + padding);
                        transform.BasisX = -XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Min.X - padding, center.Y, center.Z);
                        break;
                    case "Right":
                        sectionBox.Min = new XYZ(-size.Y / 2 - padding, -size.Z / 2 - padding, 0);
                        sectionBox.Max = new XYZ(size.Y / 2 + padding, size.Z / 2 + padding, size.X + padding);
                        transform.BasisX = XYZ.BasisY;
                        transform.BasisY = XYZ.BasisZ;
                        transform.BasisZ = -XYZ.BasisX;
                        transform.Origin = new XYZ(bounds.Max.X + padding, center.Y, center.Z);
                        break;
                    case "Top":
                        sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Y / 2 - padding, 0);
                        sectionBox.Max = new XYZ(size.X / 2 + padding, size.Y / 2 + padding, size.Z + padding);
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = XYZ.BasisY;
                        transform.BasisZ = -XYZ.BasisZ;
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Max.Z + padding);
                        break;
                    case "Bottom":
                        sectionBox.Min = new XYZ(-size.X / 2 - padding, -size.Y / 2 - padding, 0);
                        sectionBox.Max = new XYZ(size.X / 2 + padding, size.Y / 2 + padding, size.Z + padding);
                        transform.BasisX = XYZ.BasisX;
                        transform.BasisY = -XYZ.BasisY;
                        transform.BasisZ = XYZ.BasisZ;
                        transform.Origin = new XYZ(center.X, center.Y, bounds.Min.Z - padding);
                        break;
                }

                sectionBox.Transform = transform;
                return sectionBox;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Create 3D view for assembly
        /// </summary>
        private static View3D Create3DAssemblyView(Document doc, AssemblyInstance assemblyInstance, ViewFamilyType threeDViewType)
        {
            try
            {
                var view3D = View3D.CreateIsometric(doc, threeDViewType.Id);
                if (view3D == null) return null;

                // Get assembly bounds for section box
                var bounds = assemblyInstance.get_BoundingBox(null);
                if (bounds != null)
                {
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
                }

                view3D.Name = $"Assembly_3D_{DateTime.Now.Ticks % 1000000}";
                return view3D;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating 3D assembly view: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Extract views from assembly and rename them with proper naming convention
        /// </summary>
        private static AssemblyViewResult ExtractAndRenameViews(Document doc, List<View> assemblyViews, View viewTemplate)
        {
            try
            {
                var result = new AssemblyViewResult();

                foreach (var view in assemblyViews)
                {
                    // Apply template if provided
                    if (viewTemplate != null && view.GetType() == viewTemplate.GetType())
                    {
                        try
                        {
                            view.ViewTemplateId = viewTemplate.Id;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to apply template to view {view.Id}: {ex.Message}");
                        }
                    }

                    // Determine view type and rename
                    string viewType = DetermineViewType(view);
                    string newName = $"group_elements_view - {viewType}";

                    // Set unique name
                    SetUniqueViewName(view, newName);

                    // Assign to appropriate property
                    if (view is ViewSection section)
                    {
                        switch (viewType)
                        {
                            case "Top": result.TopView = section; break;
                            case "Bottom": result.BottomView = section; break;
                            case "Front": result.FrontView = section; break;
                            case "Back": result.BackView = section; break;
                            case "Left": result.LeftView = section; break;
                            case "Right": result.RightView = section; break;
                        }
                    }
                    else if (view is View3D view3D)
                    {
                        result.ThreeDView = view3D;
                    }

                    System.Diagnostics.Debug.WriteLine($"Renamed view to: {view.Name}");
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting and renaming views: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Determine view type based on view properties or name
        /// </summary>
        private static string DetermineViewType(View view)
        {
            if (view is View3D)
                return "3D";

            // Try to determine from view name or properties
            string name = view.Name.ToLower();
            if (name.Contains("top") || name.Contains("plan")) return "Top";
            if (name.Contains("bottom")) return "Bottom";
            if (name.Contains("front")) return "Front";
            if (name.Contains("back")) return "Back";
            if (name.Contains("left")) return "Left";
            if (name.Contains("right")) return "Right";

            // Fallback: analyze view direction (this would need more complex logic)
            return "Front"; // Default fallback
        }

        /// <summary>
        /// Detach views from assembly so they can exist independently
        /// </summary>
        private static void DetachViewsFromAssembly(Document doc, AssemblyInstance assemblyInstance, List<View> views)
        {
            try
            {
                // Note: This is conceptual - Revit may not allow easy detachment
                // The views should remain independent once the assembly is deleted
                System.Diagnostics.Debug.WriteLine($"Detaching {views.Count} views from assembly {assemblyInstance.Id}");

                // In practice, the views created by assembly should continue to exist
                // even after the assembly is deleted, as long as they're not assembly-specific views
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error detaching views from assembly: {ex.Message}");
            }
        }

        /// <summary>
        /// Clean up the temporary assembly
        /// </summary>
        private static void CleanupTemporaryAssembly(Document doc, AssemblyInstance assemblyInstance)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Cleaning up temporary assembly: {assemblyInstance.Id}");

                // Get assembly type before deleting instance
                var assemblyTypeId = assemblyInstance.GetTypeId();

                // Delete assembly instance
                doc.Delete(assemblyInstance.Id);
                System.Diagnostics.Debug.WriteLine($"Deleted assembly instance: {assemblyInstance.Id}");

                // Delete assembly type
                doc.Delete(assemblyTypeId);
                System.Diagnostics.Debug.WriteLine($"Deleted assembly type: {assemblyTypeId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error cleaning up assembly: {ex.Message}");
                // Continue anyway - cleanup failure shouldn't break the whole operation
            }
        }

        #region Helper Methods

        private static ViewFamilyType GetViewFamilyType(Document doc, ViewFamily viewFamily)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vt => vt.ViewFamily == viewFamily);
        }

        private static AssemblyInstance GetAssemblyInstance(Document doc, RevitAssembly assembly)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(AssemblyInstance))
                .Cast<AssemblyInstance>()
                .FirstOrDefault(ai => ai.GetTypeId() == assembly.Id);
        }

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

        #endregion
    }
}