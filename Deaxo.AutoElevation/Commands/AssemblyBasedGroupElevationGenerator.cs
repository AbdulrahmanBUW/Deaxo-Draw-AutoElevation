// AssemblyBasedGroupElevationGenerator.cs - Fixed Version
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Creates group elevations by temporarily using Revit's native Assembly system
    /// then extracting and renaming the views, and safely cleaning up the assembly
    /// FIXED VERSION - Handles missing Assembly references gracefully
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
        /// Creates group elevations - falls back to simplified approach if Assembly API is not available
        /// </summary>
        public static AssemblyViewResult CreateGroupViews(Document doc, List<Element> elements, View viewTemplate = null)
        {
            if (elements == null || elements.Count == 0)
                return null;

            try
            {
                System.Diagnostics.Debug.WriteLine("=== Attempting Assembly-Based Group Views ===");
                System.Diagnostics.Debug.WriteLine($"Processing {elements.Count} elements");

                // Check if Assembly functionality is available
                if (!IsAssemblyFunctionalityAvailable(doc))
                {
                    System.Diagnostics.Debug.WriteLine("Assembly functionality not available, falling back to simplified approach");
                    return CreateFallbackGroupViews(doc, elements, viewTemplate);
                }

                // Try to use Assembly approach
                AssemblyInstance tempAssemblyInstance = null;
                try
                {
                    // Step 1: Create temporary assembly
                    tempAssemblyInstance = CreateTemporaryAssembly(doc, elements);
                    if (tempAssemblyInstance == null)
                    {
                        System.Diagnostics.Debug.WriteLine("Failed to create temporary assembly, using fallback");
                        return CreateFallbackGroupViews(doc, elements, viewTemplate);
                    }

                    System.Diagnostics.Debug.WriteLine($"Created temporary assembly instance: {tempAssemblyInstance.Id}");

                    // Step 2: Create assembly views using Revit's native system
                    var assemblyViews = CreateAssemblyViews(doc, tempAssemblyInstance);
                    if (assemblyViews == null || assemblyViews.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine("Failed to create assembly views, using fallback");
                        return CreateFallbackGroupViews(doc, elements, viewTemplate);
                    }

                    System.Diagnostics.Debug.WriteLine($"Created {assemblyViews.Count} assembly views");

                    // Step 3: Extract and rename the views
                    var result = ExtractAndRenameViews(doc, assemblyViews, viewTemplate);

                    System.Diagnostics.Debug.WriteLine($"Successfully extracted {result.AllSectionViews.Count + (result.ThreeDView != null ? 1 : 0)} views");

                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Assembly approach failed: {ex.Message}");
                    return CreateFallbackGroupViews(doc, elements, viewTemplate);
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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in CreateGroupViews: {ex.Message}");
                // Final fallback to simplified approach
                return CreateFallbackGroupViews(doc, elements, viewTemplate);
            }
        }

        /// <summary>
        /// Check if Assembly functionality is available in this Revit version
        /// </summary>
        private static bool IsAssemblyFunctionalityAvailable(Document doc)
        {
            try
            {
                // Try to access AssemblyInstance type to see if it's available
                var assemblyInstanceType = typeof(AssemblyInstance);

                // Try to get the Create method to see if it exists
                var createMethod = assemblyInstanceType.GetMethod("Create",
                    new Type[] { typeof(Document), typeof(ICollection<ElementId>), typeof(ElementId) });

                return createMethod != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Fallback method using simplified group elevation generator
        /// </summary>
        private static AssemblyViewResult CreateFallbackGroupViews(Document doc, List<Element> elements, View viewTemplate)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Using fallback simplified group elevation generator");

                // Use the SimplifiedGroupElevationGenerator as fallback
                var simplifiedResult = SimplifiedGroupElevationGenerator.CreateGroupViews(doc, elements, viewTemplate);

                if (simplifiedResult == null) return null;

                // Convert SimplifiedGroupElevationGenerator.GroupViewResult to AssemblyViewResult
                return new AssemblyViewResult
                {
                    TopView = simplifiedResult.TopView,
                    BottomView = simplifiedResult.BottomView,
                    LeftView = simplifiedResult.LeftView,
                    RightView = simplifiedResult.RightView,
                    FrontView = simplifiedResult.FrontView,
                    BackView = simplifiedResult.BackView,
                    ThreeDView = simplifiedResult.ThreeDView
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fallback approach also failed: {ex.Message}");
                return null;
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

                if (elementIds.Count == 0) return null;

                // Get the category of the first element
                var firstElement = elements.First();
                if (firstElement.Category == null) return null;

                // Create assembly instance - Use reflection to handle different API versions
                ElementId assemblyId = null;
                try
                {
                    // Try the standard method signature first
                    var createMethod = typeof(AssemblyInstance).GetMethod("Create",
                        new Type[] { typeof(Document), typeof(ICollection<ElementId>), typeof(ElementId) });

                    if (createMethod != null)
                    {
                        // Method exists, call it
                        var result = createMethod.Invoke(null, new object[] { doc, elementIds, firstElement.Category.Id });
                        assemblyId = result as ElementId;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("AssemblyInstance.Create method not found");
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"AssemblyInstance.Create failed: {ex.Message}");
                    return null;
                }

                if (assemblyId == null || assemblyId == ElementId.InvalidElementId)
                {
                    System.Diagnostics.Debug.WriteLine("AssemblyInstance.Create returned invalid ID");
                    return null;
                }

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
                    // Get the assembly type - FIXED: Use proper method
                    var assemblyTypeId = assemblyInstance.GetTypeId();
                    var assemblyType = doc.GetElement(assemblyTypeId);

                    if (assemblyType != null)
                    {
                        System.Diagnostics.Debug.WriteLine("Found assembly type, attempting view creation");

                        // For now, fall back to manual creation since Assembly view creation API varies by version
                        var fallbackViews = CreateFallbackAssemblyViews(doc, assemblyInstance, sectionViewType);
                        if (fallbackViews != null)
                        {
                            createdViews.AddRange(fallbackViews);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Assembly view creation failed: {ex.Message}");

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
        /// Clean up the temporary assembly
        /// </summary>
        private static void CleanupTemporaryAssembly(Document doc, AssemblyInstance assemblyInstance)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Cleaning up temporary assembly: {assemblyInstance.Id}");

                // Get assembly type before deleting instance
                var assemblyTypeId = assemblyInstance.GetTypeId();

                // Delete assembly instance - FIXED: Use ElementId
                doc.Delete(assemblyInstance.Id);
                System.Diagnostics.Debug.WriteLine($"Deleted assembly instance: {assemblyInstance.Id}");

                // Delete assembly type
                if (assemblyTypeId != null && assemblyTypeId != ElementId.InvalidElementId)
                {
                    doc.Delete(assemblyTypeId);
                    System.Diagnostics.Debug.WriteLine($"Deleted assembly type: {assemblyTypeId}");
                }
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