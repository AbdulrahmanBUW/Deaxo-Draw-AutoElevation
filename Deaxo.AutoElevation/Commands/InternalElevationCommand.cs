// Complete InternalElevationCommand.cs with Enhanced Template Handling and Sheet Generation
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Deaxo.AutoElevation.UI;

namespace Deaxo.AutoElevation.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class InternalElevationCommand : IExternalCommand
    {
        // Static counter to maintain sequence across all elevations in this command execution
        private static int sequenceCounter = 1;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Reset sequence counter at start of command
                sequenceCounter = 1;

                var selectOpts = new Dictionary<string, object>()
                {
                    {"Walls", BuiltInCategory.OST_Walls},
                    {"Curtain Wall Panels", BuiltInCategory.OST_CurtainWallPanels},
                    {"Curtain Wall Mullions", BuiltInCategory.OST_CurtainWallMullions}
                };

                var allowedTypesOrCats = new List<object>();
                foreach (var kvp in selectOpts)
                {
                    var val = kvp.Value;
                    if (val is BuiltInCategory bic)
                        allowedTypesOrCats.Add(bic);
                    else if (val is BuiltInCategory[] bicArr)
                        allowedTypesOrCats.AddRange(bicArr.Cast<object>());
                    else
                        allowedTypesOrCats.Add(val);
                }

                var selFilter = new DXSelectionFilter(allowedTypesOrCats);
                IList<Reference> refs;
                try
                {
                    refs = uidoc.Selection.PickObjects(ObjectType.Element, selFilter,
                        "Select walls for internal building elevations and click Finish");
                }
                catch (OperationCanceledException)
                {
                    TaskDialog.Show("DEAXO", "Selection cancelled.");
                    return Result.Cancelled;
                }

                if (refs == null || refs.Count == 0)
                {
                    TaskDialog.Show("DEAXO", "No walls selected.");
                    return Result.Cancelled;
                }

                // Ask user for base name ONCE at the beginning
                string baseName = GetBaseNameFromUser();
                if (baseName == null)
                {
                    TaskDialog.Show("DEAXO", "Operation cancelled.");
                    return Result.Cancelled;
                }

                // Get sheet configuration from user
                var sheetConfig = GetSheetConfiguration();
                if (sheetConfig == null)
                {
                    TaskDialog.Show("DEAXO", "Operation cancelled.");
                    return Result.Cancelled;
                }

                System.Diagnostics.Debug.WriteLine($"Sheet configuration: {sheetConfig}");

                // Get view template with improved handling
                View chosenTemplate = GetViewTemplate(doc);
                if (chosenTemplate != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Using view template: {chosenTemplate.Name} (Type: {chosenTemplate.ViewType}, ID: {chosenTemplate.Id})");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No view template selected - using default formatting");
                }

                var progressWindow = new ProgressWindow();
                progressWindow.Show();
                progressWindow.UpdateStatus("Creating internal building elevations...", $"Using base name: {baseName}");

                var results = new List<string>();
                var startTime = DateTime.Now;

                // Update transaction name based on sheet option - C# 7.3 compatible
                string transactionName;
                if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                    transactionName = "DEAXO - Create Internal Building Elevations";
                else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                    transactionName = "DEAXO - Create Internal Building Elevations & Individual Sheets";
                else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                    transactionName = "DEAXO - Create Internal Building Elevations & Combined Sheet";
                else
                    transactionName = "DEAXO - Create Internal Building Elevations";

                using (Transaction t = new Transaction(doc, transactionName))
                {
                    t.Start();

                    try
                    {
                        progressWindow.UpdateProgress(0, refs.Count);

                        // Calculate total operations based on sheet option - C# 7.3 compatible
                        int totalOperations;
                        if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                            totalOperations = refs.Count; // Only views
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                            totalOperations = refs.Count * 2; // Views + individual sheets
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                            totalOperations = refs.Count + 1; // Views + one combined sheet
                        else
                            totalOperations = refs.Count;

                        int currentOperation = 0;
                        int templateAppliedCount = 0;
                        int sheetCount = 0; // Declare sheetCount variable
                        var createdElevations = new List<(ViewSection elevation, Wall wall, bool templateApplied)>();

                        for (int i = 0; i < refs.Count; i++)
                        {
                            var r = refs[i];

                            try
                            {
                                Element el = doc.GetElement(r);
                                progressWindow.UpdateStatus($"Processing wall {i + 1} of {refs.Count}...",
                                    $"Wall ID: {el.Id}, Creating: {baseName}_{sequenceCounter}");

                                if (!(el is Wall wall))
                                {
                                    results.Add($"Skipped element {el.Id}: Not a wall (Type: {el.GetType().Name})");
                                    continue;
                                }

                                // Create elevation view
                                var elevation = CreateInternalBuildingElevation(doc, wall, chosenTemplate, baseName);
                                if (elevation == null)
                                {
                                    results.Add($"Failed to create building elevation for wall {el.Id}");
                                    continue;
                                }

                                // Check if template was applied
                                bool templateApplied = (chosenTemplate != null && elevation.ViewTemplateId == chosenTemplate.Id);
                                if (templateApplied) templateAppliedCount++;

                                createdElevations.Add((elevation, wall, templateApplied));
                                results.Add($"✓ Created building elevation {elevation.Id} ({elevation.Name}) for wall {el.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                                currentOperation++;
                                progressWindow.UpdateProgress(currentOperation, totalOperations);
                                progressWindow.AddLogMessage($"Created elevation: {elevation.Name} for wall {wall.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            }
                            catch (Exception exInner)
                            {
                                var error = $"Error processing wall {r.ElementId}: {exInner.Message}";
                                results.Add($"✗ {error}");
                                progressWindow.AddLogMessage(error);
                            }
                        }

                        // Handle sheet creation based on user choice
                        if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                        {
                            progressWindow.AddLogMessage("Skipping sheet creation as requested");
                            results.Add("✓ Internal elevations created without sheets as requested");
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                        {
                            progressWindow.UpdateStatus("Creating individual sheets...", "Generating separate sheet for each elevation");
                            progressWindow.AddLogMessage("Starting individual sheet creation process");

                            foreach (var elevationTuple in createdElevations)
                            {
                                var elevationView = elevationTuple.elevation;
                                var wallElement = elevationTuple.wall;
                                var templateApplied = elevationTuple.templateApplied;

                                try
                                {
                                    progressWindow.UpdateStatus($"Creating sheet for elevation {elevationView.Name}...",
                                        $"Placing view {elevationView.Id} on new sheet");

                                    var sheet = CreateIndividualSheetForElevation(doc, elevationView, wallElement, baseName, results);
                                    if (sheet != null)
                                    {
                                        sheetCount++;
                                        results.Add($"✓ Created individual sheet {sheet.Id} ({sheet.SheetNumber}) for elevation {elevationView.Id}");
                                        progressWindow.AddLogMessage($"Created individual sheet: {sheet.SheetNumber} for elevation {elevationView.Name}");
                                    }

                                    currentOperation++;
                                    progressWindow.UpdateProgress(currentOperation, totalOperations);
                                }
                                catch (Exception sheetEx)
                                {
                                    var sheetError = $"Error creating individual sheet for elevation {elevationView.Id}: {sheetEx.Message}";
                                    results.Add($"✗ {sheetError}");
                                    progressWindow.AddLogMessage(sheetError);
                                }
                            }
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                        {
                            progressWindow.UpdateStatus("Creating combined sheet...", "Placing all elevations on single sheet");
                            progressWindow.AddLogMessage("Starting combined sheet creation process");

                            try
                            {
                                var combinedSheet = CreateCombinedSheetForElevations(doc, createdElevations, baseName);
                                if (combinedSheet != null)
                                {
                                    sheetCount = 1;
                                    results.Add($"✓ Created combined sheet {combinedSheet.Id} ({combinedSheet.SheetNumber}) with {createdElevations.Count} elevations");
                                    progressWindow.AddLogMessage($"Created combined sheet: {combinedSheet.SheetNumber} with {createdElevations.Count} elevations");
                                    System.Diagnostics.Debug.WriteLine($"✓ COMBINED SHEET: Created {combinedSheet.Id} with {createdElevations.Count} elevations");
                                }
                                else
                                {
                                    results.Add($"✗ Failed to create combined sheet");
                                    System.Diagnostics.Debug.WriteLine($"✗ COMBINED SHEET: Failed");
                                }

                                currentOperation++;
                                progressWindow.UpdateProgress(currentOperation, totalOperations);
                            }
                            catch (Exception ex)
                            {
                                results.Add($"✗ Error creating combined sheet: {ex.Message}");
                                System.Diagnostics.Debug.WriteLine($"✗ COMBINED SHEET ERROR: {ex.Message}");
                            }
                        }

                        // Update final results based on sheet option
                        if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                        {
                            results.Add($"✓ Total: {createdElevations.Count} internal elevations created (no sheets)");
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                        {
                            results.Add($"✓ Total: {createdElevations.Count} internal elevations and {sheetCount} individual sheets created");
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                        {
                            results.Add($"✓ Total: {createdElevations.Count} internal elevations on {sheetCount} combined sheet created");
                        }

                        t.Commit();

                        var duration = DateTime.Now - startTime;

                        // Show completion with appropriate title
                        string completionTitle;
                        if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                            completionTitle = "Internal Building Elevations";
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                            completionTitle = "Internal Building Elevations & Individual Sheets";
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                            completionTitle = "Internal Building Elevations & Combined Sheet";
                        else
                            completionTitle = "Internal Building Elevations";

                        progressWindow.ShowCompletion(results, completionTitle);

                        System.Diagnostics.Debug.WriteLine($"=== INTERNAL ELEVATION COMMAND COMPLETED ===");
                        System.Diagnostics.Debug.WriteLine($"Total duration: {duration.TotalSeconds:F1} seconds");
                        System.Diagnostics.Debug.WriteLine($"Template applied to {templateAppliedCount} elevations");

                        System.Threading.Thread.Sleep(1500);
                        progressWindow.Close();
                    }
                    catch (Exception ex)
                    {
                        progressWindow.ShowError($"Error in internal elevation command: {ex.Message}");
                        t.RollBack();
                        throw;
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        #region Sheet Configuration

        /// <summary>
        /// Gets sheet configuration from user
        /// </summary>
        private SheetConfigurationWindow.SheetOption? GetSheetConfiguration()
        {
            try
            {
                var configWindow = new SheetConfigurationWindow();
                bool? result = configWindow.ShowDialog();

                if (result == true)
                {
                    return configWindow.SelectedOption;
                }

                return null; // User cancelled
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting sheet configuration: {ex.Message}");
                return SheetConfigurationWindow.SheetOption.Individual; // Fallback to individual sheets
            }
        }

        #endregion

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

        #region User Input Methods

        /// <summary>
        /// Prompts user for base name input
        /// </summary>
        private string GetBaseNameFromUser()
        {
            try
            {
                // Create a simple input dialog
                var inputWindow = new BaseNameInputWindow();
                bool? result = inputWindow.ShowDialog();

                if (result == true)
                {
                    string baseName = inputWindow.BaseName;

                    // Default to "Elevation" if empty
                    if (string.IsNullOrWhiteSpace(baseName))
                    {
                        baseName = "Elevation";
                    }

                    // Clean the base name (remove invalid characters)
                    baseName = CleanViewName(baseName);

                    return baseName;
                }

                return null; // User cancelled
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting base name: {ex.Message}");
                return "Elevation"; // Fallback to default
            }
        }

        /// <summary>
        /// Clean view name by removing invalid characters
        /// </summary>
        private string CleanViewName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Elevation";

            // Remove characters that might cause issues in Revit view names
            var invalidChars = new char[] { '/', '\\', ':', '*', '?', '"', '<', '>', '|', '{', '}' };

            foreach (char c in invalidChars)
            {
                name = name.Replace(c, '_');
            }

            // Trim whitespace and limit length
            name = name.Trim();
            if (name.Length > 50) // Reasonable limit for view names
            {
                name = name.Substring(0, 50).Trim();
            }

            return string.IsNullOrWhiteSpace(name) ? "Elevation" : name;
        }

        #endregion

        #region Elevation Creation Methods

        private ViewSection CreateInternalBuildingElevation(Document doc, Wall wall, View viewTemplate, string baseName)
        {
            try
            {
                var locationCurve = wall.Location as LocationCurve;
                if (locationCurve?.Curve == null)
                {
                    return null;
                }

                var curve = locationCurve.Curve;
                var startPoint = curve.GetEndPoint(0);
                var endPoint = curve.GetEndPoint(1);
                var wallCenter = (startPoint + endPoint) / 2;
                var wallDirection = (endPoint - startPoint).Normalize();

                var wallNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();
                if (wall.Flipped) wallNormal = -wallNormal;

                var elevation = CreateBuildingElevationView(doc, wall, wallCenter, wallNormal, viewTemplate);
                if (elevation == null)
                {
                    return null;
                }

                // Use the base name with sequence number instead of wall type name
                string elevationName = $"{baseName}_{sequenceCounter}";
                SetUniqueViewName(elevation, elevationName);

                // Increment counter for next elevation
                sequenceCounter++;

                return elevation;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating internal building elevation: {ex.Message}");
                return null;
            }
        }

        private ViewSection CreateBuildingElevationView(Document doc, Wall wall, XYZ wallCenter, XYZ wallNormal, View viewTemplate)
        {
            try
            {
                var elevationViewType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vt => vt.ViewFamily == ViewFamily.Elevation);
                if (elevationViewType == null) return null;

                var internalDirection = -wallNormal;
                double safeDistance = 8.0;
                var markerPosition = wallCenter + (internalDirection * safeDistance);

                BoundingBoxXYZ wallBox = wall.get_BoundingBox(null);
                if (wallBox == null) return null;

                double midZ = (wallBox.Min.Z + wallBox.Max.Z) / 2.0;
                markerPosition = new XYZ(markerPosition.X, markerPosition.Y, midZ);

                // create marker (we will create up to 4 views using this marker)
                var elevationMarker = ElevationMarker.CreateElevationMarker(doc, elevationViewType.Id, markerPosition, 100);
                if (elevationMarker == null) return null;

                ElementId ownerViewId = ElementId.InvalidElementId;
                if (doc.ActiveView?.ViewType == ViewType.FloorPlan ||
                    doc.ActiveView?.ViewType == ViewType.CeilingPlan ||
                    doc.ActiveView?.ViewType == ViewType.AreaPlan)
                {
                    ownerViewId = doc.ActiveView.Id;
                }
                else
                {
                    var planView = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewPlan))
                        .Cast<ViewPlan>()
                        .FirstOrDefault(v => !v.IsTemplate && v.ViewType == ViewType.FloorPlan);
                    if (planView != null) ownerViewId = planView.Id;
                }

                // desired direction in XY plane
                var desired2D = new XYZ(-wallNormal.X, -wallNormal.Y, 0);
                if (desired2D.IsZeroLength())
                {
                    desired2D = new XYZ(0, 1, 0); // fallback
                }
                desired2D = desired2D.Normalize();

                ViewSection chosenView = null;
                double bestAngle = double.MaxValue;
                var createdViewIds = new List<ElementId>();

                // Try all 4 indices and pick best matching view direction
                for (int idx = 0; idx < 4; idx++)
                {
                    var v = elevationMarker.CreateElevation(doc, ownerViewId, idx) as ViewSection;
                    if (v == null) continue;

                    createdViewIds.Add(v.Id);

                    // compute horizontal angle between v.ViewDirection and desired2D
                    var cur2D = new XYZ(v.ViewDirection.X, v.ViewDirection.Y, 0);
                    if (cur2D.IsZeroLength()) continue;
                    cur2D = cur2D.Normalize();

                    // clamp dot to [-1,1] to avoid NaN from rounding
                    double dot = Math.Max(-1.0, Math.Min(1.0, cur2D.X * desired2D.X + cur2D.Y * desired2D.Y));
                    double angle = Math.Acos(dot); // angle in radians (0..PI)
                    double absAngle = Math.Abs(angle);

                    if (absAngle < bestAngle)
                    {
                        bestAngle = absAngle;
                        chosenView = v;
                    }
                }

                if (chosenView == null)
                {
                    // cleanup any created views and return
                    foreach (var id in createdViewIds) { if (id != ElementId.InvalidElementId) doc.Delete(id); }
                    return null;
                }

                // delete the other created views (keep the chosen one)
                foreach (var id in createdViewIds)
                {
                    if (id != chosenView.Id)
                    {
                        try { doc.Delete(id); } catch { /* ignore */ }
                    }
                }

                // Now proceed to set crop, template, properties on the chosen view
                SetElevationCropRegion(chosenView, wall);

                // IMPROVED TEMPLATE APPLICATION:
                if (viewTemplate != null)
                {
                    bool templateApplied = ApplyViewTemplate(chosenView, viewTemplate, $"Internal Elevation for Wall {wall.Id}");
                    System.Diagnostics.Debug.WriteLine($"Template application for wall {wall.Id}: {(templateApplied ? "Success" : "Failed")}");
                }

                SetElevationViewProperties(chosenView, wall);

                return chosenView;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void SetElevationViewProperties(ViewSection elevationView, Wall wall)
        {
            try
            {
                elevationView.DetailLevel = ViewDetailLevel.Fine;
                if (elevationView.Scale > 100)
                {
                    elevationView.Scale = 50;
                }
                elevationView.CropBoxActive = true;
                elevationView.CropBoxVisible = true;
            }
            catch { }
        }

        private void SetElevationCropRegion(ViewSection elevation, Wall wall)
        {
            try
            {
                var wallBounds = wall.get_BoundingBox(null);
                if (wallBounds == null) return;

                var locationCurve = wall.Location as LocationCurve;
                if (locationCurve?.Curve == null) return;

                // Calculate dimensions directly from bounding box (like group elevations do)
                var center = (wallBounds.Min + wallBounds.Max) / 2;
                var size = wallBounds.Max - wallBounds.Min;

                // Use actual dimensions from bounding box (like group elevation approach)
                double width = Math.Max(size.X, 10.0);   // Wall length direction
                double depth = Math.Max(size.Y, 10.0);   // Wall thickness direction  
                double height = Math.Max(size.Z, 10.0);  // Wall height direction

                // Wall endpoints for actual wall length calculation
                XYZ p0 = locationCurve.Curve.GetEndPoint(0);
                XYZ p1 = locationCurve.Curve.GetEndPoint(1);
                double wallLength = p0.DistanceTo(p1);

                // Use actual wall length instead of bounding box width if it's larger
                width = Math.Max(wallLength, width);

                // Smart padding (same as group elevations)
                double maxModelDim = Math.Max(Math.Max(width, depth), height);
                double padding = Math.Max(2.0, Math.Min(10.0, maxModelDim * 0.1)); // 10% of largest dimension

                System.Diagnostics.Debug.WriteLine($"=== WALL CROP DEBUG (GROUP STYLE) ===");
                System.Diagnostics.Debug.WriteLine($"Bounding box size: W={size.X:F1}, D={size.Y:F1}, H={size.Z:F1}");
                System.Diagnostics.Debug.WriteLine($"Wall length: {wallLength:F1}");
                System.Diagnostics.Debug.WriteLine($"Used dimensions: W={width:F1}, D={depth:F1}, H={height:F1}");
                System.Diagnostics.Debug.WriteLine($"Padding: {padding:F1}");

                // Create section box using group elevation approach
                // For internal wall elevation: looking at the wall from inside
                var sectionBox = new BoundingBoxXYZ();

                // Section box dimensions in local coordinates (like group elevations do it)
                // X = along wall (width), Y = into wall (depth), Z = up (height)
                sectionBox.Min = new XYZ(-width / 2 - padding, -depth / 2 - padding, -padding);
                sectionBox.Max = new XYZ(width / 2 + padding, depth / 2 + padding, height + padding);

                // Create transform like group elevations do (clean coordinate system)
                var wallDirection = (p1 - p0).Normalize();
                var wallNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();
                if (wall.Flipped) wallNormal = -wallNormal;

                // For internal elevation: viewer is inside looking at wall
                var transform = Transform.Identity;
                transform.BasisX = wallDirection;     // X = along wall
                transform.BasisY = XYZ.BasisZ;        // Y = up (Z becomes Y in view)  
                transform.BasisZ = wallNormal;        // Z = into wall (normal becomes depth)
                transform.Origin = center;            // Center on wall

                sectionBox.Transform = transform;

                System.Diagnostics.Debug.WriteLine($"Section box local: Min({sectionBox.Min.X:F1}, {sectionBox.Min.Y:F1}, {sectionBox.Min.Z:F1})");
                System.Diagnostics.Debug.WriteLine($"Section box local: Max({sectionBox.Max.X:F1}, {sectionBox.Max.Y:F1}, {sectionBox.Max.Z:F1})");
                System.Diagnostics.Debug.WriteLine($"Final height: {sectionBox.Max.Z - sectionBox.Min.Z:F1}");

                // Assign crop box to view
                elevation.CropBox = sectionBox;
                elevation.CropBoxActive = true;
                elevation.CropBoxVisible = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SetElevationCropRegion: {ex.Message}");
            }
        }

        private string GetWallTypeName(Wall wall)
        {
            try
            {
                var wallType = wall.Document.GetElement(wall.GetTypeId()) as WallType;
                if (wallType != null)
                {
                    var nameParam = wallType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    if (nameParam != null && nameParam.HasValue)
                    {
                        return nameParam.AsString().Replace(" ", "-");
                    }
                }
                return $"Wall-{wall.Id}";
            }
            catch
            {
                return $"Wall-{wall.Id}";
            }
        }

        /// <summary>
        /// Sets a unique view name, appending a suffix if the name already exists
        /// </summary>
        private void SetUniqueViewName(ViewSection view, string baseName)
        {
            if (view == null) return;

            string viewName = baseName;
            for (int i = 0; i < 100; i++) // Try up to 100 variations
            {
                try
                {
                    view.Name = viewName;
                    break;
                }
                catch
                {
                    // Name already exists, try with suffix
                    if (i == 0)
                    {
                        // First retry: add a short timestamp suffix
                        string suffix = DateTime.Now.ToString("mmss");
                        viewName = $"{baseName}_{suffix}";
                    }
                    else
                    {
                        // Subsequent retries: add incremental suffix
                        viewName = $"{baseName}_{i + 1}";
                    }
                }
            }
        }

        #endregion

        #region Template Selection

        /// <summary>
        /// Enhanced template selection with type information
        /// </summary>
        private View GetViewTemplate(Document doc)
        {
            var viewTemplates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate)
                .ToList();

            if (viewTemplates.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("No view templates found in document");
                return null;
            }

            // Group templates by type for better user understanding
            var sectionTemplates = viewTemplates.Where(v => v.ViewType == ViewType.Section).ToList();
            var elevationTemplates = viewTemplates.Where(v => v.ViewType == ViewType.Elevation).ToList();
            var threeDTemplates = viewTemplates.Where(v => v.ViewType == ViewType.ThreeD).ToList();

            System.Diagnostics.Debug.WriteLine($"Available templates: {elevationTemplates.Count} Elevation, {sectionTemplates.Count} Section, {threeDTemplates.Count} 3D");

            // Create display names with type information
            var templateNames = viewTemplates
                .Select(v => $"{v.Name} ({v.ViewType})")
                .OrderBy(name => name)
                .ToList();
            templateNames.Insert(0, "None");

            var templateWindow = new SelectFromDictWindow(templateNames,
                "Select ViewTemplate for Internal Building Elevations", allowMultiple: false);
            bool? result = templateWindow.ShowDialog();

            if (result == true && templateWindow.SelectedItems.Count > 0)
            {
                var selectedName = templateWindow.SelectedItems[0];
                if (selectedName != "None")
                {
                    // Extract the template name (remove the view type suffix)
                    var actualName = selectedName.Substring(0, selectedName.LastIndexOf(" ("));
                    var selectedTemplate = viewTemplates.FirstOrDefault(v => v.Name == actualName);

                    if (selectedTemplate != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"User selected template: '{selectedTemplate.Name}' (Type: {selectedTemplate.ViewType})");
                        return selectedTemplate;
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine("No template selected by user");
            return null;
        }

        #endregion

        #region Sheet Creation Methods

        /// <summary>
        /// Creates a sheet for an internal building elevation
        /// </summary>
        private ViewSheet CreateIndividualSheetForElevation(Document doc, ViewSection elevation, Wall wall, string baseName, List<string> results)
        {
            try
            {
                ElementId titleblockTypeId = GetTitleblockTypeId(doc);
                if (titleblockTypeId == null || titleblockTypeId == ElementId.InvalidElementId)
                {
                    results.Add($"Warning: No titleblock available for elevation {elevation.Id}");
                    return null;
                }

                var sheet = ViewSheet.Create(doc, titleblockTypeId);
                XYZ pos = new XYZ(0.5, 0.5, 0);

                try
                {
                    if (Viewport.CanAddViewToSheet(doc, sheet.Id, elevation.Id))
                    {
                        Viewport.Create(doc, sheet.Id, elevation.Id, pos);
                    }
                    else
                    {
                        results.Add($"Warning: Could not place elevation {elevation.Id} on sheet {sheet.Id}");
                    }
                }
                catch (Exception ex)
                {
                    results.Add($"Failed to place elevation {elevation.Id} on sheet: {ex.Message}");
                }

                // Create meaningful sheet number and name
                string wallTypeName = GetWallTypeName(wall);
                string elevationNumber = elevation.Name.Split('_').LastOrDefault() ?? (sequenceCounter - 1).ToString();
                string sheetNumber = $"DEAXO_IE_{baseName}_{elevationNumber}";
                string sheetName = $"Internal Building Elevation - {baseName}_{elevationNumber} ({wallTypeName}) - DEAXO GmbH";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);
                return sheet;
            }
            catch (Exception ex)
            {
                results.Add($"Failed to create sheet for elevation {elevation.Id}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Creates a combined sheet with all internal elevations arranged in a grid
        /// </summary>
        private ViewSheet CreateCombinedSheetForElevations(Document doc, List<(ViewSection elevation, Wall wall, bool templateApplied)> createdElevations, string baseName)
        {
            try
            {
                // Get titleblock type
                ElementId titleblockTypeId = GetTitleblockTypeId(doc);
                if (titleblockTypeId == null || titleblockTypeId == ElementId.InvalidElementId)
                {
                    System.Diagnostics.Debug.WriteLine("Warning: No titleblock available for combined sheet creation");
                    return null;
                }

                // Create sheet
                var sheet = ViewSheet.Create(doc, titleblockTypeId);
                if (sheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create combined sheet");
                    return null;
                }

                // Calculate grid layout for elevations
                int elevationCount = createdElevations.Count;
                int columns = Math.Min(4, (int)Math.Ceiling(Math.Sqrt(elevationCount))); // Max 4 columns
                int rows = (int)Math.Ceiling((double)elevationCount / columns);

                double marginX = 0.1; // 10% margin from edges
                double marginY = 0.1;
                double spacingX = 0.05; // 5% spacing between views
                double spacingY = 0.05;

                double availableWidth = 1.0 - (2 * marginX) - ((columns - 1) * spacingX);
                double availableHeight = 1.0 - (2 * marginY) - ((rows - 1) * spacingY);

                double cellWidth = availableWidth / columns;
                double cellHeight = availableHeight / rows;

                // Place elevations on combined sheet in grid pattern
                int placedViews = 0;
                for (int row = 0; row < rows && placedViews < elevationCount; row++)
                {
                    for (int col = 0; col < columns && placedViews < elevationCount; col++)
                    {
                        var elevationTuple = createdElevations[placedViews];
                        var elevation = elevationTuple.elevation;
                        var wall = elevationTuple.wall;
                        var templateApplied = elevationTuple.templateApplied;

                        try
                        {
                            if (Viewport.CanAddViewToSheet(doc, sheet.Id, elevation.Id))
                            {
                                // Calculate position in grid
                                XYZ position = new XYZ(
                                    marginX + col * (cellWidth + spacingX) + cellWidth / 2,
                                    marginY + (rows - 1 - row) * (cellHeight + spacingY) + cellHeight / 2, // Flip Y to start from top
                                    0);

                                var viewport = Viewport.Create(doc, sheet.Id, elevation.Id, position);
                                placedViews++;
                                System.Diagnostics.Debug.WriteLine($"Successfully placed elevation {elevation.Name} on combined sheet at grid position ({col}, {row})");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Cannot add elevation {elevation.Id} to combined sheet {sheet.Id}");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to place elevation {elevation.Id} on combined sheet: {ex.Message}");
                        }
                    }
                }

                // Set sheet number and name
                string sheetNumber = $"DEAXO_IE_Combined_{baseName}_{DateTime.Now.Ticks % 1000000}";
                string sheetName = $"Internal Building Elevations - Combined ({baseName}) - DEAXO GmbH";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);

                System.Diagnostics.Debug.WriteLine($"Created combined sheet with {placedViews} elevations placed in {columns}x{rows} grid");
                return sheet;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating combined sheet for elevations: {ex.Message}");
                return null;
            }
        }

        private ElementId GetTitleblockTypeId(Document doc)
        {
            var titleblockTypeId = doc.GetDefaultFamilyTypeId(new ElementId(BuiltInCategory.OST_TitleBlocks));
            if (titleblockTypeId != null && titleblockTypeId != ElementId.InvalidElementId)
                return titleblockTypeId;

            var tb = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .FirstOrDefault();

            return tb?.Id;
        }

        private void SetUniqueSheetName(ViewSheet sheet, string baseNumber, string baseName)
        {
            string sheetNumber = baseNumber;
            string sheetName = baseName;

            for (int i = 0; i < 10; i++)
            {
                try
                {
                    sheet.SheetNumber = sheetNumber;
                    sheet.Name = sheetName;
                    break;
                }
                catch
                {
                    sheetNumber = baseNumber + (i == 0 ? "*" : $"*{i}");
                    if (sheetName.Length > 200) // Revit sheet name limit
                    {
                        sheetName = baseName.Substring(0, 190) + $"...({i + 1})";
                    }
                    else
                    {
                        sheetName = baseName + (i == 0 ? " (2)" : $" ({i + 2})");
                    }
                }
            }
        }

        #endregion
    }
}