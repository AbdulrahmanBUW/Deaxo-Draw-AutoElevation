// Complete GroupElevationCommand.cs with Sheet Configuration Options
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
    public class GroupElevationCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                System.Diagnostics.Debug.WriteLine("=== GROUP ELEVATION COMMAND STARTED ===");

                // Element selection categories
                var selectOpts = new Dictionary<string, object>()
                {
                    {"Walls", BuiltInCategory.OST_Walls},
                    {"Windows", BuiltInCategory.OST_Windows},
                    {"Doors", BuiltInCategory.OST_Doors},
                    {"Columns", new BuiltInCategory[] { BuiltInCategory.OST_Columns, BuiltInCategory.OST_StructuralColumns }},
                    {"Beams/Framing", BuiltInCategory.OST_StructuralFraming},
                    {"Furniture", new BuiltInCategory[] { BuiltInCategory.OST_Furniture, BuiltInCategory.OST_FurnitureSystems }},
                    {"Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures},
                    {"Generic Models", BuiltInCategory.OST_GenericModel},
                    {"Casework", BuiltInCategory.OST_Casework},
                    {"Curtain Walls", BuiltInCategory.OST_Walls},
                    {"Lighting Fixtures", BuiltInCategory.OST_LightingFixtures},
                    {"Mass", BuiltInCategory.OST_Mass},
                    {"Parking", BuiltInCategory.OST_Parking},
                    {"MEP Elements", new BuiltInCategory[] {
                        BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeAccessory,
                        BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctAccessory,
                        BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_MechanicalEquipment
                    }},
                    {"Electrical", new BuiltInCategory[] {
                        BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_ElectricalEquipment
                    }},
                    {"Specialty", new BuiltInCategory[] {
                        BuiltInCategory.OST_Sprinklers, BuiltInCategory.OST_FireAlarmDevices,
                        BuiltInCategory.OST_CommunicationDevices, BuiltInCategory.OST_SecurityDevices
                    }},
                    {"All Loadable Families", typeof(FamilyInstance)}
                };

                // Process categories for selection filter
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

                // Element selection
                var selFilter = new DXSelectionFilter(allowedTypesOrCats);
                IList<Reference> refs;
                try
                {
                    refs = uidoc.Selection.PickObjects(ObjectType.Element, selFilter,
                        "Select elements for group elevations (6 orthographic views + 3D will be created)");
                }
                catch (OperationCanceledException)
                {
                    TaskDialog.Show("DEAXO", "Selection cancelled.");
                    return Result.Cancelled;
                }

                if (refs == null || refs.Count == 0)
                {
                    TaskDialog.Show("DEAXO", "No elements selected.");
                    return Result.Cancelled;
                }

                System.Diagnostics.Debug.WriteLine($"Selected {refs.Count} element references");

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

                // Show progress window
                var progressWindow = new ProgressWindow();
                progressWindow.Show();
                progressWindow.UpdateStatus("Preparing elements...", "Calculating bounding box and preparing view creation");

                var results = new List<string>();
                var startTime = DateTime.Now;

                // Update transaction name based on sheet option - C# 7.3 compatible
                string transactionName;
                if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                    transactionName = "DEAXO - Create Group Views";
                else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                    transactionName = "DEAXO - Create Group Views and Individual Sheets";
                else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                    transactionName = "DEAXO - Create Group Views and Combined Sheet";
                else
                    transactionName = "DEAXO - Create Group Views";

                using (Transaction t = new Transaction(doc, transactionName))
                {
                    t.Start();

                    try
                    {
                        var selectedElements = refs.Select(r => doc.GetElement(r)).ToList();
                        results.Add($"Selected {selectedElements.Count} elements for group elevation views");

                        System.Diagnostics.Debug.WriteLine($"Processing elements:");
                        foreach (var el in selectedElements)
                        {
                            System.Diagnostics.Debug.WriteLine($"  - Element {el.Id}: {el.GetType().Name}, Category: {el.Category?.Name ?? "None"}");
                        }

                        progressWindow.UpdateProgress(1, 10);
                        progressWindow.AddLogMessage($"Processing {selectedElements.Count} elements");

                        // Create group views using the enhanced approach
                        progressWindow.UpdateStatus("Creating views...", "Using enhanced group elevation generator");
                        progressWindow.UpdateProgress(3, 10);

                        System.Diagnostics.Debug.WriteLine("\n=== CALLING GROUP VIEW GENERATOR ===");

                        // Use the SimplifiedGroupElevationGenerator directly for better control
                        var groupViews = SimplifiedGroupElevationGenerator.CreateGroupViews(doc, selectedElements, chosenTemplate);

                        if (groupViews == null)
                        {
                            var errorMsg = "Failed to create group elevation views - generator returned null";
                            System.Diagnostics.Debug.WriteLine($"ERROR: {errorMsg}");
                            progressWindow.ShowError(errorMsg);
                            results.Add($"ERROR: {errorMsg}");
                            t.RollBack();
                            return Result.Failed;
                        }

                        progressWindow.UpdateProgress(6, 10);
                        progressWindow.AddLogMessage("View generation completed");

                        // Count and log created views - DETAILED ANALYSIS
                        int viewCount = 0;
                        int templateAppliedCount = 0;
                        var createdViews = new List<(string type, View view, bool templateApplied)>();

                        System.Diagnostics.Debug.WriteLine("\n=== ANALYZING CREATED VIEWS AND TEMPLATES ===");

                        // Check each view and its template status
                        if (groupViews.TopView != null)
                        {
                            viewCount++;
                            bool templateApplied = (chosenTemplate != null && groupViews.TopView.ViewTemplateId == chosenTemplate.Id);
                            if (templateApplied) templateAppliedCount++;
                            createdViews.Add(("Top", groupViews.TopView, templateApplied));
                            results.Add($"✓ Created Top view: {groupViews.TopView.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            System.Diagnostics.Debug.WriteLine($"✓ TOP VIEW: ID {groupViews.TopView.Id}, Name: {groupViews.TopView.Name}, Template: {templateApplied}");
                        }

                        if (groupViews.BottomView != null)
                        {
                            viewCount++;
                            bool templateApplied = (chosenTemplate != null && groupViews.BottomView.ViewTemplateId == chosenTemplate.Id);
                            if (templateApplied) templateAppliedCount++;
                            createdViews.Add(("Bottom", groupViews.BottomView, templateApplied));
                            results.Add($"✓ Created Bottom view: {groupViews.BottomView.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            System.Diagnostics.Debug.WriteLine($"✓ BOTTOM VIEW: ID {groupViews.BottomView.Id}, Name: {groupViews.BottomView.Name}, Template: {templateApplied}");
                        }

                        if (groupViews.LeftView != null)
                        {
                            viewCount++;
                            bool templateApplied = (chosenTemplate != null && groupViews.LeftView.ViewTemplateId == chosenTemplate.Id);
                            if (templateApplied) templateAppliedCount++;
                            createdViews.Add(("Left", groupViews.LeftView, templateApplied));
                            results.Add($"✓ Created Left view: {groupViews.LeftView.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            System.Diagnostics.Debug.WriteLine($"✓ LEFT VIEW: ID {groupViews.LeftView.Id}, Name: {groupViews.LeftView.Name}, Template: {templateApplied}");
                        }

                        if (groupViews.RightView != null)
                        {
                            viewCount++;
                            bool templateApplied = (chosenTemplate != null && groupViews.RightView.ViewTemplateId == chosenTemplate.Id);
                            if (templateApplied) templateAppliedCount++;
                            createdViews.Add(("Right", groupViews.RightView, templateApplied));
                            results.Add($"✓ Created Right view: {groupViews.RightView.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            System.Diagnostics.Debug.WriteLine($"✓ RIGHT VIEW: ID {groupViews.RightView.Id}, Name: {groupViews.RightView.Name}, Template: {templateApplied}");
                        }

                        if (groupViews.FrontView != null)
                        {
                            viewCount++;
                            bool templateApplied = (chosenTemplate != null && groupViews.FrontView.ViewTemplateId == chosenTemplate.Id);
                            if (templateApplied) templateAppliedCount++;
                            createdViews.Add(("Front", groupViews.FrontView, templateApplied));
                            results.Add($"✓ Created Front view: {groupViews.FrontView.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            System.Diagnostics.Debug.WriteLine($"✓ FRONT VIEW: ID {groupViews.FrontView.Id}, Name: {groupViews.FrontView.Name}, Template: {templateApplied}");
                        }

                        if (groupViews.BackView != null)
                        {
                            viewCount++;
                            bool templateApplied = (chosenTemplate != null && groupViews.BackView.ViewTemplateId == chosenTemplate.Id);
                            if (templateApplied) templateAppliedCount++;
                            createdViews.Add(("Back", groupViews.BackView, templateApplied));
                            results.Add($"✓ Created Back view: {groupViews.BackView.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            System.Diagnostics.Debug.WriteLine($"✓ BACK VIEW: ID {groupViews.BackView.Id}, Name: {groupViews.BackView.Name}, Template: {templateApplied}");
                        }

                        if (groupViews.ThreeDView != null)
                        {
                            viewCount++;
                            bool templateApplied = (chosenTemplate != null && groupViews.ThreeDView.ViewTemplateId == chosenTemplate.Id);
                            if (templateApplied) templateAppliedCount++;
                            createdViews.Add(("3D", groupViews.ThreeDView, templateApplied));
                            results.Add($"✓ Created 3D view: {groupViews.ThreeDView.Id} (Template: {(templateApplied ? "Applied" : "Not Applied")})");
                            System.Diagnostics.Debug.WriteLine($"✓ 3D VIEW: ID {groupViews.ThreeDView.Id}, Name: {groupViews.ThreeDView.Name}, Template: {templateApplied}");
                        }

                        progressWindow.UpdateProgress(8, 10);
                        progressWindow.AddLogMessage($"Successfully created {viewCount} views out of 7 expected");

                        if (chosenTemplate != null)
                        {
                            progressWindow.AddLogMessage($"Template '{chosenTemplate.Name}' applied to {templateAppliedCount} views");
                        }

                        System.Diagnostics.Debug.WriteLine($"\n=== VIEW CREATION SUMMARY ===");
                        System.Diagnostics.Debug.WriteLine($"TOTAL VIEWS CREATED: {viewCount}/7");
                        System.Diagnostics.Debug.WriteLine($"TEMPLATE APPLICATIONS: {templateAppliedCount}/{viewCount}");
                        System.Diagnostics.Debug.WriteLine($"SUCCESS RATE: {(viewCount / 7.0) * 100:F1}%");

                        if (viewCount == 0)
                        {
                            progressWindow.ShowError("No views were created");
                            results.Add("ERROR: Failed to create any views");
                            t.RollBack();
                            return Result.Failed;
                        }

                        // Handle sheet creation based on user choice
                        int sheetCount = 0;
                        if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                        {
                            progressWindow.AddLogMessage("Skipping sheet creation as requested");
                            results.Add("✓ Views created without sheets as requested");
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                        {
                            progressWindow.UpdateStatus("Creating individual sheets...", "Generating separate sheet for each view");
                            progressWindow.AddLogMessage("Starting individual sheet creation process");

                            foreach (var viewTuple in createdViews)
                            {
                                var type = viewTuple.type;
                                var view = viewTuple.view;
                                var templateApplied = viewTuple.templateApplied;

                                try
                                {
                                    var sheet = CreateIndividualSheetForView(doc, view, type, selectedElements);
                                    if (sheet != null)
                                    {
                                        sheetCount++;
                                        results.Add($"✓ Created individual sheet {sheet.Id} ({sheet.SheetNumber}) for {type} view {view.Id}");
                                        System.Diagnostics.Debug.WriteLine($"✓ INDIVIDUAL SHEET: Created {sheet.Id} for {type} view {view.Id}");
                                    }
                                    else
                                    {
                                        results.Add($"✗ Failed to create individual sheet for {type} view {view.Id}");
                                        System.Diagnostics.Debug.WriteLine($"✗ INDIVIDUAL SHEET: Failed for {type} view {view.Id}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    results.Add($"✗ Error creating individual sheet for {type} view: {ex.Message}");
                                    System.Diagnostics.Debug.WriteLine($"✗ INDIVIDUAL SHEET ERROR for {type}: {ex.Message}");
                                }
                            }
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                        {
                            progressWindow.UpdateStatus("Creating combined sheet...", "Placing all views on single sheet");
                            progressWindow.AddLogMessage("Starting combined sheet creation process");

                            try
                            {
                                var combinedSheet = CreateCombinedSheetForViews(doc, createdViews, selectedElements);
                                if (combinedSheet != null)
                                {
                                    sheetCount = 1;
                                    results.Add($"✓ Created combined sheet {combinedSheet.Id} ({combinedSheet.SheetNumber}) with {createdViews.Count} views");
                                    System.Diagnostics.Debug.WriteLine($"✓ COMBINED SHEET: Created {combinedSheet.Id} with {createdViews.Count} views");
                                }
                                else
                                {
                                    results.Add($"✗ Failed to create combined sheet");
                                    System.Diagnostics.Debug.WriteLine($"✗ COMBINED SHEET: Failed");
                                }
                            }
                            catch (Exception ex)
                            {
                                results.Add($"✗ Error creating combined sheet: {ex.Message}");
                                System.Diagnostics.Debug.WriteLine($"✗ COMBINED SHEET ERROR: {ex.Message}");
                            }
                        }

                        progressWindow.AddLogMessage($"Sheet creation completed: {sheetCount} sheet(s) created");

                        // Update final results based on sheet option - C# 7.3 compatible
                        if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                        {
                            results.Add($"✓ Total: {viewCount} views created (no sheets)");
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                        {
                            results.Add($"✓ Total: {viewCount} views and {sheetCount} individual sheets created");
                        }
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                        {
                            results.Add($"✓ Total: {viewCount} views on {sheetCount} combined sheet created");
                        }

                        if (chosenTemplate != null)
                        {
                            results.Add($"✓ Template '{chosenTemplate.Name}' ({chosenTemplate.ViewType}) applied to {templateAppliedCount} compatible views");
                        }

                        progressWindow.UpdateProgress(10, 10);
                        t.Commit();

                        // Show completion with appropriate title - C# 7.3 compatible
                        string completionTitle;
                        if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.None)
                            completionTitle = "Group Views";
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Individual)
                            completionTitle = "Group Views & Individual Sheets";
                        else if (sheetConfig.Value == SheetConfigurationWindow.SheetOption.Combined)
                            completionTitle = "Group Views & Combined Sheet";
                        else
                            completionTitle = "Group Views";

                        var duration = DateTime.Now - startTime;
                        progressWindow.ShowCompletion(results, completionTitle);

                        System.Diagnostics.Debug.WriteLine($"=== COMMAND COMPLETED SUCCESSFULLY ===");
                        System.Diagnostics.Debug.WriteLine($"Total duration: {duration.TotalSeconds:F1} seconds");
                        System.Diagnostics.Debug.WriteLine($"Final result count: {results.Count} messages");
                        System.Diagnostics.Debug.WriteLine($"Created {viewCount} views and {sheetCount} sheets");
                        System.Diagnostics.Debug.WriteLine($"Template applied to {templateAppliedCount} views");
                        System.Diagnostics.Debug.WriteLine($"Sheet option: {sheetConfig}");

                        // Wait a moment then close
                        System.Threading.Thread.Sleep(1500);
                        progressWindow.Close();
                    }
                    catch (Exception ex)
                    {
                        var errorMsg = $"Exception in group elevation command: {ex.Message}";
                        System.Diagnostics.Debug.WriteLine($"FATAL ERROR: {errorMsg}");
                        System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");

                        progressWindow.ShowError(errorMsg);
                        results.Add($"ERROR: {errorMsg}");
                        t.RollBack();
                        throw;
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                var errorMsg = $"Fatal error in GroupElevationCommand: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"FATAL: {errorMsg}");
                message = errorMsg;
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

        #region Sheet Creation Methods

        /// <summary>
        /// Creates an individual sheet for a single view
        /// </summary>
        private ViewSheet CreateIndividualSheetForView(Document doc, View view, string viewType, List<Element> originalElements)
        {
            try
            {
                // Get titleblock type
                ElementId titleblockTypeId = GetTitleblockTypeId(doc);
                if (titleblockTypeId == null || titleblockTypeId == ElementId.InvalidElementId)
                {
                    System.Diagnostics.Debug.WriteLine("Warning: No titleblock available for sheet creation");
                    return null;
                }

                // Create sheet
                var sheet = ViewSheet.Create(doc, titleblockTypeId);
                if (sheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create individual sheet");
                    return null;
                }

                // Place view on sheet (centered)
                try
                {
                    XYZ viewportPosition = new XYZ(0.5, 0.5, 0); // Center position

                    if (Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                    {
                        var viewport = Viewport.Create(doc, sheet.Id, view.Id, viewportPosition);
                        System.Diagnostics.Debug.WriteLine($"Successfully placed view {view.Id} on individual sheet {sheet.Id}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Cannot add view {view.Id} to individual sheet {sheet.Id}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to place view {view.Id} on individual sheet: {ex.Message}");
                }

                // Set sheet number and name
                string elementInfo = GetElementsDescription(originalElements);
                string sheetNumber = $"DEAXO_GE_{viewType}_{DateTime.Now.Ticks % 1000000}";
                string sheetName = $"Group Elevation - {viewType} View ({elementInfo}) - DEAXO GmbH";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);
                return sheet;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating individual sheet for {viewType} view: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Creates a combined sheet with all views arranged in a grid
        /// </summary>
        private ViewSheet CreateCombinedSheetForViews(Document doc, List<(string type, View view, bool templateApplied)> createdViews, List<Element> originalElements)
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

                // Calculate grid layout for views
                // Arrange views in a grid pattern: 3 columns x 3 rows (max 9 views, which covers our 7 views)
                int columns = 3;
                int rows = 3;
                double marginX = 0.1; // 10% margin from edges
                double marginY = 0.1;
                double spacingX = 0.05; // 5% spacing between views
                double spacingY = 0.05;

                double availableWidth = 1.0 - (2 * marginX) - ((columns - 1) * spacingX);
                double availableHeight = 1.0 - (2 * marginY) - ((rows - 1) * spacingY);

                double cellWidth = availableWidth / columns;
                double cellHeight = availableHeight / rows;

                // Predefined positions for common view arrangements
                var viewPositions = new Dictionary<string, XYZ>
                {
                    // Arrange orthographic views in logical positions
                    {"Top", new XYZ(marginX + cellWidth / 2, marginY + 2 * cellHeight + 2 * spacingY + cellHeight / 2, 0)}, // Top row, center
                    {"Front", new XYZ(marginX + cellWidth / 2, marginY + cellHeight + spacingY + cellHeight / 2, 0)}, // Middle row, left
                    {"Right", new XYZ(marginX + cellWidth + spacingX + cellWidth / 2, marginY + cellHeight + spacingY + cellHeight / 2, 0)}, // Middle row, center
                    {"Back", new XYZ(marginX + 2 * cellWidth + 2 * spacingX + cellWidth / 2, marginY + cellHeight + spacingY + cellHeight / 2, 0)}, // Middle row, right
                    {"Left", new XYZ(marginX + cellWidth / 2, marginY + cellHeight / 2, 0)}, // Bottom row, left
                    {"Bottom", new XYZ(marginX + cellWidth + spacingX + cellWidth / 2, marginY + cellHeight / 2, 0)}, // Bottom row, center
                    {"3D", new XYZ(marginX + 2 * cellWidth + 2 * spacingX + cellWidth / 2, marginY + cellHeight / 2, 0)} // Bottom row, right
                };

                // Place views on combined sheet
                int placedViews = 0;
                foreach (var viewTuple in createdViews)
                {
                    var type = viewTuple.type;
                    var view = viewTuple.view;
                    var templateApplied = viewTuple.templateApplied;

                    try
                    {
                        if (Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                        {
                            XYZ position = viewPositions.ContainsKey(type)
                                ? viewPositions[type]
                                : new XYZ(0.2 + (placedViews % 3) * 0.3, 0.2 + (placedViews / 3) * 0.3, 0); // Fallback grid

                            var viewport = Viewport.Create(doc, sheet.Id, view.Id, position);
                            placedViews++;
                            System.Diagnostics.Debug.WriteLine($"Successfully placed {type} view {view.Id} on combined sheet at {position}");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"Cannot add {type} view {view.Id} to combined sheet {sheet.Id}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to place {type} view {view.Id} on combined sheet: {ex.Message}");
                    }
                }

                // Set sheet number and name
                string elementInfo = GetElementsDescription(originalElements);
                string sheetNumber = $"DEAXO_GE_Combined_{DateTime.Now.Ticks % 1000000}";
                string sheetName = $"Group Elevations - Combined View ({elementInfo}) - DEAXO GmbH";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);

                System.Diagnostics.Debug.WriteLine($"Created combined sheet with {placedViews} views placed");
                return sheet;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating combined sheet: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets a description of the selected elements for sheet naming
        /// </summary>
        private string GetElementsDescription(List<Element> elements)
        {
            if (elements == null || elements.Count == 0)
                return "Elements";

            // Group by category
            var categoryGroups = elements
                .Where(e => e.Category != null)
                .GroupBy(e => e.Category.Name)
                .OrderByDescending(g => g.Count())
                .Take(3) // Top 3 categories
                .ToList();

            if (categoryGroups.Count == 0)
                return $"{elements.Count} Elements";

            var descriptions = new List<string>();
            foreach (var group in categoryGroups)
            {
                descriptions.Add($"{group.Count()} {group.Key}");
            }

            string result = string.Join(", ", descriptions);

            // Keep it reasonably short for sheet name
            if (result.Length > 50)
                result = result.Substring(0, 47) + "...";

            return result;
        }

        /// <summary>
        /// Gets the titleblock type ID for sheet creation
        /// </summary>
        private ElementId GetTitleblockTypeId(Document doc)
        {
            // Try to get default titleblock type
            var titleblockTypeId = doc.GetDefaultFamilyTypeId(new ElementId(BuiltInCategory.OST_TitleBlocks));
            if (titleblockTypeId != null && titleblockTypeId != ElementId.InvalidElementId)
                return titleblockTypeId;

            // Fallback: find any titleblock type
            var titleblockType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .FirstOrDefault();

            return titleblockType?.Id;
        }

        /// <summary>
        /// Sets unique sheet number and name
        /// </summary>
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
                    // Try with suffix if name/number already exists
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
            var threeDTemplates = viewTemplates.Where(v => v.ViewType == ViewType.ThreeD).ToList();
            var elevationTemplates = viewTemplates.Where(v => v.ViewType == ViewType.Elevation).ToList();

            System.Diagnostics.Debug.WriteLine($"Available templates: {sectionTemplates.Count} Section, {threeDTemplates.Count} 3D, {elevationTemplates.Count} Elevation");

            // Create display names with type information
            var templateNames = viewTemplates
                .Select(v => $"{v.Name} ({v.ViewType})")
                .OrderBy(name => name)
                .ToList();
            templateNames.Insert(0, "None");

            var templateWindow = new SelectFromDictWindow(templateNames,
                "Select ViewTemplate for Group Views", allowMultiple: false);
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
    }
}