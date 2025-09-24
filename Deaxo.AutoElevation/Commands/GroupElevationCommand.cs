// Enhanced GroupElevationCommand.cs with Debug Output - UPDATED: No sheet creation
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

                // Get view template
                View chosenTemplate = GetViewTemplate(doc);
                if (chosenTemplate != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Using view template: {chosenTemplate.Name} (ID: {chosenTemplate.Id})");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No view template selected");
                }

                // Show progress window
                var progressWindow = new ProgressWindow();
                progressWindow.Show();
                progressWindow.UpdateStatus("Preparing elements...", "Calculating bounding box and preparing view creation");

                var results = new List<string>();
                var startTime = DateTime.Now;

                using (Transaction t = new Transaction(doc, "DEAXO - Create Group Views"))
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
                        var createdViews = new List<(string type, View view)>();

                        System.Diagnostics.Debug.WriteLine("\n=== ANALYZING CREATED VIEWS ===");

                        if (groupViews.TopView != null)
                        {
                            viewCount++;
                            createdViews.Add(("Top", groupViews.TopView));
                            results.Add($"✓ Created Top view: {groupViews.TopView.Id}");
                            System.Diagnostics.Debug.WriteLine($"✓ TOP VIEW: ID {groupViews.TopView.Id}, Name: {groupViews.TopView.Name}");
                        }
                        else
                        {
                            results.Add($"✗ Failed to create Top view");
                            System.Diagnostics.Debug.WriteLine($"✗ TOP VIEW: FAILED");
                        }

                        if (groupViews.BottomView != null)
                        {
                            viewCount++;
                            createdViews.Add(("Bottom", groupViews.BottomView));
                            results.Add($"✓ Created Bottom view: {groupViews.BottomView.Id}");
                            System.Diagnostics.Debug.WriteLine($"✓ BOTTOM VIEW: ID {groupViews.BottomView.Id}, Name: {groupViews.BottomView.Name}");
                        }
                        else
                        {
                            results.Add($"✗ Failed to create Bottom view");
                            System.Diagnostics.Debug.WriteLine($"✗ BOTTOM VIEW: FAILED");
                        }

                        if (groupViews.LeftView != null)
                        {
                            viewCount++;
                            createdViews.Add(("Left", groupViews.LeftView));
                            results.Add($"✓ Created Left view: {groupViews.LeftView.Id}");
                            System.Diagnostics.Debug.WriteLine($"✓ LEFT VIEW: ID {groupViews.LeftView.Id}, Name: {groupViews.LeftView.Name}");
                        }
                        else
                        {
                            results.Add($"✗ Failed to create Left view");
                            System.Diagnostics.Debug.WriteLine($"✗ LEFT VIEW: FAILED");
                        }

                        if (groupViews.RightView != null)
                        {
                            viewCount++;
                            createdViews.Add(("Right", groupViews.RightView));
                            results.Add($"✓ Created Right view: {groupViews.RightView.Id}");
                            System.Diagnostics.Debug.WriteLine($"✓ RIGHT VIEW: ID {groupViews.RightView.Id}, Name: {groupViews.RightView.Name}");
                        }
                        else
                        {
                            results.Add($"✗ Failed to create Right view");
                            System.Diagnostics.Debug.WriteLine($"✗ RIGHT VIEW: FAILED");
                        }

                        if (groupViews.FrontView != null)
                        {
                            viewCount++;
                            createdViews.Add(("Front", groupViews.FrontView));
                            results.Add($"✓ Created Front view: {groupViews.FrontView.Id}");
                            System.Diagnostics.Debug.WriteLine($"✓ FRONT VIEW: ID {groupViews.FrontView.Id}, Name: {groupViews.FrontView.Name}");
                        }
                        else
                        {
                            results.Add($"✗ Failed to create Front view");
                            System.Diagnostics.Debug.WriteLine($"✗ FRONT VIEW: FAILED");
                        }

                        if (groupViews.BackView != null)
                        {
                            viewCount++;
                            createdViews.Add(("Back", groupViews.BackView));
                            results.Add($"✓ Created Back view: {groupViews.BackView.Id}");
                            System.Diagnostics.Debug.WriteLine($"✓ BACK VIEW: ID {groupViews.BackView.Id}, Name: {groupViews.BackView.Name}");
                        }
                        else
                        {
                            results.Add($"✗ Failed to create Back view");
                            System.Diagnostics.Debug.WriteLine($"✗ BACK VIEW: FAILED");
                        }

                        if (groupViews.ThreeDView != null)
                        {
                            viewCount++;
                            createdViews.Add(("3D", groupViews.ThreeDView));
                            results.Add($"✓ Created 3D view: {groupViews.ThreeDView.Id}");
                            System.Diagnostics.Debug.WriteLine($"✓ 3D VIEW: ID {groupViews.ThreeDView.Id}, Name: {groupViews.ThreeDView.Name}");
                        }
                        else
                        {
                            results.Add($"✗ Failed to create 3D view");
                            System.Diagnostics.Debug.WriteLine($"✗ 3D VIEW: FAILED");
                        }

                        progressWindow.UpdateProgress(8, 10);
                        progressWindow.AddLogMessage($"Successfully created {viewCount} views out of 7 expected");

                        System.Diagnostics.Debug.WriteLine($"\n=== VIEW CREATION SUMMARY ===");
                        System.Diagnostics.Debug.WriteLine($"TOTAL VIEWS CREATED: {viewCount}/7");
                        System.Diagnostics.Debug.WriteLine($"SUCCESS RATE: {(viewCount / 7.0) * 100:F1}%");

                        if (viewCount == 0)
                        {
                            progressWindow.ShowError("No views were created");
                            results.Add("ERROR: Failed to create any views");
                            t.RollBack();
                            return Result.Failed;
                        }

                        // NOTE: No longer creating sheets - only views as requested
                        progressWindow.AddLogMessage("Views created successfully (no sheets generated as requested)");
                        results.Add($"✓ Total: {viewCount} views created without sheets");

                        progressWindow.UpdateProgress(10, 10);
                        t.Commit();

                        // Show completion
                        var duration = DateTime.Now - startTime;
                        progressWindow.ShowCompletion(results, "Group Views");

                        System.Diagnostics.Debug.WriteLine($"=== COMMAND COMPLETED SUCCESSFULLY ===");
                        System.Diagnostics.Debug.WriteLine($"Total duration: {duration.TotalSeconds:F1} seconds");
                        System.Diagnostics.Debug.WriteLine($"Final result count: {results.Count} messages");

                        // Wait a moment then close - Results window display removed as requested
                        System.Threading.Thread.Sleep(1500);
                        progressWindow.Close();

                        // Results window display removed as requested
                        // var resultsWindow = new ResultsWindow(results, "Group Views", duration);
                        // resultsWindow.ShowDialog();
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

        // ... (keep existing helper methods but remove sheet creation methods)

        private View GetViewTemplate(Document doc)
        {
            var viewTemplates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate)
                .ToList();

            if (viewTemplates.Count == 0) return null;

            var templateNames = viewTemplates.Select(v => v.Name).ToList();
            templateNames.Insert(0, "None");

            // Updated: Single selection only with search functionality
            var templateWindow = new SelectFromDictWindow(templateNames,
                "Select ViewTemplate for Group Views", allowMultiple: false);
            bool? result = templateWindow.ShowDialog();

            if (result == true && templateWindow.SelectedItems.Count > 0)
            {
                var name = templateWindow.SelectedItems[0];
                if (name != "None")
                {
                    return viewTemplates.FirstOrDefault(v => v.Name == name);
                }
            }
            return null;
        }
    }
}