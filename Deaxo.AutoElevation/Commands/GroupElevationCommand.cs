// Updated GroupElevationCommand.cs
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
                // EXPANDED category choices - matching Single Element mode
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
                    {"Pipes", BuiltInCategory.OST_PipeCurves},
                    {"Pipe Fittings", BuiltInCategory.OST_PipeFitting},
                    {"Pipe Accessories", BuiltInCategory.OST_PipeAccessory},
                    {"Ducts", BuiltInCategory.OST_DuctCurves},
                    {"Duct Fittings", BuiltInCategory.OST_DuctFitting},
                    {"Duct Accessories", BuiltInCategory.OST_DuctAccessory},
                    {"Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures},
                    {"Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment},
                    {"Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment},
                    {"Air Terminals", BuiltInCategory.OST_DuctTerminal},
                    {"Sprinklers", BuiltInCategory.OST_Sprinklers},
                    {"Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices},
                    {"Communication Devices", BuiltInCategory.OST_CommunicationDevices},
                    {"Data Devices", BuiltInCategory.OST_DataDevices},
                    {"Nurse Call Devices", BuiltInCategory.OST_NurseCallDevices},
                    {"Security Devices", BuiltInCategory.OST_SecurityDevices},
                    {"Telephone Devices", BuiltInCategory.OST_TelephoneDevices},
                    {"All Loadable Families", typeof(FamilyInstance)}
                };

                // Process categories
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

                // Selection
                var selFilter = new DXSelectionFilter(allowedTypesOrCats);
                IList<Reference> refs;
                try
                {
                    refs = uidoc.Selection.PickObjects(ObjectType.Element, selFilter,
                        "Select elements for assembly-style orthographic views (Top, Bottom, Left, Right, Front, Back + 3D)");
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

                // Get view template
                View chosenTemplate = GetViewTemplate(doc);

                // Show progress window
                var progressWindow = new ProgressWindow();
                progressWindow.Show();
                progressWindow.UpdateStatus("Analyzing geometry...", "Determining optimal coordinate system");

                var results = new List<string>();
                var startTime = DateTime.Now;

                using (Transaction t = new Transaction(doc, "DEAXO - Create Assembly-Style Group Elevations"))
                {
                    t.Start();

                    try
                    {
                        var selectedElements = refs.Select(r => doc.GetElement(r)).ToList();
                        results.Add($"Selected {selectedElements.Count} elements for assembly-style elevation views");

                        progressWindow.UpdateProgress(1, 10);
                        progressWindow.AddLogMessage($"Processing {selectedElements.Count} elements");

                        // Create assembly-style views using the new generator
                        progressWindow.UpdateStatus("Creating orthographic views...", "Generating Top, Bottom, Left, Right, Front, Back + 3D views");
                        progressWindow.UpdateProgress(3, 10);

                        var assemblyViews = AssemblyStyleElevationGenerator.CreateAssemblyStyleViews(doc, selectedElements, chosenTemplate);

                        if (assemblyViews == null)
                        {
                            progressWindow.ShowError("Failed to analyze element geometry for view creation");
                            results.Add("ERROR: Could not analyze selected elements for view creation");
                            t.RollBack();
                            return Result.Failed;
                        }

                        progressWindow.UpdateProgress(6, 10);
                        progressWindow.AddLogMessage("Successfully analyzed geometry and created coordinate system");

                        // Count created views
                        int viewCount = 0;
                        if (assemblyViews.TopView != null) { viewCount++; results.Add($"Created Top view: {assemblyViews.TopView.Id}"); }
                        if (assemblyViews.BottomView != null) { viewCount++; results.Add($"Created Bottom view: {assemblyViews.BottomView.Id}"); }
                        if (assemblyViews.LeftView != null) { viewCount++; results.Add($"Created Left view: {assemblyViews.LeftView.Id}"); }
                        if (assemblyViews.RightView != null) { viewCount++; results.Add($"Created Right view: {assemblyViews.RightView.Id}"); }
                        if (assemblyViews.FrontView != null) { viewCount++; results.Add($"Created Front view: {assemblyViews.FrontView.Id}"); }
                        if (assemblyViews.BackView != null) { viewCount++; results.Add($"Created Back view: {assemblyViews.BackView.Id}"); }
                        if (assemblyViews.ThreeDView != null) { viewCount++; results.Add($"Created 3D Orthographic view: {assemblyViews.ThreeDView.Id}"); }

                        progressWindow.UpdateProgress(8, 10);
                        progressWindow.AddLogMessage($"Successfully created {viewCount} orthographic views");

                        if (viewCount == 0)
                        {
                            progressWindow.ShowError("No views were created");
                            results.Add("ERROR: Failed to create any orthographic views");
                            t.RollBack();
                            return Result.Failed;
                        }

                        // Create sheets for all section views
                        progressWindow.UpdateStatus("Creating sheets...", "Placing views on sheets");
                        ElementId titleblockTypeId = GetTitleblockTypeId(doc);

                        if (titleblockTypeId != null && titleblockTypeId != ElementId.InvalidElementId)
                        {
                            int sheetCounter = 1;

                            // Create sheets for all section views
                            var allSectionViews = assemblyViews.AllSectionViews;
                            foreach (var sectionView in allSectionViews)
                            {
                                if (sectionView != null)
                                {
                                    string viewType = DetermineViewType(sectionView.Name);
                                    CreateSheetForView(doc, viewType, sectionView, titleblockTypeId, sheetCounter, results);
                                    sheetCounter++;
                                }
                            }

                            // Create sheet for 3D view if it exists
                            if (assemblyViews.ThreeDView != null)
                            {
                                CreateSheetFor3DView(doc, assemblyViews.ThreeDView, titleblockTypeId, sheetCounter, results);
                            }

                            progressWindow.AddLogMessage($"Created {sheetCounter - 1} sheets with placed views");
                        }
                        else
                        {
                            results.Add("Warning: No titleblock found, views created without sheets");
                            progressWindow.AddLogMessage("Warning: Views created without sheets (no titleblock available)");
                        }

                        progressWindow.UpdateProgress(10, 10);
                        t.Commit();

                        // Show completion
                        var duration = DateTime.Now - startTime;
                        progressWindow.ShowCompletion(results, "Assembly-Style Group Elevations");

                        // Wait a moment then show detailed results
                        System.Threading.Thread.Sleep(1500);
                        progressWindow.Close();

                        // Show detailed results window
                        var resultsWindow = new ResultsWindow(results, "Assembly-Style Group Elevations", duration);
                        resultsWindow.ShowDialog();
                    }
                    catch (Exception ex)
                    {
                        progressWindow.ShowError(ex.Message);
                        results.Add($"ERROR: {ex.Message}");
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

            var templateWindow = new SelectFromDictWindow(templateNames,
                "Select ViewTemplate for Assembly-Style Group Elevations", allowMultiple: false);
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

        private void CreateSheetForView(Document doc, string viewType, ViewSection sectionView,
            ElementId titleblockTypeId, int counter, List<string> results)
        {
            try
            {
                var sheet = ViewSheet.Create(doc, titleblockTypeId);
                XYZ pos = new XYZ(0.5, 0.5, 0);

                if (Viewport.CanAddViewToSheet(doc, sheet.Id, sectionView.Id))
                    Viewport.Create(doc, sheet.Id, sectionView.Id, pos);

                string sheetNumber = $"DEAXO_ASM_{viewType}_{counter}_{DateTime.Now:HHmmss}";
                string sheetName = $"Assembly-Style {viewType} Elevation (DEAXO GmbH)";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);
                results.Add($"{viewType} Elevation -> Sheet:{sheet.Id} View:{sectionView.Id}");
            }
            catch (Exception ex)
            {
                results.Add($"Failed to create sheet for {viewType} elevation: {ex.Message}");
            }
        }

        private void CreateSheetFor3DView(Document doc, View3D view3D, ElementId titleblockTypeId, int counter, List<string> results)
        {
            try
            {
                var sheet = ViewSheet.Create(doc, titleblockTypeId);
                XYZ pos = new XYZ(0.5, 0.5, 0);

                if (Viewport.CanAddViewToSheet(doc, sheet.Id, view3D.Id))
                    Viewport.Create(doc, sheet.Id, view3D.Id, pos);

                string sheetNumber = $"DEAXO_ASM_3D_{counter}_{DateTime.Now:HHmmss}";
                string sheetName = $"Assembly-Style 3D Orthographic (DEAXO GmbH)";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);
                results.Add($"3D Orthographic -> Sheet:{sheet.Id} View:{view3D.Id}");
            }
            catch (Exception ex)
            {
                results.Add($"Failed to create sheet for 3D view: {ex.Message}");
            }
        }

        private string DetermineViewType(string viewName)
        {
            if (viewName.Contains("Top")) return "Top";
            if (viewName.Contains("Bottom")) return "Bottom";
            if (viewName.Contains("Left")) return "Left";
            if (viewName.Contains("Right")) return "Right";
            if (viewName.Contains("Front")) return "Front";
            if (viewName.Contains("Back")) return "Back";
            return "Unknown";
        }

        private void SetUniqueSheetName(ViewSheet sheet, string baseNumber, string baseName)
        {
            string sheetNumber = baseNumber;
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    sheet.SheetNumber = sheetNumber;
                    sheet.Name = baseName;
                    break;
                }
                catch
                {
                    sheetNumber += "*";
                }
            }
        }
    }
}