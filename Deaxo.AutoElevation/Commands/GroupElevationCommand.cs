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
                // Expanded category choices for group elevations
                var selectOpts = new Dictionary<string, object>()
                {
                    {"Walls", BuiltInCategory.OST_Walls},
                    {"Windows", BuiltInCategory.OST_Windows},
                    {"Doors", BuiltInCategory.OST_Doors},
                    {"Columns", new BuiltInCategory[] { BuiltInCategory.OST_Columns, BuiltInCategory.OST_StructuralColumns }},
                    {"Beams/Framing", BuiltInCategory.OST_StructuralFraming},
                    {"Pipes", BuiltInCategory.OST_PipeCurves},
                    {"Pipe Fittings", BuiltInCategory.OST_PipeFitting},
                    {"Pipe Accessories", BuiltInCategory.OST_PipeAccessory},
                    {"Ducts", BuiltInCategory.OST_DuctCurves},
                    {"Duct Fittings", BuiltInCategory.OST_DuctFitting},
                    {"Duct Accessories", BuiltInCategory.OST_DuctAccessory},
                    {"Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures},
                    {"Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment},
                    {"Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures},
                    {"Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment},
                    {"Generic Models", BuiltInCategory.OST_GenericModel},
                    {"Furniture", new BuiltInCategory[] { BuiltInCategory.OST_Furniture, BuiltInCategory.OST_FurnitureSystems }},
                    {"Lighting Fixtures", BuiltInCategory.OST_LightingFixtures},
                    {"Casework", BuiltInCategory.OST_Casework},
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
                        "Select elements for group elevation (4 views will be created around the group)");
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

                // Create group elevations
                var results = new List<string>();
                using (Transaction t = new Transaction(doc, "DEAXO - Create Group Elevations"))
                {
                    t.Start();

                    var selectedElements = refs.Select(r => doc.GetElement(r)).ToList();
                    results.Add($"Selected {selectedElements.Count} elements for group elevation");

                    // Calculate bounding box
                    BoundingBoxXYZ overallBB = ScopeBoxHelper.CalculateOverallBoundingBox(selectedElements);
                    if (overallBB == null)
                    {
                        TaskDialog.Show("DEAXO", "Could not calculate bounding box for selected elements.");
                        t.RollBack();
                        return Result.Failed;
                    }

                    results.Add($"Bounding box: Min({overallBB.Min.X:F1}, {overallBB.Min.Y:F1}, {overallBB.Min.Z:F1}) Max({overallBB.Max.X:F1}, {overallBB.Max.Y:F1}, {overallBB.Max.Z:F1})");

                    // Optional scope box creation (for reference, then deleted)
                    Element scopeBox = ScopeBoxHelper.CreateScopeBoxFromElements(doc, selectedElements);
                    if (scopeBox != null)
                        results.Add($"Created temporary scope box: {scopeBox.Id}");

                    try
                    {
                        // Create four elevations
                        var elevations = SectionGenerator.CreateGroupElevations(doc, overallBB, chosenTemplate);

                        if (elevations != null && elevations.Count > 0)
                        {
                            // Create sheets for elevations
                            ElementId titleblockTypeId = GetTitleblockTypeId(doc);
                            if (titleblockTypeId != null)
                            {
                                int sheetCounter = 1;
                                foreach (var kvp in elevations)
                                {
                                    CreateSheetForElevation(doc, kvp.Key, kvp.Value, titleblockTypeId,
                                        sheetCounter, results);
                                    sheetCounter++;
                                }
                            }
                            else
                            {
                                results.Add("Warning: No titleblock found, elevations created without sheets");
                                foreach (var kvp in elevations)
                                {
                                    results.Add($"{kvp.Key} View created: {kvp.Value.Id}");
                                }
                            }
                        }
                        else
                        {
                            results.Add("Failed to create group elevations");
                        }
                    }
                    finally
                    {
                        // Clean up scope box
                        ScopeBoxHelper.DeleteScopeBoxSafely(doc, scopeBox);
                    }

                    t.Commit();
                }

                // Show results
                var msgText = string.Join(Environment.NewLine, results.Take(50));
                TaskDialog.Show("DEAXO - Group Elevations Created",
                    string.IsNullOrWhiteSpace(msgText) ? "No new elevations created." : msgText);

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
                "Select ViewTemplate for Group Elevations", allowMultiple: false);
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

        private void CreateSheetForElevation(Document doc, string direction, ViewSection elevation,
            ElementId titleblockTypeId, int counter, List<string> results)
        {
            try
            {
                var sheet = ViewSheet.Create(doc, titleblockTypeId);
                XYZ pos = new XYZ(0.5, 0.5, 0);

                if (Viewport.CanAddViewToSheet(doc, sheet.Id, elevation.Id))
                    Viewport.Create(doc, sheet.Id, elevation.Id, pos);

                string sheetNumber = $"DEAXO_GE_{counter}_{DateTime.Now:HHmmss}";
                string sheetName = $"Group Elevation - {direction} View (DEAXO GmbH)";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);
                results.Add($"{direction} View -> Sheet:{sheet.Id} Elevation:{elevation.Id}");
            }
            catch (Exception ex)
            {
                results.Add($"Failed to create sheet for {direction} elevation: {ex.Message}");
            }
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