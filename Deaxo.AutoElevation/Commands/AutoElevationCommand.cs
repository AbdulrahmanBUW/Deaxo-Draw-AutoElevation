using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Deaxo.AutoElevation.UI;

namespace Deaxo.AutoElevation.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class AutoElevationCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Show mode selection dialog first
                var modeWindow = new ElevationModeSelectionWindow();
                bool? result = modeWindow.ShowDialog();

                if (result != true || modeWindow.SelectedMode == ElevationModeSelectionWindow.ElevationMode.None)
                {
                    return Result.Cancelled;
                }

                // Execute the appropriate elevation logic based on selection
                switch (modeWindow.SelectedMode)
                {
                    case ElevationModeSelectionWindow.ElevationMode.SingleElement:
                        return ExecuteSingleElementElevation(commandData, ref message, elements);
                    case ElevationModeSelectionWindow.ElevationMode.GroupElement:
                        return ExecuteGroupElementElevation(commandData, ref message, elements);
                    case ElevationModeSelectionWindow.ElevationMode.Internal:
                        return ExecuteInternalElevation(commandData, ref message, elements);
                    default:
                        return Result.Cancelled;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Result ExecuteSingleElementElevation(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                var selectOpts = GetStandardCategories();
                var allowedTypesOrCats = GetAllowedTypes(selectOpts);

                // Selection
                var selFilter = new DXSelectionFilter(allowedTypesOrCats);
                IList<Reference> refs = SelectElements(uidoc, selFilter, "Select elements for individual elevations and click Finish");

                if (refs == null || refs.Count == 0)
                    return Result.Cancelled;

                // Get view template
                View chosenTemplate = GetViewTemplate(doc, "Select ViewTemplate for Single Element Elevations");

                // Create elevations
                var results = new List<string>();
                using (Transaction t = new Transaction(doc, "DEAXO - Create Single Element Elevations"))
                {
                    t.Start();

                    foreach (var r in refs)
                    {
                        try
                        {
                            Element el = doc.GetElement(r);
                            var props = new ElementProperties(doc, el);
                            if (!props.IsValid) continue;

                            var created = SectionGenerator.CreateElevationOnly(doc, props);
                            if (created?.elevation == null) continue;

                            ApplyTemplateAndCreateSheet(doc, created.elevation, chosenTemplate, props, results, "SE");
                        }
                        catch (Exception exInner)
                        {
                            results.Add($"Error processing element {r.ElementId}: {exInner.Message}");
                        }
                    }

                    t.Commit();
                }

                ShowResults(results, "Single Element Elevations");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Result ExecuteGroupElementElevation(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var groupCommand = new GroupElevationCommand();
            return groupCommand.Execute(commandData, ref message, elements);
        }

        private Result ExecuteInternalElevation(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var internalCommand = new InternalElevationCommand();
            return internalCommand.Execute(commandData, ref message, elements);
        }

        // Helper methods
        private Dictionary<string, object> GetStandardCategories()
        {
            return new Dictionary<string, object>()
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
                {"Lighting Fixtures", BuiltInCategory.OST_LightingFixtures},
                {"All Loadable Families", typeof(FamilyInstance)},
                {"Electrical Fixtures, Equipment", new BuiltInCategory[] {BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_ElectricalEquipment }}
            };
        }

        private List<object> GetAllowedTypes(Dictionary<string, object> selectOpts)
        {
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
            return allowedTypesOrCats;
        }

        private IList<Reference> SelectElements(UIDocument uidoc, ISelectionFilter filter, string prompt)
        {
            try
            {
                return uidoc.Selection.PickObjects(ObjectType.Element, filter, prompt);
            }
            catch (OperationCanceledException)
            {
                TaskDialog.Show("DEAXO", "Selection cancelled.");
                return null;
            }
        }

        private View GetViewTemplate(Document doc, string title)
        {
            var viewTemplates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate)
                .ToList();

            if (viewTemplates.Count == 0) return null;

            var templateNames = viewTemplates.Select(v => v.Name).ToList();
            templateNames.Insert(0, "None");

            var templateWindow = new SelectFromDictWindow(templateNames, title, allowMultiple: false);
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

        private void ApplyTemplateAndCreateSheet(Document doc, ViewSection elevation, View template,
            ElementProperties props, List<string> results, string prefix)
        {
            // Apply template
            if (template != null)
                elevation.ViewTemplateId = template.Id;

            // Get titleblock
            ElementId titleblockTypeId = GetTitleblockTypeId(doc);
            if (titleblockTypeId == null || titleblockTypeId == ElementId.InvalidElementId)
            {
                results.Add($"No titleblock available for element {props.Element.Id}");
                return;
            }

            // Create sheet
            var sheet = ViewSheet.Create(doc, titleblockTypeId);
            XYZ pos = new XYZ(0.5, 0.5, 0);

            try
            {
                if (Viewport.CanAddViewToSheet(doc, sheet.Id, elevation.Id))
                    Viewport.Create(doc, sheet.Id, elevation.Id, pos);
            }
            catch { }

            // Name sheet uniquely
            string typeName = props.TypeName ?? props.Element.Category?.Name ?? "Element";
            string baseSheetNumber = $"DEAXO_{prefix}_{typeName}_{props.Element.Id}";
            string baseSheetName = $"{props.Element.Category?.Name} - {GetSheetTypeName(prefix)} (DEAXO GmbH)";

            SetUniqueSheetName(sheet, baseSheetNumber, baseSheetName);
            results.Add($"{props.Element.Id} -> Sheet:{sheet.Id} Elevation:{elevation.Id}");
        }

        private ElementId GetTitleblockTypeId(Document doc)
        {
            var titleblockTypeId = doc.GetDefaultFamilyTypeId(new ElementId(BuiltInCategory.OST_TitleBlocks));
            if (titleblockTypeId != null && titleblockTypeId != ElementId.InvalidElementId)
                return titleblockTypeId;

            // Fallback
            var tb = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilySymbol>()
                .FirstOrDefault();

            return tb?.Id;
        }

        private string GetSheetTypeName(string prefix)
        {
            switch (prefix)
            {
                case "SE": return "Single Elevation";
                case "GE": return "Group Elevation";
                case "IE": return "Internal Elevation";
                default: return "Elevation";
            }
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
                    sheetNumber += "*";
                }
            }
        }

        private void ShowResults(List<string> results, string title)
        {
            var msgText = string.Join(Environment.NewLine, results.Take(200));
            TaskDialog.Show($"DEAXO - {title} Created",
                string.IsNullOrWhiteSpace(msgText) ? "No new elevations created." : msgText);
        }
    }
}