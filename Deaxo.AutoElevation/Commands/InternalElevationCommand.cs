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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Categories for walls and curtain walls only
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

                // Selection with proper filter
                var selFilter = new DXSelectionFilter(allowedTypesOrCats);
                IList<Reference> refs;
                try
                {
                    refs = uidoc.Selection.PickObjects(ObjectType.Element, selFilter,
                        "Select walls or curtain walls for internal elevations and click Finish");
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

                // Get view template
                View chosenTemplate = GetViewTemplate(doc);

                // Create internal elevations
                var results = new List<string>();
                using (Transaction t = new Transaction(doc, "DEAXO - Create Internal Elevations"))
                {
                    t.Start();

                    foreach (var r in refs)
                    {
                        try
                        {
                            Element el = doc.GetElement(r);

                            // Validate element type
                            if (!(el is Wall wall))
                            {
                                results.Add($"Skipped element {el.Id}: Not a wall (Type: {el.GetType().Name})");
                                continue;
                            }

                            // Get element properties
                            var props = new ElementProperties(doc, el);
                            if (!props.IsValid)
                            {
                                results.Add($"Skipped wall {el.Id}: Invalid properties (W={props.Width:F1}, H={props.Height:F1}, D={props.Depth:F1})");
                                continue;
                            }

                            // Log processing details
                            results.Add($"Processing wall {el.Id}: W={props.Width:F1}, H={props.Height:F1}, D={props.Depth:F1}");

                            // Create internal elevation
                            var created = SectionGenerator.CreateInternalElevation(doc, props);
                            if (created?.elevation == null)
                            {
                                results.Add($"Failed to create elevation for wall {el.Id}");
                                continue;
                            }

                            var elev = created.elevation;
                            results.Add($"Successfully created elevation view {elev.Id} for wall {el.Id}");

                            // Apply view template
                            if (chosenTemplate != null)
                            {
                                elev.ViewTemplateId = chosenTemplate.Id;
                                results.Add($"Applied view template to elevation {elev.Id}");
                            }

                            // Create sheet
                            CreateSheetForElevation(doc, elev, props, results);
                        }
                        catch (Exception exInner)
                        {
                            results.Add($"Error processing wall {r.ElementId}: {exInner.Message}");
                        }
                    }

                    t.Commit();
                }

                // Show results
                var msgText = string.Join(Environment.NewLine, results.Take(100));
                TaskDialog.Show("DEAXO - Internal Elevations Created",
                    string.IsNullOrWhiteSpace(msgText) ? "No new internal elevations created." : msgText);

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
                "Select ViewTemplate for Internal Elevations", allowMultiple: false);
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

        private void CreateSheetForElevation(Document doc, ViewSection elevation, ElementProperties props, List<string> results)
        {
            try
            {
                // Get titleblock
                ElementId titleblockTypeId = GetTitleblockTypeId(doc);
                if (titleblockTypeId == null || titleblockTypeId == ElementId.InvalidElementId)
                {
                    results.Add($"No titleblock available for wall {props.Element.Id}");
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
                catch (Exception ex)
                {
                    results.Add($"Failed to place elevation {elevation.Id} on sheet: {ex.Message}");
                }

                // Name sheet uniquely
                string typeName = props.TypeName ?? props.Element.Category?.Name ?? "Wall";
                string sheetNumber = $"DEAXO_IE_{typeName}_{props.Element.Id}";
                string sheetName = $"{props.Element.Category?.Name} - Internal Elevation (DEAXO GmbH)";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);
                results.Add($"Wall {props.Element.Id} -> Sheet:{sheet.Id} Internal Elevation:{elevation.Id}");
            }
            catch (Exception ex)
            {
                results.Add($"Failed to create sheet for wall {props.Element.Id}: {ex.Message}");
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

    /// <summary>
    /// Custom selection filter for walls only - simplified and reliable
    /// </summary>
    public class WallSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem != null && !elem.ViewSpecific && elem is Wall;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }
}