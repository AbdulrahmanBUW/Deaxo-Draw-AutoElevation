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
    public class SingleElevationCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Category choices - ALL SELECTED BY DEFAULT
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
                    {"All Loadable Families", typeof(FamilyInstance)},
                    {"Electrical Fixtures, Equipment, Circuits", new BuiltInCategory[] {BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_ElectricalEquipment }}
                };

                // Use all categories by default
                var allowedTypesOrCats = new List<object>();
                foreach (var kvp in selectOpts)
                {
                    var val = kvp.Value;
                    if (val is BuiltInCategory bic) allowedTypesOrCats.Add(bic);
                    else if (val is BuiltInCategory[] bicArr)
                        allowedTypesOrCats.AddRange(bicArr.Cast<object>());
                    else allowedTypesOrCats.Add(val);
                }

                // Selection: prompt user to select elements with filter
                var selFilter = new DXSelectionFilter(allowedTypesOrCats);
                IList<Reference> refs = null;
                try
                {
                    refs = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "Select elements for individual elevations and click Finish");
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

                // Get view templates for optional selection
                var allViews = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().ToList();
                var viewTemplates = allViews.Where(v => v.IsTemplate).ToList();

                // Ask for template (single-select)
                View chosenTemplate = null;
                if (viewTemplates.Count > 0)
                {
                    var templateNames = viewTemplates.Select(v => v.Name).ToList();
                    templateNames.Insert(0, "None");

                    var templateWindow = new SelectFromDictWindow(templateNames, "Select ViewTemplate for Single Element Elevations", allowMultiple: false);
                    bool? tr = templateWindow.ShowDialog();
                    if (tr == true && templateWindow.SelectedItems.Count > 0)
                    {
                        var name = templateWindow.SelectedItems[0];
                        if (name != "None")
                        {
                            chosenTemplate = viewTemplates.FirstOrDefault(v => v.Name == name);
                        }
                    }
                }

                // Transaction: create elevations and place on sheets
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

                            // Create elevation only
                            var created = SectionGenerator.CreateElevationOnly(doc, props);
                            if (created == null || created.elevation == null) continue;

                            var elev = created.elevation;

                            // Apply view template if selected
                            if (chosenTemplate != null)
                                elev.ViewTemplateId = chosenTemplate.Id;

                            // Create sheet and place view
                            ElementId defaultTitleblockTypeId = doc.GetDefaultFamilyTypeId(new ElementId(BuiltInCategory.OST_TitleBlocks));
                            if (defaultTitleblockTypeId == null || defaultTitleblockTypeId == ElementId.InvalidElementId)
                            {
                                var tb = new FilteredElementCollector(doc)
                                    .OfClass(typeof(FamilySymbol))
                                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                    .Cast<FamilySymbol>()
                                    .FirstOrDefault();
                                if (tb != null)
                                    defaultTitleblockTypeId = tb.Id;
                            }

                            if (defaultTitleblockTypeId != null && defaultTitleblockTypeId != ElementId.InvalidElementId)
                            {
                                var vs = ViewSheet.Create(doc, defaultTitleblockTypeId);
                                XYZ pos = new XYZ(0.5, 0.5, 0);
                                try
                                {
                                    if (Viewport.CanAddViewToSheet(doc, vs.Id, elev.Id))
                                        Viewport.Create(doc, vs.Id, elev.Id, pos);
                                }
                                catch { }

                                // Name sheet (unique)
                                string typeName = props.TypeName ?? el.Category?.Name ?? "Element";
                                string sheetNumber = $"DEAXO_SE_{typeName}_{el.Id}";
                                string sheetName = $"{el.Category?.Name} - Single Elevation (DEAXO GmbH)";
                                for (int i = 0; i < 10; ++i)
                                {
                                    try
                                    {
                                        vs.SheetNumber = sheetNumber;
                                        vs.Name = sheetName;
                                        break;
                                    }
                                    catch
                                    {
                                        sheetNumber += "*";
                                    }
                                }

                                results.Add($"Element {el.Id} -> Sheet:{vs.Id} Elevation:{elev.Id}");
                            }
                        }
                        catch (Exception exInner)
                        {
                            TaskDialog.Show("DEAXO - Element Error", exInner.Message);
                        }
                    }

                    t.Commit();
                }

                // Show summary
                var msgText = string.Join(Environment.NewLine, results.Take(200));
                TaskDialog.Show("DEAXO - Single Element Elevations Created",
                    string.IsNullOrWhiteSpace(msgText) ? "No new elevations created." : msgText);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}