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

                View chosenTemplate = GetViewTemplate(doc);

                var progressWindow = new ProgressWindow();
                progressWindow.Show();
                progressWindow.UpdateStatus("Creating internal building elevations...", "Processing selected walls");

                var results = new List<string>();
                var startTime = DateTime.Now;

                using (Transaction t = new Transaction(doc, "DEAXO - Create Internal Building Elevations"))
                {
                    t.Start();

                    try
                    {
                        progressWindow.UpdateProgress(0, refs.Count);

                        for (int i = 0; i < refs.Count; i++)
                        {
                            var r = refs[i];

                            try
                            {
                                Element el = doc.GetElement(r);
                                progressWindow.UpdateStatus($"Processing wall {i + 1} of {refs.Count}...",
                                    $"Wall ID: {el.Id}");

                                if (!(el is Wall wall))
                                {
                                    results.Add($"Skipped element {el.Id}: Not a wall (Type: {el.GetType().Name})");
                                    continue;
                                }

                                var elevation = CreateInternalBuildingElevation(doc, wall, chosenTemplate);
                                if (elevation == null)
                                {
                                    results.Add($"Failed to create building elevation for wall {el.Id}");
                                    continue;
                                }

                                results.Add($"✓ Created building elevation {elevation.Id} for wall {el.Id}");

                                progressWindow.UpdateProgress(i + 1, refs.Count);
                                progressWindow.AddLogMessage($"Created building elevation for wall {wall.Id}");
                            }
                            catch (Exception exInner)
                            {
                                var error = $"Error processing wall {r.ElementId}: {exInner.Message}";
                                results.Add($"✗ {error}");
                            }
                        }

                        t.Commit();

                        var duration = DateTime.Now - startTime;
                        progressWindow.ShowCompletion(results, "Internal Building Elevations");

                        System.Threading.Thread.Sleep(1500);
                        progressWindow.Close();

                        var resultsWindow = new ResultsWindow(results, "Internal Building Elevations", duration);
                        resultsWindow.ShowDialog();
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

        private ViewSection CreateInternalBuildingElevation(Document doc, Wall wall, View viewTemplate)
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

                string wallTypeName = GetWallTypeName(wall);
                string elevationName = $"{wallTypeName}_IE";
                SetUniqueViewName(elevation, elevationName);

                return elevation;
            }
            catch (Exception ex)
            {
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

                if (viewTemplate != null)
                {
                    try
                    {
                        chosenView.ViewTemplateId = viewTemplate.Id;
                    }
                    catch { }
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

                // paddings you can tweak
                double horizontalPadding = 0.0; // left/right in view X
                double verticalPadding = 0.0;   // top/bottom in view Z
                                                // depth of crop box (how far into the model the view looks)
                double cropDepth = Math.Max(0.1, wall.Width);

                // Use current crop transform if available; otherwise use view's default
                var currentCropBox = elevation.CropBox ?? new BoundingBoxXYZ();
                Transform cropTransform = currentCropBox.Transform ?? elevation.CropBox.Transform;
                if (cropTransform == null)
                {
                    // As a fallback, use the view orientation to build a transform (rare)
                    cropTransform = Transform.Identity;
                }

                // Inverse transform: world -> crop-local
                Transform inv = null;
                try { inv = cropTransform.Inverse; }
                catch { inv = Transform.Identity; }

                // Wall endpoints in world coordinates
                XYZ p0 = locationCurve.Curve.GetEndPoint(0);
                XYZ p1 = locationCurve.Curve.GetEndPoint(1);

                // Convert wall endpoints and bounding box Z extents to crop-local coords
                XYZ p0Local = inv.OfPoint(p0);
                XYZ p1Local = inv.OfPoint(p1);

                XYZ wallMinWorld = wallBounds.Min;
                XYZ wallMaxWorld = wallBounds.Max;
                XYZ wallMinLocal = inv.OfPoint(wallMinWorld);
                XYZ wallMaxLocal = inv.OfPoint(wallMaxWorld);

                // Determine X range from endpoints (in local coords)
                double xMin = Math.Min(p0Local.X, p1Local.X) - horizontalPadding;
                double xMax = Math.Max(p0Local.X, p1Local.X) + horizontalPadding;

                // Y range centered around 0 with depth (some families put marker Y offset; center is safe)
                double yMin = -cropDepth * 0.5;
                double yMax = cropDepth * 0.5;

                // Z range from transformed wall bounding box
                double zMin = Math.Min(wallMinLocal.Z, wallMaxLocal.Z) - verticalPadding;
                double zMax = Math.Max(wallMinLocal.Z, wallMaxLocal.Z) + verticalPadding;

                var cropBox = new BoundingBoxXYZ
                {
                    Min = new XYZ(xMin, yMin, zMin),
                    Max = new XYZ(xMax, yMax, zMax),
                    Transform = cropTransform
                };

                // Assign crop box to view (inside outer transaction)
                elevation.CropBox = cropBox;
                elevation.CropBoxActive = true;
                elevation.CropBoxVisible = true;
            }
            catch
            {
                // swallow or log; keep overall process robust
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
                "Select ViewTemplate for Internal Building Elevations", allowMultiple: false);
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

        private void CreateSheetForElevation(Document doc, ViewSection elevation, Wall wall, List<string> results)
        {
            try
            {
                ElementId titleblockTypeId = GetTitleblockTypeId(doc);
                if (titleblockTypeId == null || titleblockTypeId == ElementId.InvalidElementId)
                {
                    results.Add($"Warning: No titleblock available for wall {wall.Id}");
                    return;
                }

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

                string wallTypeName = GetWallTypeName(wall);
                string sheetNumber = $"DEAXO_IE_{wallTypeName}_{wall.Id}";
                string sheetName = $"{wall.Category?.Name} - Internal Building Elevation (DEAXO GmbH)";

                SetUniqueSheetName(sheet, sheetNumber, sheetName);
                results.Add($"✓ Created sheet {sheet.Id} for building elevation {elevation.Id}");
            }
            catch (Exception ex)
            {
                results.Add($"Failed to create sheet for wall {wall.Id}: {ex.Message}");
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

        private void SetUniqueViewName(ViewSection view, string baseName)
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
                    viewName = $"{baseName}_{i + 1}";
                }
            }
        }
    }
}