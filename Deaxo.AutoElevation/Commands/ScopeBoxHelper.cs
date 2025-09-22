using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Helper class for creating and managing scope boxes for group elevations
    /// Simplified and more reliable implementation
    /// </summary>
    public static class ScopeBoxHelper
    {
        /// <summary>
        /// Creates a simple reference element instead of a complex scope box
        /// </summary>
        public static Element CreateScopeBoxFromElements(Document doc, List<Element> elements)
        {
            if (elements == null || elements.Count == 0)
                return null;

            try
            {
                BoundingBoxXYZ overallBB = CalculateOverallBoundingBox(elements);
                if (overallBB == null)
                    return null;

                // Create a simple model line as reference instead of complex scope box
                XYZ center = (overallBB.Min + overallBB.Max) / 2;
                XYZ start = center - new XYZ(1, 0, 0);
                XYZ end = center + new XYZ(1, 0, 0);

                Line line = Line.CreateBound(start, end);

                // Get the active work plane or create one
                Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, center);
                SketchPlane sp = SketchPlane.Create(doc, plane);

                ModelCurve modelLine = doc.Create.NewModelCurve(line, sp);
                return modelLine;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CreateScopeBoxFromElements failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Calculates the overall bounding box for a list of elements
        /// </summary>
        public static BoundingBoxXYZ CalculateOverallBoundingBox(List<Element> elements)
        {
            if (elements == null || elements.Count == 0)
                return null;

            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            bool hasValidBounds = false;

            foreach (var element in elements)
            {
                try
                {
                    BoundingBoxXYZ bb = element.get_BoundingBox(null);
                    if (bb != null)
                    {
                        minX = Math.Min(minX, bb.Min.X);
                        minY = Math.Min(minY, bb.Min.Y);
                        minZ = Math.Min(minZ, bb.Min.Z);
                        maxX = Math.Max(maxX, bb.Max.X);
                        maxY = Math.Max(maxY, bb.Max.Y);
                        maxZ = Math.Max(maxZ, bb.Max.Z);
                        hasValidBounds = true;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to get bounding box for element {element.Id}: {ex.Message}");
                    continue;
                }
            }

            if (!hasValidBounds)
                return null;

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }

        /// <summary>
        /// Safely deletes a scope box element
        /// </summary>
        public static void DeleteScopeBoxSafely(Document doc, Element scopeBox)
        {
            if (scopeBox == null || doc == null)
                return;

            try
            {
                doc.Delete(scopeBox.Id);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to delete scope box {scopeBox.Id}: {ex.Message}");
                // Continue execution even if deletion fails
            }
        }
    }
}