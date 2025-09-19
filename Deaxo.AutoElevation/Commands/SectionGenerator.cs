using System;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    public class ElevationResult
    {
        public ViewSection elevation;
    }

    public static class SectionGenerator
    {
        public static ElevationResult CreateElevationOnly(Document doc, ElementProperties props)
        {
            if (props == null || !props.IsValid) return null;

            var boxElev = CreateSectionBox(props, "elevation");

            ElementId sectionTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection);
            if (sectionTypeId == null || sectionTypeId == ElementId.InvalidElementId)
                throw new Exception("No section view type available in document.");

            var elev = ViewSection.CreateSection(doc, sectionTypeId, boxElev);

            // rename view
            RenameViewSafe(elev, $"{props.TypeName}_{props.Element.Id}_Elevation");

            return new ElevationResult { elevation = elev };
        }

        // Keep this method for backward compatibility if needed
        public static SectionResult CreateSections(Document doc, ElementProperties props)
        {
            var elevationResult = CreateElevationOnly(doc, props);
            if (elevationResult == null) return null;

            return new SectionResult { elevation = elevationResult.elevation, cross = null, plan = null };
        }

        private static void RenameViewSafe(View view, string newName)
        {
            if (view == null) return;
            for (int i = 0; i < 10; ++i)
            {
                try
                {
                    view.Name = newName;
                    break;
                }
                catch
                {
                    newName = newName + "*";
                }
            }
        }

        private static BoundingBoxXYZ CreateSectionBox(ElementProperties props, string mode)
        {
            var sectionBox = new BoundingBoxXYZ();
            // create transform for elevation view only
            Transform t = CreateTransform(props.Origin, props.Vector, mode);
            double W_half = props.Width / 2.0;
            double H_half = props.Height / 2.0;
            double D_half = props.Depth / 2.0;

            // Only handle elevation mode
            if (mode == "elevation")
            {
                double half = props.Width / 2.0;
                sectionBox.Min = new XYZ(-half - props.offset(), -H_half - props.offset(), 0);
                sectionBox.Max = new XYZ(half + props.offset(), H_half + props.offset(), D_half + props.offset());
            }
            else
            {
                // Fallback to elevation if invalid mode is passed
                double half = props.Width / 2.0;
                sectionBox.Min = new XYZ(-half - props.offset(), -H_half - props.offset(), 0);
                sectionBox.Max = new XYZ(half + props.offset(), H_half + props.offset(), D_half + props.offset());
            }

            sectionBox.Transform = t;
            return sectionBox;
        }

        private static Transform CreateTransform(XYZ origin, XYZ vector, string mode)
        {
            var trans = Transform.Identity;
            trans.Origin = origin;

            var v = vector.Normalize();

            // Only create transform for elevation
            trans.BasisX = v;
            trans.BasisY = XYZ.BasisZ;
            trans.BasisZ = v.CrossProduct(XYZ.BasisZ);

            return trans;
        }
    }

    public class SectionResult
    {
        public ViewSection elevation;
        public ViewSection cross;
        public ViewSection plan;
    }

    // extension to ElementProperties for offsets used above (we add a default)
    public static class ElementPropertiesExtensions
    {
        public static double offset(this ElementProperties p)
        {
            // default offset values similar to python class
            return 1.0;
        }
    }
}