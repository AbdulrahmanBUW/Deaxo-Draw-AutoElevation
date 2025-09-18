using System;
using Autodesk.Revit.DB;

namespace Deaxo.AutoElevation.Commands
{
    public class SectionResult
    {
        public ViewSection elevation;
        public ViewSection cross;
        public ViewSection plan;
    }

    public static class SectionGenerator
    {
        public static SectionResult CreateSections(Document doc, ElementProperties props)
        {
            if (props == null || !props.IsValid) return null;

            var boxElev = CreateSectionBox(props, "elevation");
            var boxCross = CreateSectionBox(props, "cross");
            var boxPlan = CreateSectionBox(props, "plan");

            ElementId sectionTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection);
            if (sectionTypeId == null || sectionTypeId == ElementId.InvalidElementId)
                throw new Exception("No section view type available in document.");

            var elev = ViewSection.CreateSection(doc, sectionTypeId, boxElev);
            var cross = ViewSection.CreateSection(doc, sectionTypeId, boxCross);
            var plan = ViewSection.CreateSection(doc, sectionTypeId, boxPlan);

            // rename views
            RenameViewSafe(elev, $"{props.TypeName}_{props.Element.Id}_Elevation");
            RenameViewSafe(cross, $"{props.TypeName}_{props.Element.Id}_Cross");
            RenameViewSafe(plan, $"{props.TypeName}_{props.Element.Id}_Plan");

            return new SectionResult { elevation = elev, cross = cross, plan = plan };
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
            // create transform similar to python logic
            Transform t = CreateTransform(props.Origin, props.Vector, mode);
            double W_half = props.Width / 2.0;
            double H_half = props.Height / 2.0;
            double D_half = props.Depth / 2.0;

            if (mode == "elevation")
            {
                double half = props.Width / 2.0;
                sectionBox.Min = new XYZ(-half - props.offset(), -H_half - props.offset(), 0);
                sectionBox.Max = new XYZ(half + props.offset(), H_half + props.offset(), D_half + props.offset());
            }
            else if (mode == "cross")
            {
                sectionBox.Min = new XYZ(-D_half - props.offset(), -H_half - props.offset(), 0);
                sectionBox.Max = new XYZ(D_half + props.offset(), H_half + props.offset(), W_half + props.offset());
            }
            else // plan
            {
                sectionBox.Min = new XYZ(-W_half - props.offset(), -D_half - props.offset(), 0);
                sectionBox.Max = new XYZ(W_half + props.offset(), D_half + props.offset(), H_half + props.offset());
            }

            sectionBox.Transform = t;
            return sectionBox;
        }

        private static Transform CreateTransform(XYZ origin, XYZ vector, string mode)
        {
            var trans = Transform.Identity;
            trans.Origin = origin;

            var v = vector.Normalize();

            if (mode == "elevation")
            {
                trans.BasisX = v;
                trans.BasisY = XYZ.BasisZ;
                trans.BasisZ = v.CrossProduct(XYZ.BasisZ);
            }
            else if (mode == "cross")
            {
                var vcross = v.CrossProduct(XYZ.BasisZ);
                trans.BasisX = vcross;
                trans.BasisY = XYZ.BasisZ;
                trans.BasisZ = vcross.CrossProduct(XYZ.BasisZ);
            }
            else // plan
            {
                trans.BasisX = -v;
                trans.BasisY = -(XYZ.BasisZ.CrossProduct(-v)).Normalize();
                trans.BasisZ = -XYZ.BasisZ;
            }

            return trans;
        }
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
