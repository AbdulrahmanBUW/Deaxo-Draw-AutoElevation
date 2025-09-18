using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Selection filter which accepts elements of specified BuiltInCategory or element Types (like FamilyInstance).
    /// Input list may contain BuiltInCategory, Type (System.Type) or other markers.
    /// </summary>
    public class DXSelectionFilter : ISelectionFilter
    {
        private readonly HashSet<ElementId> _allowedCategoryIds = new HashSet<ElementId>();
        private readonly HashSet<Type> _allowedTypes = new HashSet<Type>();

        public DXSelectionFilter(IEnumerable<object> allowed)
        {
            foreach (var a in allowed)
            {
                if (a == null) continue;
                if (a is BuiltInCategory bic)
                {
                    _allowedCategoryIds.Add(new ElementId(bic));
                }
                else if (a is BuiltInCategory[] arr)
                {
                    foreach (var b in arr) _allowedCategoryIds.Add(new ElementId(b));
                }
                else if (a is Type t)
                {
                    _allowedTypes.Add(t);
                }
                else
                {
                    // ignore unknown types
                }
            }
        }

        public bool AllowElement(Element elem)
        {
            if (elem == null) return false;
            if (elem.ViewSpecific) return false;

            // Type-check
            var et = elem.GetType();
            if (_allowedTypes.Contains(et)) return true;

            // Category-check
            var cat = elem.Category;
            if (cat != null && _allowedCategoryIds.Contains(cat.Id)) return true;

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }
}
