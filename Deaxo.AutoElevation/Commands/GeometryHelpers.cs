using Autodesk.Revit.DB;
using System;

namespace Deaxo.AutoElevation.Commands
{
    public static class GeometryHelpers
    {
        public static XYZ RotateVector(XYZ vector, double rotationRadians)
        {
            double x = vector.X;
            double y = vector.Y;
            double rz = vector.Z;

            double rx = x * Math.Cos(rotationRadians) - y * Math.Sin(rotationRadians);
            double ry = x * Math.Sin(rotationRadians) + y * Math.Cos(rotationRadians);

            return new XYZ(rx, ry, rz);
        }
    }
}
