using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace Deaxo.AutoElevation
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "DEAXO Draw";

                // Create Ribbon tab if it doesn't exist
                try { application.CreateRibbonTab(tabName); } catch { }

                // Create Ribbon panel
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Auto Elevation");

                // Get assembly path
                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                // Button 1: Internal Elevations (uses existing icons)
                PushButtonData internalButtonData = new PushButtonData(
                    "InternalElevation",
                    "Internal Elevation",
                    assemblyPath,
                    "Deaxo.AutoElevation.Commands.InternalElevationCommand"
                )
                {
                    ToolTip = "Create building elevations for internal views of selected walls",
                };

                // Load existing icons for Internal Elevations
                internalButtonData.LargeImage = LoadBitmapFromEmbeddedResource("Deaxo.AutoElevation.Resources.icon32.png");
                internalButtonData.Image = LoadBitmapFromEmbeddedResource("Deaxo.AutoElevation.Resources.icon16.png");

                // Button 2: Group Views (uses new group icons)
                PushButtonData groupButtonData = new PushButtonData(
                    "GroupViews",
                    "Group Views",
                    assemblyPath,
                    "Deaxo.AutoElevation.Commands.GroupElevationCommand"
                )
                {
                    ToolTip = "Create group elevation views (6 orthographic views + 3D) for selected elements",
                };

                // Load new group icons
                groupButtonData.LargeImage = LoadBitmapFromEmbeddedResource("Deaxo.AutoElevation.Resources.Group_Icon-32.png");
                groupButtonData.Image = LoadBitmapFromEmbeddedResource("Deaxo.AutoElevation.Resources.Group_Icon-16.png");

                // Add buttons to panel
                panel.AddItem(internalButtonData);
                panel.AddItem(groupButtonData);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("DEAXO - Ribbon Error", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// Loads an embedded PNG and converts it to BitmapSource for Revit ribbon.
        /// </summary>
        /// <param name="resourceName">Namespace + folder + filename</param>
        /// <returns>BitmapSource or null</returns>
        private BitmapSource LoadBitmapFromEmbeddedResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;

                using (var bmp = new Bitmap(stream))
                {
                    var bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(
                        bmp.GetHbitmap(),
                        IntPtr.Zero,
                        System.Windows.Int32Rect.Empty,
                        BitmapSizeOptions.FromWidthAndHeight(bmp.Width, bmp.Height)
                    );
                    bitmapSource.Freeze();
                    return bitmapSource;
                }
            }
        }
    }
}