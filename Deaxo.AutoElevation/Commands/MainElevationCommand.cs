using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Deaxo.AutoElevation.UI;

namespace Deaxo.AutoElevation.Commands
{
    /// <summary>
    /// Main unified command that shows mode selection dialog and delegates to specific commands
    /// This is optional - you can use individual commands directly from the split button
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class MainElevationCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Show mode selection dialog
                var modeWindow = new ElevationModeSelectionWindow();
                bool? result = modeWindow.ShowDialog();

                if (result != true || modeWindow.SelectedMode == ElevationModeSelectionWindow.ElevationMode.None)
                {
                    return Result.Cancelled;
                }

                // Delegate to appropriate command based on selection
                IExternalCommand selectedCommand = null;

                switch (modeWindow.SelectedMode)
                {
                    case ElevationModeSelectionWindow.ElevationMode.SingleElement:
                        selectedCommand = new SingleElevationCommand();
                        break;
                    case ElevationModeSelectionWindow.ElevationMode.GroupElement:
                        selectedCommand = new GroupElevationCommand();
                        break;
                    case ElevationModeSelectionWindow.ElevationMode.Internal:
                        selectedCommand = new InternalElevationCommand();
                        break;
                    default:
                        return Result.Cancelled;
                }

                // Execute the selected command
                return selectedCommand.Execute(commandData, ref message, elements);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}