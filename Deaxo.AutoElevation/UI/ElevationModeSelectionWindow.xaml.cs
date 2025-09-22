using System.Windows;

namespace Deaxo.AutoElevation.UI
{
    public partial class ElevationModeSelectionWindow : Window
    {
        public enum ElevationMode
        {
            None,
            SingleElement,
            GroupElement,
            Internal
        }

        public ElevationMode SelectedMode { get; private set; } = ElevationMode.None;

        public ElevationModeSelectionWindow()
        {
            InitializeComponent();
        }

        private void SingleElement_Click(object sender, RoutedEventArgs e)
        {
            SelectedMode = ElevationMode.SingleElement;
            DialogResult = true;
            Close();
        }

        private void GroupElement_Click(object sender, RoutedEventArgs e)
        {
            SelectedMode = ElevationMode.GroupElement;
            DialogResult = true;
            Close();
        }

        private void InternalElevation_Click(object sender, RoutedEventArgs e)
        {
            SelectedMode = ElevationMode.Internal;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            SelectedMode = ElevationMode.None;
            DialogResult = false;
            Close();
        }
    }
}