using System.Windows;

namespace Deaxo.AutoElevation.UI
{
    public partial class SheetConfigurationWindow : Window
    {
        public enum SheetOption
        {
            None,
            Individual,
            Combined
        }

        public SheetOption SelectedOption { get; private set; } = SheetOption.Individual;

        public SheetConfigurationWindow()
        {
            InitializeComponent();
            UpdatePreviewText();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            UpdatePreviewText();
        }

        private void UpdatePreviewText()
        {
            if (PreviewText == null) return; // During initialization

            if (NoSheetsRadio?.IsChecked == true)
            {
                PreviewText.Text = "Only elevation views will be created (no sheets)";
            }
            else if (IndividualSheetsRadio?.IsChecked == true)
            {
                PreviewText.Text = "Individual sheets will be created for each view";
            }
            else if (CombinedSheetRadio?.IsChecked == true)
            {
                PreviewText.Text = "All views will be placed on a single combined sheet";
            }
        }

        private void Continue_Click(object sender, RoutedEventArgs e)
        {
            // Determine selected option based on radio buttons
            if (NoSheetsRadio?.IsChecked == true)
                SelectedOption = SheetOption.None;
            else if (IndividualSheetsRadio?.IsChecked == true)
                SelectedOption = SheetOption.Individual;
            else if (CombinedSheetRadio?.IsChecked == true)
                SelectedOption = SheetOption.Combined;
            else
                SelectedOption = SheetOption.Individual; // Default fallback

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}