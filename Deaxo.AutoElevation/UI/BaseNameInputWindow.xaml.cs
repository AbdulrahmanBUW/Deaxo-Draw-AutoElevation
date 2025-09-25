using System;
using System.Windows;
using System.Windows.Input;

namespace Deaxo.AutoElevation.UI
{
    /// <summary>
    /// Simple input window for getting base name from user
    /// </summary>
    public partial class BaseNameInputWindow : Window
    {
        public string BaseName { get; private set; }

        public BaseNameInputWindow()
        {
            InitializeComponent();

            // Set focus to text box and select all text
            Loaded += BaseNameInputWindow_Loaded;

            // Allow Enter key to submit
            BaseNameTextBox.KeyDown += BaseNameTextBox_KeyDown;
        }

        private void BaseNameInputWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Focus the text box and select all text for easy replacement
            BaseNameTextBox.Focus();
            BaseNameTextBox.SelectAll();
        }

        private void BaseNameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Continue_Click(sender, e);
            }
            else if (e.Key == Key.Escape)
            {
                Cancel_Click(sender, e);
            }
        }

        private void Continue_Click(object sender, RoutedEventArgs e)
        {
            BaseName = BaseNameTextBox.Text?.Trim() ?? string.Empty;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            BaseName = null;
            DialogResult = false;
            Close();
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            // Don't show in taskbar
            this.ShowInTaskbar = false;
        }
    }
}