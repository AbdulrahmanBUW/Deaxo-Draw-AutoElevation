using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Deaxo.AutoElevation.UI
{
    public partial class SelectFromDictWindow : Window
    {
        public List<string> SelectedItems { get; private set; } = new List<string>();
        private readonly bool allowMultiple;
        private readonly List<string> allOptions;
        private readonly ObservableCollection<string> filteredOptions;

        public SelectFromDictWindow(List<string> options, string title, bool allowMultiple = false)
        {
            InitializeComponent();
            this.allowMultiple = false; // Force single selection only
            this.allOptions = options ?? new List<string>();
            this.filteredOptions = new ObservableCollection<string>(allOptions);

            // Set window title and header
            Title = $"DEAXO Auto Elevation - {title}";
            TitleText.Text = title;

            // Bind filtered options to ListBox
            OptionsList.ItemsSource = filteredOptions;

            // Update status text for single selection
            StatusText.Text = "Select one template or continue with default formatting";
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterOptions();
        }

        private void SearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            // No placeholder handling needed anymore
        }

        private void SearchBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // No placeholder handling needed anymore
        }

        private void FilterOptions()
        {
            string searchText = SearchBox.Text?.ToLower() ?? string.Empty;

            filteredOptions.Clear();

            var filtered = string.IsNullOrWhiteSpace(searchText)
                ? allOptions
                : allOptions.Where(option => option.ToLower().Contains(searchText));

            foreach (var option in filtered)
            {
                filteredOptions.Add(option);
            }

            UpdateStatusText();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            UpdateStatusText();
        }

        private void UpdateStatusText()
        {
            var selectedTemplate = GetSelectedTemplate();

            if (string.IsNullOrEmpty(selectedTemplate) || selectedTemplate == "None")
            {
                StatusText.Text = "No template selected - default formatting will be used";
            }
            else
            {
                StatusText.Text = $"Selected: {selectedTemplate}";
            }
        }

        private string GetSelectedTemplate()
        {
            foreach (var item in filteredOptions)
            {
                // Find the RadioButton for this item
                var container = OptionsList.ItemContainerGenerator.ContainerFromItem(item) as FrameworkElement;
                if (container != null)
                {
                    var radioButton = FindChild<RadioButton>(container);
                    if (radioButton != null && radioButton.IsChecked == true)
                        return item;
                }
            }
            return null;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            SelectedItems.Clear();

            var selectedTemplate = GetSelectedTemplate();
            if (!string.IsNullOrEmpty(selectedTemplate))
            {
                SelectedItems.Add(selectedTemplate);
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // Recursive search for child RadioButton
        private T FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T tChild) return tChild;

                var result = FindChild<T>(child);
                if (result != null) return result;
            }
            return null;
        }
    }
}