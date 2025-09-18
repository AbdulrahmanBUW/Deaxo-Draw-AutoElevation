using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Deaxo.AutoElevation.UI
{
    public partial class SelectFromDictWindow : Window
    {
        public List<string> SelectedItems { get; private set; } = new List<string>();
        private readonly bool allowMultiple;

        public SelectFromDictWindow(List<string> options, string title, bool allowMultiple)
        {
            InitializeComponent();
            this.allowMultiple = allowMultiple;

            // Set window title
            Title = title;

            // Bind options to ListBox
            OptionsList.ItemsSource = options;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            SelectedItems.Clear();
            foreach (var item in OptionsList.Items)
            {
                var container = OptionsList.ItemContainerGenerator.ContainerFromItem(item) as FrameworkElement;
                if (container != null)
                {
                    var checkBox = FindChild<CheckBox>(container);
                    if (checkBox != null && checkBox.IsChecked == true)
                        SelectedItems.Add(item.ToString());
                }
            }

            if (!allowMultiple && SelectedItems.Count > 1)
                SelectedItems = SelectedItems.Take(1).ToList();

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            SetAllCheckboxes(true);
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            SetAllCheckboxes(false);
        }

        private void SetAllCheckboxes(bool isChecked)
        {
            foreach (var item in OptionsList.Items)
            {
                var container = OptionsList.ItemContainerGenerator.ContainerFromItem(item) as FrameworkElement;
                if (container != null)
                {
                    var checkBox = FindChild<CheckBox>(container);
                    if (checkBox != null)
                        checkBox.IsChecked = isChecked;
                }
            }
        }

        // Recursive search for child CheckBox
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
