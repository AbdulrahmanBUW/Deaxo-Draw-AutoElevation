using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace Deaxo.AutoElevation.UI
{
    public partial class ProgressWindow : Window
    {
        private readonly StringBuilder logBuilder = new StringBuilder();
        private bool isCancelled = false;

        public bool IsCancelled => isCancelled;

        public ProgressWindow()
        {
            InitializeComponent();
        }

        public void UpdateStatus(string status, string detail = null)
        {
            Dispatcher.Invoke(() =>
            {
                StatusText.Text = status;
                if (!string.IsNullOrEmpty(detail))
                {
                    DetailText.Text = detail;
                }
            });
        }

        public void UpdateProgress(int current, int total)
        {
            Dispatcher.Invoke(() =>
            {
                if (total > 0)
                {
                    double percentage = (double)current / total * 100;
                    MainProgressBar.Value = percentage;
                    ProgressText.Text = $"{percentage:F0}%";
                    CountText.Text = $"{current} of {total} completed";
                }
            });
        }

        public void SetProgressRange(int minimum, int maximum)
        {
            Dispatcher.Invoke(() =>
            {
                MainProgressBar.Minimum = minimum;
                MainProgressBar.Maximum = maximum;
            });
        }

        public void AddLogMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                logBuilder.AppendLine($"[{timestamp}] {message}");
                LogText.Text = logBuilder.ToString();

                // Auto-scroll to bottom
                var scrollViewer = FindChild<ScrollViewer>(LogText.Parent as DependencyObject);
                scrollViewer?.ScrollToEnd();
            });
        }

        public void ShowCompletion(List<string> results, string title)
        {
            Dispatcher.Invoke(() =>
            {
                // Update UI to show completion
                StatusText.Text = "Completed Successfully";
                DetailText.Text = $"Created {results.Count} elevation views";
                MainProgressBar.Value = 100;
                ProgressText.Text = "100%";

                // Change cancel button to close
                CancelButton.Content = "Close";
                CancelButton.Background = System.Windows.Media.Brushes.Green;

                // Add completion log
                AddLogMessage($"Process completed successfully. Created {results.Count} items.");

                // Stop spinning animation
                var storyboard = (System.Windows.Media.Animation.Storyboard)FindResource("SpinAnimation");
                storyboard.Stop();
            });
        }

        public void ShowError(string error)
        {
            Dispatcher.Invoke(() =>
            {
                StatusText.Text = "Error Occurred";
                DetailText.Text = "An error occurred during processing";

                // Change cancel button to close
                CancelButton.Content = "Close";
                CancelButton.Background = System.Windows.Media.Brushes.Red;

                AddLogMessage($"ERROR: {error}");

                // Stop spinning animation
                var storyboard = (System.Windows.Media.Animation.Storyboard)FindResource("SpinAnimation");
                storyboard.Stop();
            });
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (CancelButton.Content.ToString() == "Close")
            {
                DialogResult = true;
                Close();
            }
            else
            {
                isCancelled = true;
                CancelButton.IsEnabled = false;
                CancelButton.Content = "Cancelling...";
                AddLogMessage("User requested cancellation...");
            }
        }

        // Helper method to find child controls
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

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            // Remove the window from Alt+Tab
            this.WindowStyle = WindowStyle.ToolWindow;
            this.ShowInTaskbar = false;
        }
    }
}