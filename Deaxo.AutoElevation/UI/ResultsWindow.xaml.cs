using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace Deaxo.AutoElevation.UI
{
    public partial class ResultsWindow : Window
    {
        private readonly List<string> results;
        private readonly string operationType;
        private readonly TimeSpan duration;

        public ResultsWindow(List<string> results, string title, TimeSpan duration)
        {
            InitializeComponent();

            this.results = results ?? new List<string>();
            this.operationType = title;
            this.duration = duration;

            InitializeWindow();
            PopulateResults();
        }

        private void InitializeWindow()
        {
            Title = $"DEAXO Auto Elevation - {operationType} Results";
            TitleText.Text = $"{operationType} Completed Successfully!";

            if (results.Count == 0)
            {
                SubtitleText.Text = "No new elevations were created";
                TitleText.Text = $"{operationType} Completed";
            }
            else if (results.Count == 1)
            {
                SubtitleText.Text = "1 elevation view and sheet have been generated";
            }
            else
            {
                SubtitleText.Text = $"{results.Count} elevation views and sheets have been generated";
            }
        }

        private void PopulateResults()
        {
            // Parse results to count different types
            int elementCount = 0;
            int viewCount = 0;
            int sheetCount = 0;
            var processedResults = new List<ResultItem>();

            foreach (var result in results)
            {
                if (result.Contains("Element") || result.Contains("Wall"))
                    elementCount++;
                if (result.Contains("Elevation:") || result.Contains("View"))
                    viewCount++;
                if (result.Contains("Sheet:"))
                    sheetCount++;

                // Create result item with timestamp
                processedResults.Add(new ResultItem
                {
                    Text = result,
                    Timestamp = DateTime.Now.ToString("HH:mm:ss")
                });
            }

            // Update statistics
            ElementCountText.Text = elementCount.ToString();
            ViewCountText.Text = viewCount.ToString();
            SheetCountText.Text = sheetCount.ToString();
            DurationText.Text = $"{duration.TotalSeconds:F1}s";

            // Populate results list
            ResultsList.ItemsSource = processedResults.Select(r => new ResultDisplayItem
            {
                Text = r.Text,
                Tag = r.Timestamp
            });
        }

        private void CopyResults_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine($"DEAXO Auto Elevation - {operationType} Results");
                sb.AppendLine($"Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"Duration: {duration.TotalSeconds:F1} seconds");
                sb.AppendLine($"Total Results: {results.Count}");
                sb.AppendLine();
                sb.AppendLine("Detailed Results:");
                sb.AppendLine(new string('-', 50));

                foreach (var result in results)
                {
                    sb.AppendLine($"• {result}");
                }

                Clipboard.SetText(sb.ToString());

                // Show confirmation
                MessageBox.Show("Results copied to clipboard!", "DEAXO Auto Elevation",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to copy results: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                    FileName = $"DEAXO_ElevationResults_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                    DefaultExt = "txt"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("DEAXO Auto Elevation - Results Export");
                    sb.AppendLine($"Operation: {operationType}");
                    sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine($"Duration: {duration.TotalSeconds:F1} seconds");
                    sb.AppendLine($"Total Results: {results.Count}");
                    sb.AppendLine();
                    sb.AppendLine("Detailed Results:");
                    sb.AppendLine(new string('=', 60));

                    for (int i = 0; i < results.Count; i++)
                    {
                        sb.AppendLine($"{i + 1:D3}. {results[i]}");
                    }

                    sb.AppendLine();
                    sb.AppendLine(new string('=', 60));
                    sb.AppendLine("Generated by DEAXO Auto Elevation Tool");
                    sb.AppendLine("© DEAXO GmbH");

                    File.WriteAllText(saveDialog.FileName, sb.ToString());

                    MessageBox.Show($"Results exported to:\n{saveDialog.FileName}",
                        "DEAXO Auto Elevation", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to export results: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        // Helper classes for data binding
        private class ResultItem
        {
            public string Text { get; set; }
            public string Timestamp { get; set; }
        }

        private class ResultDisplayItem
        {
            public string Text { get; set; }
            public string Tag { get; set; }
        }
    }
}