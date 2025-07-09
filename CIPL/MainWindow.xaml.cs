using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using System;
using System.IO;
using System.Linq;
using Windows.Storage.Pickers;
using WinRT.Interop;
using Windows.Graphics;

namespace CIPL
{
    public sealed partial class MainWindow : Window
    {
        private string selectedPath = "";

        public MainWindow()
        {
            this.InitializeComponent();
            IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            var appWindow = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(windowId);

            appWindow.Resize(new Windows.Graphics.SizeInt32(800, 800));

        }

        private async void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".xlsx");

            var hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);

            var file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                selectedPath = file.Path;
                StatusText.Text = $"Selected: {file.Name}";
            }
        }

        private void FilterAndSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedPath))
            {
                StatusText.Text = "Please select a file first.";
                return;
            }

            try
            {
                using var workbook = new XLWorkbook(selectedPath);
                var sheet = workbook.Worksheets.FirstOrDefault(s => s.Name == "CIPL");

                if (sheet == null)
                {
                    StatusText.Text = "Sheet 'CIPL' not found.";
                    return;
                }

                var outputWorkbook = new XLWorkbook();
                var outSheet = outputWorkbook.AddWorksheet("Filtered");

                int inputRow = 13;
                int outputRow = 1;

                while (!string.IsNullOrWhiteSpace(sheet.Cell(inputRow, 17).GetString())) // Q=17
                {
                    int outCol = 1;
                    for (int col = 9; col <= 17; col++) // I to Q
                    {
                        if (col == 10) // J
                        {
                            string formula = $"=I{outputRow}*H{outputRow}";
                            outSheet.Cell(outputRow, outCol).FormulaA1 = formula;
                        }
                        else
                        {
                            outSheet.Cell(outputRow, outCol).Value = sheet.Cell(inputRow, col).Value;
                        }
                        outCol++;
                    }

                    inputRow++;
                    outputRow++;
                }

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string savePath = Path.Combine(Path.GetDirectoryName(selectedPath), $"Filtered_{timestamp}.xlsx");
                outputWorkbook.SaveAs(savePath);

                StatusText.Text = $"Filtered file saved as: Filtered.xlsx";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Error: {ex.Message}";
            }
        }

            private async void MergeFinalTables_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".xlsx");

            var hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);

            var files = await picker.PickMultipleFilesAsync();
            if (files == null || files.Count == 0)
            {
                StatusText.Text = "No files selected.";
                return;
            }

            var outputWorkbook = new XLWorkbook();
            var outputSheet = outputWorkbook.AddWorksheet("Combined");

            int outputRow = 1;

            foreach (var file in files)
            {
                try
                {
                    using var wb = new XLWorkbook(file.Path);
                    var sheet = wb.Worksheets.First(); 

                    int currentRow = 1;
                    while (!sheet.Cell(currentRow, 1).IsEmpty())
                    {
                        int currentCol = 1;
                        while (!sheet.Cell(currentRow, currentCol).IsEmpty() || currentCol <= sheet.LastColumnUsed().ColumnNumber())
                        {
                            outputSheet.Cell(outputRow, currentCol).Value = sheet.Cell(currentRow, currentCol).Value;
                            currentCol++;
                        }
                        currentRow++;
                        outputRow++;
                    }
                }
                catch (Exception ex)
                {
                    StatusText.Text = $"Error in {file.Name}: {ex.Message}";
                }
            }

            // Sum D and F
            double sumD = 0;
            double sumF = 0;

            for (int row = 1; row < outputRow; row++)
            {
                var valD = outputSheet.Cell(row, 4).GetDouble(); // D = 4
                var valF = outputSheet.Cell(row, 6).GetDouble(); // F = 6
                sumD += valD;
                sumF += valF;
            }

            // Add Totals
            outputSheet.Cell(outputRow, 4).Value = sumD;
            outputSheet.Cell(outputRow, 5).Value = "Total";
            outputSheet.Cell(outputRow, 6).Value = sumF;

            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string outputPath = Path.Combine(Path.GetDirectoryName(files[0].Path), $"Combined_{timestamp}.xlsx");
                outputWorkbook.SaveAs(outputPath);
                StatusText.Text = $"Tables combined: {Path.GetFileName(outputPath)}";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Save error: {ex.Message}";
            }
        }

        private async void AttachToTemplate_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".xlsx");

            var hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);

            var file = await picker.PickSingleFileAsync();
            if (file == null)
            {
                StatusText.Text = "Please select merged file.";
                return;
            }

            try
            {
                // Getting a Template
                string templatePath = Path.Combine(AppContext.BaseDirectory, "Assets", "Template.xlsx");
                using var templateWb = new XLWorkbook(templatePath);
                var templateSheet = templateWb.Worksheet("CIPL");

                
                using var mergedWb = new XLWorkbook(file.Path);
                var mergedSheet = mergedWb.Worksheet(1);

                int readRow = 1;
                int writeRow = 13;

                int totalRows = mergedSheet.LastRowUsed().RowNumber(); 

                for (int i = 0; i < totalRows; i++)
                {
                    int col = 1;
                    int templateStyleRow = (i == totalRows - 1) ? 29 : 13;

                    // A–H (1–8)
                    for (int colA = 1; colA <= 8; colA++)
                    {
                        var templateCell = templateSheet.Cell(templateStyleRow, colA);
                        var newCell = templateSheet.Cell(writeRow + i, colA);
                        newCell.Value = ""; 
                        newCell.Style = templateCell.Style;
                    }

                    // I–Q (9–17)
                    int colOffset = 9;
                    while (!mergedSheet.Cell(readRow + i, col).IsEmpty() || col <= mergedSheet.LastColumnUsed().ColumnNumber())
                    {
                        var templateCell = templateSheet.Cell(templateStyleRow, colOffset);
                        var newCell = templateSheet.Cell(writeRow + i, colOffset);

                        newCell.Value = mergedSheet.Cell(readRow + i, col).Value;
                        newCell.Style = templateCell.Style;

                        col++;
                        colOffset++;
                    }
                }

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string savePath = Path.Combine(Path.GetDirectoryName(file.Path), $"FinalCIPL_{timestamp}.xlsx");
                templateWb.SaveAs(savePath);

                StatusText.Text = $"Template created: {Path.GetFileName(savePath)}";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Attach failed: {ex.Message}";
            }
        }
    }
}
