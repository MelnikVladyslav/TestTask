using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace тестове
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<Dictionary<string, object>> _data = new List<Dictionary<string, object>>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void GenerateColumns()
        {
            dataGrid.Columns.Clear();

            if (_data.Count > 0)
            {
                // Знаходимо всі унікальні ключі, які будуть використовуватись як стовпці
                var columnNames = _data.SelectMany(dict => dict.Keys).Distinct();

                foreach (var columnName in columnNames)
                {
                    // Створюємо новий стовпець для кожного унікального ключа
                    dataGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = columnName,
                        Binding = new System.Windows.Data.Binding($"[{columnName}]")
                    });
                }
            }
        }

        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            // Відкриваємо діалог для вибору файлу
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                string filePath = dlg.FileName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Читаємо дані з Excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    _data.Clear();
                    int startRow = 2; // Пропускаємо заголовок

                    for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var rowData = new Dictionary<string, object>();

                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            string columnName = worksheet.Cells[1, col].Text;
                            string cellValue = worksheet.Cells[row, col].Text;
                            rowData[columnName] = cellValue;
                        }

                        _data.Add(rowData);
                    }

                    GenerateColumns();
                    dataGrid.ItemsSource = _data;
                    dataGrid.Items.Refresh();
                }
            }
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            // Відкриваємо діалог для збереження файлу
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                string filePath = dlg.FileName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Зберігаємо дані у Excel
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    if (_data.Count > 0)
                    {
                        var columnNames = _data.SelectMany(dict => dict.Keys).Distinct().ToList();

                        // Записуємо заголовки
                        for (int col = 0; col < columnNames.Count; col++)
                        {
                            worksheet.Cells[1, col + 1].Value = columnNames[col];
                        }

                        // Записуємо дані
                        for (int row = 0; row < _data.Count; row++)
                        {
                            for (int col = 0; col < columnNames.Count; col++)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = _data[row][columnNames[col]];
                            }
                        }
                    }

                    package.SaveAs(new FileInfo(filePath));
                }
            }
        }
    }
}
