using Data.Models;
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
        private List<DataModel> _data = new List<DataModel>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void UpdateDataGridColumns()
        {
            // Очищаємо поточні стовпці
            dataGrid.Columns.Clear();

            if (_data.Count == 0)
                return; // Якщо немає даних, не будуємо стовпці

            // Додаємо стовпець для ID
            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "ID",
                Binding = new Binding("ID")
            });

            // Визначаємо кількість стовпців для ініціалізаційних та розрахункових даних
            var maxInitColumns = _data.Max(d => d.InitColumns.Count);
            var maxCalculatedColumns = _data.Max(d => d.CalculatedColumns.Count);

            // Додаємо стовпці для ініціалізаційних даних
            for (int i = 0; i < maxInitColumns; i++)
            {
                dataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = $"InitColumn{i + 1}",
                    Binding = new Binding($"InitColumns[{i}]")
                });
            }

            // Додаємо стовпці для розрахункових даних
            for (int i = 0; i < maxCalculatedColumns; i++)
            {
                dataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = $"CalculatedColumn{i + 1}",
                    Binding = new Binding($"CalculatedColumns[{i}]")
                });
            }
        }

        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                string filePath = dlg.FileName;

                // Встановлення контексту ліцензії EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    _data.Clear();
                    int startRow = 2; // Перший рядок — це заголовки

                    for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var dataModel = new DataModel
                        {
                            ID = worksheet.Cells[row, 1].Text
                        };

                        // Зчитуємо стовпці ініціалізації
                        for (int col = 2; col <= 6; col++)
                        {
                            dataModel.InitColumns.Add(worksheet.Cells[row, col].Text);
                        }

                        // Зчитуємо розрахункові стовпці
                        for (int col = 7; col <= 11; col++)
                        {
                            dataModel.CalculatedColumns.Add(worksheet.Cells[row, col].Text);
                        }

                        _data.Add(dataModel);
                    }

                    // Оновлюємо стовпці DataGrid
                    UpdateDataGridColumns();

                    // Оновлюємо джерело даних для DataGrid
                    dataGrid.ItemsSource = _data;
                    dataGrid.Items.Refresh();
                }
            }
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                string filePath = dlg.FileName;

                // Встановлення контексту ліцензії EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // Заповнюємо заголовки
                    worksheet.Cells[1, 1].Value = "ID";

                    // Заповнення заголовків для ініціалізаційних і розрахункових полів
                    for (int i = 0; i < 5; i++)
                    {
                        worksheet.Cells[1, i + 2].Value = $"InitColumn{i + 1}";
                        worksheet.Cells[1, i + 7].Value = $"CalculatedColumn{i + 1}";
                    }

                    // Заповнюємо дані
                    for (int row = 0; row < _data.Count; row++)
                    {
                        var data = _data[row];

                        worksheet.Cells[row + 2, 1].Value = data.ID;

                        // Заповнюємо динамічні ініціалізаційні стовпці
                        for (int i = 0; i < data.InitColumns.Count; i++)
                        {
                            worksheet.Cells[row + 2, i + 2].Value = data.InitColumns[i];
                        }

                        // Заповнюємо динамічні розрахункові стовпці
                        for (int i = 0; i < data.CalculatedColumns.Count; i++)
                        {
                            worksheet.Cells[row + 2, i + 7].Value = data.CalculatedColumns[i];
                        }
                    }

                    package.SaveAs(new FileInfo(filePath));
                }
            }
        }
    }
}
