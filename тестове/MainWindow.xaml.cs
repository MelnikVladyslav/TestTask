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
using тестове.Windows;

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
            dataGrid.Columns.Clear();

            if (_data.Count == 0)
                return; 

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "ID",
                Binding = new Binding("ID")
            });

            var maxInitColumns = _data.Max(d => d.InitColumns.Count);
            var maxCalculatedColumns = _data.Max(d => d.CalculatedColumns.Count);
            var maxUserColumns = _data.Max(d => d.UserColumns.Count);
            var maxRows = _data.Max(d => d.Rows.Count);

            for (int i = 0; i < maxInitColumns; i++)
            {
                dataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = $"InitColumn{i + 1}",
                    Binding = new Binding($"InitColumns[{i}]")
                });
            }

            for (int i = 0; i < maxCalculatedColumns; i++)
            {
                dataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = $"CalculatedColumn{i + 1}",
                    Binding = new Binding($"CalculatedColumns[{i}]")
                });
            }

            for (int i = 0; i < maxUserColumns; i++)
            {
                dataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = $"UserColumn{i + 1}",
                    Binding = new Binding($"UserColumn[{i}]")
                });
            }

            if (maxRows > 0)
            {
                var newRow = new DataModel(_data.Count);
                _data.Add(newRow);
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

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    _data.Clear();
                    int startRow = 2; 

                    for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var dataModel = new DataModel
                        {
                            ID = worksheet.Cells[row, 1].Text
                        };

                        for (int col = 2; col <= 6; col++)
                        {
                            dataModel.InitColumns.Add(worksheet.Cells[row, col].Text);
                        }

                        for (int col = 7; col <= 11; col++)
                        {
                            dataModel.CalculatedColumns.Add(worksheet.Cells[row, col].Text);
                        }

                        for (int col = 12; col <= 16; col++)
                        {
                            dataModel.UserColumns.Add(worksheet.Cells[row, col].Text);
                        }

                        _data.Add(dataModel);
                    }

                    UpdateDataGridColumns();

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

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    worksheet.Cells[1, 1].Value = "ID";

                    for (int i = 0; i < 5; i++)
                    {
                        worksheet.Cells[1, i + 2].Value = $"InitColumn{i + 1}";
                        worksheet.Cells[1, i + 7].Value = $"CalculatedColumn{i + 1}";
                        worksheet.Cells[1, i + 12].Value = $"UserColumns{i + 1}";
                    }

                    for (int row = 0; row < _data.Count; row++)
                    {
                        var data = _data[row];

                        worksheet.Cells[row + 2, 1].Value = data.ID;

                        for (int i = 0; i < data.InitColumns.Count; i++)
                        {
                            worksheet.Cells[row + 2, i + 2].Value = data.InitColumns[i];
                        }

                        for (int i = 0; i < data.CalculatedColumns.Count; i++)
                        {
                            worksheet.Cells[row + 2, i + 7].Value = data.CalculatedColumns[i];
                        }

                        for (int i = 0; i < data.UserColumns.Count; i++)
                        {
                            worksheet.Cells[row + 2, i + 12].Value = data.UserColumns[i];
                        }
                    }

                    package.SaveAs(new FileInfo(filePath));
                }
            }
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            ColumnTypeSelectionWindow columnTypeSelectionWindow = new ColumnTypeSelectionWindow();
            columnTypeSelectionWindow.ShowDialog();

            if (columnTypeSelectionWindow.IsSelected)
            {
                DataModel newRow = new DataModel { ID = (_data.Count + 1).ToString() };

                if (columnTypeSelectionWindow.SelectedColumnType == ColumnTypeSelectionWindow.ColumnType.Init)
                {
                    newRow.InitColumns.Add("New Init Data");
                }
                else if (columnTypeSelectionWindow.SelectedColumnType == ColumnTypeSelectionWindow.ColumnType.Calculated)
                {
                    newRow.CalculatedColumns.Add("New Calculated Data");
                }
                else if (columnTypeSelectionWindow.SelectedColumnType == ColumnTypeSelectionWindow.ColumnType.User)
                {
                    newRow.UserColumns.Add("New User Data");
                }
                else if (columnTypeSelectionWindow.SelectedColumnType == ColumnTypeSelectionWindow.ColumnType.Row)
                {
                    newRow.Rows.Add("New Row");
                }

                _data.Add(newRow);

                UpdateDataGridColumns();
                dataGrid.ItemsSource = _data;
                dataGrid.Items.Refresh();
            }
        }
    }
}
