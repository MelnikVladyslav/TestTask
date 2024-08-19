using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace тестове.Windows
{
    /// <summary>
    /// Interaction logic for ColumnTypeSelectionWindow.xaml
    /// </summary>
    public partial class ColumnTypeSelectionWindow : Window
    {
        public enum ColumnType { Init, Calculated, User, Row }
        public ColumnType SelectedColumnType { get; private set; }
        public bool IsSelected { get; private set; }

        public ColumnTypeSelectionWindow()
        {
            InitializeComponent();
            IsSelected = false;
        }

        private void AddInitColumn_Click(object sender, RoutedEventArgs e)
        {
            SelectedColumnType = ColumnType.Init;
            IsSelected = true;
            this.Close();
        }

        private void AddCalculatedColumn_Click(object sender, RoutedEventArgs e)
        {
            SelectedColumnType = ColumnType.Calculated;
            IsSelected = true;
            this.Close();
        }

        private void AddUserColumn_Click(object sender, RoutedEventArgs e)
        {
            SelectedColumnType = ColumnType.User;
            IsSelected = true;
            this.Close();
        }

        private void AddRow(object sender, RoutedEventArgs e)
        {
            SelectedColumnType = ColumnType.Row;
            IsSelected = true;
            this.Close();
        }
    }
}
