using GazpromTest.Models;
using GazpromTestModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
namespace GazpromTest.Views
{
    /// <summary>
    /// Логика взаимодействия для TableGrid.xaml
    /// </summary>
    public partial class TableGrid : Grid
    {
        public TableGrid()
        {
            InitializeComponent();
            ServiceManager.ExcelService.DataChanged += ExcelService_DataChanged;
            DataGrid.ItemsSource = ServiceManager.ExcelService.Data;
        }

        private void ExcelService_DataChanged(object? sender, EventArgs e)
        {
            try
            {
                DataGrid.ItemsSource = null;
                DataGrid.ItemsSource = ServiceManager.ExcelService.Data;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (e.AddedItems.Count > 0)
                {
                    ServiceManager.ExcelService.ChangeCurentInfo(e.AddedItems[0] as ObjectInfo);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.Column.Header.ToString() == "IsDefect")
            {
                e.Column.Header = "Is defect";
            }
        }

    }
}
