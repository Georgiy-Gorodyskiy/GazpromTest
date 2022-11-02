using System;
using System.Windows;
using System.Windows.Controls;
using GazpromTest.Models;
using Microsoft.Win32;

namespace GazpromTest.Views
{
    /// <summary>
    /// Логика взаимодействия для MenuGrid.xaml
    /// </summary>
    public partial class MenuGrid : Grid
    {
        public MenuGrid()
        {
            InitializeComponent();
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlss;*.csv";
                if (openFileDialog.ShowDialog() == true)
                    ExcelService.ReadFile(openFileDialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SaveFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlss;";
                if (saveFileDialog.ShowDialog() == true)
                    ExcelService.SaveFile(saveFileDialog.FileName);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CloseFileButton_Click(object sender, RoutedEventArgs e)
        {
            ExcelService.CloseFile();
        }
    }
}
