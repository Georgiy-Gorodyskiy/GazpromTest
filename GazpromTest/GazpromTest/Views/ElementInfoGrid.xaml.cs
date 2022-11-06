using GazpromTest.Models;
using GazpromTestModels;
using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace GazpromTest.Views
{
    /// <summary>
    /// Логика взаимодействия для ElementInfoGrid.xaml
    /// </summary>
    public partial class ElementInfoGrid : Grid
    {
        public ElementInfoGrid()
        {
            InitializeComponent();
            ServiceManager.ExcelService.CurentInfoChanged += ExcelService_CurentInfoChanged;
            ServiceManager.ExcelService.DataChanged += ExcelService_DataChanged;
        }

        private void ExcelService_DataChanged(object? sender, EventArgs e)
        {
            TextBlock.Text = "";
        }

        private void ExcelService_CurentInfoChanged(object? sender, Models.GazpromTestEventArgs.CurentInfoChangedEventArgs e)
        {
            TextBlock.Text = GetInfoString(e.Info);
        }
        private string GetInfoString(ObjectInfo info)
        {
            try
            {
                if (info == null)
                    return "";
                StringBuilder result = new StringBuilder("");
                result.Append("Name = ").Append(info.Name);
                result.Append(Environment.NewLine);
                result.Append("Angle = ").Append(info.Angle);
                result.Append(Environment.NewLine);
                result.Append("Distance = ").Append(info.Distance);
                result.Append(Environment.NewLine);
                result.Append("Width = ").Append(info.Width);
                result.Append(Environment.NewLine);
                result.Append("Heigth = ").Append(info.Heigth);
                result.Append(Environment.NewLine);
                result.Append("Is defect = ").Append(info.IsDefect ? "Yes" : "No");
                result.Append(Environment.NewLine);


                return result.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "";
            }
        }
    }
}
