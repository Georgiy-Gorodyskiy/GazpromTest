using GazpromTest.Models.GazpromTestEventArgs;
using GazpromTestModels;
using IronXL;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace GazpromTest.Models
{
    public static class ExcelService
    {
        public static List<ObjectInfo> Data { get; private set; } = new List<ObjectInfo>();

        public static event EventHandler DataChanged;
        public static event EventHandler<CurentInfoChangedEventArgs> CurentInfoChanged;

        public static void ChangeCurentInfo(ObjectInfo info)
        {
            try
            {
                CurentInfoChanged?.Invoke(null, new CurentInfoChangedEventArgs(info));
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        public static void ReadFile(string fileName)
        {
            try
            {
                List<ObjectInfo> newData = new List<ObjectInfo>();

                WorkBook workbook = WorkBook.Load(fileName);
                WorkSheet sheet = workbook.WorkSheets.First();

                List<string> columnNames = new List<string>();
                for (int i = 0; i < sheet.Columns.Count(); i++)
                {
                    columnNames.Add(sheet.GetCellAt(0, i).ToString().ToLower());
                }
                for (int i = 1; i < sheet.Rows.Count(); i++)
                {
                    var info = new ObjectInfo();
                    for (int j = 0; j < sheet.Columns.Count(); j++)
                    {
                        if (columnNames[j] == GetColumnName(ColumnNamesEnum.Name))
                        {
                            info.Name = sheet.GetCellAt(i, j).ToString();
                        }
                        if (columnNames[j] == GetColumnName(ColumnNamesEnum.Distance))
                        {
                            info.Distance = double.Parse(sheet.GetCellAt(i, j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        }
                        if (columnNames[j] == GetColumnName(ColumnNamesEnum.Angle))
                        {
                            info.Angle = double.Parse(sheet.GetCellAt(i, j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        }
                        if (columnNames[j] == GetColumnName(ColumnNamesEnum.Width))
                        {
                            info.Width = double.Parse(sheet.GetCellAt(i, j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        }
                        if (columnNames[j] == GetColumnName(ColumnNamesEnum.Heigth) || columnNames[j] == GetColumnName(ColumnNamesEnum.Hegth))
                        {
                            info.Heigth = double.Parse(sheet.GetCellAt(i, j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        }
                        if (columnNames[j] == GetColumnName(ColumnNamesEnum.IsDefect))
                        {
                            string isDefect = sheet.GetCellAt(i, j).ToString();
                            info.IsDefect = ParceBool(isDefect);
                        }
                    }
                    newData.Add(info);
                }
                Data.Clear();
                Data = newData;
                DataChanged.Invoke(null, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        public static void SaveFile(string fileName)
        {
            try
            {
                WorkBook workbook;
                if (fileName.ToLower().EndsWith(".xlsx"))
                {
                    workbook = WorkBook.Create(ExcelFileFormat.XLSX);
                }
                else
                {
                    if (fileName.ToLower().EndsWith(".xls"))
                        workbook = WorkBook.Create(ExcelFileFormat.XLS);
                    else
                        throw new Exception("Wrong file format: " + fileName);
                }

                WorkSheet sheet = workbook.CreateWorkSheet(GetSheetName(fileName));

                sheet["A1"].Value = "Name";
                sheet["B1"].Value = "Distance";
                sheet["C1"].Value = "Angle";
                sheet["D1"].Value = "Width";
                sheet["E1"].Value = "Heigth";
                sheet["F1"].Value = "IsDefect";

                for (int i = 0; i < Data.Count; i++)
                {
                    sheet["A" + (i + 2).ToString().ToUpper()].Value = Data[i].Name.ToString();
                    sheet["B" + (i + 2).ToString().ToUpper()].Value = Data[i].Distance.ToString();
                    sheet["C" + (i + 2).ToString().ToUpper()].Value = Data[i].Angle.ToString();
                    sheet["D" + (i + 2).ToString().ToUpper()].Value = Data[i].Width.ToString();
                    sheet["E" + (i + 2).ToString().ToUpper()].Value = Data[i].Heigth.ToString();
                    sheet["F" + (i + 2).ToString().ToUpper()].Value = Data[i].IsDefect ? "yes" : "no";
                }

                workbook.SaveAs(fileName);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        public static void CloseFile()
        {
            try
            {
                Data.Clear();
                DataChanged.Invoke(null, EventArgs.Empty);
            }
            catch(Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private static string GetSheetName(string fileName)
        {
            var tmp = fileName.Split("\\").LastOrDefault().Split('.');
            return tmp[0];
        }

        private static bool ParceBool(string isDefect)
        {
            if(isDefect.ToLower() == "yes")
                return true;
            if (isDefect.ToLower() == "no")
                return false;
            throw new Exception("Wrong IsDefectValue: " + isDefect);
        }

        private static string GetColumnName(ColumnNamesEnum name)
        {
            return name.ToString().ToLower();
        }
    }
}
