using GazpromTest.Models.GazpromTestEventArgs;
using GazpromTestModels;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
                if (fileName.EndsWith("csv"))
                {
                    newData = ParceCSV(fileName);
                    return;
                }
                else
                {
                    newData = ParceExcel(fileName);
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

        private static List<ObjectInfo> ParceCSV(string fileName)
        {
            List<ObjectInfo> newData = new List<ObjectInfo>();

            List<List<string>> lines = new List<List<string>>();
            using (var reader = new StreamReader(fileName))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var newLine = new List<string>();
                    lines.Add(newLine);
                    var values = line.Split(';');

                    foreach (var value in values)
                    {
                        newLine.Add(value);
                    }
                }
            }
            if (lines.Count == 0)
            {
                return newData;
            }


            List<string> columnNames = new List<string>();

            for (int i = 0; i < lines[0].Count; i++)
            {
                columnNames.Add(lines[0][i].ToLower());
            }
            for (int i = 1; i <= lines.Count; i++)
            {
                var info = new ObjectInfo();
                for (int j = 0; j < lines[i].Count; j++)
                {
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Name))
                    {
                        info.Name = lines[i][j];
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Distance))
                    {
                        info.Distance = double.Parse(lines[i][j].Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Angle))
                    {
                        info.Angle = double.Parse(lines[i][j].Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Width))
                    {
                        info.Width = double.Parse(lines[i][j].Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Heigth) || columnNames[j] == GetColumnName(ColumnNamesEnum.Hegth))
                    {
                        info.Heigth = double.Parse(lines[i][j].Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.IsDefect))
                    {
                        string isDefect = lines[i][j];
                        info.IsDefect = ParceBool(isDefect);
                        continue;
                    }
                }
                newData.Add(info);
            }
            return newData;
        }

        private static List<ObjectInfo> ParceExcel(string fileName)
        {
            List<ObjectInfo> newData = new List<ObjectInfo>();
            IWorkbook workbook;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }
            ISheet sheet = workbook.GetSheetAt(0);

            int noOfColumns = sheet.GetRow(0).LastCellNum;
            List<string> columnNames = new List<string>();
            var currentRow = sheet.GetRow(0);
            for (int i = 0; i < noOfColumns; i++)
            {
                columnNames.Add(currentRow.GetCell(i).StringCellValue.ToLower());
            }
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var info = new ObjectInfo();
                currentRow = sheet.GetRow(i);
                for (int j = 0; j < noOfColumns; j++)
                {
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Name))
                    {
                        info.Name = currentRow.GetCell(j).ToString();
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Distance))
                    {
                        info.Distance = double.Parse(currentRow.GetCell(j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Angle))
                    {
                        info.Angle = double.Parse(currentRow.GetCell(j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Width))
                    {
                        info.Width = double.Parse(currentRow.GetCell(j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.Heigth) || columnNames[j] == GetColumnName(ColumnNamesEnum.Hegth))
                    {
                        info.Heigth = double.Parse(currentRow.GetCell(j).ToString().Replace(',', '.'), CultureInfo.InvariantCulture);
                        continue;
                    }
                    if (columnNames[j] == GetColumnName(ColumnNamesEnum.IsDefect))
                    {
                        string isDefect = currentRow.GetCell(j).ToString();
                        info.IsDefect = ParceBool(isDefect);
                        continue;
                    }
                }
                newData.Add(info);
            }
            return newData;
        }

        public static void SaveFile(string fileName)
        {
            try
            {
                XSSFWorkbook workbook = new XSSFWorkbook();

                fileName = ToXLSX(fileName);
                XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet(GetSheetName(fileName));

                var currentRow = sheet.CreateRow(0);

                var currentCell = currentRow.CreateCell(0);
                currentCell.SetCellValue("Name");
                currentCell = currentRow.CreateCell(1);
                currentCell.SetCellValue("Distance");
                currentCell = currentRow.CreateCell(2);
                currentCell.SetCellValue("Angle");
                currentCell = currentRow.CreateCell(3);
                currentCell.SetCellValue("Width");
                currentCell = currentRow.CreateCell(4);
                currentCell.SetCellValue("Heigth");
                currentCell = currentRow.CreateCell(5);
                currentCell.SetCellValue("IsDefect");

                for (int i = 0; i < Data.Count; i++)
                {
                    currentRow = sheet.CreateRow(i+1);
                    currentCell = currentRow.CreateCell(0);
                    currentCell.SetCellValue(Data[i].Name.ToString());
                    sheet.AutoSizeColumn(0);
                    currentCell = currentRow.CreateCell(1);
                    currentCell.SetCellValue(Data[i].Distance.ToString());
                    sheet.AutoSizeColumn(1);
                    currentCell = currentRow.CreateCell(2);
                    currentCell.SetCellValue(Data[i].Angle.ToString());
                    sheet.AutoSizeColumn(2);
                    currentCell = currentRow.CreateCell(3);
                    currentCell.SetCellValue(Data[i].Width.ToString());
                    sheet.AutoSizeColumn(3);
                    currentCell = currentRow.CreateCell(4);
                    currentCell.SetCellValue(Data[i].Heigth.ToString());
                    sheet.AutoSizeColumn(4);
                    currentCell = currentRow.CreateCell(5);
                    currentCell.SetCellValue(Data[i].IsDefect ? "yes" : "no");
                    sheet.AutoSizeColumn(5);
                }

                if (!File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                //запишем всё в файл
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private static string ToXLSX(string fileName)
        {
            if (fileName.ToLower().EndsWith(".xlsx"))
            {
                return fileName;
            }
            var ind = fileName.LastIndexOf('.');
            fileName = fileName.Substring(0, ind);
            return fileName + ".xlsx";
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
