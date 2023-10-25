using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using OfficeOpenXml;
using System;
using System.Collections.ObjectModel;
using System.IO;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace PrintCommand
{
    public class ExcelReadRecive
    {
        public ObservableCollection<ObservableCollection<string>> CallingBackDataList = new ObservableCollection<ObservableCollection<string>>();
        public ObservableCollection<ObservableCollection<string>> CallingBackData(string path, int selectedSheet)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            string dataFilePath = Path.Combine(exePath, "Data");
            string FILEPATH = Path.Combine(dataFilePath, path);

            string wrokSheetName;

            using (var package = new ExcelPackage(new FileInfo(FILEPATH)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets[selectedSheet]; // 시트 선택
                wrokSheetName = worksheet.Name;
                int colCount = worksheet.Dimension.Columns; // 가로줄의 개수
                int rowCount = worksheet.Dimension.Rows; // 세로줄의 개수

                for (int col = 1; col <= colCount; col++)
                {
                    ObservableCollection<string> columnData = new ObservableCollection<string>();

                    for (int row = 1; row <= rowCount; row++) // 열 제목도 데이터로 포함시키기 위해 1부터 시작
                    {
                        string cellValue = worksheet.Cells[row, col].Text;
                        columnData.Add(cellValue);
                    }

                    CallingBackDataList.Add(columnData);
                }
            }
            return CallingBackDataList;
        }
    }
}
