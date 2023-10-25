using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.IO;

namespace PrintCommand
{
    /// <summary>
    /// ReadExcelData Command
    /// </summary>
    public class ReadExcelData
    {
        public string wrokSheetName = string.Empty;
        public ObservableCollection<ObservableCollection<string>> excelTotalData = new ObservableCollection<ObservableCollection<string>>();
        /// <summary>
        /// 지정된 경로의 엑셀 파일을 읽어옵니다.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetSelected"></param>
        public ObservableCollection<ObservableCollection<string>> ReadExcelDataList(string filePath, int sheetSelected)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetSelected]; // 시트 선택
                wrokSheetName = worksheet.Name;
                int colCount = worksheet.Dimension.Columns; // 가로줄의 개수
                int rowCount = worksheet.Dimension.Rows; // 세로줄의 개수

                //List<List<string>> excelTotalData = new List<List<string>>();

                for (int col = 1; col <= colCount; col++)
                {
                    ObservableCollection<string> columnData = new ObservableCollection<string>();

                    for (int row = 1; row <= rowCount; row++) // 열 제목도 데이터로 포함시키기 위해 1부터 시작
                    {
                        string cellValue = worksheet.Cells[row, col].Text;
                        columnData.Add(cellValue);
                    }

                    excelTotalData.Add(columnData);
                }

                return excelTotalData;
            }
        }

        /// <summary>
        /// 엑셀 파일명을 입력하면 , 해당 데이터 경로에 원하는 엑셀을 불러옵니다.
        /// </summary>
        /// <param name="fileName">엑셀 파일명을 입력하세요 </param>
        /// <returns></returns>
        public string GetRecipeFile(string fileName)
        {
            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            string dataFilePath = Path.Combine(exePath, "Data");
            string FILEPATH = Path.Combine(dataFilePath, fileName);

            return FILEPATH;
        }
        
    }
}
