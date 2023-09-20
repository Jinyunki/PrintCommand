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
        
        /// <summary>
        /// test
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetSelected"></param>
        public List<string> ReadExcelDataList(string filePath, int sheetSelected)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetSelected]; //  시트 선택

                int colCount = worksheet.Dimension.Columns; // 가로줄의 개수
                int rowCount = worksheet.Dimension.Rows; // 세로줄의 개수
                
                List<string> columnData = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    string colName = worksheet.Cells[1, col].Text;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text;
                        Console.WriteLine(cellValue + "<<<<<<< cellValue 입니다");
                        
                        columnData.Add(cellValue);
                    }
                }
                return columnData;
            }
        }
    }
}
