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
        public void ReadExcelDataList(string filePath, int sheetSelected)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetSelected]; //  시트 선택

                int colCount = worksheet.Dimension.Columns; // 가로줄의 개수
                int rowCount = worksheet.Dimension.Rows; // 세로줄의 개수

                Console.WriteLine(colCount + " <<<<< ColCount입니다.");
                Console.WriteLine(rowCount + " <<<<< RowCount입니다.");
                for (int col = 1; col <= colCount; col++)
                {
                    string colName = worksheet.Cells[1, col].Text;
                    
                    List<string> columnData = new List<string>();

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text;
                        Console.WriteLine(cellValue + "<<<<<<< cellValue 입니다");
                        
                        columnData.Add(cellValue);
                        
                    }
                        Console.WriteLine("columnData.Count :::::: " + columnData.Count);
                }
                
            }
        }
    }
}
