using System;
using OfficeOpenXml;
using System.IO;

class Program
{
    static void Main()
    {
        string filePath = "./test.xlsx";
        string sheetName = "גיליון1";
        string rangeA = "A1:A10";
        string rangeB = "B1:B10";

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

            if (worksheet != null)
            {
                // Read and print values from range A1:A10
                ExcelRange rangeValuesA = worksheet.Cells[rangeA];
                Console.WriteLine("Values in range A1:A10:");
                foreach (var cell in rangeValuesA)
                {
                    Console.WriteLine(cell.Value);
                }

                // Insert numbers 1 to 10 into range B1:B10
                int startingNumber = 11;
                int rowCount = 10;
                int column = 2; // Column B
                for (int row = 1; row <= rowCount; row++)
                {
                    worksheet.Cells[row, column].Value = startingNumber++;
                }

                package.Save();
                Console.WriteLine("Numbers inserted successfully into range B1:B10.");
            }
        }
    }
}
