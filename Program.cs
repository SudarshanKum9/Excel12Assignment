//Change 11
using System;
using ClosedXML.Excel;
using System.IO;
class Program
{
    static void Main(string[] args)
    {
        string inputFile = "InputData.xlsx";
        string outputFile = "Output.xlsx";
        if (!File.Exists(inputFile))
        {
            Console.WriteLine("InputData.xlsx not found. Please place it in the same folder as the .exe file.");
            return;
        }
        using (var workbook = new XLWorkbook(inputFile))
        {
            var worksheet = workbook.Worksheet(1); // sheet1 
            var newWorkbook = new XLWorkbook();
            var newWorksheet = newWorkbook.AddWorksheet("Result");
            newWorksheet.Cell(1, 1).Value = "FirstName";
            newWorksheet.Cell(1, 2).Value = "LastName";
            newWorksheet.Cell(1, 3).Value = "FullName";
            int lastRow = worksheet.LastRowUsed().RowNumber();
            for (int row = 2; row <= lastRow; row++)
            {
                string firstName = worksheet.Cell(row, 1).GetValue<string>();
                string lastName = worksheet.Cell(row, 2).GetValue<string>();
                string fullName = firstName + " " + lastName;
                newWorksheet.Cell(row, 1).Value = firstName;
                newWorksheet.Cell(row, 2).Value = lastName;
                newWorksheet.Cell(row, 3).Value = fullName;
            }
            newWorkbook.SaveAs(outputFile);
            Console.WriteLine("Output.xlsx created with FirstName, LastName, and FullName.");
        }
    }
}
