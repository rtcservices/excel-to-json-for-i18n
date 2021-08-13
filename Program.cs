using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateJsonFile();            
        }

        private static void CreateJsonFile()
        {
            try
            {
                string source = string.Empty;
                string destination = string.Empty;

                Console.WriteLine("Enter the file path.? \n\n[FORMAT] :- C:\\eng.xlsx\n");
                source = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(source))
                {
                    Console.WriteLine("No Records found.!");
                    return;
                }

                Console.WriteLine("Enter the JSON file name.? \n\n[FORMAT] :-  C:\\en.json\n");
                destination = Console.ReadLine();

                Application excelApp = new();

                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not installed!!");
                    return;
                }

                Workbook excelBook = excelApp.Workbooks.Open($@"{source}");
                _Worksheet excelSheet = excelBook.Sheets[1];

                Range excelRange = excelSheet.UsedRange;

                int rowCount = excelRange.Rows.Count;

                if (rowCount > 0)
                {
                    var items = new Dictionary<string, string>();

                    for (int i = 2; i <= rowCount; i++)
                    {
                        var key = (string)(excelRange.Cells[i, 1] as Range).Value2;
                        var value = (string)(excelRange.Cells[i, 2] as Range).Value2;

                        if (!string.IsNullOrWhiteSpace(key) || !string.IsNullOrWhiteSpace(value))
                            items.Add(key, value);
                    }

                    var jsonString = JsonConvert.SerializeObject(items, Formatting.Indented);

                    File.WriteAllText(destination, jsonString);

                    Console.WriteLine($"\n \n ====================================== \n\n JSON file created in {destination} \n \n ======================================");
                }

                excelBook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                var error = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                Console.WriteLine($"\n \n ***************************************************************** \n\n  Error :- {error}\n \n *****************************************************************");
            }

            Console.WriteLine("\n\nDo you want to covert another file ? Y/N");
            if (Console.ReadLine() == "Y")
            {
                CreateJsonFile();
            }
        }
    }
}
