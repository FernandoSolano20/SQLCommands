using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace SQL_Commands
{
    internal class Program
    {
        public static string filePath = @"C:\Users\fersolano\source\repos\SQL Commands\SQL Commands\Sheet1.xlsx";

        private static void Main(string[] args)
        {
            var values = ReadExcelFile();
            for (var i = 1; i < values.Count; i++) InsertInto("movie", values[0], values[i]);

            Console.ReadLine();
        }

        private static IList<List<string>> ReadExcelFile()
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine("\nReading the Excel File...");
            Console.BackgroundColor = ConsoleColor.Black;


            var xlApp = new Application();
            var xlWorkBook = xlApp.Workbooks.Open(filePath, ReadOnly: false);
            var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            var xlRange = xlWorkSheet.UsedRange;
            var totalRows = xlRange.Rows.Count;
            var totalColumns = xlRange.Columns.Count;

            IList<List<string>> sheetReaded = new List<List<string>>();

            for (var rowCount = 1; rowCount <= totalRows; rowCount++)
            {
                var row = new List<string>();
                for (var j = 1; j <= totalColumns; j++)
                    row.Add(Convert.ToString((xlRange.Cells[rowCount, j] as Range).Text));

                sheetReaded.Add(row);
            }

            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("End of the file...");

            return sheetReaded;
        }

        private static void InsertInto(string table, IList<string> config, IList<string> values)
        {
            var insertCommand = "INSERT INTO " + table + " VALUES (";

            for (var i = 0; i < config.Count; i++)
            {
                if (config[i].ToLower() == "int")
                {
                    insertCommand += values[i];
                }
                else if (config[i].ToLower() == "date")
                {
                    var dateSplit = values[i].Split('/');
                    var dateValueFormat = dateSplit[2] + "-" + dateSplit[0] + "-" + dateSplit[1];
                    insertCommand += "'" + dateValueFormat + "'";
                }
                else
                {
                    insertCommand += "'" + values[i] + "'";
                }

                if (i + 1 == config.Count)
                    insertCommand += ");";
                else
                    insertCommand += ", ";
            }

            Console.WriteLine(insertCommand);
        }
    }
}