
using DocumentFormat.OpenXml.Office.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace Excel_Macro_Detection
{
    class Program
    {
        static void Main(string[] args)
        {

            Application excelApp = new Application();
            Workbook workbook = null;

            try
            {
                Console.WriteLine("Please enter the path to the Excel file:");
                string filePath = Console.ReadLine();

                if (string.IsNullOrEmpty(filePath))
                {
                    Console.WriteLine("Error: File path cannot be empty.");
                    return;
                }

                if (File.Exists(filePath))
                {
                    Console.WriteLine("Error: File does not exist.");
                    return;
                }

                workbook = excelApp.Workbooks.Open(filePath);

                bool hasMacro = workbook.HasVBProject;
                Console.WriteLine($"Does the file contain macros? {(hasMacro ? "Yes" : "No")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                }
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
