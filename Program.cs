using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace DataRecovery
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.InitialDirectory = @"C:\Users\dimitrios.metozis\Downloads";
            fd.ShowDialog();
            
            if (fd.FileName.EndsWith("xlsx"))
            {
                //Open the workbook (or create it if it doesn't exist)
                var fi = new FileInfo(@fd.FileName);

                using (var p = new ExcelPackage(fi))
                {
                    //Get the Worksheet created in the previous codesample. 
                    var ws = p.Workbook.Worksheets["Sheet1"];
                    //Set the cell value using row and column.
                    //The style object is used to access most cells formatting and styles.
                    ws.Cells[2, 1].Style.Font.Bold = true;
                    //Save and close the package.
                    p.Save();
                }
            }
            Console.ReadLine();

        }
    }
}
