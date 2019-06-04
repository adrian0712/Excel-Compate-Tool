using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp6
{
    class Formula
    {

        public void FormulaPop()
        {
            string destfolder = @"c:\xls\diffs\";
            Excel.Application excelApp;
            excelApp = new Excel.Application();
            Excel.Workbook excelworkbook;
            Directory.EnumerateFiles(destfolder, "*.xlsx", SearchOption.AllDirectories)
                .ToList()
                .ForEach(f => Form(f));

            void Form(string excelFilePath)
            {
                try
                {
                    excelworkbook = excelApp.Workbooks.Open(excelFilePath, UpdateLinks: 0);
                    Excel.Worksheet excelWorksheet = excelworkbook.Worksheets[3];
                    Excel.Worksheet excelWorksheet1 = excelworkbook.Worksheets[1];

                    string SheetFormula = excelWorksheet1.Name;
                    string formula = "=IF('" + SheetFormula + "'!A1<>'" + SheetFormula + " (2)'!A1, 0,1)";
                    Excel.Range rng = excelWorksheet.Range["A1"];
                    Excel.Range rngdown = excelWorksheet.Range["A1:A200"];
                    Excel.Range destdown = excelWorksheet.Range["A1:A200"];
                    Excel.Range destright = excelWorksheet.Range["A1:BE200"];
                    rng.Formula = formula;
                    destdown.FillDown();
                    destright.FillRight();
                    excelworkbook.Save();
                    Console.WriteLine(formula);
                    excelWorksheet.Calculate();
                    excelworkbook.Close(true);
                    excelApp.Quit();


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }

        }
    }
}
