using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ConsoleApp6
{

    class Copy_sheet
    {

        public string OpenWb(string sourceFileName1, bool revbool, string rev)
        {
            Excel.Application excelApp;
            excelApp = new Excel.Application();
            Excel.Workbook excelworkbook;


            string sourceFileName = sourceFileName1; //Source excel file
            string folderPath = @"c:\xls\";
            string sourceFilePath = System.IO.Path.Combine(folderPath, sourceFileName);
            excelworkbook = excelApp.Workbooks.Open(sourceFilePath, UpdateLinks: 0);
            string revnew = @"c:\xls\diffs\new\";
            string revold = @"c:\xls\diffs\old\";









            foreach (Excel.Worksheet sheet in excelworkbook.Worksheets)
            {
                if (revbool)
                {
                    var newbook = excelApp.Workbooks.Add(1);
                    sheet.Copy(newbook.Sheets[1]);
                    newbook.SaveAs(revnew+sheet.Name + ".xlsx");
                    newbook.Close(true);
                }
                else
                {
                    var newbook = excelApp.Workbooks.Add(1);
                    sheet.Copy(newbook.Sheets[1]);
                    newbook.SaveAs(revold+"1"+sheet.Name + ".xlsx");
                    newbook.Close(true);
                }

            }

            excelworkbook.Close(true);
            excelApp.Quit();


            return rev;

        }





    }
}
