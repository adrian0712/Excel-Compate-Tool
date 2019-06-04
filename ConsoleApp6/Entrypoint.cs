using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Diagnostics;
using System.Windows.Forms;

namespace ConsoleApp6
{
    public class Entrypoint
    {
        public static string cellcontent;
        public static List<string> changelist = new List<string>();

        public void Start(string xlsnew, string xlsold)
        {
            Copy_sheet open = new Copy_sheet();
            Formula formulapop = new Formula();
            string runcopynew = open.OpenWb(xlsnew, true, xlsnew);
            string runcopyold = open.OpenWb(xlsold, false, xlsold);
            CopySheet();

            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
        }

        void CopySheet()
        {
            Excel.Application excelApp;
            string path1 = @"c:\xls\diffs\new\";
            string path2 = @"c:\xls\diffs\old\";
            string[] filesin1 = Directory.GetFiles(path1);
            string[] filesin2 = Directory.GetFiles(path2);

            foreach (string f in filesin1)
            {
                foreach (string f2 in filesin2)
                {
                    string subf = f.Substring(17);
                    string subf2 = f2.Substring(18);
                    if (subf.Equals(subf2))
                    {
                        form();
                    }
                    //if (f2.Contains("Switchboard") && f.Contains("Switchboard"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("IP_ADDRESS") && f.Contains("IP_ADDRESS"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("DB") && f.Contains("DB"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("HW_CONFIG") && f.Contains("HW_CONFIG"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("SERIAL_PORTS") && f.Contains("SERIAL_PORTS"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("VERSION_INFO") && f.Contains("VERSION_INFO"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("HW_OVERVIEW") && f.Contains("HW_OVERVIEW"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Fire Alarm") && f.Contains("Fire Alarm"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Alarm_Data") && f.Contains("Alarm_Data"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Station Fire Main ") && f.Contains("Station Fire Main "))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Room") && f.Contains("Room"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Escalator") && f.Contains("Escalator"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Fire Panel") && f.Contains("Fire Panel"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("UPS") && f.Contains("UPS"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("HVAC") && f.Contains("HVAC"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Lift") && f.Contains("Lift"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Tunnel Fire Main") && f.Contains("Tunnel Fire Main"))
                    //{
                    //    form();
                    //}
                    //if (f2.Contains("Pump") && f.Contains("Pump"))
                    //{
                    //    form();
                    //}
                    //else
                    //{
                    //    //Console.WriteLine("No match");
                    //}
                    //if (f2.Contains("Lighting") && f.Contains("Lighting"))
                    //{
                    //    form();
                    //}
                    //else
                    //{
                    //    //Console.WriteLine("No match");
                    //}
                    //if (f2.Contains("Water Meter") && f.Contains("Water Meter"))
                    //{
                    //    form();
                    //}
                    //else
                    //{
                    //    //Console.WriteLine("No match");
                    //}
                    void form()
                    {
                        excelApp = new Excel.Application();
                        Excel.Workbook excelworkbook1;
                        Excel.Workbook excelworkbook2;
                        excelworkbook1 = excelApp.Workbooks.Open(f, UpdateLinks: 0);
                        excelworkbook2 = excelApp.Workbooks.Open(f2);
                        Excel.Worksheet excelWorksheet2 = excelworkbook2.Worksheets[1];
                        excelWorksheet2.Copy(Type.Missing, After: excelworkbook1.Worksheets[1]);
                        excelworkbook2.Save();
                        Excel.Worksheet excelWorksheet = excelworkbook1.Worksheets[3];
                        Excel.Worksheet excelWorksheet1 = excelworkbook1.Worksheets[1];
                        string SheetFormula = excelWorksheet1.Name;
                        string diffs = "Changed";
                        string formula = "=IF('" + SheetFormula + "'!A1<>'" + SheetFormula + " (2)'!A1,\"" + diffs + "\",1)";
                        string formulacountifBE = "=COUNTIF(A1:CE700,\"" + diffs + "\")";
                        string formulacountifQ = "=COUNTIF(A1:Q700,\"" + diffs + "\")";
                        Excel.Range rng = excelWorksheet.Range["A1"];
                        Excel.Range rngBE = excelWorksheet.Range["CF1"];
                        Excel.Range rngQ = excelWorksheet.Range["CF2"];
                        Excel.Range rngdown = excelWorksheet.Range["A1:A700"];
                        Excel.Range destdown = excelWorksheet.Range["A1:A700"];
                        Excel.Range destright = excelWorksheet.Range["A1:CE700"];
                        rng.Formula = formula;
                        rngBE.Formula = formulacountifBE;
                        rngQ.Formula = formulacountifQ;
                        destdown.FillDown();
                        destright.FillRight();
                        excelWorksheet.Calculate();
                        cellcontent = "Worksheet name: " + excelWorksheet2.Name + " changes on whole: " + excelWorksheet.Range["CF1"].Value + " changes on SCADA: " + excelWorksheet.Range["CF2"].Value;                      
                        changelist.Add(cellcontent);
                        excelworkbook2.Save();
                        excelworkbook2.Close(true);
                        excelworkbook1.Close(true);
                        excelApp.Quit();


                    }

                }
            }
        }
    }
}