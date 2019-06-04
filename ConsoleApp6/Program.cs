using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Diagnostics;
using System.Windows.Forms;

namespace ConsoleApp6
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(@"c:\xls\diffs\new\");
            System.IO.DirectoryInfo di2 = new DirectoryInfo(@"c:\xls\diffs\old\");

            Delete(di);
            Delete(di2);
            
            void Delete(System.IO.DirectoryInfo path)
            {

                foreach (FileInfo file in path.GetFiles())
                {
                    file.Delete();
                }
                foreach (DirectoryInfo dir in path.GetDirectories())
                {
                    dir.Delete(true);
                }
            }

            Application.EnableVisualStyles();
            Application.Run(new Form1()); // or whatever
        }
    }
}

