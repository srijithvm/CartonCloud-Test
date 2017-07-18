using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace CartonCloud_Test
{
    class Program
    {
        static void Main(string[] args)
        {
            String outputFileContent = String.Empty;
            string inputPath = @"d:\parseit-input.xls";
            string outputPath = @"d:\output.txt";
            
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(inputPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            outputFileContent = String.Concat(outputFileContent.ToString(), "array(\r\n\t'customer_reference' => '", xlRange[5,3].Value2.ToString(), " - ",xlRange[3,9].Value2.ToString(), "', \r\n\t'records' => array(");
            //Console.WriteLine(xlRange.Cells[5, 3].Value2.ToString());
            for (int i = 9; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                {
                    outputFileContent = String.Concat(outputFileContent.ToString(),
                        "\r\n\t\tarray(\r\n\t\t\t'customer_reference' => '", xlRange.Cells[i, 5].Value2.ToString(), "',\r\n",
                        "\t\t\t'code' => '", xlRange.Cells[i, 1].Value2.ToString(), "',\r\n",
                        "\t\t\t'address_company_name' => '", xlRange.Cells[i, 2].Value2.ToString(), "',\r\n",
                        "\t\t\t'street_address' => '',\r\n",
                        "\t\t\t'suburb' => '", xlRange.Cells[i, 3].Value2.ToString(), "',\r\n",
                        "\t\t\t'state' => 'SA',\r\n",
                        "\t\t\t'postcode' => '',\r\n",
                        "\t\t\t'cartons' => (int)", xlRange.Cells[i, 7].Value2.ToString(), ",\r\n",
                        "\t\t\t'value' => (float)", xlRange.Cells[i, 6].Value2.ToString(), ",\r\n",
                        "\t\t\t'address_string' => \"", xlRange.Cells[i, 1].Value2.ToString() + "," + xlRange.Cells[i, 2].Value2.ToString() + "," + xlRange.Cells[i, 3].Value2.ToString(), "\",\r\n",
                        "\t\t\t'special_instructions' => '", xlRange.Cells[i, 8].Value2.ToString(), "'\r\n\t\t),");
                }
            }
            //Console.ReadLine();
            outputFileContent = String.Concat(outputFileContent, "\r\n\t)\r\n)");

            File.WriteAllText(outputPath, outputFileContent.ToString());

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close(false);
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
