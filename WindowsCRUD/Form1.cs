using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Microsoft.Office.Interop;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;

namespace WindowsCRUD
{
    public partial class Form1 : Form
    {
        public static string filePath = Application.StartupPath+ "\\DataBase\\RK_Excel.xlsx";
        public Form1()
        {
            InitializeComponent();
            CreateExcelFile();
        }

        private static void CreateExcelFile()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not installed in this system!");
                return;
            }

            if (!File.Exists(Application.StartupPath+"\\DataBase"))
            {
                Directory.CreateDirectory(Application.StartupPath+ "\\DataBase");
            }

            if (!File.Exists(filePath))
            {
                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "ID";
                xlWorkSheet.Cells[1, 2] = "Name";
                xlWorkSheet.Cells[2, 1] = "1001";
                xlWorkSheet.Cells[2, 2] = "Ramakrishna";
                xlWorkSheet.Cells[3, 1] = "1002";
                xlWorkSheet.Cells[3, 2] = "Praveenkumar";

                xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }


        private void AddNewRowsToExcelFile()
        {
            IList<Employee> empList = new List<Employee>() {
            new Employee(){ ID=1003, Name="Indraneel"},
            new Employee(){ ID=1004, Name="Neelohith"},
            new Employee(){ ID=1005, Name="Virat"}
            };

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            int rowNumber = xlRange.Rows.Count + 1;

            foreach (Employee emp in empList)
            {
                xlWorkSheet.Cells[rowNumber, 1] = emp.ID;
                xlWorkSheet.Cells[rowNumber, 2] = emp.Name;
                rowNumber++;
            }

            // Disable file override confirmaton message  
            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            AddNewRowsToExcelFile();
        }

        private static void DeleteRowCellFromExcelFile()
        {

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range range1 = xlWorkSheet.get_Range("A2", "B2");

            // To Delete Entire Row - below rows will shift up  
            range1.EntireRow.Delete(Type.Missing);

            Microsoft.Office.Interop.Excel.Range range2 = xlWorkSheet.get_Range("B3", "B3");
            range2.Cells.Clear();

            // To Delete Cells - Below cells will shift up  
            // range2.Cells.Delete(Type.Missing);  

            // Disable file override confirmaton message  
            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DeleteRowCellFromExcelFile();
        }

        private static void ReadExcelFile()
        {

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            int totalRows = xlRange.Rows.Count;
            int totalColumns = xlRange.Columns.Count;

            string firstValue, secondValue;

            for (int rowCount = 1; rowCount <= totalRows; rowCount++)
            {

                firstValue = Convert.ToString((xlRange.Cells[rowCount, 1] as Microsoft.Office.Interop.Excel.Range).Text);
                secondValue = Convert.ToString((xlRange.Cells[rowCount, 2] as Microsoft.Office.Interop.Excel.Range).Text);

                Console.WriteLine(firstValue + "\t" + secondValue);

            }

            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("End of the file...");
        }
    }

    public class Employee
    {
        public int ID { get; set; }
        public string Name { get; set; }
    }
}
