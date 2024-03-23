using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace ExcelGenerator
{
    class MyExcel
    {
        public static string DB_PATH = @"";
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        private static int lastRow = 0;
        private static List<string> serialNumbers = new List<string>();
        public static void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(DB_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explict cast is not required here
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }
        public static void ReadMyExcel()
        {
            for (int index = 2; index <= lastRow; index++)
            {
                serialNumbers.Add(MySheet.get_Range("A" + index.ToString()).Cells.Value);
            }
        }
        public static void WriteToExcel()
        {
            try
            {
                lastRow += 1;
                //MySheet.Cells[lastRow, 1] = emp.Name;
                //MySheet.Cells[lastRow, 2] = emp.Employee_ID;
                //MySheet.Cells[lastRow, 3] = emp.Email_ID;
                //MySheet.Cells[lastRow, 4] = emp.Number;
                //EmpList.Add(emp);
                MyBook.Save();
            }
            catch (Exception ex)
            { }

        }

        public static void CloseExcel()
        {
            MyBook.Saved = true;
            MyApp.Quit();

        }

        public static void GenerateExcel()
        {
            var Path = Directory.CreateDirectory(Environment.CurrentDirectory + "\\Excels\\");
            int counter = 1;
           foreach(string serial in serialNumbers)
            {
                Excel.Workbook MyNewBook = null;
                Excel.Application MyNewApp = null;
                Excel.Worksheet MyNewSheet = null;

                object misValue = System.Reflection.Missing.Value;
                MyNewApp = new Excel.Application();
                MyNewBook = MyNewApp.Workbooks.Add(misValue);
                MyNewSheet = (Excel.Worksheet)MyNewBook.Worksheets.get_Item(1);

                MyNewSheet.Cells[1, 1] = "Material Number";
                MyNewSheet.Cells[1, 2] = "Batch";
                MyNewSheet.Cells[1, 3] = "Stock Type";
                MyNewSheet.Cells[1, 4] = "Plant";
                MyNewSheet.Cells[1, 5] = "Stock_Doccat";
                MyNewSheet.Cells[1, 6] = "WBS";
                MyNewSheet.Cells[1, 7] = "Storage type";
                MyNewSheet.Cells[1, 8] = "Source Bin";
                MyNewSheet.Cells[1, 9] = "Quantity";
                MyNewSheet.Cells[1, 10] = "Unit";
                MyNewSheet.Cells[1, 11] = "Dest Bin";
                MyNewSheet.Cells[1, 12] = "Serial Number";

                MyNewSheet.Cells[2, 1] = "'301686339";
                MyNewSheet.Cells[2, 2] = "'0002061243";
                MyNewSheet.Cells[2, 3] = "G1";
                MyNewSheet.Cells[2, 4] = "D69V";
                MyNewSheet.Cells[2, 5] = "PJS";
                MyNewSheet.Cells[2, 6] = "N.50014943.001";
                MyNewSheet.Cells[2, 7] = "1001";
                MyNewSheet.Cells[2, 8] = "B1221";
                MyNewSheet.Cells[2, 9] = "1";
                MyNewSheet.Cells[2, 10] = "EA";
                MyNewSheet.Cells[2, 11] = "04_PALERMO_B";
                MyNewSheet.Cells[2, 12] ="'" +serial;


                MyNewBook.SaveAs(Path.FullName +"Mass UPD D69V "+counter , Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                    Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                MyNewBook.Close();
                MyNewApp.Quit();

                Marshal.ReleaseComObject(MyNewSheet);
                Marshal.ReleaseComObject(MyNewBook);
                Marshal.ReleaseComObject(MyNewApp);
                counter++;
            }

            

        }
    }

}
