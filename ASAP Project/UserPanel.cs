using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.Logging;
using Word = Microsoft.Office.Interop.Word;

namespace ASAP_Project
{
    public class UserPanel
    {
        public static void GenerateExcel()
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Sisteminizde Excel kurulu değil...");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Sıra NO";
            xlWorkSheet.Cells[1, 2] = "İsim";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "Esat";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Emre";

            xlWorkBook.SaveAs("deneme_dosya.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel dosyası c:\\deneme-dosya.xls adresinde oluşturuldu...");



            //var excelApp = new Excel.Application();
            //excelApp.Workbooks.Add();

            //var xlSheets = excelApp.Sheets as Excel.Sheets;
            //var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

            //excelApp.Visible = true;
        }

        public static void CreateReport()
        {
            //Added code version 1.0 for report genreation, 
            //experiences runtime error in Program.cs, need checking

            //We create two instances for an Excel and a Word File
            Excel.Application excelApp = new Excel.Application();
            Word.Application wordApp = new Word.Application();

            //We pick our Excel file from Pc (Emre's code)
            string oSelectedFile = "";
            System.Windows.Forms.OpenFileDialog oDlg = new System.Windows.Forms.OpenFileDialog();
            if (System.Windows.Forms.DialogResult.OK == oDlg.ShowDialog())
            {
                oSelectedFile = oDlg.FileName;

            }

            //We match our excel 
            Excel.Workbook workbook = excelApp.Workbooks.Open(oSelectedFile);
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            //We create a new Word document
            Word.Document document = wordApp.Documents.Add();

            //We scan our Excel data and add it to our newly
            //created Word document
            for (int i = 1; i <= worksheet.UsedRange.Rows.Count; i++)
            {
                for (int j = 1; j <= worksheet.UsedRange.Columns.Count; j++)
                {
                    string cellValue = worksheet.Cells[i, j].Value.ToString();
                    Word.Range range = document.Content;
                    range.InsertAfter(cellValue + "\t");
                }
                Word.Range rowRange = document.Content;
                rowRange.InsertAfter("\n");
            }

            //Save the Word document
            document.SaveAs("/report.docx");

            //Close Excel and Word documents
            workbook.Close();
            excelApp.Quit();

            document.Close();
            wordApp.Quit();
        }
    }
}
