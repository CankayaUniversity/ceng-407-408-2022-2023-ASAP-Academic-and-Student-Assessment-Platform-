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
        public static void GenerateExcel(int Student_no, int Midterm_no, int Homework_no ,
            int Lab_no, int Quiz_no, int Project_no, int Lesson_output_no , bool isCatalog , bool isFinal ,
            int[] Midterm_Q_no, int[] Homework_Q_no,
            int[] Quiz_Q_no, int Final_Q_no = 1)
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Sisteminizde Excel kurulu değil...");
                return;
            }

            Excel.Workbook xlWorkBook;

            //First we create the sheet responsible for Student Information holding
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet xlStudentSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlStudentSheet.Name = "Students";
            xlStudentSheet.Cells[1, 1] = "Id";
            xlStudentSheet.Cells[1, 2] = "Student ID";
            xlStudentSheet.Cells[1, 3] = "Student Name";
            xlStudentSheet.Cells[1, 4] = "Student Surname";
            xlStudentSheet.Cells[1, 5] = "Age";
            xlStudentSheet.Cells[1, 6] = "Email";
            xlStudentSheet.Cells[1, 8] = "GPA";
            xlStudentSheet.Cells[1, 9] = "CumGPA";

            for (int i = 2; i < Student_no + 2; i++)
            {
                xlStudentSheet.Cells[i, 1] = i - 1;
            }

            //Now we create Midterm excel
            Excel.Worksheet[] xlMidtermSheet = new Excel.Worksheet[Midterm_no];
            for (int i = 0; i < Midterm_no; i++)
            {
                xlMidtermSheet[i] = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            }

            for (int i = 0; i < xlMidtermSheet.Length; i++)
            {
                Excel.Worksheet sheet = xlMidtermSheet[i];
                sheet.Name = "Midterm-" + (i + 1).ToString();
                sheet.Cells[1, 1] = "Id";
                sheet.Cells[1, 2] = "Student ID";
                sheet.Cells[1, 3] = "Student Name";
                sheet.Cells[1, 4] = "Student Surname";
                for (int k=5;k<Midterm_Q_no[i] + 5; k++)
                {
                    sheet.Cells[1, k] = "Question-" + (k - 4).ToString();
                }
                for (int j = 2; j < Student_no + 2; j++)
                {
                    sheet.Cells[j, 1] = j - 1;
                }
            }

            //Homework sheet(s)
            Excel.Worksheet[] xlHomeworkSheet = new Excel.Worksheet[Homework_no];
            for (int i = 0; i < Homework_no; i++)
            {
                xlHomeworkSheet[i] = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            }

            for (int i = 0; i < xlHomeworkSheet.Length; i++)
            {
                Excel.Worksheet sheet = xlHomeworkSheet[i];
                sheet.Name = "Midterm-" + (i + 1).ToString();
                sheet.Cells[1, 1] = "Id";
                sheet.Cells[1, 2] = "Student ID";
                sheet.Cells[1, 3] = "Student Name";
                sheet.Cells[1, 4] = "Student Surname";
                for (int k = 5; k < Homework_Q_no[i] + 5; k++)
                {
                    sheet.Cells[1, k] = "Question-" + (k - 4).ToString();
                }
                for (int j = 2; j < Student_no + 2; j++)
                {
                    sheet.Cells[j, 1] = j - 1;
                }
            }


            Excel.Worksheet xlLabSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            Excel.Worksheet xlQuizSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            Excel.Worksheet xlProjectSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            Excel.Worksheet xlLessonOutputSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();

            xlLabSheet.Name = "Labs";
            xlQuizSheet.Name = "Quizs";
            xlProjectSheet.Name = "Projects";
            xlLessonOutputSheet.Name = "Lesson Outputs";

            xlApp.Visible = true;
        }
        public static void CreateReport()
        {
            //Added code version 1.0 for report genreation, 
            //experiences runtime error in Program.cs, need checking

            //We create two instances for an Excel and a Word File
            Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

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
                    string cellValue = worksheet.Cells[i, j].Value;
                    if (cellValue == null)
                    {
                        continue;
                    }
                    else
                    {
                        cellValue = worksheet.Cells[i, j].Value.ToString();
                    }
                    Word.Range range = document.Content;
                    range.InsertAfter(cellValue + "\t");
                }
                Word.Range rowRange = document.Content;
                rowRange.InsertAfter("\n");
            }

            //Save the Word document

            //Close Excel and Word documents
            workbook.Close();
            excelApp.Quit();

            wordApp.Visible = true;
        }
    }
}
