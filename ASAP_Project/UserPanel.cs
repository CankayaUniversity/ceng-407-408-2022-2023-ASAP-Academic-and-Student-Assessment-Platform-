using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Windows;

namespace ASAP_Project
{
    public class UserPanel
    {
        public void GenerateExcel(int Student_no, int Midterm_no, int Homework_no,
            int Lab_no, int Quiz_no, int Project_no, int Lesson_output_no, bool isCatalog, bool isFinal,
            int[] Midterm_Q_no, int[] Homework_Q_no,
            int Final_Q_no)
        {
            // Student sayfasına öğrenci bilgileri girilecek, o bilgiler diğer sayfalara otomatik olarak doldurulacak *** :D eheee
            // Student sayfasına, en son öğrencinin altına sınav notlarının ortalamasının gözükmesi (zamanın olursa bak)

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;

            if (xlApp == null)
            {
                MessageBox.Show("Sisteminizde Excel kurulu değil...");
                return;
            }

            Excel.Workbook xlWorkBook;

            //First we create the sheet responsible for Student Information holding
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            

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
                int k;
                for (k = 5; k < Midterm_Q_no[i] + 5; k++)
                {
                    sheet.Cells[1, k] = "Question-" + (k - 4).ToString();
                }
                sheet.Cells[1, k] = "Total Grade";
                for (int j = 2; j < Student_no + 2; j++)
                {
                    sheet.Cells[j, 1] = j - 1;
                }
                //for loop for function
                for (int a = 2; a < Student_no + 2; a++)
                {
                    Excel.Range functionRange = sheet.Range[sheet.Cells[a, k].Address];
                    functionRange.Locked = false;
                    string formulaString = "=SUM(" + sheet.Cells[a, 5].Address + ":" + sheet.Cells[a, k - 1].Address + ")";
                    functionRange.Formula = formulaString;
                }


            }

            //Homework sheet(s)
            //Hep çalışıyordu
            //bi anda eksik çalışmaya başladı
            //hatalı değil eksik
            //mesela hw no 2 dersek, 1. sheet çıkıyor ama 2. sheet çıkmıyor
            Excel.Worksheet[] xlHomeworkSheet = new Excel.Worksheet[Homework_no];
            for (int i = 0; i < Homework_no; i++)
            {
                xlHomeworkSheet[i] = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            }

            for (int i = 0; i < xlHomeworkSheet.Length; i++)
            {
                Excel.Worksheet sheet = xlHomeworkSheet[i];
                sheet.Name = "Homework-" + (i + 1).ToString();
                sheet.Cells[1, 1] = "Id";
                sheet.Cells[1, 2] = "Student ID";
                sheet.Cells[1, 3] = "Student Name";
                sheet.Cells[1, 4] = "Student Surname";
                int k;
                for (k = 5; k < Homework_Q_no[i] + 5; k++)
                {
                    sheet.Cells[1, k] = "Question-" + (k - 4).ToString();
                }
                sheet.Cells[1, k] = "Total Grade";
                for (int j = 2; j < Student_no + 2; j++)
                {
                    sheet.Cells[j, 1] = j - 1;
                }
                //for loop for function
                for (int a = 2; a < Student_no + 2; a++)
                {
                    Excel.Range functionRange = sheet.Range[sheet.Cells[a, k].Address];
                    functionRange.Locked = false;
                    string formulaString = "=SUM(" + sheet.Cells[a, 5].Address + ":" + sheet.Cells[a, k - 1].Address + ")";
                    functionRange.Formula = formulaString;
                }
            }

            //Final sheet
            if(isFinal == true)
            {
                Excel.Worksheet FinalSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                FinalSheet.Columns.AutoFit();
                FinalSheet.Name = "Final";
                FinalSheet.Cells[1, 1] = "Id";
                FinalSheet.Cells[1, 2] = "Student ID";
                FinalSheet.Cells[1, 3] = "Student Name";
                FinalSheet.Cells[1, 4] = "Student Surname";
                int k;
                for (k = 5; k < Final_Q_no + 5; k++)
                {
                    FinalSheet.Cells[1, k] = "Question-" + (k - 4).ToString();
                }
                FinalSheet.Cells[1, k] = "Total Grade";
                for (int j = 2; j < Student_no + 2; j++)
                {
                    FinalSheet.Cells[j, 1] = j - 1;
                }
                for (int a = 2; a < Student_no + 2; a++)
                {
                    Excel.Range functionRange = FinalSheet.Range[FinalSheet.Cells[a, k].Address];
                    string formulaString = "=SUM(" + FinalSheet.Cells[a, 5].Address + ":" + FinalSheet.Cells[a, k - 1].Address + ")";
                    functionRange.Formula = formulaString;
                }
            }

            //Labs
            Excel.Worksheet[] xlLabSheet = new Excel.Worksheet[Lab_no];
            for (int i = 0; i < Lab_no; i++)
            {
                xlLabSheet[i] = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            }
            for (int i = 0; i < xlLabSheet.Length; i++)
            {
                Excel.Worksheet sheet = xlLabSheet[i];
                sheet.Name = "Lab-" + (i + 1).ToString();
                sheet.Cells[1, 1] = "Id";
                sheet.Cells[1, 2] = "Student ID";
                sheet.Cells[1, 3] = "Student Name";
                sheet.Cells[1, 4] = "Student Surname";
                int k = 5;
                sheet.Cells[1, k] = "Grade";
                for (int j = 2; j < Student_no + 2; j++)
                {
                    sheet.Cells[j, 1] = j - 1;
                }
            }
            //Quizzes
            Excel.Worksheet[] xlQuizSheet = new Excel.Worksheet[Quiz_no];
            for (int i = 0; i < Quiz_no; i++)
            {
                xlQuizSheet[i] = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            }
            for (int i = 0; i < Quiz_no; i++)
            {
                Excel.Worksheet sheet = xlQuizSheet[i];
                sheet.Name = "Quiz-" + (i + 1).ToString();
                sheet.Cells[1, 1] = "Id";
                sheet.Cells[1, 2] = "Student ID";
                sheet.Cells[1, 3] = "Student Name";
                sheet.Cells[1, 4] = "Student Surname";
                int k = 5;
                sheet.Cells[1, k] = "Total Grade";
                for (int j = 2; j < Student_no + 2; j++)
                {
                    sheet.Cells[j, 1] = j - 1;
                }
            }
            //Projects
            Excel.Worksheet[] xlProjectSheet = new Excel.Worksheet[Project_no];
            for (int i = 0; i < Project_no; i++)
            {
                xlProjectSheet[i] = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            }
            for (int i = 0; i < Lab_no; i++)
            {
                Excel.Worksheet sheet = xlProjectSheet[i];
                sheet.Name = "Project-" + (i + 1).ToString();
                sheet.Cells[1, 1] = "Id";
                sheet.Cells[1, 2] = "Student ID";
                sheet.Cells[1, 3] = "Student Name";
                sheet.Cells[1, 4] = "Student Surname";
                int k = 5;
                sheet.Cells[1, k] = "Grade";
                for (int j = 2; j < Student_no + 2; j++)
                {
                    sheet.Cells[j, 1] = j - 1;
                }
            }
            Excel.Worksheet xlLessonOutputSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlLessonOutputSheet.Name = "Lesson Output";

            //We create Student sheet
            Excel.Worksheet xlStudentSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlStudentSheet.Columns.AutoFit();
            xlStudentSheet.Name = "Students";
            xlStudentSheet.Cells[1, 1] = "Id";
            xlStudentSheet.Cells[1, 2] = "Student ID";
            xlStudentSheet.Cells[1, 3] = "Student Name";
            xlStudentSheet.Cells[1, 4] = "Student Surname";
            int num; // columns for midterms and total midterm
            for (num = 5; num < Midterm_no + 5; num++)
            {
                xlStudentSheet.Cells[1, num] = "Midterm-" + (num - 4).ToString();
            }
            xlStudentSheet.Cells[1, num] = "Midterm Total Grade";
            int temp = num;// columns for hws and total homeworks
            //midterm total grade and connection between midterm-n sheets
            for (int a = 2; a < Student_no + 2; a++)
            {
                Excel.Range functionRange = xlStudentSheet.Range[xlStudentSheet.Cells[a, num].Address];
                string formulaString = "=SUM(" + xlStudentSheet.Cells[a, 5].Address + ":" + xlStudentSheet.Cells[a, num - 1].Address + ")";
                functionRange.Formula = formulaString;
                for (int k = 5; k < num; k++)
                {
                    Excel.Range functionRange2 = xlStudentSheet.Range[xlStudentSheet.Cells[a, num - 2].Address];
                    string formul = "='Midterm-" + (k - 4) + "'!" + Convert.ToChar('A' + (4 + Midterm_Q_no[k - 4]) - 1) + a ;
                    functionRange2.Formula = formul;
                }
            }
            //Homework and total grade input
            for (num = num + 1; num < Homework_no + 8; num++)
            {
                xlStudentSheet.Cells[1, num] = "Homework-" + (num - 8).ToString();
            }
            xlStudentSheet.Cells[1, num] = "Homework Total Grade";
            //formula generation
            for (int a = 2; a < Student_no + 2; a++)
            {
                Excel.Range functionRange = xlStudentSheet.Range[xlStudentSheet.Cells[a, num].Address];
                string formulaString = "=SUM(" + xlStudentSheet.Cells[a, temp + 1].Address + ":" + xlStudentSheet.Cells[a, num - 1].Address + ")";
                functionRange.Formula = formulaString;
            }
            num++;
            xlStudentSheet.Cells[1, num] = "Final";

            for (int i = 2; i < Student_no + 2; i++)
            {
                xlStudentSheet.Cells[i, 1] = i - 1;
            }

            xlApp.Visible = true;
        }
        public static void CreateReport()
        {
            
        }
    }
}
