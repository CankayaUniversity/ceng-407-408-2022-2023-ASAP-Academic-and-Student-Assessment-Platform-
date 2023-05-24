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
    /// <summary>
    /// We have a really long code both in Generate Excel and in
    /// CreateReport. I thought about creating functions to get rid of extra work
    /// but for excel read-write in C#, (If we used another language, for example Python, I would have used functions to do repeating actions of these functions)
    /// that just doesn't work. So I wrote each code one by one copied some parts if needed.
    /// This is just an explanation for why these codes are way too long.
    /// - Tan :D
    /// </summary>
    public class UserPanel
    {
        public static String[,] Name_taker(Excel.Workbook wb, ref int Student_Count)
        {
            String[,] StuInfo = null;
            //Now this will be the code we take from admin page to drive and to this snippet of code
            int totalWorksheets = wb.Worksheets.Count;

            for (int i = totalWorksheets; i > 0; i--)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)wb.Worksheets[i];
                if (worksheet.Name == "Students")
                {
                    StuInfo = new String[Student_Count, 3];
                    for (int j = 2; j < Student_Count + 2; j++)
                    {
                        for (int k = 2; k < 3 + 2; k++) //to make it more understandable I wrote 3+2 instead of 5
                        {
                            StuInfo[j - 2, k - 2] = Convert.ToString(worksheet.Cells[j, k].Value);
                        }
                    }
                    break;
                }
            }
            return StuInfo;
        }

        public static void Name_giver(Excel.Worksheet worksheet, int Student_no, String[,] info)
        {
            //This is from the name_taker code, but doesn't relies on the other for and if statements
            //And directly writes on the worksheet provided, so no return options needed.
            for (int j = 2; j < Student_no + 2; j++)
            {
                for(int k = 2; k < 3 + 2; k++) 
                {
                    worksheet.Cells[j, k].Value = info[j - 2, k - 2];
                }
            }
        }
        /// <summary>
        /// This one generates an Excel from scratch
        /// - Tan :D
        /// </summary>
        public void GenerateExcel(String[,] info, int Student_no, int Midterm_no, int Homework_no,
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
                MessageBox.Show("There is No Excel in your system!!");
                return;
            }

            Excel.Workbook xlWorkBook;

            //First we create the sheet responsible for Student Information holding
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);


            //Projects
            Excel.Worksheet[] xlProjectSheet = new Excel.Worksheet[Project_no];
            for (int i = 0; i < Project_no; i++)
            {
                xlProjectSheet[i] = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            }
            for (int i = 0; i < Project_no; i++)
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
                Name_giver(sheet, Student_no, info);
            }
            //Dc Table generation for projects
            Excel.Worksheet xlProjectsDC = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlProjectsDC.Name = "Project Constraints";
            xlProjectsDC.Cells[1, 1] = "Lesson Output No.";
            for (int i = 1; i < Project_no + 1; i++)
            {
                xlProjectsDC.Cells[1, i + 1] = "Project-" + (i).ToString();
            }
            for (int j = 2; j < Lesson_output_no + 2; j++)
            {
                xlProjectsDC.Cells[j, 1] = j - 1;
            }
            Excel.Worksheet xlProjectsGrading = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            //A table to hold grades of all projects
            xlProjectsGrading.Name = "Project Grading";
            for (int a = 1; a < Project_no + 1; a++)
            {
                xlProjectsGrading.Cells[1, a] = "Project-" + (a).ToString();
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
                Name_giver(sheet, Student_no, info);
            }
            //Dc Table generation for Quizzes
            Excel.Worksheet xlQuizzesDC = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlQuizzesDC.Name = "Quiz Constraints";
            xlQuizzesDC.Cells[1, 1] = "Lesson Output No.";
            for (int i = 1; i < Quiz_no + 1; i++)
            {
                xlQuizzesDC.Cells[1, i + 1] = "Quiz-" + (i).ToString();
            }
            for (int j = 2; j < Lesson_output_no + 2; j++)
            {
                xlQuizzesDC.Cells[j, 1] = j - 1;
            }
            Excel.Worksheet xlQuizzesGrading = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            //A table to hold grades of all Quizzes
            xlQuizzesGrading.Name = "Quiz Grading";
            for (int a = 1; a < Quiz_no + 1; a++)
            {
                xlQuizzesGrading.Cells[1, a] = "Quiz-" + (a).ToString();
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
                Name_giver(sheet, Student_no, info);
            }
            //Dc Table generation for Labs
            Excel.Worksheet xlLabsDC = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlLabsDC.Name = "Lab Constraints";
            xlLabsDC.Cells[1, 1] = "Lesson Output No.";
            for (int i = 1; i < Lab_no + 1; i++)
            {
                xlLabsDC.Cells[1, i + 1] = "Lab-" + (i).ToString();
            }
            for (int j = 2; j < Lesson_output_no + 2; j++)
            {
                xlLabsDC.Cells[j, 1] = j - 1;
            }
            Excel.Worksheet xlLabsGrading = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            //A table to hold grades of all Labs
            xlLabsGrading.Name = "Lab Grading";
            for (int a = 1; a < Lab_no + 1; a++)
            {
                xlLabsGrading.Cells[1, a] = "Lab-" + (a).ToString();
            }
            xlWorkBook.Save();

            //Final sheet
            if (isFinal == true)
            {
                Excel.Worksheet FinalSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
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
                Name_giver(FinalSheet, Student_no, info);
                //Formulation to add questions into the total score
                for (int a = 2; a < Student_no + 2; a++)
                {
                    Excel.Range functionRange = FinalSheet.Range[FinalSheet.Cells[a, k].Address];
                    string formulaString = "=SUM(" + FinalSheet.Cells[a, 5].Address + ":" + FinalSheet.Cells[a, k - 1].Address + ")";
                    functionRange.Formula = formulaString;
                }
                //This one creates the Final-n grading constraints
                //which has Final-n DC - Question table
                Excel.Worksheet xlFinalDC = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlFinalDC.Name = "Final" + " Constraints";
                xlFinalDC.Cells[1, 1] = "Lesson Output No.";

                for (int a = 2; a < Final_Q_no + 2; a++)
                {
                    xlFinalDC.Cells[1, a] = "Question-" + (a - 1).ToString();
                }
                for (int j = 2; j < Lesson_output_no + 2; j++)
                {
                    xlFinalDC.Cells[j, 1] = j - 1;
                }
                //This one creates the homweork-n grading 
                //which has homework-n Question - full_grade table
                Excel.Worksheet xlFinalGrading = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlFinalGrading.Name = "Final " + "Grading";
                for (int a = 1; a < Final_Q_no + 1; a++)
                {
                    xlFinalGrading.Cells[1, a] = "Question-" + (a - 1).ToString();
                    k = a + 1;
                }
                xlFinalGrading.Cells[1, k] = "Total grade";
                //These parts can be copy-pasted to other parts
            }


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
                Name_giver(sheet, Student_no, info);
                //Formulation to add questions into the total score
                for (int a = 2; a < Student_no + 2; a++)
                {
                    Excel.Range functionRange = sheet.Range[sheet.Cells[a, k].Address];
                    functionRange.Locked = false;
                    string formulaString = "=SUM(" + sheet.Cells[a, 5].Address + ":" + sheet.Cells[a, k - 1].Address + ")";
                    functionRange.Formula = formulaString;
                }
                //This one creates the Homework-n grading constraints
                //which has Howework-n DC - Question table
                Excel.Worksheet xlHomeworkDC = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlHomeworkDC.Name = "Homework-" + (i + 1).ToString() + " Constraints";
                xlHomeworkDC.Cells[1, 1] = "Lesson Output No.";

                for (int a = 2; a < Homework_Q_no[i] + 2; a++)
                {
                    xlHomeworkDC.Cells[1, a] = "Question-" + (a - 1).ToString();
                }
                for (int j = 2; j < Lesson_output_no + 2; j++)
                {
                    xlHomeworkDC.Cells[j, 1] = j - 1;
                }
                //This one creates the homweork-n grading 
                //which has homework-n Question - full_grade table
                Excel.Worksheet xlHomeworkGrading = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlHomeworkGrading.Name = "Homework-" + (i + 1).ToString() + " Grading";
                for (int a = 1; a < Homework_Q_no[i] + 1; a++)
                {
                    xlHomeworkGrading.Cells[1, a] = "Question-" + (a - 1).ToString();
                    k = a + 1;
                }
                xlHomeworkGrading.Cells[1, k] = "Total grade";
                //These parts can be copy-pasted to other parts
            }

            //midterm sheet kodu
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
                Name_giver(sheet, Student_no, info);
                //Formulation to add questions into the total score
                for (int a = 2; a < Student_no + 2; a++)
                {
                    Excel.Range functionRange = sheet.Range[sheet.Cells[a, k].Address];
                    functionRange.Locked = false;
                    string formulaString = "=SUM(" + sheet.Cells[a, 5].Address + ":" + sheet.Cells[a, k - 1].Address + ")";
                    functionRange.Formula = formulaString;
                }
                //This one creates the Midterm-n grading constraints
                //which has midterm-n DC - Question table
                Excel.Worksheet xlMidtermDC = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlMidtermDC.Name = "Midterm-" + (i + 1).ToString() + " Constraints";
                xlMidtermDC.Cells[1, 1] = "Lesson Output No.";
                
                for(int a = 2; a < Midterm_Q_no[i] + 2; a++)
                {
                    xlMidtermDC.Cells[1, a] = "Question-" + (a - 1).ToString();
                }
                for (int j = 2; j < Lesson_output_no + 2; j++)
                {
                    xlMidtermDC.Cells[j, 1] = j - 1;
                }
                //This one creates the Midterm-n grading 
                //which has midterm-n Question - full_grade table
                Excel.Worksheet xlMidtermGrading = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                xlMidtermGrading.Name = "Midterm-" + (i + 1).ToString() + " Grading";
                for (int a = 1; a < Midterm_Q_no[i] + 1; a++)
                {
                    xlMidtermGrading.Cells[1, a] = "Question-" + (a - 1).ToString();
                    k = a + 1;
                }
                xlMidtermGrading.Cells[1, k] = "Total grade";
                //These parts can be copy-pasted to other parts
            }

            //We create Student sheet
            Excel.Worksheet xlStudentSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
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
                /*for (int k = 5; k < num; k++)
                {
                    Excel.Range functionRange2 = xlStudentSheet.Range[xlStudentSheet.Cells[a, num - 2].Address];
                    string formul = "='Midterm-" + (k - 4) + "'!" + Convert.ToChar('A' + (4 + Midterm_Q_no[k - 4]) - 1) + a ;
                    functionRange2.Formula = formul;
                }*/
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
            Name_giver(xlStudentSheet, Student_no, info);

            xlApp.Visible = true;
        }


        //Calculator for HW,Midterms and Final
        private static double[,] ExcelCalculator(Excel.Workbook wb, Excel.Worksheet worksheet, int Counter, String name)
        {
            
            //we have 3 templates in generate excel
            //1- Midterm and Homeworks (and Final, but difference is it is only one so no numbers)
            //(DC Table for each Midterm and Homework, and for the one Final)
            //2- Labs ,Projects and Quizes (we make a DC table for their count)
            if (name == "Midterm-" || name == "Homework-" || name == "Final")
            {

                int Question_no = 0;
                for (int i = 5; i < worksheet.Cells.Rows.Count; i++)
                {
                    if (worksheet.Cells[1, i].Value == "Question-" + (i - 4).ToString())
                    {
                        Question_no++;
                    }
                    else if (worksheet.Cells[1, i].Value == null)
                    {
                        break;
                    }
                }
                //We get student_no of this midterm
                int Student_no = 0;
                for (int i = 2; i < worksheet.Cells.Columns.Count; i++)
                {
                    if (worksheet.Cells[i, 1].Value == i - 1)
                    {
                        Student_no++;
                    }
                    else if (worksheet.Cells[i, 1].Value == null)
                    {
                        break;
                    }
                }
                String[,] info = Name_taker(wb, ref Student_no);
                //We create an array to hold Questions of each student
                int[,] questionScores = new int[Student_no, Question_no];
                //then we fill this array
                for (int j = 2; j < Student_no + 2; j++)
                {
                    for (int k = 5; k < Question_no + 5; k++)
                    {
                        questionScores[j - 2, k - 5] = Convert.ToInt32(worksheet.Cells[j, k].Value);
                    }
                }
                //Now we take information from Ders  Ciktisi(Lesson output) table
                //Since we will want users to load the excel they generated from us, it will has its own special template
                //So this code is designed in order to work for that
                //We must show users error messages if they try to upload specially created folders.
                //I MIGHT NEED HELP IN THIS :)
                //- TAN :D
                Excel.Worksheet worksheetDC = null;
                Excel.Worksheet worksheetGrades = null;
                int DC_no = 0;
                int[,] DCArray = null;
                int[] GradesArray = new int[Question_no];
                //Here we get the other 2 worksheet we will use and their values
                //from constranints and expected grades list for midterm-n
                if (name == "Midterm-" || name == "Homework-")
                {
                    foreach (Excel.Worksheet worksheet2 in wb.Worksheets)
                    {
                        if (worksheet2.Name == name + Counter.ToString() + " Constraints")
                        {
                            worksheetDC = worksheet2;
                            for (int i = 2; i < worksheet2.Cells.Columns.Count; i++)
                            {
                                if (worksheet.Cells[i, 1].Value == i - 1)
                                {
                                    DC_no++;
                                }
                                else if (worksheet.Cells[i, 1].Value == null)
                                {
                                    break;
                                }

                            }
                            DCArray = new int[DC_no, Question_no];
                        }
                        else if (worksheet2.Name == name + Counter.ToString() + " Grading")
                        {
                            worksheetGrades = worksheet2;
                        }
                    }
                }else if(name == "Final")
                {
                    foreach (Excel.Worksheet worksheet2 in wb.Worksheets)
                    {
                        if (worksheet2.Name == name + " Constraints")
                        {
                            worksheetDC = worksheet2;
                            for (int i = 2; i < worksheet.Cells.Columns.Count; i++)
                            {
                                if (worksheet.Cells[i, 1].Value == i - 1)
                                {
                                    DC_no++;
                                }
                                else if (worksheet.Cells[i, 1].Value == null)
                                {
                                    break;
                                }

                            }
                            DCArray = new int[DC_no, Question_no];
                        }
                        else if (worksheet2.Name == name + " Grading")
                        {
                            worksheetGrades = worksheet2;
                        }
                    }
                }
                //We input DC of that midterm inside our dynamic DC-lesson array
                for (int j = 2; j < DC_no + 2; j++)
                {
                    for (int k = 2; k < Question_no + 2; k++)
                    {
                        DCArray[j - 2, k - 2] = Convert.ToInt32(worksheetDC.Cells[j, k].Value);
                    }
                }
                //We input that midterms max grades question by question to the Point array
                for (int j = 1; j < Question_no + 1; j++)
                {

                    GradesArray[j - 1] = Convert.ToInt32(worksheetGrades.Cells[2, j].Value);
                }

                //Now we start calculations
                double[,] Student_DC = new double[Student_no, DC_no];
                int sum_grade;

                //This one calculates each DC for all students and stores them in an Student-DC array
                for (int i = 0; i < Student_no; i++)
                {
                    for (int j = 0; j < DC_no; j++)
                    {
                        double sum = 0;
                        double total_DC = 0;
                        for (int k = 0; k < Question_no; k++)
                        {
                            if (DCArray[j, k] == 1)
                            {
                                sum = sum + (questionScores[i, k] / Convert.ToDouble(GradesArray[k]));
                                total_DC = total_DC + 1.0;
                            }
                        }

                        Student_DC[i, j] = sum / total_DC;

                    }
                }

                Excel.Worksheet xlStudentDCSheet = (Excel.Worksheet)wb.Worksheets.Add();
                if (name == "Midterm-" || name == "Homework-")
                {
                    xlStudentDCSheet.Name = name + Counter.ToString() + " Student - DC";
                }
                else if (name == "Final")
                {
                    xlStudentDCSheet.Name = name + " Student - DC";
                }
                xlStudentDCSheet.Cells[1, 1] = "Id";
                xlStudentDCSheet.Cells[1, 2] = "Student ID";
                xlStudentDCSheet.Cells[1, 3] = "Student Name";
                xlStudentDCSheet.Cells[1, 4] = "Student Surname";
                Name_giver(xlStudentDCSheet, Student_no, info);
                for (int x = 5; x < DC_no + 5; x++)
                {
                    xlStudentDCSheet.Cells[1, x] = "DC" + (x - 4).ToString();
                }
                for (int i = 2; i < Student_no + 2; i++)
                {
                    //We enter student info to that code as well
                    //To do so, we must get the student data for the lesson.
                    //If we are going to use this on many places, we might need a function which extracts 
                    //The student information from the drive excel created by the admin
                    xlStudentDCSheet.Cells[i, 1] = i - 1;
                    for (int j = 5; j < DC_no + 5; j++)
                    {
                        xlStudentDCSheet.Cells[i, j] = Student_DC[i - 2, j - 5];
                    }
                }

                //Now here we will store some values for a global operation
                return Student_DC;
            }
            else if(name =="Project" || name == "Quiz" || name == "Lab")//Now for labs,quizzes and projects
            {
                int totalWorksheets = wb.Worksheets.Count;
                Counter++;
                int Event_no = 0;
                for (int i = totalWorksheets; i > 0; i--)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[i];
                    if(ws.Name == name + "-" + Counter.ToString())
                    {
                        Event_no++;
                        Counter++;
                    }
                }
                //We get student_no of this event
                int Student_no = 0;
                for (int i = 2; i < worksheet.Cells.Columns.Count; i++)
                {
                    if (worksheet.Cells[i, 1].Value == i - 1)
                    {
                        Student_no++;
                    }
                    else if (worksheet.Cells[i, 1].Value == null)
                    {
                        break;
                    }
                }
                String[,] info = Name_taker(wb, ref Student_no);
                //We create an array to hold Event(quiz,lab,project) scores of each student
                double[,] EventScores = new double[Student_no, Event_no];
                int col = 0;
                Counter = 1;
                //then we fill this array
                for (int i = totalWorksheets; i > 0; i--)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[i];
                    if (ws.Name == name + "-" + Counter.ToString())
                    {
                        for(int k = 0; k < EventScores.GetLength(0); k++)
                        {
                            //ws.Cells[5,k + 2].Value = EventScores[k,col];
                            //Ben Tan, güdümlü bir malım
                            EventScores[k, col] = ws.Cells[k + 2,5].Value;
                        }
                        col++;
                        Counter++;
                    }
                }
                //Now we take information from Ders  Ciktisi(Lesson output) table
                //Since we will want users to load the excel they generated from us, it will has its own special template
                //So this code is designed in order to work for that
                //We must show users error messages if they try to upload specially created folders.
                //I MIGHT NEED HELP IN THIS :)
                //- TAN :D
                Excel.Worksheet worksheetDC = null;
                Excel.Worksheet worksheetGrades = null;
                int DC_no = 0;
                int[,] DCArray = null;
                int[] GradesArray = new int[Event_no];
                //Here we get the other 2 worksheet we will use and their values
                //from constranints and expected grades list for midterm-n

                foreach (Excel.Worksheet worksheet2 in wb.Worksheets)
                {
                    if (worksheet2.Name == name + " Constraints")
                    {
                        worksheetDC = worksheet2;
                        for (int i = 2; i < worksheet2.Cells.Columns.Count; i++)
                        {
                            if (worksheet.Cells[i, 1].Value == i - 1)
                            {
                                DC_no++;
                            }
                            else if (worksheet.Cells[i, 1].Value == null)
                            {
                                break;
                            }

                        }
                        DCArray = new int[DC_no, Event_no];
                    }
                    if (worksheet2.Name == name + " Grading")
                    {
                        worksheetGrades = worksheet2;
                    }
                }//We input DC of that midterm inside our dynamic DC-lesson array

                for (int j = 2; j < DC_no + 2; j++)
                {
                    for (int k = 2; k < Event_no + 2; k++)
                    {
                        DCArray[j - 2, k - 2] = Convert.ToInt32(worksheetDC.Cells[j, k].Value);
                    }
                }
                //We input that midterms max grades question by question to the Point array
                for (int j = 0; j < Event_no; j++)
                {
                    int num = Convert.ToInt32(worksheetGrades.Cells[2, j + 1].Value);
                    GradesArray[j] = num;
                }

                //Now we start calculations
                double[,] Student_DC = new double[Student_no, DC_no];

                //This one calculates each DC for all students and stores them in an Student-DC array
                for (int i = 0; i < Student_no; i++)
                {
                    for (int j = 0; j < DC_no; j++)
                    {
                        double sum = 0;
                        double total_DC = 0;
                        for (int k = 0; k < Event_no; k++)
                        {
                            if (DCArray[j, k] == 1)
                            {
                                sum = sum + (EventScores[i, k] / Convert.ToDouble(GradesArray[k]));
                                total_DC = total_DC + 1.0;
                            }
                        }

                        Student_DC[i, j] = sum / total_DC;

                    }
                }

                Excel.Worksheet xlStudentDCSheet = (Excel.Worksheet)wb.Worksheets.Add();
                xlStudentDCSheet.Name = name + " Student - DC";
                xlStudentDCSheet.Cells[1, 1] = "Id";
                xlStudentDCSheet.Cells[1, 2] = "Student ID";
                xlStudentDCSheet.Cells[1, 3] = "Student Name";
                xlStudentDCSheet.Cells[1, 4] = "Student Surname";
                Name_giver(xlStudentDCSheet, Student_no, info);
                for (int x = 5; x < DC_no + 5; x++)
                {
                    xlStudentDCSheet.Cells[1, x] = "DC" + (x - 4).ToString();
                }
                for (int i = 2; i < Student_no + 2; i++)
                {
                    //We enter student info to that code as well
                    //To do so, we must get the student data for the lesson.
                    //If we are going to use this on many places, we might need a function which extracts 
                    //The student information from the drive excel created by the admin
                    xlStudentDCSheet.Cells[i, 1] = i - 1;
                    for (int j = 5; j < DC_no + 5; j++)
                    {
                        xlStudentDCSheet.Cells[i, j] = Student_DC[i - 2, j - 5];
                    }
                }
                return Student_DC;
            }
            else
            {
                return null;
            }
        }

        //A Linked list node :D
        //each node holds an 2D Double type array
        public class LinkedListNode
        {
            //Reminder Note from Tan to Tan
            //this Array2D is like Array2D[Student_number, DC_number] ([rows, columns])
            public double[,] Array2D { get; set; }
            public int Rows => Array2D.GetLength(0);
            public int Columns => Array2D.GetLength(1);
            public LinkedListNode Next { get; set; }
        }

        //This simply has a Linked List Addition.
        //I haven't added other functions so far since I will not need them
        //But more functions are likely to come :D
        public class LinkedList
        {
            private LinkedListNode head;
            //To specify which linkedlist is what type
            //Will be useful to tell if we have a midterm or a quiz at hand for example
            public String name;

            public LinkedList(String name)
            {
                this.name = name;
            }

            //returns head DC No
            public int returnHeadColumns()
            {
                return head.Columns;            }

            //Adds a node with a 2D array
            public void Add(double[,] array2D)
            {
                LinkedListNode newNode = new LinkedListNode();
                newNode.Array2D = array2D;

                if (head == null)
                {
                    head = newNode;
                }
                else
                {
                    LinkedListNode current = head;

                    while (current.Next != null)
                    {
                        current = current.Next;
                    }

                    current.Next = newNode;
                }
            }

            //sets head to null
            //No worries for memory leaks since C# HAS Garbage Collector
            public void SetHeadToNull()
            {
                head = null;
            }
            //returns total Node count of the LinkedList
            public int Length()
            {
                int count = 0;
                LinkedListNode current = head;

                while (current != null)
                {
                    count++;
                    current = current.Next;
                }

                return count;
            }
            //returns contents of its Nth node
            public LinkedListNode GetNthNode(int n)
            {
                if (n < 1)
                    throw new ArgumentException("Invalid value for 'n'. Must be greater than or equal to 1.");

                LinkedListNode current = head;
                int count = 1;

                while (current != null)
                {
                    if (count == n)
                        return current;

                    current = current.Next;
                    count++;
                }

                throw new ArgumentOutOfRangeException("n", "The linked list does not contain the specified index.");
            }
        }

        public static LinkedList MidtermDC = new LinkedList("Midterm");
        public static LinkedList FinalDC = new LinkedList("Final");
        public static LinkedList HomeworkDC = new LinkedList("Homework");
        public static LinkedList QuizDC = new LinkedList("Quiz");
        public static LinkedList ProjectDC = new LinkedList("Project");
        public static LinkedList LabDC = new LinkedList("Lab");

        /// <summary>
        /// THIS ONE CREATES A REPORT FROM AN EXISING EXCEL
        /// IT MAKES NECESARRY CALCULATIONS INTERNALLY AND IMPLEMENTS THEM TO EXCEL FILE NEWLY CREATED
        /// IN THIS ONE I CHOOSE A LOCAL FILE SO IT CHANGES THAT
        /// BUT WE WILL PICK IT UP FROM DRIVE SO IT WİLL SAVE THE FILE TO THE PC INSTEAD (GOD I HOPE SO)
        /// - TAN :D
        /// </summary>
        /// 
        public void CreateReport(Excel.Workbook wb, Excel.Application application)
        {
            MidtermDC.SetHeadToNull();
            FinalDC.SetHeadToNull();
            HomeworkDC.SetHeadToNull();
            QuizDC.SetHeadToNull();
            ProjectDC.SetHeadToNull();
            LabDC.SetHeadToNull();

            int Homework_counter = 0;
            int Midterm_counter = 1;
            int Final_counter = 1;
            //We check each sheet of the loaded file  statement
            // Get the total number of worksheets
            int totalWorksheets = wb.Worksheets.Count;
            bool quiz = false;
            bool project = false;
            bool lab = false;

            // Iterate through the worksheets collection in reverse order
            for (int i = totalWorksheets; i > 0; i--)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)wb.Worksheets[i];

                // Check if the worksheet is one of the desired worksheets
                // Perform calculations and editing inside this if statement
                if (worksheet.Name == "Midterm" + "-" + Midterm_counter.ToString())
                {
                    MidtermDC.Add(ExcelCalculator(wb, worksheet, Midterm_counter, "Midterm-"));
                    Midterm_counter++;
                    
                }
                else if (worksheet.Name == "Final")
                {

                    FinalDC.Add(ExcelCalculator(wb, worksheet, Final_counter, "Final"));
                }
                else if (worksheet.Name == "Homework" + "-" + (Homework_counter + 1).ToString())
                {
                    Homework_counter = Homework_counter + 1;
                    HomeworkDC.Add(ExcelCalculator(wb, worksheet, Homework_counter, "Homework-"));
                }
                //One time only entry for quizzes, labs and projects
                else if (worksheet.Name == "Quiz-" + 1.ToString() && !quiz)
                {
                    int quizCount = 0;
                    QuizDC.Add(ExcelCalculator(wb, worksheet, quizCount, "Quiz"));
                    quiz = true;
                }
                else if (worksheet.Name == "Lab-" + 1.ToString() && !lab)
                {
                    int labCount = 0; 
                    LabDC.Add(ExcelCalculator(wb, worksheet, labCount, "Lab"));
                    lab = true;
                }
                else if (worksheet.Name == "Project-" + 1.ToString() && !project)
                {
                    int projectCount = 0;
                    ProjectDC.Add(ExcelCalculator(wb, worksheet, projectCount, "Project"));
                    project = true;
                }
            }
            wb.Save();
            ///////CONSTRUCTION/////////////////
            Excel.Worksheet MainDC = (Excel.Worksheet)wb.Worksheets.Add();
            int current_column = 1;
            MainDC.Name = "Total DC Contribution";
            MainDC.Cells[1, current_column].Value = "Lesson Output";
            current_column =  current_column + 1;
            EditMainDCExcel(MainDC, MidtermDC, ref current_column);//We will  declare like this for 6 times for all potential lesson objects
            EditMainDCExcel(MainDC, HomeworkDC, ref current_column);
            EditMainDCExcel(MainDC, FinalDC, ref current_column);
            EditMainDCExcel(MainDC, QuizDC, ref current_column);
            EditMainDCExcel(MainDC, LabDC, ref current_column);
            EditMainDCExcel(MainDC, ProjectDC, ref current_column);
            ///////CONSTRUCTION/////////////////
            MidtermDC.SetHeadToNull();
            FinalDC.SetHeadToNull();
            HomeworkDC.SetHeadToNull();
            QuizDC.SetHeadToNull();
            ProjectDC.SetHeadToNull();
            LabDC.SetHeadToNull();

            application.Visible = true;
            wb.Save();
        }

        public static void EditMainDCExcel(Excel.Worksheet worksheet, LinkedList linkedList, ref int current_column)
        {
            /*To fill all nodes I will sure need a nev variable. because these will work for one
             type of LinkedList type (ex. Midterm, Final, Quiz etc) each time. What if I also keep an account of 
            the current column I will write to. We can call It here with a reference so we can actually create 
            a sort of progress of columns.
            Lets call this integer value "current_column"*/
            int Lenght = linkedList.Length();
            LinkedListNode node = new LinkedListNode();
            
            for (int counter = 1; counter <= Lenght; counter++)
            {
                //Now we will fill the members of this DC array
                //(Which can be Midterm or HWs)
                //each array index will represent Dc-1 to Dc-n
                //Also a reminder Note from Tan to Tan
                //this Array2D we use is like Array2D[Student_number, DC_number] ([rows, columns])
                double[] DCSumHolder = new double[linkedList.returnHeadColumns()];
                node = linkedList.GetNthNode(counter);
                int Student_number = node.Array2D.GetLength(0);
                double sum = 0.0;
                for (int j = 0; j < node.Columns; j++)
                {
                    for (int i = 0; i < node.Rows; i++)
                    {
                        sum = sum + node.Array2D[i, j];
                    }
                    DCSumHolder[j] = sum / Convert.ToDouble(node.Rows); //total sum of all student contributions to DCj / Student_no
                }
                if (linkedList.name == "Midterm" || linkedList.name == "Homework")
                {
                    worksheet.Cells[1, current_column].Value = linkedList.name + "-" + counter;
                }
                else if (linkedList.name == "Final" || linkedList.name == "Quiz" || linkedList.name == "Lab" || linkedList.name == "Project")
                {
                    worksheet.Cells[1, current_column].Value = linkedList.name;
                }
                for (int i = 2; i < node.Array2D.GetLength(1) + 2; i++)
                {

                    worksheet.Cells[i, current_column] = DCSumHolder[i - 2];
                }
                current_column = current_column + 1;
            }
        }
    }
}
