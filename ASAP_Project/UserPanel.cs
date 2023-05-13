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
        /// <summary>
        /// This one generates an Excel from scratch
        /// - Tan :D
        /// </summary>
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
            xlProjectsGrading.Name = "Projects " + " Grading";
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
            xlQuizzesGrading.Name = "Quiz " + " Grading";
            for (int a = 1; a < Project_no + 1; a++)
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
            xlLabsGrading.Name = "Lab " + " Grading";
            for (int a = 1; a < Lab_no + 1; a++)
            {
                xlLabsGrading.Cells[1, a] = "Lab-" + (a).ToString();
            }

            //Final sheet
            if (isFinal == true)
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
                xlFinalDC.Name = "Final " + " Constraints";
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
                xlFinalGrading.Name = "Final " + " Grading";
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
                for (int a = 1; a < Midterm_Q_no[i] + 1; a++)
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

            xlApp.Visible = true;
        }

        //Calculator for HW,Midterms and Final
        private static int ExcelCalculator(Excel.Workbook wb, Excel.Worksheet worksheet, int Counter, String name)
        {
            //We calculate question no of this midterm
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
            foreach (Excel.Worksheet worksheet2 in wb.Worksheets)
            {
                if (worksheet2.Name == name + Counter.ToString() + " Constraints")
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
                else if (worksheet2.Name == name + Counter.ToString() + " Grading")
                {
                    worksheetGrades = worksheet2;
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
            for(int i = 0; i < Student_no; i++)
            {
                for(int j = 0; j < DC_no; j++)
                {
                    double sum = 0;
                    double total_DC = 0;
                    for(int k = 0; k < Question_no; k++)
                    {
                        if(DCArray[j, k] == 1)
                        {
                            sum = sum + (questionScores[i, k] / Convert.ToDouble(GradesArray[k]));
                            total_DC = total_DC + 1.0;
                        }
                    }

                    Student_DC[i, j] = sum / total_DC;

                }
            }

            Excel.Worksheet xlStudentDCSheet = (Excel.Worksheet)wb.Worksheets.Add();
            xlStudentDCSheet.Name = name + Counter.ToString() + " Student - DC";
            xlStudentDCSheet.Cells[1, 1] = "Id";
            xlStudentDCSheet.Cells[1, 2] = "Student ID";
            xlStudentDCSheet.Cells[1, 3] = "Student Name";
            xlStudentDCSheet.Cells[1, 4] = "Student Surname";
            for (int x = 5; x < DC_no + 5; x++)
            {
                xlStudentDCSheet.Cells[1, x] = "DC" + (x - 4).ToString();
            }
            for(int i = 2; i <Student_no + 2; i++)
            {
                //We enter student info to that code as well
                //To do so, we must get the student data for the lesson.
                //If we are going to use this on many places, we might need a function which extracts 
                //The student information from the drive excel created by the admin
                xlStudentDCSheet.Cells[i, 1] = i - 1;
                for (int j = 5; j < DC_no + 5; j++)
                {
                    xlStudentDCSheet.Cells[i, j]= Student_DC[i - 2, j - 5];
                }
            }
            Counter = Counter + 1;
            return Counter;
        }


        /// <summary>
        /// THIS ONE CREATES A REPORT FROM AN EXISING EXCEL
        /// IT MAKES NECESARRY CALCULATIONS INTERNALLY AND IMPLEMENTS THEM TO EXCEL FILE NEWLY CREATED
        /// IN THIS ONE I CHOOSE A LOCAL FILE SO IT CHANGES THAT
        /// BUT WE WILL PICK IT UP FROM DRIVE SO IT WİLL SAVE THE FILE TO THE PC INSTEAD (GOD I HOPE SO)
        /// - TAN :D
        /// </summary>
        /// 


        // verilerin tutulacağı alan:
        public class Veriler
        {
            private double[,] finalSinavVerileri;
            private double[,,] vizeSinavVerileri;
            private double[,,] odevVerileri;
            private double[,,] labVerileri;

            public Veriler(double[,] finalSinavVerileri, double[,,] vizeSinavVerileri, double[,,] odevVerileri, double[,,] labVerileri)
            {
                this.finalSinavVerileri = finalSinavVerileri;
                this.vizeSinavVerileri = vizeSinavVerileri;
                this.odevVerileri = odevVerileri;
                this.labVerileri = labVerileri;
            }

            public double[,] FinalSinavVerileri
            {
                get { return finalSinavVerileri; }
                set { finalSinavVerileri = value; }
            }

            public double[,,] VizeSinavVerileri
            {
                get { return vizeSinavVerileri; }
                set { vizeSinavVerileri = value; }
            }

            public double[,,] OdevVerileri
            {
                get { return odevVerileri; }
                set { odevVerileri = value; }
            }

            public double[,,] LabVerileri
            {
                get { return labVerileri; }
                set { labVerileri = value; }
            }
        }


        public static void CreateReport()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.ShowDialog();

            //This part takes values from excel 
            //In the order of midterm-n , midterm-n constraints and midterm-gradeing constraints
            //does necessary calculations and writes the result
            Excel.Application application = new Excel.Application();
            Excel.Workbook wb = application.Workbooks.Open(openFileDialog.FileName);
            int Homework_counter = 1;
            int Midterm_counter = 1;
            //We check each sheet of the loaded file  statement
            // Get the total number of worksheets
            int totalWorksheets = wb.Worksheets.Count;

            // Iterate through the worksheets collection in reverse order
            for (int i = totalWorksheets; i > 0; i--)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)wb.Worksheets[i];

                // Check if the worksheet is one of the desired worksheets
                // Perform calculations and editing inside this if statement
                if (worksheet.Name == "Midterm-" + Midterm_counter.ToString())
                {
                    Midterm_counter = ExcelCalculator(wb, worksheet, Midterm_counter, "Midterm-");
                }
            }

            //Now for Homeworks
            /*foreach (Excel.Worksheet worksheet in wb.Worksheets)
            {

                if (worksheet.Name == "Homework-" + Homework_counter.ToString())
                {
                    Homework_counter = ExcelCalculator(wb, worksheet, Homework_counter, "Homework-");
                }
            }*/
            application.Visible = true;
            wb.Save();
        }
    }
}
