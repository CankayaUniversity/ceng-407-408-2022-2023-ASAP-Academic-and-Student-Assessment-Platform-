using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ASAP_Project
{
    /// <summary>
    /// LoginWindow.xaml etkileşim mantığı
    /// </summary>
    public partial class LoginWindow : System.Windows.Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void button_login_Click(object sender, RoutedEventArgs e)
        {
            
            MemoryStream userdata = GoogleDrive.GetFile("UserInfo.xlsx");
            // Converts MemoryStream to byte[]
            byte[] excelData = userdata.ToArray();

            // Saves byte[] as a temporary file
            string tempFilePath = System.IO.Path.GetTempFileName();
            File.WriteAllBytes(tempFilePath, excelData);

            // Opens the temporary file with Excel Interop
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(tempFilePath);
            int choice = 0;
            foreach (Excel.Worksheet worksheet in wb.Sheets)
            {
                if (worksheet.Name == "Account")
                {
                    for(int i = 2; i < worksheet.Rows.Count; i++)
                    {
                        if(worksheet.Cells[i,2].Value == textbox_username.Text)
                        {
                            if(Convert.ToString(worksheet.Cells[i, 3].Value) == passwordbox_password.Password)
                            {
                                UserData.username = worksheet.Cells[i,2].Value;
                                UserData.role = worksheet.Cells[i,4].Value;
                                //assigns if the user is a user or an admin
                                //goes to the main page with username and user type
                                //in main page we open the specific buttons according to if we have an admin or
                                //a user on board.
                                //From Tan to Emre :D
                                choice = 1;
                                break;
                            }
                            else
                            {
                                MessageBox.Show("Your username or password is wrong");
                                break;
                            }                           
                        }                       
                    }
                }
            }
            if(choice == 1)
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                this.Close();
            }
           
        }




    }
}
