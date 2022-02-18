using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace StudentCardReader
{
    public partial class Form1 : Form
    {
        private String excelFilename;
        private String dataFilename;
        private List<String> studentIDList = new List<string>();
        private bool isFoundStudent = false;
        private int ticketNum = 0;
        public Form1()
        {
            InitializeComponent();
            Form2 f2 = new Form2();
            f2.ShowDialog();
            dataFilename = f2.fileDir;
            label7.Text = dataFilename;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 19)
            {
                label2.Text = "Scanning!";
                button2.Enabled = true;
            }
            else 
            {
                label2.Text = "Please scan the card";
                button2.Enabled = false;
                isFoundStudent = false;
                label9.Text = "";
                label3.Text = "Student ID";
                label4.Text = "Student Name";
                label6.Text = "Okey Username";
                label8.Text = "Student Email Address";
                textBox2.Enabled = false;
                textBox2.Text = "";
                label10.Text = "";
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Length > 0)
            {
                bool canConvert = int.TryParse(textBox2.Text, out ticketNum);

                if (canConvert)
                {
                    string dirName = Path.GetDirectoryName(dataFilename) + @"\datasave.xlsx";
                    label7.Text = dirName;
                    if (File.Exists(dirName))
                    {
                        // This path is a file
                        Console.WriteLine("file exist");

                        Microsoft.Office.Interop.Excel.Application xlApp1 = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(dirName);
                        Microsoft.Office.Interop.Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
                        Microsoft.Office.Interop.Excel.Range xlRange1 = xlWorksheet1.UsedRange;

                        int rowCount1 = xlRange1.Rows.Count;
                        int colCount1 = xlRange1.Columns.Count;

                        xlWorksheet1.Cells[rowCount1 + 1, 1] = label3.Text;
                        xlWorksheet1.Cells[rowCount1 + 1, 2] = label4.Text;
                        xlWorksheet1.Cells[rowCount1 + 1, 3] = label6.Text;
                        xlWorksheet1.Cells[rowCount1 + 1, 4] = label8.Text;
                        xlWorksheet1.Cells[rowCount1 + 1, 5] = ticketNum;

                        //cleanup  
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        //rule of thumb for releasing com objects:  
                        //  never use two dots, all COM objects must be referenced and released individually  
                        //  ex: [somthing].[something].[something] is bad  

                        //release com objects to fully kill excel process from running in the background  
                        Marshal.ReleaseComObject(xlRange1);
                        Marshal.ReleaseComObject(xlWorksheet1);

                        //close and release
                        xlWorkbook1.Save();
                        xlWorkbook1.Close();
                        Marshal.ReleaseComObject(xlWorkbook1);

                        //quit and release  
                        xlApp1.Quit();
                        Marshal.ReleaseComObject(xlApp1);

                        MessageBox.Show("New data added");
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel.Application xlApp1 = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook xlWorkbook1;
                        Microsoft.Office.Interop.Excel._Worksheet xlWorksheet1;
                        object misValue = System.Reflection.Missing.Value;

                        xlWorkbook1 = xlApp1.Workbooks.Add(misValue);
                        xlWorksheet1 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook1.Worksheets.get_Item(1);

                        xlWorksheet1.Cells[1, 1] = "Student ID";
                        xlWorksheet1.Cells[1, 2] = "Name";
                        xlWorksheet1.Cells[1, 3] = "UserID";
                        xlWorksheet1.Cells[1, 4] = "Email Address";
                        xlWorksheet1.Cells[1, 5] = "# of Ticket";
                        xlWorksheet1.Cells[2, 1] = label3.Text; 
                        xlWorksheet1.Cells[2, 2] = label4.Text;
                        xlWorksheet1.Cells[2, 3] = label6.Text;
                        xlWorksheet1.Cells[2, 4] = label8.Text;
                        xlWorksheet1.Cells[2, 5] = ticketNum;

                        //cleanup  
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        //rule of thumb for releasing com objects:  
                        //  never use two dots, all COM objects must be referenced and released individually  
                        //  ex: [somthing].[something].[something] is bad  

                        //release com objects to fully kill excel process from running in the background  
                        //Marshal.ReleaseComObject(xlRange1);
                        Marshal.ReleaseComObject(xlWorksheet1);

                        xlWorkbook1.Worksheets[1].Name = "MySheet";//Renaming the Sheet1 to MySheet
                        xlWorkbook1.SaveAs(dirName);
                        //close and release  
                        xlWorkbook1.Close();
                        Marshal.ReleaseComObject(xlWorkbook1);

                        //quit and release  
                        xlApp1.Quit();
                        Marshal.ReleaseComObject(xlApp1);

                        MessageBox.Show("New file has created");
                    }
                }
                else
                {
                    label10.Text = "Only number format is allowed to enter.";
                    textBox2.Text = "";
                }
                
            }
            else 
            {
                label10.Text = "Please enter the number of tickets.";
                textBox2.Text = "";
            }
            
            //label3.Text = String.Join(",\n", studentIDList);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            char[] trimChar = { ';', '?', '=' };
            string result = textBox1.Text.Trim(trimChar);
            studentIDList.Add(result);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(dataFilename);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int statusCount = 0;

            for (int i = 1; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 1].Value2.ToString() == result)
                {
                    label3.Text = xlRange.Cells[i, 2].Value2.ToString();
                    label4.Text = xlRange.Cells[i, 3].Value2.ToString() + ", " + xlRange.Cells[i, 4].Value2.ToString();
                    label6.Text = xlRange.Cells[i, 11].Value2.ToString();
                    label8.Text = xlRange.Cells[i, 12].Value2.ToString();

                    isFoundStudent = true;
                    label9.Text = "";
                    textBox2.Enabled = true;
                    break;
                }
                else if (xlRange.Cells[i, 1].Value2.ToString() != result)
                {
                    isFoundStudent = false;
                    if (statusCount == 0)
                    {
                        label9.Text = "Searching .";
                    }
                    else if (statusCount == 50)
                    {
                        label9.Text = "Searching . .";
                    }
                    else if (statusCount == 100)
                    {
                        label9.Text = "Searching . . .";
                        statusCount = 0;
                    }
                    statusCount++;
                }
            }

            if (!isFoundStudent)
            {
                label9.Text = "No data found! ";
                textBox2.Enabled = false;
            }

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            excelFilename = saveFileDialog1.FileName;

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }
         //https://www.c-sharpcorner.com/article/read-excel-file-in-c-sharp-winform/

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }
    }
}
