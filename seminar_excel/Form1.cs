using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace seminar_excel
{
    public partial class Form1 : Form
    {
        private Form2 form2;
        public Form1()
        {
            InitializeComponent();
            textBox2.Text = "0";
            textBox4.Text = "0";
            textBox5.Text = "0";
            linkLabel1.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path;
            string[] files = null;
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    // shows the path to the selected folder in the folder dialog
                    textBox1.Text = fbd.SelectedPath;
                path = fbd.SelectedPath;
            }
            else
            {
                path = textBox1.Text;
            }
            try
            {
                files = Directory.GetFiles(path);

                int fileCount = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories).Length;
                int per = int.Parse(textBox4.Text);
                int hr = int.Parse(textBox2.Text) * 3600 + int.Parse(textBox5.Text) * 60;
                int filenum = 0;

                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }


                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                /*zoom
                xlWorkSheet.Cells[1, 1] = "Name (Original Name)";
                xlWorkSheet.Cells[1, 2] = "User Email";
                xlWorkSheet.Cells[1, 3] = "Join Time";
                xlWorkSheet.Cells[1, 4] = "Leave Time";
                xlWorkSheet.Cells[1, 5] = "Duration (Minutes)";
                xlWorkSheet.Cells[1, 6] = "Guest";
                xlWorkSheet.Cells[1, 7] = "Recording Consent";*/

                /*ms team*/
                xlWorkSheet.Cells[1, 1] = "ชื่อเต็ม";
                xlWorkSheet.Cells[1, 2] = "เวลาเข้าร่วม";
                xlWorkSheet.Cells[1, 3] = "เวลาออก";
                xlWorkSheet.Cells[1, 4] = "ระยะเวลา";
                xlWorkSheet.Cells[1, 5] = "อีเมล";
                xlWorkSheet.Cells[1, 6] = "บทบาท";
                xlWorkSheet.Cells[1, 7] = "ID ผู้เข้าร่วม(UPN)";
                xlWorkSheet.Cells[1, 10] = "count(สะสม)";
                xlWorkSheet.Cells[1, 11] = "count(รวม)";
                xlWorkSheet.Cells[1, 12] = "เวลา(สะสม)";
                int miniTime = (hr * per) / 100;
                xlWorkSheet.Cells[1, 13] = "ผ่านโดยเวลาขั้นต่ำ = " + miniTime / 3600 + "h " + (miniTime / 60) % 60 + "m " + miniTime % 60 + "s ";
                xlWorkSheet.Cells[1, 14] = per + "% ผ่าน";
                xlWorkSheet.Cells[1, 15] = "ผ่าน/ไม่ผ่าน";



                foreach (var item in files)
                {
                    filenum++;
                    Console.WriteLine(item.ToString());
                    Microsoft.Office.Interop.Excel.Workbook sheet2 = xlApp.Workbooks.Open(item);
                    Microsoft.Office.Interop.Excel.Worksheet x2 = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    Excel.Range userRange2 = x2.UsedRange;
                    int countRecords2 = userRange2.Rows.Count;
                    int alldata = countRecords2 - 8;

                    Excel.Range userRange = xlWorkSheet.UsedRange;
                    int countRecords = userRange.Rows.Count;
                    int add = countRecords + 1;
                    int header = 9;

                    Console.WriteLine(alldata);


                    string nameF = Path.GetFileName(item);

                    for (int b = 0; b < alldata; b++)
                    {
                        xlWorkSheet.Cells[add + b, 1] = x2.Cells[header + b, 1];
                        xlWorkSheet.Cells[add + b, 2] = x2.Cells[header + b, 2];
                        xlWorkSheet.Cells[add + b, 3] = x2.Cells[header + b, 3];
                        xlWorkSheet.Cells[add + b, 4] = x2.Cells[header + b, 4];
                        xlWorkSheet.Cells[add + b, 5] = x2.Cells[header + b, 5];
                        xlWorkSheet.Cells[add + b, 6] = x2.Cells[header + b, 6];
                        xlWorkSheet.Cells[add + b, 7] = x2.Cells[header + b, 7];

                        //Console.WriteLine(fileCount + "     "+ b + "/" + alldata);
                        textBox3.Text = ("1/2 รวมไฟล์ " + filenum + "/" + fileCount + "     " + b + "/" + alldata);
                    }
                }
                Excel.Range userRange3 = xlWorkSheet.UsedRange;
                int countRecords3 = userRange3.Rows.Count;
                Excel.Range userRange4 = xlWorkSheet.get_Range("A2", "G" + countRecords3);
                userRange4.Sort(userRange4.Columns[7], Excel.XlSortOrder.xlAscending);

                int sumdata = countRecords3 - 1;
                string d;
                int i = 0;
                int s = 0;
                string[] numh;
                string[] numm = null;
                string[] nums = null;
                for (int b = 1; b <= sumdata; b++)
                {
                    i++;
                    userRange4.Cells[b, 10] = i;
                    d = ((Excel.Range)userRange4.Cells[b, 4]).Value2.ToString();
                    // date to h m s
                    if (d.IndexOf('h') != -1)
                    {
                        numh = d.Split('h');
                        s += Convert.ToInt32(numh[0]) * 3600;
                        if (d.IndexOf('m') != -1)
                        {
                            numm = numh[1].Split('m');
                            s += Convert.ToInt32(numm[0]) * 60;
                        }
                    }
                    if (d.IndexOf('h') == -1 && d.IndexOf('m') != -1)
                    {
                        numm = d.Split('m');
                        s += Convert.ToInt32(numm[0]) * 60;
                        if (d.IndexOf('s') != -1)
                        {
                            nums = numm[1].Split('s');
                            s += Convert.ToInt32(nums[0]);
                        }

                    }
                    if (d.IndexOf('h') == -1 && d.IndexOf('m') == -1 && d.IndexOf('s') != -1)
                    {
                        nums = d.Split('s');
                        s += Convert.ToInt32(nums[0]);
                    }
                    else
                    {
                        s += 0;
                    }
                    userRange4.Cells[b, 12] = s / 3600 + "h " + (s / 60) % 60 + "m " + s % 60 + "s ";

                    if ((string)userRange4.Cells[b, 7].Value != (string)userRange4.Cells[b + 1, 7].Value || (string)userRange4.Cells[b, 7].Value == null)
                    {
                        userRange4.Cells[b, 11] = i;
                        userRange4.Cells[b, 13] = s / 3600 + "h " + (s / 60) % 60 + "m " + s % 60 + "s ";
                        userRange4.Cells[b, 14] = ((s * 100) / hr) + " %";
                        if (((s * 100) / hr) >= per)
                        {
                            ;
                            userRange4.Cells[b, 15] = "ผ่าน";
                        }
                        else
                        {
                            userRange4.Cells[b, 15] = "ไม่ผ่าน";
                        }
                        i = 0;
                        s = 0;
                    }

                    //Console.WriteLine((string)userRange4.Cells[b, 7].Value);
                    //Console.WriteLine(b + "/" + sumdata);
                    textBox3.Text = ("2/2 คัดแยกผู้ผ่านอบรม " + b + "/" + sumdata);

                }


                xlWorkBook.SaveAs(path + "\\.สรุป1.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                textBox3.Text = ("เสร็จสิ้น " + path + "\\.สรุป1.xlsx");
                linkLabel1.Text = (path + "\\.สรุป1.xlsx");
            }
            catch (COMException)
            {
                textBox3.Text = "โปรดปิดไฟล์ Excel ที่กำลังจะใช้งาน และปิด excel จาก Task Manager / ลบไฟล์ .สรุป.xlsx";
                //Console.WriteLine("123");
            }
            catch (ArgumentException)
            {
                //Console.WriteLine("123");
                textBox3.Text = "โปรดเลือกหรือใส่ตำแหน่ง folder ให้ถูกต้อง";
            }
            catch (DirectoryNotFoundException)
            {
                //Console.WriteLine("123");
                textBox3.Text = "โปรดเลือกหรือใส่ตำแหน่ง folder ให้ถูกต้อง";
            }            
            catch (Exception ex)
            {
                textBox3.Text = ex.Message;
            }
        }
        /*
        private void button2_Click(object sender, EventArgs e)
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet2 = xlApp.Workbooks.Open(textBox3.Text);
            Microsoft.Office.Interop.Excel.Worksheet x2 = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;



            Excel.Range userRange3 = x2.UsedRange;
            int countRecords2 = userRange3.Rows.Count;
            Excel.Range userRange2 = x2.get_Range("A2", "G" + countRecords2);
            userRange2.Sort(userRange2.Columns[1], Excel.XlSortOrder.xlAscending);


            xlApp.Quit();

        }
        */
        /*
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

    openFileDialog1.InitialDirectory = "c:\\" ;
    openFileDialog1.Filter = "Database files (*.mdb, *.accdb)|*.mdb;*.accdb" ;
    openFileDialog1.FilterIndex = 0;
    openFileDialog1.RestoreDirectory = true ;

    if(openFileDialog1.ShowDialog() == DialogResult.OK)
    {
        string selectedFileName = openFileDialog1.FileName;
        //...
    }
        }
        */
        private void button2_Click_1(object sender, EventArgs e)
        {
            string path;
            string[] files;
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    // shows the path to the selected folder in the folder dialog
                    textBox1.Text = fbd.SelectedPath;
                path = fbd.SelectedPath;
            }
            else
            {
                path = textBox1.Text;
            }
            try
            {

            
            files = Directory.GetFiles(path);


            int per = int.Parse(textBox4.Text);
            int hr = int.Parse(textBox2.Text) * 60 + int.Parse(textBox5.Text);
            int filenum = 0;
            int fileCount = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories).Length;

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


            /*zoom*/
            xlWorkSheet.Cells[1, 1] = "Name (Original Name)";
            xlWorkSheet.Cells[1, 2] = "User Email";
            xlWorkSheet.Cells[1, 3] = "Join Time";
            xlWorkSheet.Cells[1, 4] = "Leave Time";
            xlWorkSheet.Cells[1, 5] = "Duration (Minutes)";
            xlWorkSheet.Cells[1, 6] = "Guest";
            xlWorkSheet.Cells[1, 7] = "Recording Consent";
            xlWorkSheet.Cells[1, 8] = "ID";
            xlWorkSheet.Cells[1, 10] = "count(สะสม)";
            xlWorkSheet.Cells[1, 11] = "count(ทั้งหมด)";
            xlWorkSheet.Cells[1, 12] = "เวลา(สะสม)นาที่";
            int miniTime = (hr * per) / 100;
            xlWorkSheet.Cells[1, 13] = "ผ่านโดยเวลาขั้นต่ำ = " + miniTime + " นาที";
            xlWorkSheet.Cells[1, 14] = per + "% ผ่าน";
            xlWorkSheet.Cells[1, 15] = "ผ่าน/ไม่ผ่าน";




            /*ms team*/
            //xlWorkSheet.Cells[1, 1] = "ชื่อเต็ม";
            //xlWorkSheet.Cells[1, 2] = "เวลาเข้าร่วม";
            //xlWorkSheet.Cells[1, 3] = "เวลาออก";
            //xlWorkSheet.Cells[1, 4] = "ระยะเวลา";
            //xlWorkSheet.Cells[1, 5] = "อีเมล";
            //xlWorkSheet.Cells[1, 6] = "บทบาท";
            //xlWorkSheet.Cells[1, 7] = "ID ผู้เข้าร่วม(UPN)";
            //xlWorkSheet.Cells[1, 14] = per + "% ผ่าน";


            foreach (var item in files)
            {
                filenum++;
                //Console.WriteLine(item.ToString());
                Microsoft.Office.Interop.Excel.Workbook sheet2 = xlApp.Workbooks.Open(item);
                Microsoft.Office.Interop.Excel.Worksheet x2 = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                Excel.Range userRange2 = x2.UsedRange;
                int countRecords2 = userRange2.Rows.Count;
                int alldata = countRecords2 - 1;

                Excel.Range userRange = xlWorkSheet.UsedRange;
                int countRecords = userRange.Rows.Count;
                int add = countRecords + 1;
                int header = 2;


                //Console.WriteLine(alldata);

                string nameF = Path.GetFileName(item);

                for (int b = 0; b < alldata; b++)
                {

                    //if (((Excel.Range)x2.Cells[header + b, 1]).Value2.ToString().Substring(0, 1).contains("[0-9]+") == true)
                    //{
                    xlWorkSheet.Cells[add + b, 1] = x2.Cells[header + b, 1];
                        xlWorkSheet.Cells[add + b, 2] = x2.Cells[header + b, 2];
                        xlWorkSheet.Cells[add + b, 3] = x2.Cells[header + b, 3];
                        xlWorkSheet.Cells[add + b, 4] = x2.Cells[header + b, 4];
                        xlWorkSheet.Cells[add + b, 5] = x2.Cells[header + b, 5];
                        xlWorkSheet.Cells[add + b, 6] = x2.Cells[header + b, 6];
                        xlWorkSheet.Cells[add + b, 7] = x2.Cells[header + b, 7];

                    //}


                    //Console.WriteLine("file" + nameF + "  " + b + "/" + alldata);
                    textBox3.Text = ("1/3 รวมไฟล์ " + filenum + "/" + fileCount + "     " + b + "/" + alldata);
                }
            }
            Excel.Range userRange3 = xlWorkSheet.UsedRange;
            int countRecords3 = userRange3.Rows.Count;
            Excel.Range userRange4 = xlWorkSheet.get_Range("A2", "G" + countRecords3);
            userRange4.Sort(userRange4.Columns[1], Excel.XlSortOrder.xlAscending);
            Excel.Range userRange5 = xlWorkSheet.get_Range("C2", "D" + countRecords3);
            userRange5.NumberFormat = "d/m/yyyy h:mm";

            int sumdata = countRecords3 - 1;
            string d;
            int i = 0;
            int s = 0;
            int m = 0;
            string c2 = "";
            int cint = 0;
            for (int b = 1; b <= sumdata; b++)
            {
                textBox3.Text = ("2/3 คัดแยกชื่อผู้อบรม " + b + "/" + sumdata);
                //if (((Excel.Range)xlWorkSheet.Cells[1, b + 1]).Value2.ToString().Substring(0, 1).contains("[0-9]+") == true)

                string c1 = ((Excel.Range)xlWorkSheet.Cells[b+1, 1]).Value2.ToString();
                m += Convert.ToInt32(((Excel.Range)userRange3.Cells[b+1, 5]).Value2.ToString());
                int a = c1.Length;
                for (int i1 = 1; i1 <= a; i1++)
                    {
                        c2 = c1.Substring(0, 1);
                        c1 = c1.Substring(1);
                        if (int.TryParse(c2, out int value))
                        {
                            cint = int.Parse(c2.Substring(0, 1)) + (cint * 10);
                        }
                        else
                            break;
                    }
                Console.WriteLine("item" + c1);
                if (cint == 999)
                    {
                        cint = 0;
                    }
                Console.WriteLine("cint" + cint);
                if (cint != 0)
                {
                    Console.WriteLine("rec");
                    userRange3.Cells[b + 1, 8] = cint.ToString();
                }
                else
                {
                    userRange3.Cells[b + 1, 8] = ((Excel.Range)userRange3.Cells[b + 1, 1]).Value2.ToString();
                }
                cint = 0;


                
            }
                for (int b = 2; b <= sumdata; b++)
            {
                i++;
                s += Convert.ToInt32(((Excel.Range)userRange3.Cells[b, 5]).Value2.ToString());
                userRange3.Cells[b, 10] = i;
                userRange3.Cells[b, 12] = s;
                if (((Excel.Range)userRange3.Cells[b, 8]).Value2.ToString() != ((Excel.Range)userRange3.Cells[b+1, 8]).Value2.ToString())
                {
                    userRange3.Cells[b, 11] = i;
                    userRange3.Cells[b, 13] = s;
                    userRange3.Cells[b, 14] = ((s * 100) / hr) + " %";
                    if ((s*100)/hr >= per)
                    {
                        userRange3.Cells[b, 15] = "ผ่าน";
                    }
                    else
                    {
                        userRange3.Cells[b, 15] = "ไม่ผ่าน";
                    }
                    i = 0;
                    s = 0;
                }
                textBox3.Text = ("3/3 คัดแยกผู้ผ่านอบรม " + b + "/" + sumdata);
                Console.WriteLine(((Excel.Range)userRange3.Cells[b+1, 8]).Value2.ToString() + ((Excel.Range)userRange3.Cells[b+1, 8]).Value2.ToString().GetType());                
            }

            if (((Excel.Range)userRange3.Cells[sumdata, 8]).Value2.ToString() == ((Excel.Range)userRange3.Cells[sumdata + 1, 8]).Value2.ToString())
            {
                i++;
            }
            else{
                i = 1;
                s = 0;
            }


            s += Convert.ToInt32(((Excel.Range)userRange3.Cells[sumdata + 1, 5]).Value2.ToString());
            userRange3.Cells[sumdata + 1, 10] = i;
            userRange3.Cells[sumdata + 1, 11] = i;
            userRange3.Cells[sumdata + 1, 12] = s;
            userRange3.Cells[sumdata + 1, 13] = s;
            userRange3.Cells[sumdata + 1, 14] = ((s * 100) / hr) + " %";
            if ((s * 100) / hr >= per)
            {
                userRange3.Cells[sumdata + 1, 15] = "ผ่าน";
            }
            else
            {
                userRange3.Cells[sumdata + 1, 15] = "ไม่ผ่าน";
            }



            xlWorkBook.SaveAs(path + "\\.สรุป1.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();



            textBox3.Text = ("เสร็จสิ้น " + path + "\\.สรุป1.xlsx");
                Console.WriteLine(path + "\\.สรุป1.xlsx");
                linkLabel1.Text = (path + "\\.สรุป1.xlsx");
        }
            catch (COMException)
            {
                textBox3.Text = "โปรดปิดไฟล์ Excel ที่กำลังจะใช้งาน และปิด excel จาก Task Manager / ลบไฟล์ .สรุป.xlsx";
                //Console.WriteLine("123");
            }
            catch (ArgumentException)
            {
                //Console.WriteLine("123");
                textBox3.Text = "โปรดเลือกหรือใส่ตำแหน่ง folder ให้ถูกต้อง";
            }
            catch (DirectoryNotFoundException)
            {
                //Console.WriteLine("123");
                textBox3.Text = "โปรดเลือกหรือใส่ตำแหน่ง folder ให้ถูกต้อง";
            } 
            catch (Exception ex)
            {
                textBox3.Text = ex.Message;
            }
        }


        /*
private void button4_Click(object sender, EventArgs e)
{
   string s = textBox1.Text ;
   string[] cities = new string[3] { "Mumbai", "London", "New York" };
   string[] slist = {"123asd","12fg","1h","j","1","  ", "asd321","999asd"};
   string c1 = "";
   string c2 = "";
   int cint = 0;
   string[] slists = { };
   string s1 = "131asd";
   string s2 = "23fg";
   string s3 = "45h";
   string s4 = "7";
   string s5 = "j";
   string s6 = "  b";
   int b = 0;

   foreach(var item in slist)
   {
       int a = item.Length;
       c1 = item;
       for(int i = 1; i <= a; i++)
       {
           c2 = c1.Substring(0,1);
           c1 = c1.Substring(1);
           if (int.TryParse(c2, out int value))
           {
               cint = int.Parse(c2.Substring(0, 1)) + (cint * 10);
           }
           else
               break;
       }
       Console.WriteLine("item" + item);
       if (cint == 999)
       {
           cint = 0;
       }
       Console.WriteLine("cint"+ cint);
       cint = 0;
   }

}

private void button5_Click(object sender, EventArgs e)
{
   string path;
   if (String.IsNullOrEmpty(textBox1.Text))
   {
       FolderBrowserDialog fbd = new FolderBrowserDialog();
       if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
           // shows the path to the selected folder in the folder dialog
           textBox1.Text = fbd.SelectedPath;
       path = fbd.SelectedPath;
   }
   else
   {
       path = textBox1.Text;
   }
   textBox3.Text = path;
}
*/
private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
{
   this.linkLabel1.LinkVisited = true;

   // Navigate to a URL.
   System.Diagnostics.Process.Start(linkLabel1.Text);
}

        private void button3_Click(object sender, EventArgs e)
        {
            form2 = new Form2();
            form2.ShowDialog();
        }
    }
}