using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Excel_IO
{
    public partial class Form1 : Form
    {
        // be able to let the user select the path
        // be able to let the user set the time
        // basically create the initial form for setup
        
        public Form1()
        {
            InitializeComponent();
            this.Size = new Size(310, 200);
            this.CenterToScreen();
            beginExcel();
            timer2.Start();
            countdownLbl.Text = "";
        }

        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        Microsoft.Office.Interop.Excel.Range oRng;

        object _missingValue = System.Reflection.Missing.Value;

        string file = @"";
        int cellRow = 1;
        int numberOfEntries;
        int timeLeft;

        public void ChooseFolder()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label4.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void beginExcel()
        {
            
            if (!System.IO.File.Exists(file))
            {
                MessageBox.Show("Click 'OK' to create the log file!");
                openExcel();
                //create file
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                getSheet();
                fillHeaders();
            }
            else
            {
                openExcel();
                //open existing file
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(file, _missingValue, false, _missingValue, _missingValue, _missingValue, true, _missingValue, _missingValue, true, _missingValue, _missingValue, _missingValue));
                getSheet();
                fillHeaders();
            }
        }
        
        //Start Excel and get Application object.
        private void openExcel()
        {
            oXL = new Microsoft.Office.Interop.Excel.Application();
            
            oXL.Visible = false;
            oXL.UserControl = true;
        }

        private void closeExcel()
        {
            
            oXL.DisplayAlerts = false;
            oXL.Visible = false;
            oXL.UserControl = true;
            //oWB.Save();
            oWB.SaveAs(file, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
            oXL.Quit();
        }

        private void getSheet()
        {
            bool exists = false; //can be a 1 if exists or 0 if doesnt

            //check current sheets if one is made for today
            foreach (Microsoft.Office.Interop.Excel._Worksheet sheet in oWB.Worksheets)
            {
                //this is for a brand new file
                if (sheet.Name == "Sheet1")
                {
                    sheet.Name = DateTime.Now.ToString("MM.dd.yy");
                }

                if (sheet.Name == DateTime.Now.ToString("MM.dd.yy"))
                {
                    exists = true;
                    oSheet = sheet;
                }

            }

            if (exists == false)
            {
                Excel.Worksheet newWorksheet = oWB.Worksheets.Add();
                newWorksheet.Name = DateTime.Now.ToString("MM.dd.yy");
                oSheet = newWorksheet;
            }
            
        }

        private void fillHeaders()
        {
            oSheet.Cells[1, 1] = "First Name";
            oSheet.Cells[1, 2] = "Last Name";
            oSheet.Cells[1, 3] = "Employee ID";

            cellRow = cellRow + 1;

            //Format A1:D1 as bold, vertical alignment = center.
            oSheet.get_Range("A1", "C1").Font.Bold = true;
            oSheet.get_Range("A1", "C1").VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            //AutoFit columns A:D.
            oRng = oSheet.get_Range("A1", "C1");
            oRng.EntireColumn.AutoFit();

            oRng = oSheet.get_Range("F1");
            oRng.Formula = "=COUNTIF(A:A,\"*\")";
            oRng.Style.locked = true;
            
            numberOfEntries = Convert.ToInt16(oSheet.Cells[1, 6].value2);
        }

        private void logInBtn_Click(object sender, EventArgs e)
        {
            cellRow = numberOfEntries + 1;

            if (firstTxt.Text != "")
            {
                oSheet.Cells[cellRow, 1] = firstTxt.Text.ToString();
                firstTxt.Text = "";
            }
            else
            {
                oSheet.Cells[cellRow, 1] = "null";
            }
            if (lastTxt.Text != "")
            {
                oSheet.Cells[cellRow, 2] = lastTxt.Text.ToString();
                lastTxt.Text = "";
            }
            else
            {
                oSheet.Cells[cellRow, 2] = "null";
            }
            if (empTxt.Text != "") 
            {
                oSheet.Cells[cellRow, 3] = empTxt.Text.ToString();
                empTxt.Text = "";
            }
            else
            {
                oSheet.Cells[cellRow, 3] = "null";
            }
            
            firstTxt.Enabled = false;
            lastTxt.Enabled = false;
            empTxt.Enabled = false;
            logInBtn.Enabled = false;

            numberOfEntries = Convert.ToInt16(oSheet.Cells[1, 6].value2);
            timeLeft = 5;
            timer1.Start();
            MoveCursor();
            this.WindowState = FormWindowState.Minimized;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
            closeExcel();
        }

        private void MoveCursor()
        {
            // Set the Current cursor, move the cursor's Position,
            // and set its clipping rectangle to the form. 
            
            Cursor.Position = new Point(Cursor.Position.X - 50, Cursor.Position.Y - 50);
            Cursor.Clip = new Rectangle(this.Location, this.Size);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Alt | Keys.P))
            {
                this.Size = new Size(310, 359);
                MoveCursor();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void unlockCursor()
        {
            Cursor.Clip = new Rectangle(1, 1, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Size = new Size(310, 200);
            MoveCursor();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Stop();
            button3.PerformClick();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (timeLeft > 0)
            {
                timeLeft = timeLeft - 1;
                countdownLbl.Text = timeLeft + " seconds";
            }
            else
            {
                timer1.Stop();
                countdownLbl.Text = "times up!";
                
                firstTxt.Enabled = true;
                lastTxt.Enabled = true;
                empTxt.Enabled = true;
                logInBtn.Enabled = true;
                this.Show();
                this.WindowState = FormWindowState.Normal;
                button3.PerformClick();
            }
        }

        protected override void WndProc(ref Message message)
        {
            const int WM_SYSCOMMAND = 0x0112;
            const int SC_MOVE = 0xF010;

            switch (message.Msg)
            {
                case WM_SYSCOMMAND:
                    int command = message.WParam.ToInt32() & 0xfff0;
                    if (command == SC_MOVE)
                        return;
                    break;
            }
            base.WndProc(ref message);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            unlockCursor();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MoveCursor();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("This will exit the program and open the log file in Excel. Proceed?", "Application Exit", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                closeExcel();
                oXL = new Microsoft.Office.Interop.Excel.Application();

                oXL.Visible = true;
                oXL.UserControl = true;
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(file, _missingValue, false, _missingValue, _missingValue, _missingValue, true, _missingValue, _missingValue, true, _missingValue, _missingValue, _missingValue));
                this.Close();
            }
        }
    }
}
