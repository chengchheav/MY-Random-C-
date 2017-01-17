using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace MyRandom
{
    public partial class FormListAndShowData : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string FilePath = string.Empty;
        private int rowAmount = 0;
        public FormListAndShowData()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            this.FilePath = filePath;
            string extension = Path.GetExtension(filePath);
            string header = "YES";// rbHeaderYes.Checked ? "YES" : "NO";
            string conStr, sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();
                        Console.WriteLine(dt);

                        //Populate DataGridView.
                        dataGridView1.DataSource = dt;
                        this.rowAmount = dataGridView1.RowCount;
                    }
                }
            }
        }

        private void FormListAndShowData_Load(object sender, EventArgs e)
        {
            //this.TopMost = true;
            //this.FormBorderStyle = FormBorderStyle.Fixed3D;
            //this.WindowState = FormWindowState.Maximized;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {

                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                var column1 = row.Cells["No"].Value.ToString();
                var column2 = row.Cells["QA"].Value.ToString();
                column2.Replace("\n", "\r\n");
                var column3 = row.Cells["Correct"].Value.ToString();
                txtNo.Text = column1;
                //string qa = row.Cells["QA"].Value.ToString();
                //qa = qa.Replace("\n", "\r\n");
                //txtQA.Text = column2;
                //richTextBox1.AppendText(column2);
                richTextBox1.Text = column2;

                txtCorrect.Text = column3;
            }
        }

        private void btnRandom_Click(object sender, EventArgs e)
        {
            if(this.FilePath !="")
            {
                var filepath = this.FilePath;
                btnRandom.ForeColor = Color.White;
                btnRandom.BackColor = Color.LightGray;
                btnRandom.Text = "Randoming...";
                Cursor.Current = Cursors.WaitCursor;
                int result = 0;
                int.TryParse(txtResult.Text, out result);
                var ln = @"D:\MyResult\test.xlsx";
                //DataRow lastRow = dataGridView1.Rows[dataGridView1.Rows.Count - 1];
                //int lastRow = 
                Excel.Range[] rows = RandomRows(result, this.FilePath,1,1,this.rowAmount);

                DataTable dt = new DataTable();

                bool ColumnsCreated = false;

                foreach (Excel.Range row in rows)
                {
                    object[,] values = row.Value;

                    int columnCount = values.Length;

                    if (!ColumnsCreated)
                    {
                        for (int i = 0; i < columnCount; i++)
                        {
                            //DataColumn dc = new DataColumn(String.Format("Column {0}", i));
                            DataColumn dc = new DataColumn(String.Format("Column {0}", i));
                            dt.Columns.Add(dc);
                            ColumnsCreated = true;
                        }
                    }

                    DataRow dr = dt.NewRow();

                    for (int i = 0; i < columnCount; i++)
                    {
                        dr[String.Format("Column {0}", i)] = values[1, i + 1];
                    }

                    dt.Rows.Add(dr);
                }
                dt.Columns[0].ColumnName = "No";
                dt.Columns[1].ColumnName = "QA";
                dt.Columns[2].ColumnName = "Correct";
                //RandomExcelRows.DataSource = dt;
                dataGridView1.DataSource = dt;
                killExcelProccess();
                btnRandom.Text = "Random";
                btnRandom.ForeColor = Color.Black;
                btnRandom.BackColor = Color.Gainsboro;
            }
            else
            {
                MessageBox.Show("Please select excel file 1.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public Excel.Range[] RandomRows(int randomRowsToGet, string worksheetLocation, int worksheetNumber = 1, int lowestRow = 1, int highestRow = 100)
        {
            Excel.Range[] rows = new Excel.Range[randomRowsToGet];

            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(worksheetLocation);
            Excel.Worksheet worksheet = workbook.Worksheets[worksheetNumber];

            List<int> rowNumbers = new List<int>();

            bool allUniqueNumbers = false;

            Random random = new Random();

            while (!allUniqueNumbers)
            {
                int nextNumber = random.Next(lowestRow, highestRow);

                if (!rowNumbers.Contains(nextNumber))
                    rowNumbers.Add(nextNumber);

                if (rowNumbers.Count == randomRowsToGet)
                    allUniqueNumbers = true;
            }

            for (int i = 0; i < randomRowsToGet; i++)
            {
                rows[i] = worksheet.UsedRange.Rows[rowNumbers[i]];
            }

            Marshal.ReleaseComObject(excel);

            return rows;
        }

        private void killExcelProccess()
        {
            //Kill the all EXCEL obj from the Task Manager(Process)
            System.Diagnostics.Process[] objProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");

            if (objProcess.Length > 0)
            {
                System.Collections.Hashtable objHashtable = new System.Collections.Hashtable();

                // check to kill the right process
                foreach (System.Diagnostics.Process processInExcel in objProcess)
                {
                    if (objHashtable.ContainsKey(processInExcel.Id) == false)
                    {
                        processInExcel.Kill();
                    }
                }
                objProcess = null;
            }

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void btnCSV_Click(object sender, EventArgs e)
        {
            bool isCompleted = false;
            if (dataGridView1.RowCount > 0)
            {
                try
                {
                    int widthCulomn = 60;
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook wb = excel.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
                    Excel.Worksheet ws = (Excel.Worksheet)excel.ActiveSheet;
                    //var excel = new Microsoft.Office.Interop.Excel.Application();
                    //Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
                    //Worksheet ws = (Worksheet)excel.ActiveSheet;
                    
                    //excel.Visible = true;//this show the excel file.
                    //row1
                    //merging cells
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]].Merge();
                    //bold entire row
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]].EntireRow.Font.Bold = true;
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]].Cells.Font.Size = 14;
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]].Cells = "Random Q & A";
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]].RowHeight = 30;
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]].Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    //end

                    //row 2
                    ws.Range[ws.Cells[2, 1], ws.Cells[2, 3]].Merge();
                    ws.Range[ws.Cells[2, 1], ws.Cells[2, 3]].RowHeight = 20;
                    ws.Range[ws.Cells[2, 1], ws.Cells[2, 3]].Cells = DateTime.Now;
                    ws.Range[ws.Cells[2, 1], ws.Cells[2, 3]].Columns.NumberFormat = "DD-MMM-YYYY HH:mm:ss";
                    ws.Range[ws.Cells[2, 1], ws.Cells[2, 3]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    ws.Range[ws.Cells[2, 1], ws.Cells[2, 3]].Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    //end

                    //header table
                    ws.Cells[3, 1] = "No";
                    ws.Cells[3, 1].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //ws.Range[ws.Cells[3, 2], ws.Cells[3, 4]].Merge();
                    ws.Cells[3, 2].Cells = "Q&A";
                    ws.Columns[2].ColumnWidth = widthCulomn;
                    ws.Cells[3, 2].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    ws.Cells[3, 3] = "Correct";
                    ws.Cells[3, 3].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //ws.Cells[3, 1].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#737373");
                    //ws.Cells[3, 2].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#737373");
                    //end 

                    //body table
                    for (int j = 4; j <= dataGridView1.RowCount + 3; j++)
                    {
                        for (int i = 1; i <= dataGridView1.ColumnCount; i++)
                        {
                            var val= dataGridView1.Rows[j - 4].Cells[i - 1].Value;
                            ws.Cells[j, i] = val;
                            if (i % 2 != 0)
                            {
                                ws.Cells[j, i].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                                ws.Cells[j, i].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                //ws.Cells[j, i].Style.Font.Size = 20;
                            }

                            /*for set color to row*/
                            /*if (j % 2 == 0)
                            {
                                ws.Cells[j, i].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#d3d3d3");
                                ws.Cells[j, i].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#d3d3d3");
                            }
                            else
                            {
                                ws.Cells[j, i].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#eee");
                                ws.Cells[j, i].Interior.Color = System.Drawing.ColorTranslator.FromHtml("#eee");
                            }*/
                            /*end*/
                        }
                    }
                    //end 

                    //set border 
                    Excel.Range last = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    Excel.Range cellRange = ws.get_Range("A3", last);
                    Excel.Borders xborders = cellRange.Borders;
                    xborders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xborders.Weight = 2d;
                    isCompleted = true;
                    //lblExporting.Text = "Export Completed.";
                    if (isCompleted)
                    {
                        excel.Visible = true;//this show the excel file.
                        //lblExporting.Text = "";
                    }
                }
                catch (Exception exp)
                {
                    throw exp;
                }
            }
            else
            {
                MessageBox.Show("No data to export.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txtResult_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtResult_MouseClick(object sender, MouseEventArgs e)
        {
            txtResult.SelectAll();
            txtResult.Focus();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void FormListAndShowData_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dialog = MessageBox.Show("you wish to close application ?", "Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (dialog == DialogResult.OK)
            {
                Application.Exit();
            }
            else
            {
                e.Cancel = true;
            }
        }
    }
}
