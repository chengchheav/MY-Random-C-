using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PagedList;
namespace MyRandom
{
    public partial class NewDesign : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string FilePath = string.Empty;
        private int rowAmount = 0;
        int pageNumber = 1;
        DataTable AllLists = null;
        IPagedList list;
        public NewDesign()
        {
            InitializeComponent();
        }
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void txtResultNew_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void txtResultNew_MouseClick(object sender, MouseEventArgs e)
        {
            txtResultNew.Select();
            txtResultNew.Focus();
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
            string header = "NO";// rbHeaderYes.Checked ? "YES" : "NO";
            string conStr, sheetName;
            bool supportFile = true;
            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
                default:
                    supportFile = false;
                    MessageBox.Show("File not support.");
                    break;
            }
            if (supportFile)
            {
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
                            //cmd.CommandText = "SELECT top 10 * From [" + sheetName + "]";//select with limi
                            cmd.Connection = con;
                            con.Open();
                            oda.SelectCommand = cmd;
                            oda.Fill(dt);
                            con.Close();
                            this.AllLists = dt;
                            //Bind all row and column from excel to data gridview
                            //dataGridView1.DataSource = dt;

                            //Object cellValue = dt.Rows[i][j];

                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                //Create the new row first and get the index of the new row
                                int rowIndex = this.dataGridView1.Rows.Add();
                                //Obtain a reference to the newly created DataGridViewRow 
                                var row = this.dataGridView1.Rows[rowIndex];
                                for (int j = 0; j < dt.Columns.Count; j++)
                                {
                                    //get cell value
                                    var cellValue = dt.Rows[i][j];
                                    //add value to cell of data gridview
                                    row.Cells[j].Value = cellValue;

                                }
                            }
                            this.rowAmount = dataGridView1.RowCount;
                        }
                    }
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.Clear();
                this.FilePath = string.Empty;
                this.rowAmount = 0;
            }
            catch (Exception exp)
            {
                throw exp;
            }
        }

        private void btnRandom_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            int result = 0;
            int.TryParse(txtResultNew.Text, out result);
            if (result < this.rowAmount)
            {
                if (result != 0)
                {
                    if (this.FilePath != "")
                    {
                        var filepath = this.FilePath;
                        btnRandom.Text = "Randoming...";
                        Cursor.Current = Cursors.WaitCursor;
                        Excel.Range[] rows = RandomRows(result, this.FilePath, 1, 1, this.rowAmount);

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

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            //Create the new row first and get the index of the new row
                            int rowIndex = this.dataGridView1.Rows.Add();
                            //Obtain a reference to the newly created DataGridViewRow 
                            var row = this.dataGridView1.Rows[rowIndex];
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                //get cell value
                                var cellValue = dt.Rows[i][j];
                                //add value to cell of data gridview
                                row.Cells[j].Value = cellValue;
                            }
                        }

                        //RandomExcelRows.DataSource = dt;
                        //dataGridView1.DataSource = dt;
                        killExcelProccess();
                        btnRandom.Text = "Random";
                    }
                    else
                    {
                        MessageBox.Show("Please select excel file 1.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Result must greater than 0.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
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

        private async void btnCsvNew_Click(object sender, EventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;
            bool isCompleted = false;
            if (dataGridView1.RowCount > 0)
            {
                try
                {
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = dataGridView1.RowCount;
                    int index = 1;
                    int widthCulomn = 60;
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook wb = excel.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
                    Excel.Worksheet ws = (Excel.Worksheet)excel.ActiveSheet;
                    //header table
                    ws.Cells[1, 1] = "No";
                    ws.Columns[2].ColumnWidth = 7;
                    ws.Cells[1, 1].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    ws.Cells[1, 2] = "NoOfQuestion";
                    ws.Columns[2].ColumnWidth = 15;
                    ws.Cells[1, 2].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    ws.Cells[1, 3].Cells = "Q&A";
                    ws.Columns[3].ColumnWidth = widthCulomn;
                    ws.Cells[1, 3].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    ws.Cells[1, 4] = "Correct";
                    ws.Cells[1, 4].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;



                    //end 

                    //body table
                    for (int j = 2; j <= dataGridView1.RowCount + 1; j++)
                    {
                        for (int i = 0; i <= dataGridView1.ColumnCount; i++)
                        {
                            if (i != 0)
                            {
                                var val = dataGridView1.Rows[j - 2].Cells[i - 1].Value;
                                ws.Cells[j, i + 1] = val;
                                if (i % 2 != 0)
                                {
                                    ws.Cells[j, i + 1].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                                    ws.Cells[j, i + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                }
                            }
                            else
                            {
                                ws.Cells[j, i + 1] = index;
                                index++;
                                ws.Cells[j, i + 1].Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                                ws.Cells[j, i + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            }

                        }
                        progressBar1.Value = j - 1;
                    }
                    //end 

                    //set border 
                    Excel.Range last = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    Excel.Range cellRange = ws.get_Range("A1", last);
                    Excel.Borders xborders = cellRange.Borders;
                    xborders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xborders.Weight = 2d;
                    isCompleted = true;
                    await Task.Delay(3000);
                    progressBar1.Value = 0;
                    //label5.Text = "";
                    if (isCompleted)
                    {
                        excel.Visible = true;//this show the excel file.
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

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void NewDesign_Load(object sender, EventArgs e)
        {
            lblPage.Visible = false;
            btnPrevious.Visible = false;
            btnNext.Visible = false;
        }
    }
}
