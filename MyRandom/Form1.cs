using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace MyRandom
{
    public partial class Form1 : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        public Form1()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }
        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void btnStat_Click(object sender, EventArgs e)
        {
            Excel.Range[] rows = RandomRows(5, @"D:\MyResult\test.xlsx", 3);

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

            RandomExcelRows.DataSource = dt;
            killExcelProccess();
        }

        public Excel.Range[] RandomRows(int randomRowsToGet, string worksheetLocation, int worksheetNumber = 1, int lowestRow = 1, int highestRow = 10)
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

        private void openFileDialog1_FileOk_1(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
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
                    }
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            RandomExcelRows.Rows.Clear();
        }
    }
}
