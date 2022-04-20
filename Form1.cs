using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

// Below references (libraries, .dll) you should download and install first then use, via "Nuget-pocket Manager" or references features 
using ExcelDataReader;
using System.Data.Common;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp_tesTask_Excel_1_TableDataPrct
{
    public partial class Form1 : Form
    {
        //
        // Creating Interface to import Data from Excel-file to DataGridView 
        //

        private string fileName = string.Empty;

        private DataTableCollection tableCollection = null;
        public Form1()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;

                    Text = fileName;

                    OpenExcelFile(fileName);
                }
                else 
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                { 
                    UseHeaderRow=true
                }

            });

            tableCollection = db.Tables;
            toolStripComboBox1.Items.Clear();

            foreach (DataTable table in tableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

            dataGridView1.DataSource = table;
        }

        //
        // Filling 2D Array from DataGridView 
        //

        private void button6_Click_1(object sender, EventArgs e)
        {
            string[,] result_arr = new string[dataGridView1.RowCount, dataGridView1.ColumnCount];

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (i == dataGridView1.RowCount - 1)
                    {
                        result_arr[0, j] = dataGridView1.Columns[j].HeaderText;
                        
                    }
                    if (i < dataGridView1.RowCount - 1)
                    {
                        result_arr[i + 1, j] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                    }
                }
            }
            
            //
            // Array-Data-Processing and Printing our results
            //

            string[] result = new string[result_arr.Length];
            textBox1.Text.Trim();
            int correctnumber = 1;
            int ur_prorab = 1;
            for (int i = 0; i < result_arr.GetLength(0); i++)
            {
                if (result_arr[i, 1] == textBox1.Text)
                {
                    result[0] = result_arr[i, 1] + " | ";
                    correctnumber++;
                    for (int j = 0; j < result_arr.GetLength(1); j++)
                    {
                        if (result_arr[i, j] == "+")
                        {
                            result[ur_prorab] = result_arr[0, j] + " | ";
                            ur_prorab++;
                        }
                    }
                }
               
            }
           
            if (correctnumber == 1)
            {
                MessageBox.Show("Your number was wrong", "Some title", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (textBox2.Text.Trim()==" ")
            {
                textBox2.Clear();
            }

            string separator = " ";
            string q = String.Join(separator, result);
            textBox2.Text = q.Trim();
        }

        // to avoid mistakes in result We block pressing Enter in our small texbox  
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
            }
        }
    }
}
