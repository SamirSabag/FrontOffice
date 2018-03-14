using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using Library;

namespace FOApp
{
    public partial class Form1 : Form
    {
        SqlDataAdapter sda;
        SqlCommandBuilder scb;
        DataTable dt;
        public static string TableName;
        static string conString = ConfigurationManager.ConnectionStrings["FOApp.Properties.Settings.SettingString"].ConnectionString; 
        SqlConnection con = new SqlConnection(conString);

        public Form1()
        {
            Utility.Datafiller();
            InitializeComponent();
        }

        public void Display_table()
        {
            try
            {
                TableName = listBoxTrackers.SelectedItem.ToString();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT [ID_],[VENDOR],[INVOICE #],[CURR]," +
                    "[SCANNING DATE],[RECEPTION DATE],[TOTAL],[RATE],[DOC ID]," +
                    "[RESCAN CAUSED BY],[R DOC ID],[RESCN DATE],[NOTES] FROM " + TableName + "";
                sda = new SqlDataAdapter(cmd.CommandText, con );
                dt = new DataTable();
                sda.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                con.Close();
                MessageBox.Show(ex.Message);
            }
        }
        private void listBoxTrackers_SelectedIndexChanged(object sender, EventArgs e)
        {
            Display_table();
        }
  
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                 scb = new SqlCommandBuilder(sda);
                sda.Update(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
   
        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog();
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = opfd.FileName;
            }           
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            importdatafromexcel(txtPath.Text);
            Display_table();
        }
        public void importdatafromexcel(string excelfilepath)
        {
            try
            {
                string ssqltable;
                ssqltable = listBoxTrackers.SelectedItem.ToString();
                string myexceldataquery = "select * from [Tracker$]";
                try
                {
                    string sexcelconnectionstring = @"Provider=Microsoft.ACE.OLEDB.12.0;data source=" + excelfilepath +
                    ";extended properties=" + "\"excel 8.0;hdr=yes;\"";
                    string ssqlconnectionstring = @"Data Source = SABAGSA1\SQLEXPRESS; Initial Catalog = 
                   FOTrackers2018; Integrated Security = True";
                    string sclearsql = "delete from " + ssqltable;
                    string resedsql = " DBCC CHECKIDENT (" + ssqltable + ", RESEED, 0)";
                    SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                    SqlCommand sqlcmd = new SqlCommand(sclearsql, sqlconn);
                    sqlconn.Open();
                    sqlcmd.ExecuteNonQuery();
                    sqlconn.Close();
                    SqlCommand sqlcmd2 = new SqlCommand(resedsql, sqlconn);
                    sqlconn.Open();
                    sqlcmd2.ExecuteNonQuery();
                    sqlconn.Close();
                    OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
                    OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                    oledbconn.Open();
                    OleDbDataReader dr = oledbcmd.ExecuteReader();
                    SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                    bulkcopy.DestinationTableName = ssqltable;
                    while (dr.Read())
                    {
                        bulkcopy.WriteToServer(dr);
                    }
                    MessageBox.Show("Import Data to  SQL Successfully");
                    oledbconn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch(Exception ex2)
            {
                MessageBox.Show("Select the destination table first");
            }
        }
    
      
        private void btnSearch_Click(object sender, EventArgs e)
        {
            TableName = listBoxTrackers.SelectedItem.ToString();
            FindInvoiceForm FindForm = new FindInvoiceForm();
            FindForm.ShowDialog();            
        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            //if user clicked Shift+Ins or Ctrl+V (paste from clipboard)
            try
            {
                if ((e.Shift && e.KeyCode == Keys.Insert) || (e.Control && e.KeyCode == Keys.V))
                {
                    char[] rowSplitter = { '\r', '\n' };
                    char[] columnSplitter = { '\t' };
                    //get the text from clipboard
                    IDataObject dataInClipboard = Clipboard.GetDataObject();
                    string stringInClipboard = (string)dataInClipboard.GetData(DataFormats.Text);
                    //split it into lines
                    string[] rowsInClipboard = stringInClipboard.Split(rowSplitter, StringSplitOptions.RemoveEmptyEntries);
                    //get the row and column of selected cell in grid
                    int r = dataGridView1.SelectedCells[0].RowIndex;
                    int c = dataGridView1.SelectedCells[0].ColumnIndex;
                    //add rows into grid to fit clipboard lines
                    if (dataGridView1.Rows.Count < (r + rowsInClipboard.Length))
                    {
                        for (int i = 0; i < rowsInClipboard.Length; i++)
                        {
                            DataRow dr = dt.NewRow();
                            dr["NOTES"] = "";
                            dt.Rows.Add(dr);
                        }
                    }
                    // loop through the lines, split them into cells and place the values in the corresponding cell.
                    for (int iRow = 0; iRow < rowsInClipboard.Length; iRow++)
                    {
                        //split row into cell values
                        string[] valuesInRow = rowsInClipboard[iRow].Split(columnSplitter);
                        //cycle through cell values
                        for (int iCol = 0; iCol < valuesInRow.Length; iCol++)
                        {
                            //assign cell value, only if it within columns of the grid
                            if (dataGridView1.ColumnCount - 1 >= c + iCol)
                            {
                            
                                dataGridView1.Rows[r + iRow].Cells[c + iCol].Value = valuesInRow[iCol];
                            }
                        }                    
                    }
                }                            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs anError)
        {
   
            MessageBox.Show("Error: Please check if the input value corresponds to the cells format");
            
            if ((anError.Exception) is ConstraintException)
            {
                DataGridView view = (DataGridView)sender;
                view.Rows[anError.RowIndex].ErrorText = "an error";
                view.Rows[anError.RowIndex].Cells[anError.ColumnIndex].ErrorText = "an error";

                anError.ThrowException = false;
            }
        }
    }
    }