using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FOApp;

namespace FOApp
{
    public partial class FindInvoiceForm : Form
    {
        SqlDataAdapter sda;
        SqlCommandBuilder scb;
        DataTable dt;
        public string TableName2 = Form1.TableName;
        static string conString = ConfigurationManager.ConnectionStrings["FOApp.Properties.Settings.SettingString"].ConnectionString;
        SqlConnection con = new SqlConnection(conString);

        public void Display_table()
        {
            try
            {
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM " + TableName2 + " WHERE ([INVOICE #] = '" + txtSearchInvoice.Text + "');";

                sda = new SqlDataAdapter(cmd.CommandText, con);
                dt = new DataTable();
                sda.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                con.Close();
                MessageBox.Show("First select the Tracker from the droplist");

            }
        }
        public FindInvoiceForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Display_table();



        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            

        }

        private void FindInvoiceForm_Load(object sender, EventArgs e)
        {
            
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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT [INVOICE #],COUNT([INVOICE #]) AS 'Coincidence count'  " +
                    " FROM " + TableName2 + " GROUP BY[INVOICE #]" + "HAVING COUNT([INVOICE #]) > 1";
                sda = new SqlDataAdapter(cmd.CommandText, con);
                dt = new DataTable();
                sda.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}