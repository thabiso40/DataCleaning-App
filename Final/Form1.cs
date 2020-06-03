using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using Spire.Xls;

namespace Final
{
    public partial class Form1 : Form
    {
        OpenFileDialog open = new OpenFileDialog();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.panel1.Visible = false;
            this.dataGridView1.Visible = false;
            this.dataGridView2.Visible = false;
            this.label1.Visible = false;
            this.label2.Visible = false;
        }

        private void loadExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.panel1.Visible = true;
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (open.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = open.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //fetches excel
            try { 
            string constr = string.Empty;
            FileInfo file = new FileInfo(textBox1.Text);
            string extention = file.Extension;
            switch (extention)
            {
                case ".xls":
                    constr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox1.Text + ";Extended Properties= 'Excel 8.0;HDR=YES;IMEX=1'";
                    break;
                case ".xlsx":
                    constr = @"provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox1.Text + ";Extended Properties= 'Excel 12.0;HDR=YES;IMEX=1'";
                    break;
                default:
                    constr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox1.Text + ";Extended Properties= 'Excel 8.0;HDR=YES;IMEX=1'";
                    break;


            }
            string sqlconstr = @"Server=DESKTOP-L80DKPC\SQLEXPRESS;Database=WILData; Trusted_Connection=True";
            //execute query to erase previous data
            string sclearSql = "delete  from Table12";
            SqlConnection sqlconn = new SqlConnection(sqlconstr);
            SqlCommand sqlcmd = new SqlCommand(sclearSql, sqlconn);
            SqlCommand delclean = new SqlCommand("delete from clean", sqlconn);
            SqlCommand delDup = new SqlCommand("delete from duplicate", sqlconn);
            sqlconn.Open();
            sqlcmd.ExecuteNonQuery();
            delclean.ExecuteNonQuery();
            delDup.ExecuteNonQuery();
            sqlconn.Close();
            //copy bulk data from excel to database
            OleDbConnection MyConnection = new OleDbConnection(constr);
            OleDbCommand MyCommand = new OleDbCommand("select * from [" + textBox2.Text + "$]", MyConnection);
            MyConnection.Open();
            OleDbDataReader dr = MyCommand.ExecuteReader();
            SqlBulkCopy bulkcopy = new SqlBulkCopy(sqlconstr);
            bulkcopy.DestinationTableName = "Table12";
                
                do
                {
                    bulkcopy.WriteToServer(dr);
                } while (dr.Read());
            MessageBox.Show("successful ", "Export To DataBase", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            MyConnection.Close();
        }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);


            }
}

        private void viewToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cleanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable tableclean = new DataTable();
                this.dataGridView1.Visible = true;
            this.label1.Visible = true;
            SqlConnection CONNECT = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\FinalData.mdf;Integrated Security=True");
            SqlCommand command = new SqlCommand("SELECT * FROM clean",CONNECT);
            SqlDataAdapter adapt = new SqlDataAdapter(command.CommandText, CONNECT);
            adapt.Fill(tableclean);
            dataGridView1.DataSource = tableclean;
            
            int i = dataGridView1.RowCount;
            label3.Text= "Total Records: " + i.ToString();
        }

        private void duplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable tableclean = new DataTable();
            this.dataGridView2.Visible = true;
            this.label2.Visible = true;
            SqlConnection CONNECT = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\FinalData.mdf;Integrated Security=True");
            SqlCommand command = new SqlCommand("SELECT * FROM duplicate", CONNECT);
            SqlDataAdapter adapt = new SqlDataAdapter(command.CommandText, CONNECT);
            adapt.Fill(tableclean);
            dataGridView2.DataSource = tableclean;
            int i = dataGridView2.RowCount;
            label4.Text ="Total Records: " + i.ToString();
        }

        private void exportExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //export to excel

                
                SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\FinalData.mdf;Integrated Security=True");
                SqlCommand commandor = new SqlCommand("SELECT * FROM Table12", con);//uncleaned
                SqlCommand command = new SqlCommand("SELECT * FROM clean", con);//Clean data
                SqlCommand comm = new SqlCommand("SELECT * FROM duplicate", con);//Duplicats
                SqlCommand com = new SqlCommand("SELECT * FROM InvalidID", con);//invalidID

                SqlDataAdapter firstadapter = new SqlDataAdapter(commandor.CommandText, con);
                SqlDataAdapter dataAdapter = new SqlDataAdapter(command.CommandText, con);
                SqlDataAdapter dataAdapter1 = new SqlDataAdapter(comm.CommandText, con);
                SqlDataAdapter datareaderinvalidId = new SqlDataAdapter(com.CommandText, con);

                string file = open.FileName;
                //datatables tobe filled with data from sql
                DataTable t = new DataTable();
                DataTable table = new DataTable();
                DataTable table2 = new DataTable();
                DataTable tableinvalid = new DataTable();

                dataAdapter.Fill(t);
                firstadapter.Fill(table);
                dataAdapter1.Fill(table2);
                datareaderinvalidId.Fill(tableinvalid);

                // workbooks to export excel
                Workbook book = new Workbook();
                book.CreateEmptySheets(4);

                Worksheet sheet1 = book.Worksheets[0];
                sheet1.Name = textBox2.Text;
                //style sheet
                sheet1.Range["A1:N1"].Style.Font.IsBold = true;
                sheet1.Range["A1:N1"].Style.Color = Color.Gray;
                sheet1.InsertDataTable(table,true,1,1);
                //------------------------------------------------------------
                Worksheet sheet = book.Worksheets[1];
                sheet.Name="clean";
                //style sheet
                sheet.Range["A1:N1"].Style.Font.IsBold = true;
                sheet.Range["A1:N1"].Style.Color = Color.Gray;
                sheet.InsertDataTable(t, true, 1, 1);

                Worksheet sheet2 = book.Worksheets[2];
                sheet2.Name = "Duplicates";
                //style sheet
                sheet2.Range["A1:N1"].Style.Font.IsBold = true;
                sheet2.Range["A1:N1"].Style.Color = Color.Gray;
                sheet2.InsertDataTable(table2, true, 1, 1);

                Worksheet sheet3 = book.Worksheets[3];
                sheet3.Name = "InvalidID";
                //style sheet
                sheet3.Range["A1:N1"].Style.Font.IsBold = true;
                sheet3.Range["A1:N1"].Style.Color = Color.Gray;
                sheet3.InsertDataTable(tableinvalid, true, 1, 1);

                book.SaveToFile(file, ExcelVersion.Version2010);
                
                System.Diagnostics.Process.Start(file);
                MessageBox.Show("successfully exported to excel ", "Export To DataBase", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
        }

        private void cleanDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // sorting data
            try
            {
                SqlConnection sqlconn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\FinalData.mdf;Integrated Security=True");
               
                SqlCommand cmd = new SqlCommand("copy", sqlconn);
                cmd.CommandType = CommandType.StoredProcedure;
               

                
                SqlCommand comm = new SqlCommand("DeleteDuplicate", sqlconn);
                comm.CommandType = CommandType.StoredProcedure;
                

                
                SqlCommand com = new SqlCommand("Invalid", sqlconn);
                com.CommandType = CommandType.StoredProcedure;

                SqlCommand del = new SqlCommand("DEL", sqlconn);
                del.CommandType = CommandType.StoredProcedure;

                //sqlconn.Open();
                //cmd.ExecuteNonQuery();
                //comm.ExecuteNonQuery();
                //com.ExecuteNonQuery();
                //sqlconn.Close();
                sqlconn.Open();
                cmd.ExecuteNonQuery();
                sqlconn.Close();

                sqlconn.Open();
                comm.ExecuteNonQuery();
                sqlconn.Close();

                sqlconn.Open();
                com.ExecuteNonQuery();
                sqlconn.Close();

                sqlconn.Open();
                del.ExecuteNonQuery();
                sqlconn.Close();


                MessageBox.Show("successfully cleaned ", "Export To DataBase", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void composeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
    }
}
