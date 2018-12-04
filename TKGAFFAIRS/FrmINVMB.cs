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
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using FastReport;
using FastReport.Data;

namespace TKGAFFAIRS
{
    public partial class FrmINVMB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        string STATUS=null;
        int result;
        Thread TD;

        public FrmINVMB()
        {
            InitializeComponent();

            comboBox1load();
        }

        #region FUNCTION

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [KIND],[NAME] FROM [TKGAFFAIRS].[dbo].[INVKIND]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("KIND", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "KIND";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void SEARCHINVMB()
        {
            ds.Clear();

            StringBuilder NAME = new StringBuilder();
            

            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                NAME.AppendFormat(@" WHERE ( [MB001] LIKE '%{0}%' OR [MB002] LIKE '%{0}%')", textBox1.Text);
            }

          

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [KIND] AS '類別',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVMB]");
                sbSql.AppendFormat(@"  {0}", NAME.ToString());
                sbSql.AppendFormat(@"  ORDER BY [KIND], [MB001] ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();


                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox2.Text = row.Cells["品號"].Value.ToString();
                    textBox3.Text = row.Cells["品名"].Value.ToString();
                    textBox4.Text = row.Cells["規格"].Value.ToString();
                    comboBox1.Text = row.Cells["類別"].Value.ToString();

                }
                else
                {
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    comboBox1.Text = null;

                }
            }
        }

        public void SETTEXT()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            comboBox1.Text = null;

            textBox2.ReadOnly = false;
        }

        public void SETTEXT2()
        {
            textBox2.ReadOnly = true;
        }
        public void ADDINVMB()
        {
            try
            {
                
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKGAFFAIRS].[dbo].[INVMB]");
                sbSql.AppendFormat(" ([MB001],[MB002],[MB003],[KIND] )");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')",textBox2.Text,textBox3.Text,textBox4.Text, comboBox1.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEINVMB()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
               
                sbSql.AppendFormat(" UPDATE [TKGAFFAIRS].[dbo].[INVMB]");
                sbSql.AppendFormat(" SET [MB002]='{0}',[MB003]='{1}',[KIND]='{2}'", textBox3.Text, textBox4.Text, comboBox1.Text);
                sbSql.AppendFormat("WHERE [MB001] ='{0}' ", textBox2.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DEL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKGAFFAIRS].[dbo].[INVMB]");
                sbSql.AppendFormat(" WHERE [MB001]='{0}'", textBox2.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHINVMB();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETTEXT();
            STATUS = "ADD";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(STATUS.Equals("ADD"))
            {
                ADDINVMB();
            }
            else if(STATUS.Equals("UPDATE"))
            {
                UPDATEINVMB();
            }

            STATUS = null; ;

            SETTEXT2();
            SEARCHINVMB();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            STATUS = null;
            string message = textBox2.Text + " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DEL();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHINVMB();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            STATUS = "UPDATE";
        }


        #endregion


    }
}
