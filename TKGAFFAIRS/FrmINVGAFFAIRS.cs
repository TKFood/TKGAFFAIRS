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
    public partial class FrmINVGAFFAIRS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapterTEMP = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderTEMP = new SqlCommandBuilder();

        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();


        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet dsTEMP = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        string STATUS = null;
        int result;
        Thread TD;

        public FrmINVGAFFAIRS()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            
        }

        #region FUNCTION

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ME001";
            comboBox1.DisplayMember = "ME001";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT [MB001] ,[MB002] ,[MB003] FROM [TKGAFFAIRS].[dbo].INVMB ORDER BY [MB001]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MB001";
            comboBox2.DisplayMember = "MB001";
            sqlConn.Close();

           
        }


        public void SEARCHINVGAFFAIRS()
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

                sbSql.AppendFormat(@"  SELECT [MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',SUM([NUM]) AS '庫存數量',SUM([MONEY]) AS '庫存金額'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(@"  {0}", NAME.ToString());
                sbSql.AppendFormat(@"  GROUP BY [MB001],[MB002],[MB003]");
                sbSql.AppendFormat(@"  ORDER BY [MB001] ");
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = FINDCMSME();
        }

        public string FINDCMSME()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

             
                sbSql.AppendFormat(@"  SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME001 LIKE '%{0}%' ORDER BY ME001",comboBox1.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["ME002"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = FINDCMSMV();
        }

        public string FINDCMSMV()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV001='{0}'", textBox3.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["MV002"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            FINDINVMB();
        }

        public void FINDINVMB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"   SELECT [MB001],[MB002] ,[MB003] FROM [TKGAFFAIRS].[dbo].INVMB  WHERE MB001='{0}'", comboBox2.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    textBox5.Text = null;
                    textBox6.Text = null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {
                        textBox5.Text = dsTEMP.Tables["dsTEMP"].Rows[0]["MB002"].ToString();
                        textBox6.Text = dsTEMP.Tables["dsTEMP"].Rows[0]["MB003"].ToString();                       

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

        public void SEARCHINVGAFFAIRS2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [DATES] AS '日期',[DEP] AS '部門',[DEPNAME] AS '部門名',[WID] AS '工號',[NAME] AS '姓名',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[NUM] AS '數量',[MONEY] AS '金額',[ID]");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(@"  WHERE [NUM]>0");
                sbSql.AppendFormat(@"  AND [DATES]='{0}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();


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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    dateTimePicker1.Value = Convert.ToDateTime(row.Cells["日期"].Value.ToString());
                    comboBox1.Text= row.Cells["部門"].Value.ToString();
                    comboBox2.Text = row.Cells["品號"].Value.ToString();
                    textBox2.Text = row.Cells["部門名"].Value.ToString();
                    textBox3.Text = row.Cells["工號"].Value.ToString();
                    textBox4.Text = row.Cells["姓名"].Value.ToString();
                    textBox5.Text = row.Cells["品名"].Value.ToString();
                    textBox6.Text = row.Cells["規格"].Value.ToString();
                    textBox7.Text = row.Cells["數量"].Value.ToString();
                    textBox8.Text = row.Cells["金額"].Value.ToString();
                    textBoxID1.Text = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    textBoxID1.Text = null;

                }
            }
        }
        public void SETTEXT1()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = "0";
            textBox8.Text = "0";
            textBoxID1.Text = null;

        }

        public void ADDINVGAFFAIRS1()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" ([ID],[DATES],[DEP],[DEPNAME],[WID],[NAME],[MB001],[MB002],[MB003],[NUM],[MONEY])");
                sbSql.AppendFormat(" VALUES (NEWID(),'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')",dateTimePicker1.Value.ToString("yyyy/MM/dd"),comboBox1.Text,textBox2.Text,textBox3.Text,textBox4.Text,comboBox2.Text,textBox5.Text,textBox6.Text,textBox7.Text,textBox8.Text);
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

        public void UPDATEGAFFAIRS1()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE  [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" SET [DATES]='{0}',[DEP]='{1}',[DEPNAME]='{2}',[WID]='{3}',[NAME]='{4}',[MB001]='{5}',[MB002]='{6}',[MB003]='{7}',[NUM]='{8}',[MONEY]='{9}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"), comboBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, comboBox2.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text);
                sbSql.AppendFormat(" WHERE [ID]='{0}'",textBoxID1.Text);
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

        public void DELGAFFAIRS1()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'", textBoxID1.Text);
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
            SEARCHINVGAFFAIRS();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETTEXT1();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(textBoxID1.Text))
            {
                ADDINVGAFFAIRS1();
            }
            else if(!string.IsNullOrEmpty(textBoxID1.Text))
            {
                UPDATEGAFFAIRS1();
            }

            SEARCHINVGAFFAIRS2();
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            string message = textBox5.Text + " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELGAFFAIRS1();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHINVGAFFAIRS2();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            SEARCHINVGAFFAIRS2();
        }



        #endregion

        
    }
}
