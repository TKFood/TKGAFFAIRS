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
    public partial class FrmBUYUPDATE : Form
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
        int result;
        Thread TD;

        string STATUS = null;
        string BUYNO;
        string OLDBUYNO;
        string CHECKYN = "N";

        public FrmBUYUPDATE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [BUYDATES] AS '請購日期',[BUYNO] AS '請購編號',[NAME] AS '請購人員',[DEP] AS '請購部門'");
                sbSql.AppendFormat(@"  ,[BUYNAME] AS '品名',[SPEC] AS '規格',[VENDOR] AS '供應商',[NUM] AS '數量',[UNIT] AS '單位'");
                sbSql.AppendFormat(@"  ,[PRICES] AS '單價',[TMONEY] AS '總價',[INDATES] AS '到貨日期',[CHECKNUM] AS '驗收數量'");
                sbSql.AppendFormat(@"  ,[SIGN] AS '簽名',[REMARK] AS '備考'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[BUYITEM]");
                sbSql.AppendFormat(@"  WHERE [BUYDATES]>='{0}' AND [BUYDATES]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ");
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
        public void Search2()
        {
            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [BUYDATES] AS '請購日期',[BUYNO] AS '請購編號',[NAME] AS '請購人員',[DEP] AS '請購部門'");
                sbSql.AppendFormat(@"  ,[BUYNAME] AS '品名',[SPEC] AS '規格',[VENDOR] AS '供應商',[NUM] AS '數量',[UNIT] AS '單位'");
                sbSql.AppendFormat(@"  ,[PRICES] AS '單價',[TMONEY] AS '總價',[INDATES] AS '到貨日期',[CHECKNUM] AS '驗收數量'");
                sbSql.AppendFormat(@"  ,[SIGN] AS '簽名',[REMARK] AS '備考'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[BUYITEM]");
                sbSql.AppendFormat(@"  WHERE [BUYDATES]>='{0}' AND [BUYDATES]<='{1}'", dateTimePicker5.Value.ToString("yyyy/MM/dd"), dateTimePicker6.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();

                if (CHECKYN.Equals("N"))
                {
                    //建立一個DataGridView的Column物件及其內容
                    DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                    dgvc.Width = 40;
                    dgvc.Name = "選取";

                    this.dataGridView2.Columns.Insert(0, dgvc);
                    CHECKYN = "Y";
                }

                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
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

        public void UPDATE()
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [TKGAFFAIRS].[dbo].[BUYITEM]");
                sbSql.AppendFormat(" SET [BUYDATES]='{0}',[BUYNO]='{1}',[NAME]='{2}',[DEP]='{3}',[BUYNAME]='{4}',[SPEC]='{5}',[VENDOR]='{6}'", dateTimePicker3.Value.ToString("yyyy/MM/dd"), textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text);
                sbSql.AppendFormat(" ,[NUM]={0},[UNIT]='{1}',[PRICES]={2},[TMONEY]={3},[INDATES]='{4}',[CHECKNUM]={5},[SIGN]='',[REMARK]='{6}'", textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, dateTimePicker4.Value.ToString("yyyy/MM/dd"), textBox13.Text, textBox14.Text);
                sbSql.AppendFormat(" WHERE [BUYNO]='{0}'", OLDBUYNO);
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
        public string GETBUYNO()
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX([BUYNO]),'000000000000') AS BUYNO");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[BUYITEM] ");
                sbSql.AppendFormat(@"  WHERE [BUYNO] LIKE 'B{0}%'", dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        BUYNO = SETBUYNO(ds4.Tables["TEMPds4"].Rows[0]["BUYNO"].ToString());
                        return BUYNO;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SETBUYNO(string BUYNO)
        {
            if (BUYNO.Equals("000000000000"))
            {
                return "B" + dateTimePicker3.Value.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(BUYNO.Substring(9, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return "B" + dateTimePicker3.Value.ToString("yyyyMMdd") + temp.ToString();
            }
        }
        public void ADD()
        {
            try
            {
                textBox2.Text = GETBUYNO();
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKGAFFAIRS].[dbo].[BUYITEM]");
                sbSql.AppendFormat(" ([BUYDATES],[BUYNO],[NAME],[DEP],[BUYNAME],[SPEC],[VENDOR],[NUM],[UNIT],[PRICES],[TMONEY],[INDATES],[CHECKNUM],[SIGN],[REMARK])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}',{9},{10},'{11}',{12},'{13}','{14}')", dateTimePicker3.Value.ToString("yyyy/MM/dd"), textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, dateTimePicker4.Value.ToString("yyyy/MM/dd"), textBox13.Text, null, textBox14.Text);
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

                sbSql.AppendFormat(" DELETE [TKGAFFAIRS].[dbo].[BUYITEM]");
                sbSql.AppendFormat(" WHERE [BUYNO]='{0}'", textBox2.Text);
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
        public void SETSTATUS()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = "0";
            textBox9.Text = null;
            textBox10.Text = "0";
            textBox11.Text = "0";
            textBox13.Text = "0";
            textBox14.Text = null;


            //textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBox8.ReadOnly = false;
            textBox9.ReadOnly = false;
            textBox10.ReadOnly = false;
            textBox11.ReadOnly = false;
            textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;

        }
        public void SETSTATUS2()
        {
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBox8.ReadOnly = false;
            textBox9.ReadOnly = false;
            textBox10.ReadOnly = false;
            textBox11.ReadOnly = false;
            textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;
        }

        public void SETSTAUSFIANL()
        {
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox5.ReadOnly = true;
            textBox6.ReadOnly = true;
            textBox7.ReadOnly = true;
            textBox8.ReadOnly = true;
            textBox9.ReadOnly = true;
            textBox10.ReadOnly = true;
            textBox11.ReadOnly = true;
            textBox13.ReadOnly = true;
            textBox14.ReadOnly = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    dateTimePicker3.Value = Convert.ToDateTime(row.Cells["請購日期"].Value.ToString());
                    dateTimePicker4.Value = Convert.ToDateTime(row.Cells["到貨日期"].Value.ToString());
                    textBox2.Text = row.Cells["請購編號"].Value.ToString();
                    textBox3.Text = row.Cells["請購人員"].Value.ToString();
                    textBox4.Text = row.Cells["請購部門"].Value.ToString();
                    textBox5.Text = row.Cells["品名"].Value.ToString();
                    textBox6.Text = row.Cells["規格"].Value.ToString();
                    textBox7.Text = row.Cells["供應商"].Value.ToString();
                    textBox8.Text = row.Cells["數量"].Value.ToString();
                    textBox9.Text = row.Cells["單位"].Value.ToString();
                    textBox10.Text = row.Cells["單價"].Value.ToString();
                    textBox11.Text = row.Cells["總價"].Value.ToString();
                    textBox13.Text = row.Cells["驗收數量"].Value.ToString();
                    textBox14.Text = row.Cells["備考"].Value.ToString();


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
                    textBox9.Text = null;
                    textBox10.Text = null;
                    textBox11.Text = null;
                    textBox13.Text = null;
                    textBox14.Text = null;

                }
            }
        }

        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\總務課每日採購請購項目總表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            string BUYNOSERIAL = GETBUYNOSERIAL();


            FASTSQL.AppendFormat(@" SELECT [BUYDATES] AS '請購日期',[BUYNO] AS '請購編號',[NAME] AS '請購人員',[DEP] AS '請購部門' ");
            FASTSQL.AppendFormat(@"  ,[BUYNAME] AS '品名',[SPEC] AS '規格',[VENDOR] AS '供應商',[NUM] AS '數量',[UNIT] AS '單位'");
            FASTSQL.AppendFormat(@"  ,[PRICES] AS '單價',[TMONEY] AS '總價',[INDATES] AS '到貨日期',[CHECKNUM] AS '驗收數量'");
            FASTSQL.AppendFormat(@"  ,[SIGN] AS '簽名',[REMARK] AS '備考'");
            FASTSQL.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[BUYITEM]");
            FASTSQL.AppendFormat(@" WHERE [BUYDATES]>='{0}' AND [BUYDATES]<='{1}' ", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  AND [BUYNO] IN ({0})", BUYNOSERIAL.ToString());
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public string GETBUYNOSERIAL()
        {
            string BUYNOSERIAL = null;

            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    try
                    {
                        BUYNOSERIAL = BUYNOSERIAL + "'" + dr.Cells["請購編號"].Value.ToString() + "',";
                    }
                    catch
                    {

                    }

                    finally
                    {

                    }
                }
            }
            return BUYNOSERIAL = BUYNOSERIAL + "''";
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            STATUS = "ADD";
            SETSTATUS();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            STATUS = "EDIT";
            OLDBUYNO = textBox2.Text;
            SETSTATUS2();
        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (STATUS.Equals("EDIT"))
            {
                UPDATE();
            }
            else if (STATUS.Equals("ADD"))
            {
                ADD();
            }

            STATUS = null;

            SETSTAUSFIANL();

            Search();
        }

        private void button4_Click(object sender, EventArgs e)
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

            Search();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            Search2();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
