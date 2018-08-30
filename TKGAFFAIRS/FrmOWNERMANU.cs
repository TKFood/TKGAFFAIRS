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
    public partial class FrmOWNERMANU : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;
        string BUYNO;
        string OLDBUYNO;

        public FrmOWNERMANU()
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001,ME002    ");
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

            label7.Text = dt.Rows[0]["ME002"].ToString();


        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME001='{0}'    ", comboBox1.Text.ToString());
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);

            sqlConn.Close();

            if(dt.Rows.Count>0)
            {
                label7.Text = dt.Rows[0]["ME002"].ToString();
            }
            else
            {
                label7.Text = "DEP";
            }
            
        }
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            
        }

        public void Search()
        {
            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if(!string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSqlQuery.AppendFormat(@" AND [ID]='{0}'  ",textBox1.Text);
                }

                if (!string.IsNullOrEmpty(textBox2.Text))
                {
                    sbSqlQuery.AppendFormat(@" AND [NAME]='{0}'  ", textBox2.Text);
                }

                sbSql.AppendFormat(@"  SELECT [ID] AS '工號',[NAME] AS '保管人',[DEP] AS '部門',[DEPNAME] AS '單位',[CREATEDATES] AS '建立日期'");
                sbSql.AppendFormat(@"  ,[CLASS] AS '分類',[NO] AS '流水號',[OWNNAME] AS '保管品名',[BRAND] AS '廠牌',[SPEC] AS '規格'");
                sbSql.AppendFormat(@"  ,[PRICES] AS '原價',[NUM] AS '數量',[GIVENAME] AS '發放人',[REMARK] AS '備註'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[OWNERMANU]");
                sbSql.AppendFormat(@"  WHERE DEP='{0}'",comboBox1.Text.ToString());
                sbSql.AppendFormat(sbSqlQuery.ToString());
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

            textBox1.Text = null;
            textBox2.Text = null;
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion

        
    }
}
