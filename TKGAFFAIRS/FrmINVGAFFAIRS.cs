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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHINVGAFFAIRS();
        }

        #endregion

      
    }
}
