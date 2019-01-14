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
    public partial class FrmOWNERMANUsub : Form
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
        SqlDataAdapter adapterTEMP = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderTEMP = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet  dsTEMP2 = new DataSet();

        public FrmOWNERMANUsub()
        {
            InitializeComponent();
        }

        public FrmOWNERMANUsub(string ID1, string ID2, string ID3, string ID4, string ID5, string ID6, string ID7, string ID8)
        {
            InitializeComponent();

            textBox1.Text = ID1;
            textBox2.Text = ID2;
            textBox3.Text = ID3;
            textBox4.Text = ID4;
            textBox5.Text = ID5;
            textBox6.Text = ID6;
            textBox7.Text = ID7;
            textBox8.Text = ID8;
        }

        #region FUNCTION
        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            SEARCHNAME1();

            FINDDEP1();
        }

        public void SEARCHNAME1()
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

                sbSql.AppendFormat(@" SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV001='{0}' ", textBox9.Text);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    textBox10.Text = null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        textBox10.Text = ds4.Tables["TEMPds4"].Rows[0]["MV002"].ToString();


                    }

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

        public void FINDDEP1()
        {
            DataSet dsTEMP2 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" SELECT [CnName] ,[Department].[DepartmentId],[Department].[Code],[Name],[Department].[Code] AS 'DEPID'     FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] WHERE [Employee].DepartmentId=[Department].DepartmentId AND [Employee].Code='{0}'", textBox9.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP2.Clear();
                adapterTEMP.Fill(dsTEMP2, "dsTEMP2");
                sqlConn.Close();


                if (dsTEMP2.Tables["dsTEMP2"].Rows.Count == 0)
                {
                    textBox11.Text = null;
                    textBox12.Text = null;
                }
                else
                {
                    if (dsTEMP2.Tables["dsTEMP2"].Rows.Count >= 1)
                    {
                        textBox12.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["Name"].ToString();
                        textBox11.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["DEPID"].ToString();

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
        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            //SEARCHNAME2();

            //FINDDEP2();
        }
        public void SEARCHNAME2()
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

                sbSql.AppendFormat(@" SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV002='{0}' ", textBox10.Text);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    textBox9.Text = null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        textBox9.Text = ds4.Tables["TEMPds4"].Rows[0]["MV002"].ToString();


                    }

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

        public void FINDDEP2()
        {
            DataSet dsTEMP2 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" SELECT [CnName] ,[Department].[DepartmentId],[Department].[Code],[Name],[Department].[Code] AS 'DEPID'     FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] WHERE [Employee].DepartmentId=[Department].DepartmentId AND [Employee].CnName='{0}'", textBox10.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP2.Clear();
                adapterTEMP.Fill(dsTEMP2, "dsTEMP2");
                sqlConn.Close();


                if (dsTEMP2.Tables["dsTEMP2"].Rows.Count == 0)
                {
                    textBox11.Text = null;
                    textBox12.Text = null;
                }
                else
                {
                    if (dsTEMP2.Tables["dsTEMP2"].Rows.Count >= 1)
                    {
                        textBox12.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["Name"].ToString();
                        textBox11.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["DEPID"].ToString();

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
        public void UPDATEOWNERMANU()
        {

        }

      
        #endregion

        #region BUTTON


        private void button7_Click(object sender, EventArgs e)
        {
            UPDATEOWNERMANU();
        }

        #endregion

       
    }
}
