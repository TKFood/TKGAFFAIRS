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

namespace TKGAFFAIRS
{
    public partial class FrmBUY : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        public FrmBUY()
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
                sbSql.AppendFormat(@"  WHERE [BUYDATES]>='20180801' AND [BUYDATES]<='20180831'");
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion
    }
}
