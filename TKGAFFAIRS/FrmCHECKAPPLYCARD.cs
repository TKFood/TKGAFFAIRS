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
using System.Xml;

namespace TKGAFFAIRS
{
    public partial class FrmCHECKAPPLYCARD : Form
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

        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string TaskId;


        public FrmCHECKAPPLYCARD()
        {
            InitializeComponent();

            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
            timer1.Start();
        }


        #region FUNCTION
        private void timer1_Tick(object sender, EventArgs e)
        {
            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SEARCHHREngFrm001(textBox1.Text.Trim());
        }

        public void  SEARCHHREngFrm001(string CARDNO)
        {

            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [HREngFrm001User] AS '申請人',[HREngFrm001Rank] AS '職級',[HREngFrm001OutDate] AS '外出日期',[HREngFrm001Transp] AS '交通工具',[HREngFrm001LicPlate] AS '車牌',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間',[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間'");
                sbSql.AppendFormat(@"  ,[TaskId] AS 'TaskId',[HREngFrm001SN] AS '表單編號',[HREngFrm001Date] AS '申請日期',[HREngFrm001UsrDpt] AS '部門',[HREngFrm001Location] AS '外出地點',[HREngFrm001Agent] AS '代理人',[HREngFrm001Cause] AS '外出原因',[HREngFrm001FF] AS '是否由公司出發',[HREngFrm001CH] AS '是否回廠	',[CRADNO] AS '卡號'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(@"  WHERE [HREngFrm001OutDate]='{0}' AND [CRADNO]='{1}'",DateTime.Now.ToString("yyyy/MM/dd"), CARDNO);
                sbSql.AppendFormat(@"  ORDER BY [HREngFrm001DefOutTime]");
                sbSql.AppendFormat(@"  ");
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
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 10);
                      
                        dataGridView1.AutoResizeColumns();

                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[i];
                            row.Height = 60;
                        }
                            


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

        }

        #endregion

        
    }
}
