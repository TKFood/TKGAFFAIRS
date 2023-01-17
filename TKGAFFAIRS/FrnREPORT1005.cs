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
using TKITDLL;

namespace TKGAFFAIRS
{
    public partial class FrnREPORT1005 : Form
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

        public FrnREPORT1005()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDATES, string EDATES)
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\1005.雜項採購單總表.frx");

            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL(SDATES, EDATES);
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string SDATES,string EDATES)
        {
            StringBuilder FASTSQL = new StringBuilder();
          
            FASTSQL.AppendFormat(@"  
                                SELECT 
                                 [ID] AS '採購單單號'
                                ,[GA001] AS '請購日期'
                                ,[GA002] AS '原請購單號'
                                ,[GA003] AS '請購人員'
                                ,[GA004] AS '請購部門'
                                ,[GA005] AS '品名'
                                ,[GA006] AS '規格'
                                ,[GA007] AS '單位'
                                ,[GA008] AS '數量' 
                                ,[GA009] AS '單價'
                                ,[GA010] AS '總價'
                                ,[GA011] AS '供應商'
                                ,[GA012] AS '預計到貨日期'
                                ,[GA013] AS '付款方式'
                                ,[GA014] AS '付款天數'
                                ,[GA015] AS '到貨日期'
                                ,[GA016] AS '驗收數量'
                                ,[GA017] AS '請購序號'
                                ,[GA018] AS '是否已議價'
                                ,[GA099] AS '備註'
                                ,[GA999] AS '負責採購人員'
                                FROM [TKGAFFAIRS].[dbo].[BUYITEMREPORTS]
                                WHERE [GA015]>='{0}' AND [GA015]<='{1}'
                                ", SDATES,EDATES);

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button7_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
        }
        #endregion
    }
}
