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
    public partial class FrmREPORT1004 : Form
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

        public FrmREPORT1004()
        {
            InitializeComponent();

            SETDATE();
        }

        #region FUNCTION

        public void SETDATE()
        {
            dateTimePicker1.Value = Convert.ToDateTime(DateTime.Now.Year +"/"+ DateTime.Now.Month + "/01");
            dateTimePicker2.Value = DateTime.Now.AddMonths(1).AddDays(-DateTime.Now.AddMonths(1).Day);
        }
        public void SETFASTREPORT(string SDATES, string EDATES)
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\1004.總務修繕單.frx");

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

        public string SETFASETSQL(string SDATES, string EDATES)
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"                                 
                                SELECT 
                                [ID]
                                ,[DOC_NBR] AS '表單編號'
                                ,[GAFrm004SN] 
                                ,[GAFrm004SI] AS '申請人'
                                ,[GAFrm004SD] AS '申請單位'
                                ,[GAFrm004Applydates] AS '申請日期'
                                ,[GAFrm004EXdates] 
                                ,[GAFrm004DN] AS '設備名稱'
                                ,[GAFrm004ER] AS '異常情形'
                                FROM [TKGAFFAIRS].[dbo].[UOFGAFIXSNEW]
                                WHERE [GAFrm004Applydates]>='{0}' AND [GAFrm004Applydates]<='{1}'
                                ORDER BY [DOC_NBR],[GAFrm004DN]
                                ", SDATES, EDATES);

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
