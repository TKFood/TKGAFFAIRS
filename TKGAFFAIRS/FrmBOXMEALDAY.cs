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
    public partial class FrmBOXMEALDAY : Form
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

        public FrmBOXMEALDAY()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\當日伙食統計表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyy/MM/dd"));
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT [ID] AS '工號',[NAME] AS '姓名',SUM([NUM]) AS '訂餐量',[MEAL].[MEALNAME] AS '午/晚餐',[MEALDISH].[DISHNAME] AS '葷/素','' AS '用餐'");
            FASTSQL.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER]");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEALDISH] ON  [MEALDISH].[DISH]=[LOCALEMPORDER].[DISH]");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEAL] ON [MEAL].[MEAL]=[LOCALEMPORDER].[MEAL]");
            FASTSQL.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATE],112)>='{0}' AND CONVERT(NVARCHAR,[DATE],112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  GROUP BY [ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
            FASTSQL.AppendFormat(@"  ORDER BY [ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
