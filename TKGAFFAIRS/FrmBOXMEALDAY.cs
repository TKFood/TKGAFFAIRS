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
            string SQL2;
            Report report1 = new Report();
            report1.Load(@"REPORT\當日伙食統計表.frx");

            //20210902密
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
            TableDataSource Table1 = report1.GetDataSource("Table1") as TableDataSource;
            SQL = SETFASETSQL();
            //SQL2 = SETFASETSQL2();
            Table.SelectCommand = SQL;

            //Table1.SelectCommand = SQL2;

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyy/MM/dd"));
            report1.Preview = previewControl1;
            report1.Show();

        }
        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112) AS '日期',[ID] AS '工號',[NAME] AS '姓名',SUM([NUM]) AS '訂餐量',[MEAL].[MEALNAME] AS '午/晚餐',[MEALDISH].[DISHNAME] AS '葷/素','' AS '用餐',ME002 AS '部門'");
            FASTSQL.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER]");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEALDISH] ON  [MEALDISH].[DISH]=[LOCALEMPORDER].[DISH]");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEAL] ON [MEAL].[MEAL]=[LOCALEMPORDER].[MEAL]");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSMV ON MV001=[ID]");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSME ON ME001=MV004");
            FASTSQL.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATE],112)>='{0}' AND CONVERT(NVARCHAR,[DATE],112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));            
            FASTSQL.AppendFormat(@"  GROUP BY CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112),ME002,[ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
            FASTSQL.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112),ME002,[ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
            FASTSQL.AppendFormat(@"   ");

            return FASTSQL.ToString();
        }

        //public string SETFASETSQL()
        //{
        //    StringBuilder FASTSQL = new StringBuilder();

        //    FASTSQL.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112) AS '日期',[ID] AS '工號',[NAME] AS '姓名',SUM([NUM]) AS '訂餐量',[MEAL].[MEALNAME] AS '午/晚餐',[MEALDISH].[DISHNAME] AS '葷/素','' AS '用餐',ME002 AS '部門'");
        //    FASTSQL.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEALDISH] ON  [MEALDISH].[DISH]=[LOCALEMPORDER].[DISH]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEAL] ON [MEAL].[MEAL]=[LOCALEMPORDER].[MEAL]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSMV ON MV001=[ID]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSME ON ME001=MV004");
        //    FASTSQL.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATE],112)>='{0}' AND CONVERT(NVARCHAR,[DATE],112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        //    FASTSQL.AppendFormat(@"  AND (ME001 NOT LIKE '103%' AND  [ID]  NOT IN ('160033','170018','180014','200012') )");
        //    FASTSQL.AppendFormat(@"  GROUP BY CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112),ME002,[ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
        //    FASTSQL.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112),ME002,[ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
        //    FASTSQL.AppendFormat(@"   ");

        //    return FASTSQL.ToString();
        //}

        //public string SETFASETSQL2()
        //{
        //    StringBuilder FASTSQL = new StringBuilder();

        //    FASTSQL.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112) AS '日期',[ID] AS '工號',[NAME] AS '姓名',SUM([NUM]) AS '訂餐量',[MEAL].[MEALNAME] AS '午/晚餐',[MEALDISH].[DISHNAME] AS '葷/素','' AS '用餐',ME002 AS '部門'");
        //    FASTSQL.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEALDISH] ON  [MEALDISH].[DISH]=[LOCALEMPORDER].[DISH]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TKBOXEDMEAL].[dbo].[MEAL] ON [MEAL].[MEAL]=[LOCALEMPORDER].[MEAL]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSMV ON MV001=[ID]");
        //    FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSME ON ME001=MV004");
        //    FASTSQL.AppendFormat(@"  WHERE CONVERT(NVARCHAR,[DATE],112)>='{0}' AND CONVERT(NVARCHAR,[DATE],112)<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        //    FASTSQL.AppendFormat(@"  AND (ME001  LIKE '103%' OR  [ID]  IN ('160033','170018','180014','200012') )");
        //    FASTSQL.AppendFormat(@"  GROUP BY CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112),ME002,[ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
        //    FASTSQL.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112),ME002,[ID],[NAME],[MEAL].[MEALNAME] ,[MEALDISH].[DISHNAME]");
        //    FASTSQL.AppendFormat(@"   ");

        //    return FASTSQL.ToString();
        //}


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
