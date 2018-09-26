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
    public partial class FrmBOXMEALORDER : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder InsertsbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string OrderBoxed;
        int rownum = 0;
        DateTime startdt;
        DateTime enddt;
        DateTime startdinnerdt;
        DateTime enddinnerdt;
        DateTime comdt;
        string InputID;
        string CardNo;
        string EmployeeID;
        string Name;
        string Meal;
        string Dish;
        string OrderCancel;
        string QueryMeal;
        string Lang = "CH";
        string lastdate = null;
        int messagetime = 3000;

        public FrmBOXMEALORDER()
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

                
                sbSql.AppendFormat(@" SELECT [LOCALEMPORDER].[ID] AS '工號',[LOCALEMPORDER].[NAME] AS '姓名',[MEAL].[MEALNAME] AS '餐別',[MEALDISH].[DISHNAME] AS '葷素' ");
                sbSql.AppendFormat(@" ,[LOCALEMPORDER].[SERNO],[LOCALEMPORDER].[CARDNO],[LOCALEMPORDER].[DATE],[LOCALEMPORDER].[MEAL],[LOCALEMPORDER].[DISH],[LOCALEMPORDER].[NUM] ");
                sbSql.AppendFormat(@" FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] ");
                sbSql.AppendFormat(@" LEFT JOIN [TKBOXEDMEAL].[dbo].[MEAL] ON [MEAL].[MEAL]=[LOCALEMPORDER].[MEAL] ");
                sbSql.AppendFormat(@" LEFT JOIN [TKBOXEDMEAL].[dbo].[MEALDISH] ON [MEALDISH].[DISH]=[LOCALEMPORDER].[DISH] ");
                sbSql.AppendFormat(@" WHERE CONVERT(NVARCHAR,[LOCALEMPORDER].[DATE],112)='{0}'  ",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@" ORDER BY [LOCALEMPORDER].[ID] ");
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
