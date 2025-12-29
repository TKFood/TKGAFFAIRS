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
    public partial class FrmBOXMEAL : Form
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


        public FrmBOXMEAL()
        {
            InitializeComponent();
        }
        private void FrmBOXMEAL_Load(object sender, EventArgs e)
        {
            DataTable DT = FIND_BOXEDMEAL();
            if(DT!=null && DT.Rows.Count>=1)
            {
                textBox1.Text = DT.Rows[0]["PARANAME"].ToString();
            }
            else
            {
                textBox1.Text = "0";
            }
            
        }
        #region FUNCTION

        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\伙食月統計.frx");

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
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL() 
        {
            int MONEYS = int.Parse(textBox1.Text);
            StringBuilder FASTSQL = new StringBuilder();
                               
              
            FASTSQL.AppendFormat(@" 

                             SELECT [ID]+[NAME] AS '姓名',SUBSTRING(CONVERT(NVARCHAR,[DATE],112),5,4) AS '日期',SUM([NUM]) AS '數量' 
                             , MEALNAME
                            ,{2}*SUM([NUM]) AS '金額' 
                            FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER],[TKBOXEDMEAL].[dbo].[MEAL]
                            WHERE 1=1
                            AND [LOCALEMPORDER].MEAL=[MEAL].MEAL
                            AND CONVERT(NVARCHAR,[DATE],112)>='{0}' AND CONVERT(NVARCHAR,[DATE],112)<='{1}'
                            GROUP BY [ID]+[NAME], MEALNAME,CONVERT(NVARCHAR,[DATE],112)

                            ", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"),MONEYS);


            return FASTSQL.ToString();
        }

        public DataTable FIND_BOXEDMEAL()
        {
            ds.Clear();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                sbSql.AppendFormat(@"  
                                    SELECT  
                                    [ID]
                                    ,[KIND]
                                    ,[PARAID]
                                    ,[PARANAME]
                                    FROM [TKGAFFAIRS].[dbo].[TBPARA]
                                    WHERE [KIND]='BOXEDMEAL'
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds.Tables["TEMPds1"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }
        #endregion

        #region BUTTON
        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion

       
    }
}
