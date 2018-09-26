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
        private void FrmBOXMEALORDER_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timer1.Interval = 1000;
            timer1.Start();
        }
      

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString();

        }
        public void Search()
        {
            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                
                sbSql.AppendFormat(@" SELECT [LOCALEMPORDER].[ID] AS '工號',[LOCALEMPORDER].[NAME] AS '姓名',[MEAL].[MEALNAME] AS '餐別',[MEALDISH].[DISHNAME] AS '葷素' ,[LOCALEMPORDER].[NUM] AS '訂餐量'");
                sbSql.AppendFormat(@" ,[LOCALEMPORDER].[SERNO],[LOCALEMPORDER].[CARDNO],[LOCALEMPORDER].[DATE],[LOCALEMPORDER].[MEAL],[LOCALEMPORDER].[DISH] ");
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

        public void ORDERAdd(string Meal, string Dish, string OrderBoxed)
        {
            try
            {

                InsertsbSql.Clear();
                sbSql.Clear();
                //ADD COPTC

                if (Meal.Equals("10+20"))
                {
                    DataSet ds1 = new DataSet();
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE  CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}'  ", EmployeeID, "10");

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter.Fill(ds1, "TEMPds1");
                    sqlConn.Close();

                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {
                        //InsertsbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') AND [EATNUM]=0", EmployeeID, Meal);
                        Meal = "10";
                        InsertsbSql.Append(" ");
                        InsertsbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] ([SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}','{3}',GETDATE(),'{4}','{5}',1) ",dateTimePicker1.Value.ToString("yyyyMMdd") + DateTime.Now.ToString("HHmmss"), EmployeeID, Name, CardNo, Meal, Dish);
                    }

                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}'  ", EmployeeID, "20");

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter.Fill(ds1, "TEMPds1");
                    sqlConn.Close();
                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {
                        Meal = "20";
                        InsertsbSql.Append(" ");
                        InsertsbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] ([SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}','{3}',GETDATE(),'{4}','{5}',1) ", dateTimePicker1.Value.ToString("yyyyMMdd") + DateTime.Now.ToString("HHmmss"), EmployeeID, Name, CardNo, Meal, Dish);
                    }

                }
                else
                {
                    DataSet ds1 = new DataSet();
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),'{2}', 112) AND [ID]='{0}' AND  [MEAL]='{1}' ", EmployeeID, Meal,dateTimePicker1.Value.ToString("yyyyMMdd"));

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter.Fill(ds1, "TEMPds1");
                    if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                    {
                        Name = ds1.Tables["TEMPds1"].Rows[0][2].ToString();
                    }
                    sqlConn.Close();

                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {
                        InsertsbSql.Append(" ");
                        //InsertsbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' AND [EATNUM]=0 ", EmployeeID, Meal);
                        InsertsbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] ([SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7}) ", dateTimePicker1.Value.ToString("yyyyMMdd")+ DateTime.Now.ToString("HHmmss"), EmployeeID, Name, CardNo, dateTimePicker1.Value.ToString("yyyyMMdd"), Meal, Dish,textBox3.Text);
                    }
                    else
                    {
                        //AutoClosingMessageBox.Show("已經訂過餐了!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + "已經訂過餐了!!!!");
                    }

                }


                if (!string.IsNullOrEmpty(InsertsbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();
                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = InsertsbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();
                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        SHOWMESSAGE(Name + " 訂餐失敗!!");

                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        SHOWMESSAGE(Name + " 訂餐成功!!" + " 訂了: " + OrderBoxed.ToString());
                    }

                    sqlConn.Close();
                }
         

               
            }
            catch
            {

            }
            finally
            {

            }
            
        }

        public void SHOWMESSAGE(String mess)
        {
            MessageBox.Show(mess);
        }

        public void OrderCanel(string Meal, string Dish, string OrderBoxed)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD COPTC

                if (Meal.Equals("10+20"))
                {
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),'{2}', 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') AND [DISH]='{1}' ", EmployeeID, Dish,dateTimePicker1.Value.ToString("yyyyMMdd"));
                }
                else
                {
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),'{3}', 112) AND [ID]='{0}' AND  [MEAL]='{1}' AND [DISH]='{2}' ", EmployeeID, Meal, Dish, dateTimePicker1.Value.ToString("yyyyMMdd"));
                }

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();
                if (result == 0)
                {
                    tran.Rollback();    //交易取消

                    SHOWMESSAGE(Name + " 取消訂餐失敗!!");

                }
                else
                {
                    tran.Commit();      //執行交易  

                    SHOWMESSAGE(Name + " 取消訂餐成功!!" + " 您取消了: " + OrderBoxed.ToString());

                }

                sqlConn.Close();
                Search();
            }
            catch
            {

            }
            finally
            {

            }
            
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
          
            SearchEmplyee();
        }

        public void SearchEmplyee()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT TOP 1  [EmployeeID],[CardNo],[Name] FROM [TKBOXEDMEAL].[dbo].[VEMPLOYEE] WHERE [EmployeeID]='{0}' OR [CardNo]='{0}'", textBox1.Text.ToString());
                sbSql.AppendFormat(@" UNION ALL");
                sbSql.AppendFormat(@" SELECT TOP 1  ME001 collate Chinese_Taiwan_Stroke_CI_AI,ME001 collate Chinese_Taiwan_Stroke_CI_AI,ME002 collate Chinese_Taiwan_Stroke_CI_AI FROM [TK].dbo.[CMSME]  WHERE ME001='{0}'", textBox1.Text.ToString());
                sbSql.AppendFormat(@" ");
                sbSql.AppendFormat(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

                    //SHOWMESSAGE("沒有此員工!!");

                    //textBox1.Text = null;
                    textBox2.Text = textBox1.Text;
                    


                }
                else
                {
                    EmployeeID = ds1.Tables["TEMPds1"].Rows[0][0].ToString();
                    CardNo = ds1.Tables["TEMPds1"].Rows[0][1].ToString();
                    Name = ds1.Tables["TEMPds1"].Rows[0][2].ToString();

                    textBox2.Text = ds1.Tables["TEMPds1"].Rows[0][2].ToString();

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

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "10";
                Dish = "1";
                EmployeeID = textBox1.Text;
                Name = textBox2.Text;
                ORDERAdd(Meal, Dish, OrderBoxed);
            }   
           

            Search();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "10";
                Dish = "2";
                EmployeeID = textBox1.Text;
                Name = textBox2.Text;
                ORDERAdd(Meal, Dish, OrderBoxed);
            }

            Search();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "20";
                Dish = "1";
                EmployeeID = textBox1.Text;
                Name = textBox2.Text;
                ORDERAdd(Meal, Dish, OrderBoxed);
            }
            Search();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "20";
                Dish = "2";
                EmployeeID = textBox1.Text;
                Name = textBox2.Text;
                ORDERAdd(Meal, Dish, OrderBoxed);
            }
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "10";
                Dish = "1";
                EmployeeID = textBox1.Text;

                OrderCanel(Meal, Dish, OrderBoxed);
            }

            Search();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "10";
                Dish = "2";
                EmployeeID = textBox1.Text;

                OrderCanel(Meal, Dish, OrderBoxed);
            }

            Search();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "20";
                Dish = "1";
                EmployeeID = textBox1.Text;

                OrderCanel(Meal, Dish, OrderBoxed);
            }
            Search();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                Meal = "20";
                Dish = "2";
                EmployeeID = textBox1.Text;

                OrderCanel(Meal, Dish, OrderBoxed);
            }
            Search();
        }


        #endregion


    }
}
