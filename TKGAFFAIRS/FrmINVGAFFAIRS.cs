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
    public partial class FrmINVGAFFAIRS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapterTEMP = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderTEMP = new SqlCommandBuilder();

        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();


        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet dsTEMP = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        string STATUS = null;
        int result;
        Thread TD;

        public FrmINVGAFFAIRS()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();

        }

        #region FUNCTION

        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ME002";
            comboBox1.DisplayMember = "ME002";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT [MB001] ,[MB002] ,[MB003] FROM [TKGAFFAIRS].[dbo].INVMB ORDER BY [KIND],[MB001]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MB002";
            comboBox2.DisplayMember = "MB002";
            sqlConn.Close();

           
        }

        public void comboBox3load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ME002";
            comboBox3.DisplayMember = "ME002";
            sqlConn.Close();


        }

        public void comboBox4load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT [MB001] ,[MB002] ,[MB003] FROM [TKGAFFAIRS].[dbo].INVMB ORDER BY [KIND],[MB001]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "MB002";
            comboBox4.DisplayMember = "MB002";
            sqlConn.Close();


        }

        public void SEARCHINVGAFFAIRS()
        {
            ds.Clear();

            StringBuilder NAME = new StringBuilder();


            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                NAME.AppendFormat(@" WHERE ( [MB001] LIKE '%{0}%' OR [MB002] LIKE '%{0}%')", textBox1.Text);
            }



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

                sbSql.AppendFormat(@"  SELECT [MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',SUM([NUM]) AS '庫存數量',AVG([MONEY]) AS '平均單價',SUM([TOTALMONEY])  AS '庫存金額'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(@"  {0}", NAME.ToString());
                sbSql.AppendFormat(@"  GROUP BY [MB001],[MB002],[MB003]");
                sbSql.AppendFormat(@"  ORDER BY [MB001] ");
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = FINDCMSME();
        }

        public string FINDCMSME()
        {
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

             
                sbSql.AppendFormat(@"  SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 LIKE '%{0}%' ORDER BY ME001",comboBox1.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["ME001"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }

        public string FINDCMSME2()
        {
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


                sbSql.AppendFormat(@"  SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 LIKE '%{0}%' ORDER BY ME001", comboBox3.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["ME001"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }


        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = FINDCMSMV();
            FINDDEP1();
        }

        public string FINDCMSMV()
        {
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


                sbSql.AppendFormat(@"  SELECT TOP 1 [CnName] ,[DepartmentId],[Code]   FROM [HRMDB].[dbo].[Employee] WHERE Code='{0}'", textBox3.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["CnName"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }
        
        public void FINDDEP1()
        {
            DataSet dsTEMP2 = new DataSet();

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


                sbSql.AppendFormat(@" SELECT [CnName] ,[Department].[DepartmentId],[Department].[Code],[Name],[Department].[Code] AS 'DEPID'     FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] WHERE [Employee].DepartmentId=[Department].DepartmentId AND [Employee].Code='{0}'", textBox3.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP2.Clear();
                adapterTEMP.Fill(dsTEMP2, "dsTEMP2");
                sqlConn.Close();


                if (dsTEMP2.Tables["dsTEMP2"].Rows.Count == 0)
                {
                    comboBox1.Text = null;
                    textBox2.Text = null;
                }
                else
                {
                    if (dsTEMP2.Tables["dsTEMP2"].Rows.Count >= 1)
                    {
                        comboBox1.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["Name"].ToString();
                        textBox2.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["DEPID"].ToString();

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

        public void FINDDEP2()
        {
            DataSet dsTEMP2 = new DataSet();

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


                sbSql.AppendFormat(@" SELECT [CnName] ,[Department].[DepartmentId],[Department].[Code],[Name],[Department].[Code] AS 'DEPID'     FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] WHERE [Employee].DepartmentId=[Department].DepartmentId AND [Employee].Code='{0}'", textBox10.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP2.Clear();
                adapterTEMP.Fill(dsTEMP2, "dsTEMP2");
                sqlConn.Close();


                if (dsTEMP2.Tables["dsTEMP2"].Rows.Count == 0)
                {
                    comboBox3.Text = null;
                    textBox9.Text = null;
                }
                else
                {
                    if (dsTEMP2.Tables["dsTEMP2"].Rows.Count >= 1)
                    {
                        comboBox3.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["Name"].ToString();
                        textBox9.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["DEPID"].ToString();

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
        public string FINDCMSMV1()
        {
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


                sbSql.AppendFormat(@"  SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV001='{0}'", textBox3.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["MV002"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }
        public string FINDCMSMV2()
        {
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


                sbSql.AppendFormat(@"  SELECT TOP 1 [CnName] ,[DepartmentId],[Code]   FROM [HRMDB].[dbo].[Employee] WHERE Code='{0}'", textBox10.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["CnName"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            FINDINVMB();
        }

        public void FINDINVMB()
        {
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


                sbSql.AppendFormat(@"   SELECT [MB001],[MB002] ,[MB003] FROM [TKGAFFAIRS].[dbo].INVMB  WHERE MB002='{0}'", comboBox2.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    textBox5.Text = null;
                    textBox6.Text = null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {
                        textBox5.Text = dsTEMP.Tables["dsTEMP"].Rows[0]["MB001"].ToString();
                        textBox6.Text = dsTEMP.Tables["dsTEMP"].Rows[0]["MB003"].ToString();                       

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
        

        public void SEARCHINVGAFFAIRS2()
        {
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

                sbSql.AppendFormat(@"  SELECT [DATES] AS '日期',[DEP] AS '部門',[DEPNAME] AS '部門名',[WID] AS '工號',[NAME] AS '姓名',[KINID] AS '類別',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[NUM] AS '數量',[MONEY] AS '單價',[TOTALMONEY]  AS '金額',[ID]");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(@"  WHERE [NUM]>0");
                sbSql.AppendFormat(@"  AND [DATES]='{0}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();


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

        public void SEARCHINVGAFFAIRS3()
        {
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

                sbSql.AppendFormat(@"  SELECT [DATES] AS '日期',[DEP] AS '部門',[DEPNAME] AS '部門名',[WID] AS '工號',[NAME] AS '姓名',[KINID] AS '類別',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[NUM] AS '數量',[MONEY] AS '單價',[TOTALMONEY]  AS '金額',[ID]");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(@"  WHERE [NUM]<0");
                sbSql.AppendFormat(@"  AND [DATES]='{0}'", dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds3.Tables["ds3"];
                        dataGridView3.AutoResizeColumns();


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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    dateTimePicker1.Value = Convert.ToDateTime(row.Cells["日期"].Value.ToString());
                    comboBox1.Text= row.Cells["部門名"].Value.ToString();
                    comboBox2.Text = row.Cells["品名"].Value.ToString();
                    textBox2.Text = row.Cells["部門"].Value.ToString();
                    textBox3.Text = row.Cells["工號"].Value.ToString();
                    textBox4.Text = row.Cells["姓名"].Value.ToString();
                    textBox5.Text = row.Cells["品號"].Value.ToString();
                    textBox6.Text = row.Cells["規格"].Value.ToString();
                    textBox7.Text = row.Cells["數量"].Value.ToString();
                    textBox8.Text = row.Cells["單價"].Value.ToString();
                    textBox16.Text = row.Cells["金額"].Value.ToString();
                    textBoxID1.Text = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    textBoxID1.Text = null;

                }
            }
        }
        public void SETTEXT1()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = "0";
            textBox8.Text = "0";
            textBoxID1.Text = null;

        }

        public void SETTEXT2()
        {
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = "0";
            textBox15.Text = "0";
            textBoxID2.Text = null;

        }

        public void ADDINVGAFFAIRS1()
        {
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" ([ID],[DATES],[DEP],[DEPNAME],[WID],[NAME],[MB001],[MB002],[MB003],[NUM],[MONEY],[KINID],[TOTALMONEY])");
                sbSql.AppendFormat(" VALUES (NEWID(),'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','採購','{10}')",dateTimePicker1.Value.ToString("yyyy/MM/dd"), textBox2.Text, comboBox1.Text,textBox3.Text,textBox4.Text, textBox5.Text, comboBox2.Text,textBox6.Text,textBox7.Text,textBox8.Text, textBox16.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


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

        public void ADDINVGAFFAIRS2()
        {
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" ([ID],[DATES],[DEP],[DEPNAME],[WID],[NAME],[MB001],[MB002],[MB003],[NUM],[MONEY],[KINID],[TOTALMONEY])");
                sbSql.AppendFormat(" VALUES (NEWID(),'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','領用','{10}')", dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox9.Text, comboBox3.Text, textBox10.Text, textBox11.Text, textBox12.Text, comboBox4.Text, textBox13.Text, textBox14.Text, textBox17.Text, textBox15.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


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
        public void UPDATEGAFFAIRS1()
        {
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE  [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" SET [DATES]='{0}',[DEP]='{1}',[DEPNAME]='{2}',[WID]='{3}',[NAME]='{4}',[MB001]='{5}',[MB002]='{6}',[MB003]='{7}',[NUM]='{8}',[MONEY]='{9}',[TOTALMONEY]='{10}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), textBox2.Text, comboBox1.Text, textBox3.Text, textBox4.Text, textBox5.Text, comboBox2.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox16.Text);
                sbSql.AppendFormat(" WHERE [ID]='{0}'",textBoxID1.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


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

        public void UPDATEGAFFAIRS2()
        {
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE  [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" SET [DATES]='{0}',[DEP]='{1}',[DEPNAME]='{2}',[WID]='{3}',[NAME]='{4}',[MB001]='{5}',[MB002]='{6}',[MB003]='{7}',[NUM]='{8}',[MONEY]='{9}',[TOTALMONEY]='{10}'", dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox9.Text, comboBox3.Text, textBox10.Text, textBox11.Text, textBox12.Text, comboBox4.Text, textBox13.Text, textBox14.Text, textBox17.Text, textBox15.Text);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", textBoxID2.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


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
        public void DELGAFFAIRS1()
        {
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'", textBoxID1.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
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
        public void DELGAFFAIRS2()
        {
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'", textBoxID2.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    dateTimePicker2.Value = Convert.ToDateTime(row.Cells["日期"].Value.ToString());
                    comboBox3.Text = row.Cells["部門名"].Value.ToString();
                    comboBox4.Text = row.Cells["品名"].Value.ToString();
                    textBox9.Text = row.Cells["部門"].Value.ToString();
                    textBox10.Text = row.Cells["工號"].Value.ToString();
                    textBox11.Text = row.Cells["姓名"].Value.ToString();
                    textBox12.Text = row.Cells["品號"].Value.ToString();
                    textBox13.Text = row.Cells["規格"].Value.ToString();
                    textBox14.Text = row.Cells["數量"].Value.ToString();
                    textBox15.Text = row.Cells["金額"].Value.ToString();
                    textBoxID2.Text = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    textBox9.Text = null;
                    textBox10.Text = null;
                    textBox11.Text = null;
                    textBox12.Text = null;
                    textBox13.Text = null;
                    textBox14.Text = null;
                    textBox15.Text = null;
                    textBoxID2.Text = null;


                }
            }
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox9.Text = FINDCMSME2();
        }



        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            FINDINVMB2();
            textBox15.Text = CALNUMCOST();
            textBox17.Text = CALNUMCOST2();
        }
        public void FINDINVMB2()
        {
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


                sbSql.AppendFormat(@"   SELECT [MB001],[MB002] ,[MB003] FROM [TKGAFFAIRS].[dbo].INVMB  WHERE MB002='{0}'", comboBox4.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    textBox12.Text = null;
                    textBox13.Text = null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {
                        textBox12.Text = dsTEMP.Tables["dsTEMP"].Rows[0]["MB001"].ToString();
                        textBox13.Text = dsTEMP.Tables["dsTEMP"].Rows[0]["MB003"].ToString();

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
            textBox11.Text = FINDCMSMV2();
            FINDDEP2();
        }

        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\品號入庫及領用.frx");

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

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyy/MM/dd"));
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT CONVERT(nvarchar,[DATES],112) AS '日期',[DEP] AS '部門',[DEPNAME] AS '部門名',[WID] AS '工號',[NAME] AS '姓名',[KINID] AS '類別',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[NUM] AS '數量',[MONEY] AS '單價',[TOTALMONEY] AS '金額',[ID]");
            FASTSQL.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
            FASTSQL.AppendFormat(@"  WHERE [DATES]>='{0}' AND [DATES]<='{1}'",dateTimePicker3.Value.ToString("yyyy/MM/dd"), dateTimePicker4.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@" ORDER BY [DATES],[DEP] ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {
            if(comboBox5.Text.Equals("用品盤存月表"))
            {
                string SQL;
                Report report1 = new Report();
                report1.Load(@"REPORT\用品盤存月表.frx");

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
                SQL = SETFASETSQL2();
                Table.SelectCommand = SQL;
                
                report1.Preview = previewControl2;
                report1.Show();
            }

            else if(comboBox5.Text.Equals("用品盤存明細表"))
            {
                string SQL;
                Report report1 = new Report();
                report1.Load(@"REPORT\用品盤存明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                SQL = SETFASETSQL3();
                Table.SelectCommand = SQL;

                report1.Preview = previewControl2;
                report1.Show();
            }

            else if(comboBox5.Text.Equals("財務每月用品統計表"))
            {
                string SQL;
                Report report1 = new Report();
                report1.Load(@"REPORT\財務每月用品統計表.frx");

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
                SQL = SETFASETSQL4();
                Table.SelectCommand = SQL;

                report1.Preview = previewControl2;
                report1.Show();
            }

        }

        public string SETFASETSQL2()
        {
            DateTime dt = new DateTime();
            dt = Convert.ToDateTime(dateTimePicker7.Value.ToString("yyyy/MM")+"/1");

            DateTime LASTdt = dt.AddDays(1 - dt.Day).AddDays(-1);
            DateTime FirstDay = dt.AddDays(-dt.Day + 1);
            DateTime LastDay = dt.AddMonths(1).AddDays(-dt.AddMonths(1).Day);


            StringBuilder FASTSQL = new StringBuilder();
            FASTSQL.AppendFormat(@"  SELECT 品號,品名,規格,單位,期初數量,期初金額,本期進貨數量,本期進貨金額,本期領料數量,本期領料金額,(期初數量+本期進貨數量-本期領料數量) AS 期末數量,(期初金額+本期進貨金額-本期領料金額) AS 期末金額");
            FASTSQL.AppendFormat(@"  FROM (");
            FASTSQL.AppendFormat(@"  SELECT [MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[UNIT] AS '單位'");
            FASTSQL.AppendFormat(@"  ,(SELECT ISNULL(SUM(NUM),0) FROM  [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [INVGAFFAIRS].MB001=[INVMB].[MB001] AND DATES<='{0}') AS '期初數量'", LASTdt.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"  ,(SELECT ISNULL(SUM(TOTALMONEY),0) FROM  [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [INVGAFFAIRS].MB001=[INVMB].[MB001] AND DATES<='{0}') AS '期初金額'", LASTdt.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"  ,(SELECT ISNULL(SUM(NUM),0) FROM  [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [INVGAFFAIRS].MB001=[INVMB].[MB001] AND KINID='採購' AND DATES>='{0}' AND DATES<='{1}') AS '本期進貨數量'", FirstDay.ToString("yyyy/MM/dd"), LastDay.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"  ,(SELECT ISNULL(SUM(TOTALMONEY),0) FROM  [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [INVGAFFAIRS].MB001=[INVMB].[MB001]  AND KINID='採購'  AND DATES>='{0}' AND DATES<='{1}') AS '本期進貨金額'", FirstDay.ToString("yyyy/MM/dd"), LastDay.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"  ,(SELECT ISNULL(SUM(NUM),0)*-1 FROM  [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [INVGAFFAIRS].MB001=[INVMB].[MB001] AND KINID='領用' AND DATES>='{0}' AND DATES<='{1}') AS '本期領料數量'", FirstDay.ToString("yyyy/MM/dd"), LastDay.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"  ,(SELECT ISNULL(SUM(TOTALMONEY),0)*-1 FROM  [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [INVGAFFAIRS].MB001=[INVMB].[MB001]  AND KINID='領用'  AND DATES>='{0}' AND DATES<='{1}') AS '本期領料金額'", FirstDay.ToString("yyyy/MM/dd"), LastDay.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVMB]");
            FASTSQL.AppendFormat(@"  ) AS TEMP");
            FASTSQL.AppendFormat(@"  ORDER BY 品號");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public string SETFASETSQL3()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT CONVERT(nvarchar,[DATES],112) AS '日期',[DEP] AS '部門',[DEPNAME] AS '部門名',[WID] AS '工號',[NAME] AS '姓名',[KINID] AS '類別',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[NUM] AS '數量',[MONEY] AS '單價',[TOTALMONEY] AS '金額',[ID]");
            FASTSQL.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS]");
            FASTSQL.AppendFormat(@"  WHERE [DATES]>='{0}' AND [DATES]<='{1}'", dateTimePicker5.Value.ToString("yyyy/MM/dd"), dateTimePicker6.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@" ORDER BY [DATES],[DEP] ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public string SETFASETSQL4()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"  SELECT [INVMB].[KIND] AS '類別',[DEP] AS '部門',[DEPNAME] AS '部門名',SUM([TOTALMONEY])*-1 AS '金額'");
            FASTSQL.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS],[TKGAFFAIRS].[dbo].[INVMB]");
            FASTSQL.AppendFormat(@"  WHERE [INVGAFFAIRS].[MB001]= [INVMB].[MB001]");
            FASTSQL.AppendFormat(@"  AND [DATES]>='{0}' AND [DATES]<='{1}'", dateTimePicker5.Value.ToString("yyyy/MM/dd"), dateTimePicker6.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"  AND KINID='領用'");
            FASTSQL.AppendFormat(@"  GROUP BY  [INVMB].[KIND],[DEP],[DEPNAME]");
            FASTSQL.AppendFormat(@"  ORDER BY  [INVMB].[KIND],[DEP],[DEPNAME]");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            FINDINVMB2();
            textBox15.Text = CALNUMCOST();
            textBox17.Text = CALNUMCOST2();
        }
        public string CALNUMCOST()
        {
            if(!string.IsNullOrEmpty(comboBox4.Text))
            {
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


                    sbSql.AppendFormat(@"   SELECT CONVERT(DECIMAL(16,4),SUM([TOTALMONEY])/SUM(NUM))  AS COST FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [MB001]='{0}'", textBox12.Text.ToString());


                    adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                    sqlConn.Open();
                    dsTEMP.Clear();
                    adapterTEMP.Fill(dsTEMP, "dsTEMP");
                    sqlConn.Close();


                    if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                        {
                            return ( Convert.ToInt32(textBox14.Text)*Convert.ToDecimal(dsTEMP.Tables["dsTEMP"].Rows[0]["COST"].ToString())).ToString();                           

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
            return null;
        }

        public string CALNUMCOST2()
        {
            if (!string.IsNullOrEmpty(comboBox4.Text))
            {
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


                    sbSql.AppendFormat(@"   SELECT CONVERT(DECIMAL(16,4),SUM([TOTALMONEY])/SUM(NUM))  AS COST FROM [TKGAFFAIRS].[dbo].[INVGAFFAIRS] WHERE [MB001]='{0}'", textBox12.Text.ToString());


                    adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                    sqlConn.Open();
                    dsTEMP.Clear();
                    adapterTEMP.Fill(dsTEMP, "dsTEMP");
                    sqlConn.Close();


                    if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                        {
                            return ( Convert.ToDecimal(dsTEMP.Tables["dsTEMP"].Rows[0]["COST"].ToString())).ToString();

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
            return null;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = FINDCMSMV3();
        }
        public string FINDCMSMV3()
        {
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


                sbSql.AppendFormat(@"  SELECT TOP 1 [CnName] ,[DepartmentId],[Code]   FROM [HRMDB].[dbo].[Employee] WHERE CnName='{0}' ", textBox4.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["Code"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            textBox10.Text = FINDCMSMV4();
        }

        public string FINDCMSMV4()
        {
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


                sbSql.AppendFormat(@"  SELECT TOP 1 [CnName] ,[DepartmentId],[Code]   FROM [HRMDB].[dbo].[Employee] WHERE [CnName]='{0}'", textBox11.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP.Clear();
                adapterTEMP.Fill(dsTEMP, "dsTEMP");
                sqlConn.Close();


                if (dsTEMP.Tables["dsTEMP"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (dsTEMP.Tables["dsTEMP"].Rows.Count >= 1)
                    {

                        return dsTEMP.Tables["dsTEMP"].Rows[0]["Code"].ToString();

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            return null;
        }



        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHINVGAFFAIRS();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETTEXT1();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(textBox7.Text) > 0 && Convert.ToDecimal(textBox8.Text) > 0 && !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(textBox5.Text))
            {
                if (string.IsNullOrEmpty(textBoxID1.Text))
                {
                    ADDINVGAFFAIRS1();
                }
                else if (!string.IsNullOrEmpty(textBoxID1.Text))
                {
                    UPDATEGAFFAIRS1();
                }

            }
            else
            {
                MessageBox.Show("數量或金額不得小於0 或部門 或品號 沒有資料");
            }

            SEARCHINVGAFFAIRS2();
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            string message = textBox5.Text + " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELGAFFAIRS1();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHINVGAFFAIRS2();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            SEARCHINVGAFFAIRS2();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SETTEXT2();
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(textBox14.Text) < 0 && Convert.ToDecimal(textBox15.Text) < 0&&!string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrEmpty(textBox12.Text))
            {
                if (string.IsNullOrEmpty(textBoxID2.Text))
                {
                    ADDINVGAFFAIRS2();
                }
                else if (!string.IsNullOrEmpty(textBoxID2.Text))
                {
                    UPDATEGAFFAIRS2();
                }

            }
            else
            {
                MessageBox.Show("數量或金額不得大於0 或部門 或品號 沒有資料 ");
            }

            SEARCHINVGAFFAIRS3();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string message = textBox12.Text + " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELGAFFAIRS2();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHINVGAFFAIRS3();
        }
        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHINVGAFFAIRS3();
        }


        private void button12_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }






        #endregion

      
    }
}
