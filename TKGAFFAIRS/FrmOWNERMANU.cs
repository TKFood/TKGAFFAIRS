﻿using System;
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
    public partial class FrmOWNERMANU : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapterTEMP = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderTEMP = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = null;

        string BUYNO;
        string OLDBUYNO;

        string NO;
        string OLDID;
        string OLDDEP;
        string OLDCLASS;
        string OLDNO;

        string NOID;
        string OWNNAME;
        string BRAND;
        string SPEC;
        string ID;
        string NAME;
        string DEP;
        string DEPNAME;


        public FrmOWNERMANU()
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001,ME002    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ME001";
            comboBox1.DisplayMember = "ME001";
            sqlConn.Close();

            label7.Text = dt.Rows[0]["ME002"].ToString();


        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME001='{0}'    ", comboBox1.Text.ToString());
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);

            sqlConn.Close();

            if(dt.Rows.Count>0)
            {
                label7.Text = dt.Rows[0]["ME002"].ToString();
            }
            else
            {
                label7.Text = "DEP";
            }
            
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001,ME002    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ME001";
            comboBox2.DisplayMember = "ME001";
            sqlConn.Close();

            textBox5.Text = dt.Rows[0]["ME002"].ToString();


        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME  WHERE ME001='{0}'    ", comboBox2.Text.ToString());
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            
            sqlConn.Close();

            
            if (dt.Rows.Count > 0)
            {
                textBox5.Text = dt.Rows[0]["ME002"].ToString();
            }
            else
            {
                textBox5.Text = null;
            }

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
            Sequel.AppendFormat(@"SELECT [CLASS],[CLASSNAME] FROM [TKGAFFAIRS].[dbo].[CLASSBRAND] ORDER BY CLASS");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("CLASS", typeof(string));
            dt.Columns.Add("CLASSNAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "CLASS";
            comboBox3.DisplayMember = "CLASS";
            sqlConn.Close();

            textBox6.Text = dt.Rows[0]["CLASSNAME"].ToString();


        }
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
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
            Sequel.AppendFormat(@"SELECT [CLASS],[CLASSNAME] FROM [TKGAFFAIRS].[dbo].[CLASSBRAND] WHERE CLASS='{0}' ORDER BY CLASS",comboBox3.Text);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("CLASS", typeof(string));
            dt.Columns.Add("CLASSNAME", typeof(string));
            da.Fill(dt);
            
            sqlConn.Close();

            if (dt.Rows.Count > 0)
            {
                textBox6.Text = dt.Rows[0]["CLASSNAME"].ToString();
            }
            else
            {
                textBox6.Text = null;
            }
            
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME002 NOT LIKE '%停用%' ORDER BY ME001,ME002    ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ME001";
            comboBox4.DisplayMember = "ME001";
            sqlConn.Close();

            label20.Text = dt.Rows[0]["ME002"].ToString();


        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
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
            Sequel.AppendFormat(@"SELECT ME001,ME002 FROM [TK].dbo.CMSME WHERE ME001='{0}'    ", comboBox4.Text.ToString());
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            da.Fill(dt);

            sqlConn.Close();

            if (dt.Rows.Count > 0)
            {
                label20.Text = dt.Rows[0]["ME002"].ToString();
            }
            else
            {
                label20.Text = "DEP";
            }
        }
        public void Search()
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

                if(!string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSqlQuery.AppendFormat(@" AND [ID]='{0}'  ",textBox1.Text);
                }

                if (!string.IsNullOrEmpty(textBox2.Text))
                {
                    sbSqlQuery.AppendFormat(@" AND [NAME]='{0}'  ", textBox2.Text);
                }

                sbSql.AppendFormat(@"  SELECT [ID] AS '工號',[NAME] AS '保管人',[DEP] AS '部門',[DEPNAME] AS '單位',[CREATEDATES] AS '建立日期'");
                sbSql.AppendFormat(@"  ,[CLASS] AS '分類',[CLASSNAME] AS '分類名',[NO] AS '流水號',[OWNNAME] AS '保管品名',[BRAND] AS '廠牌',[SPEC] AS '規格'");
                sbSql.AppendFormat(@"  ,[PRICES] AS '原價',[NUM] AS '數量',[GIVENAME] AS '發放人',[REMARK] AS '備註'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[OWNERMANU]");
                sbSql.AppendFormat(@"  WHERE DEP='{0}'",comboBox1.Text.ToString());
                sbSql.AppendFormat(sbSqlQuery.ToString());
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

            textBox1.Text = null;
            textBox2.Text = null;
        }

        public void SETSTATUS()
        {
           
            textBox3.Text = null;
            textBox4.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = "0";
            textBox12.Text = "0";
            textBox13.Text = null;
            textBox14.Text = null;


            //textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;           
            textBox8.ReadOnly = false;
            textBox9.ReadOnly = false;
            textBox10.ReadOnly = false;
            textBox11.ReadOnly = false;
            textBox12.ReadOnly = false;
            textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;

        }
        public void SETSTATUS2()
        {
            
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;           
            textBox8.ReadOnly = false;
            textBox9.ReadOnly = false;
            textBox10.ReadOnly = false;
            textBox11.ReadOnly = false;
            textBox12.ReadOnly = false;
            textBox13.ReadOnly = false;
            textBox14.ReadOnly = false;
        }

        public void ADD()
        {
            try
            {
                textBox7.Text =comboBox3.Text+'-'+ GETNO();
                //add ZWAREWHOUSEPURTH
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
                sbSql.AppendFormat(" INSERT INTO [TKGAFFAIRS].[dbo].[OWNERMANU]");
                sbSql.AppendFormat(" ([ID],[NAME],[DEP],[DEPNAME],[CREATEDATES],[CLASS],[CLASSNAME],[NO],[OWNNAME],[BRAND],[SPEC],[PRICES],[NUM],[GIVENAME],[REMARK])");
                sbSql.AppendFormat(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',{11},{12},'{13}','{14}')",textBox3.Text, textBox4.Text,comboBox2.Text, textBox5.Text,dateTimePicker1.Value.ToString("yyyy/MM/dd"),comboBox3.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text);
                sbSql.AppendFormat(" ");
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
        public string GETNO()
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX([NO]),'00000000') AS NO");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[OWNERMANU]");               
                sbSql.AppendFormat(@"  WHERE [CLASS]='{0}'",comboBox3.Text);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        NO = SETNO(ds4.Tables["TEMPds4"].Rows[0]["NO"].ToString());
                        return NO;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SETNO(string NO)
        {
            if (NO.Equals("00000000"))
            {
                return  "00001";
            }

            else
            {
                NO = NO.Substring(3, 5);
                int serno = Convert.ToInt16(NO);
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(5, '0');
                return temp.ToString();
            }
        }
        public void UPDATE()
        {
            try
            {

                //add ZWAREWHOUSEPURTH
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
   
                sbSql.AppendFormat(@" UPDATE [TKGAFFAIRS].[dbo].[OWNERMANU]");
                sbSql.AppendFormat(@" SET [ID]='{0}',[NAME]='{1}',[DEP]='{2}',[DEPNAME]='{3}',[CREATEDATES]='{4}',[CLASS]='{5}',[CLASSNAME]='{6}'", textBox3.Text, textBox4.Text, comboBox2.Text, textBox5.Text, dateTimePicker1.Value.ToString("yyyy/MM/dd"), comboBox3.Text, textBox6.Text);
                sbSql.AppendFormat(@" ,[NO]='{0}',[OWNNAME]='{1}',[BRAND]='{2}',[SPEC]='{3}',[PRICES]='{4}',[NUM]='{5}',[GIVENAME]='{6}',[REMARK]='{7}'", textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text);
                sbSql.AppendFormat(@"  WHERE [ID]='{0}' AND [DEP]='{1}' AND [CLASS]='{2}' AND [NO]='{3}'",OLDID,OLDDEP,OLDCLASS,OLDNO);
                sbSql.AppendFormat(@" ");
                sbSql.AppendFormat(@" ");
                sbSql.AppendFormat(@" ");

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

        public void SETSTAUSFIANL()
        {
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox8.ReadOnly = true;
            textBox9.ReadOnly = true;
            textBox10.ReadOnly = true;
            textBox11.ReadOnly = true;
            textBox12.ReadOnly = true;
            textBox13.ReadOnly = true;
            textBox14.ReadOnly = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    dateTimePicker1.Value = Convert.ToDateTime(row.Cells["建立日期"].Value.ToString());

                    textBox3.Text = row.Cells["工號"].Value.ToString();
                    textBox4.Text = row.Cells["保管人"].Value.ToString();
                    comboBox2.Text = row.Cells["部門"].Value.ToString();
                    textBox5.Text = row.Cells["單位"].Value.ToString();
                    comboBox3.Text = row.Cells["分類"].Value.ToString();
                    textBox6.Text = row.Cells["分類名"].Value.ToString();
                    textBox7.Text = row.Cells["流水號"].Value.ToString();
                    textBox8.Text = row.Cells["保管品名"].Value.ToString();
                    textBox9.Text = row.Cells["廠牌"].Value.ToString();
                    textBox10.Text = row.Cells["規格"].Value.ToString();
                    textBox11.Text = row.Cells["原價"].Value.ToString();
                    textBox12.Text = row.Cells["數量"].Value.ToString();
                    textBox13.Text = row.Cells["發放人"].Value.ToString();
                    textBox14.Text = row.Cells["備註"].Value.ToString();

                    NOID = row.Cells["流水號"].Value.ToString();
                    OWNNAME = row.Cells["保管品名"].Value.ToString();
                    BRAND = row.Cells["廠牌"].Value.ToString();
                    SPEC = row.Cells["規格"].Value.ToString();
                    ID =row.Cells["工號"].Value.ToString();
                    NAME = row.Cells["保管人"].Value.ToString();
                    DEP = row.Cells["部門"].Value.ToString();
                    DEPNAME = row.Cells["單位"].Value.ToString();

                }
                else
                {

                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    textBox9.Text = null;
                    textBox10.Text = null;
                    textBox11.Text = null;
                    textBox12.Text = null;
                    textBox13.Text = null;
                    textBox14.Text = null;

                    NOID = null;
                    OWNNAME = null;
                    BRAND = null;
                    SPEC = null;
                    ID = null;
                    NAME = null;
                    DEP = null;
                    DEPNAME = null;
                }
            }
        }

        public void DEL()
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

                sbSql.AppendFormat(" DELETE [TKGAFFAIRS].[dbo].[OWNERMANU]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}' AND [DEP]='{1}' AND [CLASS]='{2}' AND [NO]='{3}'", textBox3.Text, comboBox2.Text, comboBox3.Text, textBox7.Text);
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
        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\文具、個人手工具保管卡.frx");

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
            StringBuilder FASTSQL = new StringBuilder();

            sbSqlQuery.Clear();

            if (!string.IsNullOrEmpty(textBox15.Text))
            {
                sbSqlQuery.AppendFormat(@" AND [ID]='{0}'  ", textBox15.Text);
            }

            if (!string.IsNullOrEmpty(textBox16.Text))
            {
                sbSqlQuery.AppendFormat(@" AND [NAME]='{0}'  ", textBox16.Text);
            }

            FASTSQL.AppendFormat(@"  SELECT [ID] AS '工號',[NAME] AS '保管人',[DEP] AS '部門',[DEPNAME] AS '單位',[CREATEDATES] AS '建立日期'");
            FASTSQL.AppendFormat(@"  ,[CLASS] AS '分類',[CLASSNAME] AS '分類名',[NO] AS '流水號',[OWNNAME] AS '保管品名',[BRAND] AS '廠牌',[SPEC] AS '規格'");
            FASTSQL.AppendFormat(@"  ,[PRICES] AS '原價',[NUM] AS '數量',[GIVENAME] AS '發放人',[REMARK] AS '備註'");
            FASTSQL.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[OWNERMANU]");
            FASTSQL.AppendFormat(@"  WHERE DEP='{0}'", comboBox4.Text.ToString());
            FASTSQL.AppendFormat(sbSqlQuery.ToString());
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SEARCHNAME1();
            
            FINDDEP2();
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


                sbSql.AppendFormat(@" SELECT [CnName] ,[Department].[DepartmentId],[Department].[Code],[Name],[Department].[Code] AS 'DEPID'     FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] WHERE [Employee].DepartmentId=[Department].DepartmentId AND [Employee].Code='{0}'", textBox1.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP2.Clear();
                adapterTEMP.Fill(dsTEMP2, "dsTEMP2");
                sqlConn.Close();


                if (dsTEMP2.Tables["dsTEMP2"].Rows.Count == 0)
                {
                    label7.Text = null;
                    comboBox1.Text = null;
                }
                else
                {
                    if (dsTEMP2.Tables["dsTEMP2"].Rows.Count >= 1)
                    {
                        label7.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["Name"].ToString();
                        comboBox1.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["DEPID"].ToString();

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
        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            SEARCHNAME2();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            SEARCHNAME3();
            FINDDEP4();
        }

        public void FINDDEP4()
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
                    comboBox2.Text = null;
                    textBox5.Text = null;
                }
                else
                {
                    if (dsTEMP2.Tables["dsTEMP2"].Rows.Count >= 1)
                    {
                        comboBox2.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["DEPID"].ToString();
                        textBox5.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["Name"].ToString();

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
        public void SEARCHNAME1()
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.Clear();

                sbSql.AppendFormat(@" SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV001='{0}' ",textBox1.Text);
                sbSql.AppendFormat(@"   ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        textBox2.Text = ds4.Tables["TEMPds4"].Rows[0]["MV002"].ToString();
                        

                    }
                    
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
        public void SEARCHNAME2()
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.Clear();

                sbSql.AppendFormat(@" SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV001='{0}' ", textBox15.Text);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        textBox16.Text = ds4.Tables["TEMPds4"].Rows[0]["MV002"].ToString();


                    }

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
        public void SEARCHNAME3()
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.Clear();

                sbSql.AppendFormat(@" SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV001='{0}' ", textBox3.Text);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        textBox4.Text = ds4.Tables["TEMPds4"].Rows[0]["MV002"].ToString();


                    }

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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            SEARCHNAME4();

            FINDDEP3();
        }

        public void SEARCHNAME4()
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.Clear();

                sbSql.AppendFormat(@" SELECT MV001,MV002 FROM [TK].dbo.CMSMV WHERE MV002='{0}' ", textBox2.Text);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        textBox1.Text = ds4.Tables["TEMPds4"].Rows[0]["MV001"].ToString();


                    }

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

        public void FINDDEP3()
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


                sbSql.AppendFormat(@" SELECT [CnName] ,[Department].[DepartmentId],[Department].[Code],[Name],[Department].[Code] AS 'DEPID'     FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] WHERE [Employee].DepartmentId=[Department].DepartmentId AND [Employee].[CnName]='{0}'", textBox2.Text.ToString());


                adapterTEMP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderTEMP = new SqlCommandBuilder(adapterTEMP);
                sqlConn.Open();
                dsTEMP2.Clear();
                adapterTEMP.Fill(dsTEMP2, "dsTEMP2");
                sqlConn.Close();


                if (dsTEMP2.Tables["dsTEMP2"].Rows.Count == 0)
                {
                    label7.Text = null;
                    comboBox1.Text = null;
                }
                else
                {
                    if (dsTEMP2.Tables["dsTEMP2"].Rows.Count >= 1)
                    {
                        label7.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["Name"].ToString();
                        comboBox1.Text = dsTEMP2.Tables["dsTEMP2"].Rows[0]["DEPID"].ToString();

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
        private void button2_Click(object sender, EventArgs e)
        {
            STATUS = "ADD";
            SETSTATUS();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            STATUS = "EDIT";
            OLDID =textBox3.Text;
            OLDDEP = comboBox2.Text;
            OLDCLASS = comboBox3.Text;
            OLDNO = textBox7.Text;

            SETSTATUS2();
        }

        private void button5_Click(object sender, EventArgs e)
        {
           
            if (STATUS.Equals("EDIT"))
            {
                UPDATE();
            }
            else if (STATUS.Equals("ADD"))
            {
                ADD();
            }

            STATUS = null;

            SETSTAUSFIANL();

            Search();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            STATUS = null;
            string message = textBox3.Text+ "的"+textBox7.Text + " 要刪除了?";

            DialogResult dialogResult = MessageBox.Show(message.ToString(), "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DEL();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            Search();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }


        private void button7_Click(object sender, EventArgs e)
        {

            FrmOWNERMANUsub FrmOWNERMANUsub = new FrmOWNERMANUsub(NOID, OWNNAME, BRAND, SPEC, ID, NAME, DEP, DEPNAME);
            FrmOWNERMANUsub.ShowDialog();

            Search();
        }




        #endregion


    }
}
