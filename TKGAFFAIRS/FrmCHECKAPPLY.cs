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
    public partial class FrmCHECKAPPLY : Form
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

        public FrmCHECKAPPLY()
        {
            InitializeComponent();

            timer1.Enabled = true;
            timer1.Interval = 1000*60;
            timer1.Start();
        }

        #region FUNCTION
        private void timer1_Tick(object sender, EventArgs e)
        {
            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            dateTimePicker3.Value = DateTime.Now;
            dateTimePicker4.Value = DateTime.Now;
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

                StringBuilder query = new StringBuilder();

                if(comboBox1.Text.Equals("未完成"))
                {
                    query.AppendFormat(@"  AND (ISNULL([HREngFrm001OutTime] ,'')='' OR ISNULL([HREngFrm001BakTime]  ,'')='')");
                }
                else if(comboBox1.Text.Equals("全部"))
                {
                    query.AppendFormat(@"  ");
                }

                sbSql.AppendFormat(@"  SELECT [HREngFrm001SN] AS '表單編號	',[HREngFrm001OutDate] AS '預計日期',[HREngFrm001User] AS '申請人',[HREngFrm001UsrDpt] AS '部門',[HREngFrm001Rank] AS '職級',[HREngFrm001Agent] AS '代理人',[HREngFrm001Transp] AS '交通工具',[HREngFrm001Location] AS '外出地點'	,[HREngFrm001Cause] AS '外出原因',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間',[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間'	");
                sbSql.AppendFormat(@"  ,[TaskId]");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(@"  WHERE [HREngFrm001OutDate]>='{0}' AND [HREngFrm001OutDate]<='{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  {0}",query.ToString());
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

        public void INSERTUOFDATETIME()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["UOFdbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

              
                sbSql.AppendFormat(@"  SELECT [TASK_ID],[TASK_SEQ],[BEGIN_TIME],[END_TIME],[TASK_STATUS],[TASK_RESULT],[DOC_NBR],[FLOW_TYPE],[FLOW_ID],[FORM_VERSION_ID],[SOURCE_DOC_ID],[CURRENT_DOC_ID],[FORM_STATUS],[USER_GUID],[USER_GROUP_ID],[USER_JOB_TITLE_ID],[ATTACH_ID],[URGENT_LEVEL],[CURRENT_SIGNER],[LOCK_STATUS],[CURRENT_DOC],[FILING_STATUS],[CURRENT_SITE_ID],[IS_APPLICANT_GETBACK],[APPLICANT_COMMENT],[DISPLAY_TITLE],[MESSAGE_CONTENT],[DEFAULT_IQY_USERS],[AGENT_USER],[CANCEL_FORM_REASON],[CANCEL_USER],[JSON_DISPLAY]");
                sbSql.AppendFormat(@"  FROM [UOF].[dbo].[TB_WKF_TASK]");
                sbSql.AppendFormat(@"  WHERE TASK_ID='75ccea35-3a28-418c-9411-22aa42b124b7'");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();

               
                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        XmlDocument Xmldoc = new XmlDocument();
                        Xmldoc.LoadXml(ds.Tables["ds"].Rows[0]["CURRENT_DOC"].ToString());

                        XmlNode node = Xmldoc.SelectSingleNode("Form/FormFieldValue/FieldItem[@fieldId='HREngFrm001OutTime']");
                        XmlElement element = (XmlElement)node;
                        element.SetAttribute("fieldValue", "11:22");

                        UPDATEAPPLY(Xmldoc);
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

        public void UPDATEAPPLY(XmlDocument Xmldoc)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["UOFdbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [UOF].[dbo].[TB_WKF_TASK]");
                sbSql.AppendFormat(" SET  CURRENT_DOC=@CURRENT_DOC");
                sbSql.AppendFormat(" WHERE TASK_ID='75ccea35-3a28-418c-9411-22aa42b124b7'");
                sbSql.AppendFormat(" ");

                cmd.Parameters.AddWithValue("@CURRENT_DOC", Xmldoc.OuterXml);

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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    TaskId = row.Cells["TaskId"].Value.ToString();                  

                }
                else
                {
                    TaskId = null;                   

                }
            }
        }

        public void INSERTHREngFrm001HREngFrm001OutTime(string TaskId)
        {
            if(!string.IsNullOrEmpty(TaskId))
            {
                UPDATEHREngFrm001HREngFrm001OutTime(TaskId, dateTimePicker3.Value.ToString("HH:mm"));
            }
        }

        public void UPDATEHREngFrm001HREngFrm001OutTime(string TaskId, string HREngFrm001OutTime)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                
                sbSql.AppendFormat(" UPDATE [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(" SET [HREngFrm001OutTime]='{0}'", HREngFrm001OutTime);
                sbSql.AppendFormat(" WHERE TaskId='{0}'", TaskId);
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


        public void INSERTHREngFrm001HREngFrm001BakTime(string TaskId)
        {
            if (!string.IsNullOrEmpty(TaskId))
            {
                UPDATEHREngFrm001HREngFrm001BakTime(TaskId, dateTimePicker4.Value.ToString("HH:mm"));
            }
        }

        public void UPDATEHREngFrm001HREngFrm001BakTime(string TaskId, string HREngFrm001OutTime)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(" SET [HREngFrm001BakTime]='{0}'", HREngFrm001OutTime);
                sbSql.AppendFormat(" WHERE TaskId='{0}'", TaskId);
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

        public void INSERTUOFHREngFrm001HREngFrm001OutTime(string TaskId)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["UOFdbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [TASK_ID],[TASK_SEQ],[BEGIN_TIME],[END_TIME],[TASK_STATUS],[TASK_RESULT],[DOC_NBR],[FLOW_TYPE],[FLOW_ID],[FORM_VERSION_ID],[SOURCE_DOC_ID],[CURRENT_DOC_ID],[FORM_STATUS],[USER_GUID],[USER_GROUP_ID],[USER_JOB_TITLE_ID],[ATTACH_ID],[URGENT_LEVEL],[CURRENT_SIGNER],[LOCK_STATUS],[CURRENT_DOC],[FILING_STATUS],[CURRENT_SITE_ID],[IS_APPLICANT_GETBACK],[APPLICANT_COMMENT],[DISPLAY_TITLE],[MESSAGE_CONTENT],[DEFAULT_IQY_USERS],[AGENT_USER],[CANCEL_FORM_REASON],[CANCEL_USER],[JSON_DISPLAY]");
                sbSql.AppendFormat(@"  FROM [UOF].[dbo].[TB_WKF_TASK]");
                sbSql.AppendFormat(@"  WHERE TASK_ID='{0}'", TaskId);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        XmlDocument Xmldoc = new XmlDocument();
                        Xmldoc.LoadXml(ds.Tables["ds"].Rows[0]["CURRENT_DOC"].ToString());

                        XmlNode node = Xmldoc.SelectSingleNode("Form/FormFieldValue/FieldItem[@fieldId='HREngFrm001OutTime']");
                        XmlElement element = (XmlElement)node;
                        element.SetAttribute("fieldValue", dateTimePicker3.Value.ToString("HH:mm"));

                        UPDATETUOFHREngFrm001HREngFrm001OutTime(TaskId, Xmldoc);
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

        public void UPDATETUOFHREngFrm001HREngFrm001OutTime(string TaskId, XmlDocument Xmldoc)
        {
            SqlCommand cmd = new SqlCommand();

            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["UOFdbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [UOF].[dbo].[TB_WKF_TASK]");
                sbSql.AppendFormat(" SET  CURRENT_DOC=@CURRENT_DOC");
                sbSql.AppendFormat(" WHERE TASK_ID='{0}'",TaskId);
                sbSql.AppendFormat(" ");

                cmd.Parameters.AddWithValue("@CURRENT_DOC", Xmldoc.OuterXml);

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

        public void INSERTUOFHREngFrm001HREngFrm001BakTime(string TaskId)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["UOFdbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [TASK_ID],[TASK_SEQ],[BEGIN_TIME],[END_TIME],[TASK_STATUS],[TASK_RESULT],[DOC_NBR],[FLOW_TYPE],[FLOW_ID],[FORM_VERSION_ID],[SOURCE_DOC_ID],[CURRENT_DOC_ID],[FORM_STATUS],[USER_GUID],[USER_GROUP_ID],[USER_JOB_TITLE_ID],[ATTACH_ID],[URGENT_LEVEL],[CURRENT_SIGNER],[LOCK_STATUS],[CURRENT_DOC],[FILING_STATUS],[CURRENT_SITE_ID],[IS_APPLICANT_GETBACK],[APPLICANT_COMMENT],[DISPLAY_TITLE],[MESSAGE_CONTENT],[DEFAULT_IQY_USERS],[AGENT_USER],[CANCEL_FORM_REASON],[CANCEL_USER],[JSON_DISPLAY]");
                sbSql.AppendFormat(@"  FROM [UOF].[dbo].[TB_WKF_TASK]");
                sbSql.AppendFormat(@"  WHERE TASK_ID='{0}'", TaskId);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        XmlDocument Xmldoc = new XmlDocument();
                        Xmldoc.LoadXml(ds.Tables["ds"].Rows[0]["CURRENT_DOC"].ToString());

                        XmlNode node = Xmldoc.SelectSingleNode("Form/FormFieldValue/FieldItem[@fieldId='HREngFrm001BakTime']");
                        XmlElement element = (XmlElement)node;
                        element.SetAttribute("fieldValue", dateTimePicker4.Value.ToString("HH:mm"));

                        UPDATETUOFHREngFrm001HREngFrm001OutTime(TaskId, Xmldoc);
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

        public void UPDATETUOFHREngFrm001HREngFrm001BakTime(string TaskId, XmlDocument Xmldoc)
        {
            SqlCommand cmd = new SqlCommand();

            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["UOFdbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [UOF].[dbo].[TB_WKF_TASK]");
                sbSql.AppendFormat(" SET  CURRENT_DOC=@CURRENT_DOC");
                sbSql.AppendFormat(" WHERE TASK_ID='{0}'", TaskId);
                sbSql.AppendFormat(" ");

                cmd.Parameters.AddWithValue("@CURRENT_DOC", Xmldoc.OuterXml);

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
            string SQL2;
            Report report1 = new Report();
            report1.Load(@"REPORT\刷卡記錄.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;            
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

            FASTSQL.AppendFormat(@"   SELECT [HREngFrm001User] AS '人員',[HREngFrm001Date] AS '日期',[HREngFrm001OutTime] AS '時間',[HREngFrm001Cause] AS '外出原因',[MODIFYCASUE] AS '記錄'");
            FASTSQL.AppendFormat(@"   FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
            FASTSQL.AppendFormat(@"   WHERE [HREngFrm001Cause]='可自由外出人員'");
            FASTSQL.AppendFormat(@"   AND [HREngFrm001Date]>='{0}' AND [HREngFrm001Date]<='{1}'",dateTimePicker5.Value.ToString("yyyy/MM/dd"), dateTimePicker6.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"   UNION");
            FASTSQL.AppendFormat(@"   SELECT [HREngFrm001User] AS '人員',[HREngFrm001OutDate] AS '日期',[HREngFrm001OutTime] AS '時間',[HREngFrm001Cause] AS '外出原因','外出' AS '記錄'");
            FASTSQL.AppendFormat(@"   FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
            FASTSQL.AppendFormat(@"   WHERE [HREngFrm001Cause]<>'可自由外出人員'");
            FASTSQL.AppendFormat(@"   AND ISNULL([HREngFrm001OutTime],'')<>''");
            FASTSQL.AppendFormat(@"   AND [HREngFrm001OutDate]>='{0}' AND [HREngFrm001OutDate]<='{1}'", dateTimePicker5.Value.ToString("yyyy/MM/dd"), dateTimePicker6.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"   UNION");
            FASTSQL.AppendFormat(@"   SELECT [HREngFrm001User] AS '人員',[HREngFrm001OutDate] AS '日期',[HREngFrm001BakTime] AS '時間',[HREngFrm001Cause] AS '外出原因','回廠' AS '記錄'");
            FASTSQL.AppendFormat(@"   FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
            FASTSQL.AppendFormat(@"   WHERE [HREngFrm001Cause]<>'可自由外出人員'");
            FASTSQL.AppendFormat(@"   AND ISNULL([HREngFrm001BakTime],'')<>''");
            FASTSQL.AppendFormat(@"   AND [HREngFrm001OutDate]>='{0}' AND [HREngFrm001OutDate]<='{1}'", dateTimePicker5.Value.ToString("yyyy/MM/dd"), dateTimePicker6.Value.ToString("yyyy/MM/dd"));
            FASTSQL.AppendFormat(@"   ");
            FASTSQL.AppendFormat(@"   ORDER BY [HREngFrm001Date],[HREngFrm001OutTime] ");
            FASTSQL.AppendFormat(@"   ");
            FASTSQL.AppendFormat(@"   ");


            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            INSERTUOFDATETIME();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            INSERTHREngFrm001HREngFrm001OutTime(TaskId);
            INSERTUOFHREngFrm001HREngFrm001OutTime(TaskId);


            Search();

            MessageBox.Show("完成"); 
        }

        private void button4_Click(object sender, EventArgs e)
        {
            INSERTHREngFrm001HREngFrm001BakTime(TaskId);
            INSERTUOFHREngFrm001HREngFrm001BakTime(TaskId);

            Search();

            MessageBox.Show("完成");
        }



        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }


        #endregion

    }
}
