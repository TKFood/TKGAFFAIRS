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

        public FrmCHECKAPPLY()
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

              
                sbSql.AppendFormat(@"  SELECT [HREngFrm001Name] AS '姓名',[HREngFrm001Dpt] AS '部門',[HREngFrm001Date] AS '日期',[HREngFrm001TITLE] AS '職稱',[HREngFrm001Agent] AS '代理人' ");
                sbSql.AppendFormat(@"  ,[HREngFrm001Transp] AS '交通工具',[HREngFrm001Location] AS '外出地點',[HREngFrm001Cause] AS '外出原因',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間' ");
                sbSql.AppendFormat(@"  ,[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間',[HREngFrm001UsrDpt] AS '申請人部門',[HREngFrm001User] AS '申請人姓名'");
                sbSql.AppendFormat(@"  ,[TaskId],[HREngFrm001SN]");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(@"  WHERE [HREngFrm001Date]>='{0}' AND [HREngFrm001Date]<='{1}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  AND (ISNULL([HREngFrm001OutTime] ,'')='' OR ISNULL([HREngFrm001DefBakTime]  ,'')='')");
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
                sbSql.AppendFormat(@"  FROM [UOFTEST].[dbo].[TB_WKF_TASK]");
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

                sbSql.AppendFormat(" UPDATE [UOFTEST].[dbo].[TB_WKF_TASK]");
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

        #endregion

       
    }
}
