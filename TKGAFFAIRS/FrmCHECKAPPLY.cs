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

        public void INSERTUOFDATETIME()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["UOFdbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

              
                sbSql.AppendFormat(@"  SELECT [DOC_ID],[CONTENT],[FORM_VERSION_ID],[MODIFIER],[TASK_ID]");
                sbSql.AppendFormat(@"  FROM [UOFTEST].[dbo].[TB_WKF_DOC]");
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
                        Xmldoc.LoadXml(ds.Tables["ds"].Rows[0]["CONTENT"].ToString());

                        XmlNode node = Xmldoc.SelectSingleNode("Form/FormFieldValue/FieldItem[@fieldId='HREngFrm001OutTime']");
                        XmlElement element = (XmlElement)node;
                        element.SetAttribute("fieldValue", "11:22");

                        XmlDocument Xmldoc2 = Xmldoc;


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

        private void button2_Click(object sender, EventArgs e)
        {
            INSERTUOFDATETIME();
        }

        #endregion
    }
}
