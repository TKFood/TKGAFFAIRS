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
    public partial class FrmASSETS : Form
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
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        public FrmASSETS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            Report report1 = new Report();
            report1.Load(@"REPORT\資產QRCODE.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
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
          
            FASTSQL.AppendFormat(@"  SELECT ASTMB.MB001 AS '資產編號',ASTMA.MA002 AS '類別',ASTMB.MB002 AS '資產名稱',ASTMB.MB003 AS '規格'");
            FASTSQL.AppendFormat(@"  ,CONVERT(NVARCHAR,ASTMB.MB012)+ASTMB.MB011 AS '數量',CMSME.ME002 AS '部門名稱',ASTMC.MC006 AS '放置地點'");
            FASTSQL.AppendFormat(@"  ,CMSMV.MV001,CMSMV.MV002 AS '保管人'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.ASTMA ASTMA, [TK].dbo.ASTMB ASTMB, [TK].dbo.ASTMC ASTMC,[TK].dbo.CMSME CMSME,[TK].dbo.CMSMV CMSMV");
            FASTSQL.AppendFormat(@"  WHERE ASTMA.MA001=ASTMB.MB006");
            FASTSQL.AppendFormat(@"  AND ASTMC.MC002=CMSME.ME001 ");
            FASTSQL.AppendFormat(@"  AND ASTMB.MB001=ASTMC.MC001");
            FASTSQL.AppendFormat(@"  AND CMSMV.MV001=ASTMC.MC003");
            FASTSQL.AppendFormat(@"  AND ASTMC.MC003='{0}'",textBox1.Text);
            FASTSQL.AppendFormat(@"  ORDER BY ASTMB.MB001");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
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
