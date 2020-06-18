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
    public partial class FrmCHECKAPPLYCARD : Form
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


        public FrmCHECKAPPLYCARD()
        {
            InitializeComponent();

            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
            timer1.Start();
        }


        #region FUNCTION
        private void timer1_Tick(object sender, EventArgs e)
        {
            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {

        }
        #endregion

      
    }
}
