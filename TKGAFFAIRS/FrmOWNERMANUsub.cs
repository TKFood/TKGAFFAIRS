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
    public partial class FrmOWNERMANUsub : Form
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

        public FrmOWNERMANUsub()
        {
            InitializeComponent();
        }

        public FrmOWNERMANUsub(string ID1, string ID2, string ID3, string ID4, string ID5, string ID6, string ID7, string ID8)
        {
            InitializeComponent();

            textBox1.Text = ID1;
            textBox2.Text = ID2;
            textBox3.Text = ID3;
            textBox4.Text = ID4;
            textBox5.Text = ID5;
            textBox6.Text = ID6;
            textBox7.Text = ID7;
            textBox8.Text = ID8;
        }

        #region FUNCTION
        public void UPDATEOWNERMANU()
        {

        }
        #endregion

        #region BUTTON


        private void button7_Click(object sender, EventArgs e)
        {
            UPDATEOWNERMANU();
        }
        #endregion
    }
}
