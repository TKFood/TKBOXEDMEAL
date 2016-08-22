using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace TKBOXEDMEAL
{
    public partial class frmORDER : Form
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
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string Name;
        string OrderBoxed;
        int rownum = 0;

        public frmORDER()
        {
            InitializeComponent();
        }

        #region FUNCTION

        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            label5.Text = "訂餐成功!";
            label4.Text = Name.ToString()+" 您訂了: " + OrderBoxed.ToString();
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            label5.Text = "訂餐成功!";
            label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();
        }
    }
}
