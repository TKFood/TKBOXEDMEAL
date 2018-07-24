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
using System.Threading;
using System.Globalization;
using System.Resources;
using System.Reflection;

namespace TKBOXEDMEAL
{
    public partial class frmORDERLOCAL : Form
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

        public frmORDERLOCAL()
        {
            CultureInfo CI = new CultureInfo("zh");
            System.Threading.Thread.CurrentThread.CurrentUICulture = CI;

            InitializeComponent();
        }

        #region FUNCTION
        public void CreateResourceManager(Control Control, string Language)
        {
            CultureInfo info = new System.Globalization.CultureInfo(Language);
            Thread.CurrentThread.CurrentUICulture = info;//變更文化特性
            this._ResourceManager = new ComponentResourceManager(Control.GetType());
            this._ResourceManager.ApplyResources(Control, "$this");
            this.ApplyForm(Control);
        }

        public void ApplyForm(Control control)
        {
            foreach (Control ctrl in control.Controls)
            {
                this._ResourceManager.ApplyResources(ctrl, ctrl.Name);
                if (ctrl.HasChildren)
                {
                    ApplyForm(ctrl);
                }
            }
        }



        #endregion


        #region BUTTON
        private void button11_Click(object sender, EventArgs e)
        {
            CreateResourceManager(this, "zh-TW");

            this.WindowState = FormWindowState.Normal;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            CreateResourceManager(this, "vi-VN");

            this.WindowState = FormWindowState.Normal;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;
        }

        #endregion
    }
}
