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
        DateTime startdt;
        DateTime enddt;
        DateTime comdt;

        public frmORDER()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"SELECT [ID] AS '編號',[NAME]  AS '名稱',CONVERT(VARCHAR(5),[STARTORDERTIME] ,108)  AS '訂餐開始時間',CONVERT(VARCHAR(5),[ENDORDERTIME] ,108)   AS '訂餐結束時間', CONVERT(VARCHAR(5),[STARTEATTIME] ,108)  AS '用餐開始時間',CONVERT(VARCHAR(5),[ENDEATTIME] ,108)   AS '用餐結束時間' FROM [{0}].[dbo].[BOXEDMEALSET]  ", sqlConn.Database.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                label1.Text = "資料筆數:" + ds.Tables["TEMPds"].Rows.Count.ToString();

                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView1.DataSource = ds.Tables["TEMPds"];
                    dataGridView1.AutoResizeColumns();
                    //rownum = ds.Tables["TEMPds"].Rows.Count - 1;
                    dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];
                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    DataRow[] result = ds.Tables["TEMPds"].Select("名稱='午餐'");
                    foreach (DataRow row in result)
                    {
                        startdt = Convert.ToDateTime(row[2].ToString());
                        enddt = Convert.ToDateTime(row[3].ToString());
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

        private void frmORDER_Load(object sender, EventArgs e)
        {
            Search();
        }

        public void SetString()
        {
            Name = "aa";
            OrderBoxed = "ok";
        }
        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //comdt = DateTime.Now;
            SetString();

            comdt = Convert.ToDateTime("09:10");
            if (DateTime.Compare(startdt, comdt) <0&& DateTime.Compare(enddt, comdt) > 0)
            {
                label5.Text = "訂餐成功!";
                label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();
            }
            else
            {
                label5.Text = "超過可點餐時間!";
                //label4.Text = "";
            }
           
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            SetString();
            comdt = Convert.ToDateTime("09:10");
            if (DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0)
            {
                button3.Visible = true;
                button4.Visible = true;
                button5.Visible = true;
                button6.Visible = true;
                button7.Visible = true;
                button8.Visible = true;

                label5.Text = "訂餐成功!";
                label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();
            }
            else
            {
                label5.Text = "超過可點餐時間!";
                //label4.Text = "";
            }
        }

        
    }
}
