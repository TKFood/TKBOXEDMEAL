using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Media;

namespace TKBOXEDMEAL
{
    public partial class frmRUN : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlConnection sqlConn2 = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlTransaction tran2;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string NowDay;
        int rownum = 0;
        DateTime StartLunchdt;
        DateTime EndLunchdt;
        DateTime StartDinnerdt;
        DateTime EndDinnerdt;
        DateTime comdt;
        TimeSpan ts;
        string ID;
        string orderMeal;

        public frmRUN()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void PLAYMP3()
        {
            WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
            wplayer.URL = @"\\Server2003\PROG更新\TKBOXEDMEAL\mp3\BEE.mp3";
            wplayer.controls.play();
        }
        private void frmRUN_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timer1.Interval = 1000;
            timer1.Start();
            StartLunchdt = DateTime.Now;
            StartDinnerdt = DateTime.Now;

            Search();
            textBox1.Select();

            //comdt = DateTime.Now;
            comdt = Convert.ToDateTime("12:10");

        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString();

            //enddt = DateTime.Now;
            //ts = enddt.Subtract(startdt);

            
            //if(ts.TotalMilliseconds>=2000 && label4.Text.ToString().Equals("用餐愉快!"))
            //{
            //    label4.Text = "請刷卡!";
            //    startdt = DateTime.Now;
            //}
        }

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

                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView1.DataSource = ds.Tables["TEMPds"];                    
                    dataGridView1.DefaultCellStyle.Font = new Font("新細明體", 20);
                    dataGridView1.AutoResizeColumns();
                    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    //rownum = ds.Tables["TEMPds"].Rows.Count - 1;
                    dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];
                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    DataRow[] result = ds.Tables["TEMPds"].Select("名稱='午餐'");
                    foreach (DataRow row in result)
                    {
                        StartLunchdt = Convert.ToDateTime(row[4].ToString());
                        EndLunchdt = Convert.ToDateTime(row[5].ToString());
                    }
                    DataRow[] result2 = ds.Tables["TEMPds"].Select("名稱='晚餐'");
                    foreach (DataRow row in result2)
                    {
                        StartDinnerdt = Convert.ToDateTime(row[4].ToString());
                        EndDinnerdt = Convert.ToDateTime(row[5].ToString());
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

        public void GetMeal(string ID,string MEAL)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"SELECT  [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [{0}].[dbo].[EMPORDER] WITH (NOLOCK) WHERE CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND ([ID]='{1}' OR [CARDNO]='{1}' ) AND [MEAL]='{2}' AND [EATNUM]=0   ", sqlConn.Database.ToString(),ID,MEAL);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT  [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [{0}].[dbo].[EMPORDER] WITH (NOLOCK) WHERE CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND ([ID]='{1}' OR [CARDNO]='{1}' ) AND [MEAL]='{2}'   ", sqlConn.Database.ToString(), ID, MEAL);

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds2.Clear();
                    adapter.Fill(ds2, "TEMPds2");
                    sqlConn.Close();

                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        label4.Text = "沒有訂餐記錄!";
                        //System.Media.SystemSounds.Beep.Play();
                        PLAYMP3();
                    }
                    else
                    {
                        if(ds2.Tables["TEMPds2"].Rows[0][8].ToString().Equals("1"))
                        {
                            label4.Text = "已經用過餐了!";
                            //System.Media.SystemSounds.Beep.Play();
                            PLAYMP3();
                        }
                    }

                }
                else
                {
                    UPDATEEAT(ID,MEAL);
                }

            }
            catch
            {

            }
            finally
            {

            }

            
        }
        public void UPDATEEAT(string ID,string MEAL)
        { 
            try
            {               
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                if (sqlConn.State != ConnectionState.Open)
                {
                    sqlConn.Open();
                }
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
         
                sbSql.AppendFormat(" UPDATE [TKBOXEDMEAL].[dbo].[EMPORDER] SET [EATNUM]=1 WHERE CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND ([ID]='{0}' OR [CARDNO]='{0}' ) AND [MEAL]='{1}' AND [EATNUM]=0 ", ID, MEAL);
                

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();
                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    label4.Text = "無法用餐!";
                    //System.Media.SystemSounds.Beep.Play();
                    PLAYMP3();
                }
                else
                {
                    label4.Text = "用餐愉快!";
                    tran.Commit();      
                    
                }

                sqlConn.Close();
                Search();
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
            //System.Media.SystemSounds.Exclamation.Play();
            //System.Media.SystemSounds.Hand.Play();




            ID = textBox1.Text.ToString();
            if (DateTime.Compare(StartLunchdt, comdt) < 0 && DateTime.Compare(EndLunchdt, comdt) > 0)
            {
                GetMeal(ID,"10");
            }
            else if (DateTime.Compare(StartDinnerdt, comdt) < 0 && DateTime.Compare(EndDinnerdt, comdt) > 0)
            {
                GetMeal(ID, "20");
            }
            else
            {
                label4.Text = "非用餐時間";
                PLAYMP3();
            }

            textBox1.Text = null;
            textBox1.Select();
        }
        #endregion


    }
}
