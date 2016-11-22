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
        string Lang = "CH";
        int messagetime = 3000;
        string Name;
        string mess=null;

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
           

            Search();
            textBox1.Select();

            comdt = DateTime.Now;
            //comdt = Convert.ToDateTime("12:10");

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

                if (Lang.Equals("CH"))
                {
                    sbSql.AppendFormat(@"SELECT [ID] AS '編號',[NAME]  AS '名稱',CONVERT(VARCHAR(5),[STARTORDERTIME] ,108)  AS '訂餐開始時間',CONVERT(VARCHAR(5),[ENDORDERTIME] ,108)   AS '訂餐結束時間', CONVERT(VARCHAR(5),[STARTEATTIME] ,108)  AS '用餐開始時間',CONVERT(VARCHAR(5),[ENDEATTIME] ,108)   AS '用餐結束時間' FROM [{0}].[dbo].[BOXEDMEALSET]  ", sqlConn.Database.ToString());
                }
                else if (Lang.Equals("VN"))
                {
                    sbSql.AppendFormat(@"SELECT [ID] AS 'số',[VNNAME]  AS 'tên',CONVERT(VARCHAR(5),[STARTORDERTIME] ,108)  AS 'Thứ tự thời gian bắt đầu',CONVERT(VARCHAR(5),[ENDORDERTIME] ,108)   AS 'Đặt End Time', CONVERT(VARCHAR(5),[STARTEATTIME] ,108)  AS 'Ăn Start Time',CONVERT(VARCHAR(5),[ENDEATTIME] ,108)   AS 'Kết thúc thời gian bữa ăn' FROM [{0}].[dbo].[BOXEDMEALSET]  ", sqlConn.Database.ToString());

                }

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
                    dt = ds.Tables["TEMPds"];
                    dataGridView1.DataSource = dt;
                    dataGridView1.DefaultCellStyle.Font = new Font("新細明體", 20);
                    dataGridView1.AutoResizeColumns();
                    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    //rownum = ds.Tables["TEMPds"].Rows.Count - 1;
                    dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];
                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    

                    if (Lang.Equals("CH"))
                    {
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

                    else if (Lang.Equals("VN"))
                    {
                        DataRow[] result = ds.Tables["TEMPds"].Select("tên='bữa ăn trưa'");
                        foreach (DataRow row in result)
                        {
                            StartLunchdt = Convert.ToDateTime(row[4].ToString());
                            EndLunchdt = Convert.ToDateTime(row[5].ToString());
                        }
                        DataRow[] result2 = ds.Tables["TEMPds"].Select("tên='bữa tối'");
                        foreach (DataRow row in result2)
                        {
                            StartDinnerdt = Convert.ToDateTime(row[4].ToString());
                            EndDinnerdt = Convert.ToDateTime(row[5].ToString());
                        }

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
                    if (ds2.Tables["TEMPds2"].Rows.Count > 0)
                    {
                        Name = ds2.Tables["TEMPds2"].Rows[0][2].ToString();
                    }
                    sqlConn.Close();

                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {                       
                        if (Lang.Equals("CH"))
                        {
                            //label4.Text = "沒有訂餐記錄!";
                            //AutoClosingMessageBox.Show("沒有訂餐記錄!!", "TITLE", messagetime);
                            SHOWMESSAGE(Name + mess+"沒有訂餐記錄!!");
                        }
                        else if (Lang.Equals("VN"))
                        {
                            //label4.Text = "Không có hồ sơ đặt hàng!";
                            //AutoClosingMessageBox.Show("Không có hồ sơ đặt hàng!!", "TITLE", messagetime);
                            SHOWMESSAGE(Name + mess + "Không có hồ sơ đặt hàng!!");
                        }
                        //System.Media.SystemSounds.Beep.Play();
                        PLAYMP3();
                    }
                    else
                    {
                        if(ds2.Tables["TEMPds2"].Rows[0][8].ToString().Equals("1"))
                        {                            
                            if (Lang.Equals("CH"))
                            {
                                //label4.Text = "已經用過餐了!";
                                //AutoClosingMessageBox.Show("已經用過餐了!!", "TITLE", messagetime);
                                SHOWMESSAGE(Name + " 已經用過餐了!!");
                            }
                            else if (Lang.Equals("VN"))
                            {
                                //label4.Text = "Đã ăn lên!";
                                //AutoClosingMessageBox.Show("Đã ăn lên!!", "TITLE", messagetime);
                                SHOWMESSAGE(Name + "Đã ăn lên!");
                            }
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
        public void SearchEmplyee()
        {
            try
            {
                if(!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT TOP 1  [EmployeeID],[CardNo],[Name]FROM [TKBOXEDMEAL].[dbo].[VEMPLOYEE] WHERE [EmployeeID]='{1}' OR [CardNo]='{1}'", sqlConn.Database.ToString(), textBox1.Text.ToString());

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter.Fill(ds1, "TEMPds1");
                    sqlConn.Close();

                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {

                        if (Lang.Equals("CH"))
                        {
                            //label5.Text = "沒有此員工!";
                            //label4.Text = "";
                            //AutoClosingMessageBox.Show("沒有此員工!!", "TITLE", messagetime);
                            //SHOWMESSAGE("沒有此員工!!");
                            mess = "沒有此員工!!";
                        }
                        else if (Lang.Equals("VN"))
                        {
                            //label5.Text = "Không có nhân viên!";
                            //label4.Text = "";
                            //AutoClosingMessageBox.Show("Không có nhân viên!!", "TITLE", messagetime);
                            //SHOWMESSAGE("Không có nhân viên!!");
                            mess = "Không có nhân viên!!";
                        }

                        textBox1.Text = "";

                        Name = null;
                        PLAYMP3();
                    }
                    else
                    {

                        Name = ds1.Tables["TEMPds1"].Rows[0][2].ToString();

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
                    if (Lang.Equals("CH"))
                    {
                        //label4.Text = "無法用餐!";
                        //AutoClosingMessageBox.Show("無法用餐!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + " 無法用餐!!");
                    }
                    else if (Lang.Equals("VN"))
                    {
                        //label4.Text = "không thể để ăn cơm trưa!";
                        //AutoClosingMessageBox.Show("không thể để ăn cơm trưa!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + " không thể để ăn cơm trưa!!");
                    }
                    
                    //System.Media.SystemSounds.Beep.Play();
                    PLAYMP3();
                }
                else
                {                    
                    if (Lang.Equals("CH"))
                    {
                         //label4.Text = "用餐愉快!";
                         //AutoClosingMessageBox.Show("用餐愉快!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + " 請取餐，祝用餐愉快!!");
                    }
                    else if (Lang.Equals("VN"))
                    {
                        //label4.Text = "Thưởng thức bữa ăn của bạn!";
                        //AutoClosingMessageBox.Show("Thưởng thức bữa ăn của bạn!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + " Thưởng thức bữa ăn của bạn!!");
                    }
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

        public void SetLang()
        {
            if(Lang.Equals("CH"))
            {
                label4.Text = "請刷卡!";
                label5.Text = "請刷卡";
                button1.Text = "用餐";
            }
            else if(Lang.Equals("VN"))
            {
                label4.Text = "Hãy swipe!";
                label5.Text = "Hãy swipe";
                button1.Text = "ăn cơm trưa";
            }
        }

        public class AutoClosingMessageBox
        {
            System.Threading.Timer _timeoutTimer;
            string _caption;
            AutoClosingMessageBox(string text, string caption, int timeout)
            {
                _caption = caption;
                _timeoutTimer = new System.Threading.Timer(OnTimerElapsed,
                    null, timeout, System.Threading.Timeout.Infinite);
                using (_timeoutTimer)
                    MessageBox.Show(text, caption);
            }
            public static void Show(string text, string caption, int timeout)
            {
                new AutoClosingMessageBox(text, caption, timeout);
            }
            void OnTimerElapsed(object state)
            {
                IntPtr mbWnd = FindWindow("#32770", _caption); // lpClassName is #32770 for MessageBox
                if (mbWnd != IntPtr.Zero)
                    SendMessage(mbWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                _timeoutTimer.Dispose();
            }
            const int WM_CLOSE = 0x0010;
            [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
            static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
            [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
            static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
        }

        public void SHOWMESSAGE(String mess)
        {
            String text = mess;
            Message message = new Message(text);
            message.Show();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //System.Media.SystemSounds.Exclamation.Play();
            //System.Media.SystemSounds.Hand.Play();
            mess = null;
            Name = null;
            SearchEmplyee();

            comdt = DateTime.Now;

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
                if (Lang.Equals("CH"))
                {
                    //label4.Text = "非用餐時間";
                    //AutoClosingMessageBox.Show("非用餐時間" , "TITLE", messagetime);
                    SHOWMESSAGE(Name + " 非用餐時間!!");
                }
                else if (Lang.Equals("VN"))
                {
                    //label4.Text = "Hiện không bữa ăn";
                    //AutoClosingMessageBox.Show("Hiện không bữa ăn", "TITLE", messagetime);
                    SHOWMESSAGE(Name + " Hiện không bữa ăn!!");
                }
                
                PLAYMP3();
            }

            textBox1.Text = null;
            textBox1.Select();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Lang = "CH";
            ds.Tables["TEMPds"].Columns.Clear();
            Search();
            SetLang();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Lang = "VN";
            ds.Tables["TEMPds"].Columns.Clear();
            Search();
            SetLang();
        }


        #endregion

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //if ((textBox1.Text.Length == 11))
            //{
            //    button1.PerformClick();
            //}
          
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }
    }
}
