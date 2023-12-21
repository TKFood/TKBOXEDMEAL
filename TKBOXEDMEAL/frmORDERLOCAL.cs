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
using TKITDLL;

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

        private void frmORDERLOCAL_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timer1.Interval = 1000;
            timer1.Start();

            Search();
            textBox1.Select();
            comdt = DateTime.Now;
        }

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

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString();
        }
        public void PLAYMP3()
        {
            WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
            wplayer.URL = @"\\Server2003\PROG更新\TKBOXEDMEAL\mp3\BEE.mp3";
            wplayer.controls.play();
        }

        private void frmORDER_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timer1.Interval = 1000;
            timer1.Start();

            Search();
            textBox1.Select();
            comdt = DateTime.Now;
            //comdt = Convert.ToDateTime("10:10");
        }

        public void Search()
        {

            try
            {     
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                if (Lang.Equals("CH"))
                {
                    sbSql.AppendFormat(@"SELECT [NAME]  AS '名稱',CONVERT(VARCHAR(5),[STARTORDERTIME] ,108)  AS '訂餐開始時間',CONVERT(VARCHAR(5),[ENDORDERTIME] ,108)   AS '訂餐結束時間' FROM [{0}].[dbo].[BOXEDMEALSET]  ", sqlConn.Database.ToString());
                }
                else if (Lang.Equals("VN"))
                {
                    sbSql.AppendFormat(@"SELECT [VNNAME]  AS 'tên',CONVERT(VARCHAR(5),[STARTORDERTIME] ,108)  AS 'Thứ tự thời gian bắt đầu',CONVERT(VARCHAR(5),[ENDORDERTIME] ,108)   AS 'Đặt End Time' FROM [{0}].[dbo].[BOXEDMEALSET]  ", sqlConn.Database.ToString());

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
                            startdt = Convert.ToDateTime(row[1].ToString());
                            enddt = Convert.ToDateTime(row[2].ToString());
                        }
                        DataRow[] result2 = ds.Tables["TEMPds"].Select("名稱='晚餐'");
                        foreach (DataRow row2 in result2)
                        {
                            startdinnerdt = Convert.ToDateTime(row2[1].ToString());
                            enddinnerdt = Convert.ToDateTime(row2[2].ToString());
                        }

                    }

                    else if (Lang.Equals("VN"))
                    {
                        DataRow[] result = ds.Tables["TEMPds"].Select("tên='bữa ăn trưa'");
                        foreach (DataRow row in result)
                        {
                            startdt = Convert.ToDateTime(row[1].ToString());
                            enddt = Convert.ToDateTime(row[2].ToString());
                        }
                        DataRow[] result2 = ds.Tables["TEMPds"].Select("tên='bữa tối'");
                        foreach (DataRow row2 in result2)
                        {
                            startdinnerdt = Convert.ToDateTime(row2[1].ToString());
                            enddinnerdt = Convert.ToDateTime(row2[2].ToString());
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




        public void SetOrderButton()
        {

            if (DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0)
            //if(1==1)
            {
                if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    textBox1.Text = textBox1.Text.TrimStart('0').ToString();
                    InputID = textBox1.Text.ToString();

                    SearchEmplyee();

                    if (!string.IsNullOrEmpty(Name))
                    {
                        button3.Visible = true;                     
                        button6.Visible = true;                      

                        //button1.Visible = false;
                        //button9.Visible = false;
                        button2.Visible = false;
                    }
                    else
                    {

                    }

                }

            }
            else
            {
                if (Lang.Equals("CH"))
                {
                    //label5.Text = "超過可點餐時間!!";
                    //AutoClosingMessageBox.Show("超過可點餐時間!!" , "TITLE", messagetime);
                    //SHOWMESSAGE("超過可點餐時間!!");

                }
                else if (Lang.Equals("VN"))
                {
                    //label5.Text = "Vượt quá thời gian bữa ăn!";
                    //AutoClosingMessageBox.Show("Vượt quá thời gian bữa ăn!!", "TITLE", messagetime);                  
                    //SHOWMESSAGE("Vượt quá thời gian bữa ăn!!");
                }
                PLAYMP3();
                //label4.Text = "";
            }

            if ((DateTime.Compare(startdinnerdt, comdt) < 0 && DateTime.Compare(enddinnerdt, comdt) > 0))
            //if(1==1)
            {
                if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    textBox1.Text = textBox1.Text.TrimStart('0').ToString();
                    InputID = textBox1.Text.ToString();

                    SearchEmplyee();

                    if (!string.IsNullOrEmpty(Name))
                    {
                        
                        button4.Visible = true;
                        button5.Visible = false;                       
                        button7.Visible = true;
                        button8.Visible = false;
                        button14.Visible = true;
                        button15.Visible = true;

                        //button1.Visible = false;
                        //button9.Visible = false;
                        button2.Visible = false;
                    }
                    else
                    {

                    }

                }

            }
            else
            {
                if (Lang.Equals("CH"))
                {
                    //label5.Text = "超過可點餐時間!!";
                    //AutoClosingMessageBox.Show("超過可點餐時間!!" , "TITLE", messagetime);
                    //SHOWMESSAGE("超過可點餐時間!!");

                }
                else if (Lang.Equals("VN"))
                {
                    //label5.Text = "Vượt quá thời gian bữa ăn!";
                    //AutoClosingMessageBox.Show("Vượt quá thời gian bữa ăn!!", "TITLE", messagetime);                  
                    //SHOWMESSAGE("Vượt quá thời gian bữa ăn!!");
                }
                PLAYMP3();
                //label4.Text = "";
            }

        }


        public void SetCancelButton()
        {
            if ((DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0 || (DateTime.Compare(startdinnerdt, comdt) < 0 && DateTime.Compare(enddinnerdt, comdt) > 0)))
            {
                if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    InputID = textBox1.Text.ToString();
                    SearchEmplyee();

                    if (!string.IsNullOrEmpty(Name))
                    {
                        button3.Visible = true;
                        button4.Visible = true;
                        button5.Visible = false;
                        button6.Visible = true;
                        button7.Visible = true;
                        button8.Visible = false;
                        button14.Visible = true;
                        button15.Visible = true;

                        //button1.Visible = false;
                        button2.Visible = false;
                    }
                }
            }
            else
            {

                if (Lang.Equals("CH"))
                {
                    //label5.Text = "超過可取消點餐時間!";
                    //AutoClosingMessageBox.Show("超過可點餐時間!!", "TITLE", messagetime);
                    SHOWMESSAGE("超過可點餐時間!!");
                }
                else if (Lang.Equals("VN"))
                {
                    //label5.Text = "Qua thời gian để hủy bỏ các bữa ăn!";
                    //AutoClosingMessageBox.Show("Vượt quá thời gian bữa ăn!!", "TITLE", messagetime);
                    SHOWMESSAGE("Vượt quá thời gian bữa ăn!!");
                }
                PLAYMP3();
                //label4.Text = "";
            }
        }

        public void SetCancel()
        {
            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button14.Visible = false;
            button15.Visible = false;

            textBox1.Text = null;
            textBox1.Select();

          
        }

        public void SearchEmplyee()
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"SELECT TOP 1  [EmployeeID],[CardNo],[Name] FROM [TKBOXEDMEAL].[dbo].[EMPLOYEE] WHERE [EmployeeID]='{1}' OR [CardNo]='{1}'", sqlConn.Database.ToString(), textBox1.Text.ToString());

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
                        SHOWMESSAGE("沒有此員工!!");
                    }
                    else if (Lang.Equals("VN"))
                    {
                        //label5.Text = "Không có nhân viên!";
                        //label4.Text = "";
                        //AutoClosingMessageBox.Show("Không có nhân viên!!", "TITLE", messagetime);
                        SHOWMESSAGE("Không có nhân viên!!");
                    }

                    textBox1.Text = "";
                    EmployeeID = null;
                    Name = null;
                    CardNo = null;
                    Meal = null;
                    PLAYMP3();
                }
                else
                {
                    EmployeeID = ds1.Tables["TEMPds1"].Rows[0][0].ToString();
                    CardNo = ds1.Tables["TEMPds1"].Rows[0][1].ToString();
                    Name = ds1.Tables["TEMPds1"].Rows[0][2].ToString();

                }

            }
            catch
            {

            }
            finally
            {

            }

        }
        public void ORDERAdd(string Meal, string Dish, string OrderBoxed)
        {
            try
            {

                InsertsbSql.Clear();
                sbSql.Clear();
                //ADD COPTC

                if (Meal.Equals("10+20"))
                {
                    DataSet ds1 = new DataSet();
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE  CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}'  ", EmployeeID, "10");

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter.Fill(ds1, "TEMPds1");
                    sqlConn.Close();

                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {
                        //InsertsbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') AND [EATNUM]=0", EmployeeID, Meal);
                        Meal = "10";
                        InsertsbSql.Append(" ");
                        InsertsbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] ([SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}',N'{2}','{3}',GETDATE(),'{4}','{5}',1) ", DateTime.Now.ToString("yyyyMMddHHmmss") ,EmployeeID, Name, CardNo, Meal, Dish);
                    }

                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}'  ", EmployeeID, "20");

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter.Fill(ds1, "TEMPds1");
                    sqlConn.Close();
                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {
                        Meal = "20";
                        InsertsbSql.Append(" ");
                        InsertsbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] ([SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}',N'{2}','{3}',GETDATE(),'{4}','{5}',1) ", DateTime.Now.ToString("yyyyMMddHHmmss"), EmployeeID, Name, CardNo, Meal, Dish);
                    }

                }
                else 
                {
                    DataSet ds1 = new DataSet();
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM] FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' ", EmployeeID, Meal);

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds1.Clear();
                    adapter.Fill(ds1, "TEMPds1");
                    if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                    {
                        Name = ds1.Tables["TEMPds1"].Rows[0][2].ToString();
                    }
                    sqlConn.Close();

                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {
                        InsertsbSql.Append(" ");
                        //InsertsbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' AND [EATNUM]=0 ", EmployeeID, Meal);
                        InsertsbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] ([SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}',N'{2}','{3}',GETDATE(),'{4}','{5}',1) ", DateTime.Now.ToString("yyyyMMddHHmmss"), EmployeeID, Name, CardNo, Meal, Dish);
                    }
                    else
                    {
                        //AutoClosingMessageBox.Show("已經訂過餐了!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + "已經訂過餐了!!!!");
                    }

                }


                if (!string.IsNullOrEmpty(InsertsbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();
                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = InsertsbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        if (Lang.Equals("CH"))
                        {                           
                            SHOWMESSAGE(Name + " 訂餐失敗!!");
                        }
                        else if (Lang.Equals("VN"))
                        {
                            SHOWMESSAGE(Name + " đặt hàng không!!");
                        }
                        PLAYMP3();
                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        if (Lang.Equals("CH"))
                        {  
                            SHOWMESSAGE(Name + " 訂餐成功!!" + " 訂了: " + OrderBoxed.ToString());

                        }
                        else if (Lang.Equals("VN"))
                        {                           
                            SHOWMESSAGE(Name + " thành công đặt phòng!!" + "  bạn đặt: " + OrderBoxed.ToString());
                        }
                    }

                    sqlConn.Close();
                }
                else
                {
                    if (Lang.Equals("CH"))
                    {
                        label5.Text = "";
                        
                    }
                    else if (Lang.Equals("VN"))
                    {
                        label5.Text = "";
                        
                    }
                }

                Search();
            }
            catch
            {

            }
            finally
            {

            }
            textBox1.Select();
        }

        public void OrderCanel(string Meal, string Dish, string OrderBoxed)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD COPTC

                if (Meal.Equals("10+20"))
                {
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') AND [DISH]='{1}' ", EmployeeID, Dish);
                }
                else
                {
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' AND [DISH]='{2}' ", EmployeeID, Meal, Dish);
                }

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
                        //label5.Text = "取消訂餐失敗!";
                        //label4.Text = "";
                        //AutoClosingMessageBox.Show("取消訂餐失敗!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + " 取消訂餐失敗!!");
                    }
                    else if (Lang.Equals("VN"))
                    {
                        //label5.Text = "Hủy bỏ Đặt Không!";
                        //label4.Text = "";
                        //AutoClosingMessageBox.Show("Hủy bỏ Đặt Không!!", "TITLE", messagetime);
                        SHOWMESSAGE(Name + " Hủy bỏ Đặt Không!!");
                    }
                    PLAYMP3();
                }
                else
                {
                    tran.Commit();      //執行交易  
                    if (Lang.Equals("CH"))
                    {
                        //label5.Text = "取消訂餐成功!";
                        //label4.Text = Name.ToString() + " 您取消了: " + OrderBoxed.ToString();

                        //AutoClosingMessageBox.Show("取消訂餐成功!!"+ Name.ToString() + " 您取消了: " + OrderBoxed.ToString(), "TITLE", messagetime);
                        SHOWMESSAGE(Name + " 取消訂餐成功!!" + " 您取消了: " + OrderBoxed.ToString());
                    }
                    else if (Lang.Equals("VN"))
                    {
                        //label5.Text = "Hủy bỏ thành công Reservation!";
                        //label4.Text = Name.ToString() + " bạn đã huỷ: " + OrderBoxed.ToString();
                        //AutoClosingMessageBox.Show("Hủy bỏ thành công Reservation!!" + Name.ToString() + "  bạn đã huỷ: " + OrderBoxed.ToString(), "TITLE", messagetime);
                        SHOWMESSAGE(Name + " Hủy bỏ thành công Reservation!!" + "   bạn đã huỷ: " + OrderBoxed.ToString());
                    }

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
            textBox1.Select();
        }

        public void OrderLast()
        {

            //if ((DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0 || (DateTime.Compare(startdinnerdt, comdt) < 0 && DateTime.Compare(enddinnerdt, comdt) > 0)))
            //{
            //    try
            //    {

            //        lastdate = SearchLastDate();

            //        connectionString = ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString;
            //        sqlConn = new SqlConnection(connectionString);

            //        sqlConn.Close();
            //        sqlConn.Open();
            //        tran = sqlConn.BeginTransaction();

            //        sbSql.Clear();
            //        //ADD COPTC
            //        //sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE  CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND [ID]='{0}' ", EmployeeID);
            //        sbSql.AppendFormat(" INSERT INTO [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM])");
            //        sbSql.AppendFormat(" SELECT [ID],[NAME],[CARDNO],GETDATE()AS  [DATE],[MEAL],[DISH],[NUM],0 AS [EATNUM] ");
            //        sbSql.AppendFormat(" FROM [TKBOXEDMEAL].[dbo].[EMPORDER]");
            //        sbSql.AppendFormat(" WHERE [ID]='{0}' AND  CONVERT(varchar(100),[DATE],112)='{1}'", EmployeeID, lastdate);
            //        sbSql.AppendFormat(" AND NOT EXISTS");
            //        sbSql.AppendFormat(" (SELECT [ID],[MEAL],[DISH] FROM [TKBOXEDMEAL].[dbo].[EMPORDER] AS EMP2");
            //        sbSql.AppendFormat(" WHERE EMP2.[ID]=EMPORDER.[ID]");
            //        sbSql.AppendFormat(" AND EMP2.[MEAL]=EMPORDER.[MEAL]");
            //        sbSql.AppendFormat(" AND EMP2.[DISH]=EMPORDER.[DISH]");
            //        sbSql.AppendFormat(" AND CONVERT(varchar(100), EMP2.[DATE],112)=CONVERT(varchar(100),getdate(),112))");
            //        sbSql.AppendFormat(" ");


            //        cmd.Connection = sqlConn;
            //        cmd.CommandTimeout = 60;
            //        cmd.CommandText = sbSql.ToString();
            //        cmd.Transaction = tran;
            //        result = cmd.ExecuteNonQuery();
            //        if (result == 0)
            //        {
            //            tran.Rollback();    //交易取消

            //            if (Lang.Equals("CH"))
            //            {
            //                //label5.Text = "訂餐失敗!";
            //                //label4.Text = "";
            //                //AutoClosingMessageBox.Show("訂餐失敗!!", "TITLE", messagetime);
            //                SHOWMESSAGE(Name + " 訂餐失敗!!");
            //            }
            //            else if (Lang.Equals("VN"))
            //            {
            //                //label5.Text = "đặt hàng không!";
            //                //label4.Text = "";
            //                //AutoClosingMessageBox.Show("đặt hàng không!!", "TITLE", messagetime);
            //                SHOWMESSAGE(Name + " 訂餐失敗!!");
            //            }
            //            PLAYMP3();
            //        }
            //        else
            //        {
            //            tran.Commit();      //訂餐成功  
            //            OrderBoxed = SearchMeal();
            //            if (!string.IsNullOrEmpty(OrderBoxed))
            //            {
            //                if (Lang.Equals("CH"))
            //                {
            //                    //label5.Text = "訂餐成功!";
            //                    //label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();
            //                    //AutoClosingMessageBox.Show("訂餐成功!!"+ Name.ToString() + " 您訂了: " + OrderBoxed.ToString(), "TITLE", messagetime);
            //                    SHOWMESSAGE(Name + " 訂餐成功!!" + " 您訂了: " + OrderBoxed.ToString());
            //                }
            //                else if (Lang.Equals("VN"))
            //                {
            //                    //label5.Text = "thành công đặt phòng!";
            //                    //label4.Text = Name.ToString() + " bạn đặt: " + OrderBoxed.ToString();
            //                    //AutoClosingMessageBox.Show("thành công đặt phòng!!" + Name.ToString() + " bạn đặt: " + OrderBoxed.ToString(), "TITLE", messagetime);
            //                    SHOWMESSAGE(Name + " thành công đặt phòng!!" + " bạn đặt: " + OrderBoxed.ToString());
            //                }
            //            }

            //        }

            //        sqlConn.Close();
            //        Search();

            //        textBox1.Text = null;
            //    }
            //    catch
            //    {

            //    }
            //    finally
            //    {

            //    }

            //}
            //else
            //{

            //    if (Lang.Equals("CH"))
            //    {
            //        //label5.Text = "超過可點餐時間!";
            //        //label4.Text = "";
            //        //AutoClosingMessageBox.Show("超過可點餐時間!!", "TITLE", messagetime);
            //        SHOWMESSAGE(Name + " 超過可點餐時間!!");
            //    }
            //    else if (Lang.Equals("VN"))
            //    {
            //        //label5.Text = "Vượt quá thời gian bữa ăn!";
            //        //label4.Text = "";
            //        //AutoClosingMessageBox.Show("Vượt quá thời gian bữa ăn!!", "TITLE", messagetime);
            //        SHOWMESSAGE(Name + " Vượt quá thời gian bữa ăn!!");
            //    }

            //    PLAYMP3();
            //}
            //textBox1.Select();
        }

        public string SearchLastDate()
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT TOP 1 CONVERT(varchar(100),[DATE],112) AS DATE    FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER] WHERE [ID]='{0}' ORDER BY [DATE] DESC ", EmployeeID);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        return ds.Tables["TEMPds1"].Rows[0]["DATE"].ToString();
                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }

        }
        public string SearchMeal()
        {           QueryMeal = null;
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"SELECT [MEAL],[DISH] FROM [{0}].[dbo].[LOCALEMPORDER] WHERE CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND [ID]='{1}'", sqlConn.Database.ToString(), EmployeeID);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();

                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

                }
                else
                {
                    for (int i = 0; i < ds2.Tables["TEMPds2"].Rows.Count; i++)
                    {
                        if (ds2.Tables["TEMPds2"].Rows[i][0].ToString().Equals("10"))
                        {
                            if (Lang.Equals("CH"))
                            {
                                QueryMeal = QueryMeal + "中餐-";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                QueryMeal = QueryMeal + "Trung Quốc-";
                            }
                        }
                        if (ds2.Tables["TEMPds2"].Rows[i][0].ToString().Equals("20"))
                        {

                            if (Lang.Equals("CH"))
                            {
                                QueryMeal = QueryMeal + "晚餐-";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                QueryMeal = QueryMeal + "bữa tối-";
                            }
                        }
                        if (ds2.Tables["TEMPds2"].Rows[i][1].ToString().Equals("1"))
                        {

                            if (Lang.Equals("CH"))
                            {
                                QueryMeal = QueryMeal + "葷 ";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                QueryMeal = QueryMeal + "món thịt";
                            }
                        }
                        if (ds2.Tables["TEMPds2"].Rows[i][1].ToString().Equals("2"))
                        {

                            if (Lang.Equals("CH"))
                            {
                                QueryMeal = QueryMeal + "素";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                QueryMeal = QueryMeal + "chay";
                            }
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

            return QueryMeal;
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

        public void SEARCHORDER()
        {
            string mess = null;
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT  [SERNO],[ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM]
                                    FROM [TKBOXEDMEAL].[dbo].[LOCALEMPORDER]
                                    WHERE CONVERT(varchar(100),[DATE],112) =CONVERT(varchar(100),GETDATE(),112)
                                    AND (ID='{0}' OR CARDNO='{0}')"
                                    , textBox1.Text.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds3.Clear();
                adapter.Fill(ds3, "TEMPds3");
                sqlConn.Close();

                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {
                    //AutoClosingMessageBox.Show("沒有訂餐記錄" + mess, "TITLE", messagetime);
                    SHOWMESSAGE(Name + " 沒有訂餐記錄!!");
                }
                else
                {


                    foreach (DataRow dr in ds3.Tables["TEMPds3"].Rows)
                    {
                        if (dr["MEAL"].ToString().Equals("10"))
                        {
                            if (Lang.Equals("CH"))
                            {
                                mess = mess + "中餐 ";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                mess = mess + "Ăn trưa ";
                            }

                        }
                        else if (dr["MEAL"].ToString().Equals("20"))
                        {
                            if (Lang.Equals("CH"))
                            {
                                mess = mess + "晚餐 ";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                mess = mess + "Ăn trưa ";
                            }

                        }
                        else if (dr["MEAL"].ToString().Equals("30"))
                        {
                            if (Lang.Equals("CH"))
                            {
                                mess = mess + "延後加班的晚餐 ";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                mess = mess + "Bữa tối tăng ca ";
                            }

                        }

                        if (dr["DISH"].ToString().Equals("1"))
                        {
                            if (Lang.Equals("CH"))
                            {
                                mess = mess + "葷 ";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                mess = mess + "thịt";
                            }

                        }
                        else if (dr["DISH"].ToString().Equals("2"))
                        {
                            if (Lang.Equals("CH"))
                            {
                                mess = mess + "素 ";
                            }
                            else if (Lang.Equals("VN"))
                            {
                                mess = mess + "Surat";
                            }
                        }

                    }

                    // AutoClosingMessageBox.Show("您訂了" + mess, "TITLE", messagetime);
                    SHOWMESSAGE(Name + "您訂了" + mess);
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADD_LOG_LOCALEMPORDER(string KEY)
        {

            StringBuilder SQLADD = new StringBuilder();
           
            try
            {
                DataTable DT = SearchEmplyee_ID(KEY);

                if (DT != null && DT.Rows.Count >= 1)
                {
                    string ID = DT.Rows[0]["EmployeeID"].ToString();
                    string NAME = DT.Rows[0]["Name"].ToString();
                    string CARDNO = DT.Rows[0]["CardNo"].ToString();

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    SQLADD.AppendFormat(@"
                                    INSERT INTO  [TKBOXEDMEAL].[dbo].[LOG_LOCALEMPORDER]
                                    (
                                    [ID]
                                    ,[NAME]
                                    ,[CARDNO]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    )
                                    ", ID, NAME, CARDNO);

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = SQLADD.ToString();
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



                    sqlConn.Close();
                }
               
            }
            catch { }
            finally { }
         
        
        }

        public DataTable SearchEmplyee_ID(string KEY)
        {
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnlocal"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"SELECT TOP 1  [EmployeeID],[CardNo],[Name] FROM [TKBOXEDMEAL].[dbo].[EMPLOYEE] WHERE [EmployeeID]='{1}' OR [CardNo]='{1}'", sqlConn.Database.ToString(), KEY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["TEMPds1"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }

        }
        public void SHOWMESSAGE(String mess)
        {
            String text = mess;
            Message message = new Message(text);
            message.Show();
        }

        #endregion


        #region BUTTON
        private void button11_Click(object sender, EventArgs e)
        {
            CreateResourceManager(this, "zh-TW");

            this.WindowState = FormWindowState.Normal;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;

            Lang = "CH";
            ds.Tables["TEMPds"].Columns.Clear();
            Search();

            SetCancel();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            CreateResourceManager(this, "vi-VN");

            this.WindowState = FormWindowState.Normal;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;

            Lang = "VN";
            ds.Tables["TEMPds"].Columns.Clear();
            Search();

            SetCancel();
        }


        private void button10_Click(object sender, EventArgs e)
        {
            SetCancel();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;

            if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
            {
                InputID = textBox1.Text.ToString();
                SearchEmplyee();
                if (!string.IsNullOrEmpty(Name))
                {
                    OrderLast();
                }
            }
            SetCancel();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;
            OrderCancel = "Order";

            ADD_LOG_LOCALEMPORDER(textBox1.Text.Trim());
            SetOrderButton();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;
            OrderCancel = "Cancel";
            SetCancelButton();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;
            Meal = "10";
            Dish = "1";

            if (Lang.Equals("CH"))
            {
                OrderBoxed = "中餐-葷";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Ăn trưa -thịt";
            }

            if (OrderCancel.Equals("Order"))
            {
                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }



            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;

            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;

            Meal = "20";
            Dish = "1";

            if (Lang.Equals("CH"))
            {
                OrderBoxed = "晚餐-葷";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Ăn tối - C";
            }
            if (OrderCancel.Equals("Order"))
            {
                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }


            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;

            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;

            Meal = "10+20";
            Dish = "1";

            if (Lang.Equals("CH"))
            {
                OrderBoxed = "中/晚餐-葷";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Ăn trưa / Ăn tối - bẩn ";
            }

            if (OrderCancel.Equals("Order"))
            {

                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }

            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;

            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;

            Meal = "10";
            Dish = "2";

            if (Lang.Equals("CH"))
            {
                OrderBoxed = "中餐-素";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Thực phẩm Trung Quốc - Surat";
            }

            if (OrderCancel.Equals("Order"))
            {
                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }

            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;


            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;

            Meal = "20";
            Dish = "2";
            if (Lang.Equals("CH"))
            {
                OrderBoxed = "晚餐-素";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Ăn tối - Surat";
            }


            if (OrderCancel.Equals("Order"))
            {

                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }

            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;

            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            comdt = DateTime.Now;

            Meal = "10+20";
            Dish = "2";
            if (Lang.Equals("CH"))
            {
                OrderBoxed = "中/晚餐-素";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Ăn trưa / Ăn tối - Surat";
            }


            if (OrderCancel.Equals("Order"))
            {

                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }

            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;

            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SearchEmplyee();
            SEARCHORDER();
            SetCancel();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DateTime omdt = DateTime.Now;

            Meal = "30";
            Dish = "1";

            if (Lang.Equals("CH"))
            {
                OrderBoxed = "延後加班晚餐-葷";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Bữa tối tăng ca - Ăn mặn ";
            }
            if (OrderCancel.Equals("Order"))
            {
                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }


            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button14.Visible = false;
            button15.Visible = false;

            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DateTime omdt = DateTime.Now;

            Meal = "30";
            Dish = "2";

            if (Lang.Equals("CH"))
            {
                OrderBoxed = "延後加班晚餐-素";
            }
            else if (Lang.Equals("VN"))
            {
                OrderBoxed = "Bữa tối tăng ca - Ăn chay";
            }
            if (OrderCancel.Equals("Order"))
            {
                if (!string.IsNullOrEmpty(Name))
                {
                    ORDERAdd(Meal, Dish, OrderBoxed);
                }
            }
            else if (OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }


            //button1.Visible = true;
            button2.Visible = true;
            //button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button14.Visible = false;
            button15.Visible = false;

            textBox1.Text = "";
            EmployeeID = null;
            Name = null;
            CardNo = null;
            Meal = null;

            SetCancel();
        }
        #endregion


    }
}
