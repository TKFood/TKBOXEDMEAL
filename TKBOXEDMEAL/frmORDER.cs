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
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
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

        public frmORDER()
        {
            InitializeComponent();

        }

        #region FUNCTION
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
                            startdt = Convert.ToDateTime(row[2].ToString());
                            enddt = Convert.ToDateTime(row[3].ToString());
                        }
                        DataRow[] result2 = ds.Tables["TEMPds"].Select("名稱='晚餐'");
                        foreach (DataRow row2 in result2)
                        {
                            startdinnerdt = Convert.ToDateTime(row2[2].ToString());
                            enddinnerdt = Convert.ToDateTime(row2[3].ToString());
                        }

                    }

                    else if (Lang.Equals("VN"))
                    {
                        DataRow[] result = ds.Tables["TEMPds"].Select("tên='bữa ăn trưa'");
                        foreach (DataRow row in result)
                        {
                            startdt = Convert.ToDateTime(row[2].ToString());
                            enddt = Convert.ToDateTime(row[3].ToString());
                        }
                        DataRow[] result2 = ds.Tables["TEMPds"].Select("tên='bữa tối'");
                        foreach (DataRow row2 in result2)
                        {
                            startdinnerdt = Convert.ToDateTime(row2[2].ToString());
                            enddinnerdt = Convert.ToDateTime(row2[3].ToString());
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

            if ((DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0|| (DateTime.Compare(startdinnerdt, comdt) < 0 && DateTime.Compare(enddinnerdt, comdt) > 0)))
            {
                if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    InputID = textBox1.Text.ToString();
                    SearchEmplyee();
                    
                    if (!string.IsNullOrEmpty(Name))
                    {
                        button3.Visible = true;
                        button4.Visible = true;
                        button5.Visible = true;
                        button6.Visible = true;
                        button7.Visible = true;
                        button8.Visible = true;

                        button1.Visible = false;
                        button9.Visible = false;
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
                    label5.Text = "超過可點餐時間!!";
                }
                else if (Lang.Equals("VN"))
                {
                    label5.Text = "Vượt quá thời gian bữa ăn!";
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
                        button5.Visible = true;
                        button6.Visible = true;
                        button7.Visible = true;
                        button8.Visible = true;

                        button1.Visible = false;
                        button2.Visible = false;
                    }
                }                
            }
            else
            {
                
                if (Lang.Equals("CH"))
                {
                    label5.Text = "超過可取消點餐時間!";
                }
                else if (Lang.Equals("VN"))
                {
                    label5.Text = "Qua thời gian để hủy bỏ các bữa ăn!";
                }
                PLAYMP3();
                //label4.Text = "";
            }
        }

        public void SetCancel()
        {
            button1.Visible = true;
            button2.Visible = true;
            button9.Visible = true;

            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;

            textBox1.Text = null;
            textBox1.Select();
        }

        public void SearchEmplyee()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"SELECT TOP 1  [EmployeeID],[CardNo],[Name]FROM [TKBOXEDMEAL].[dbo].[VEMPLOYEE] WHERE [EmployeeID]='{1}' OR [CardNo]='{1}'", sqlConn.Database.ToString(), InputID);

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
                        label5.Text = "沒有此員工!";
                        label4.Text = "";
                    }
                    else if (Lang.Equals("VN"))
                    {
                        label5.Text = "Không có nhân viên!";
                        label4.Text = "";
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
        public void ORDERAdd(string Meal,string Dish, string OrderBoxed)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD COPTC
                
                if (Meal.Equals("10+20"))
                {
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') ", EmployeeID, Meal);
                    Meal = "10";
                    sbSql.Append(" ");                    
                    sbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}',GETDATE(),'{3}','{4}',1) ", EmployeeID, Name, CardNo, Meal,Dish);

                    Meal = "20";
                    sbSql.Append(" ");                 
                    sbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}',GETDATE(),'{3}','{4}',1) ", EmployeeID, Name, CardNo, Meal, Dish);

                }              
                else
                {
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' ", EmployeeID, Meal);
                    sbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}',GETDATE(),'{3}','{4}',1) ", EmployeeID, Name, CardNo, Meal, Dish);
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
                        label5.Text = "訂餐失敗!";
                        label4.Text = "";
                    }
                    else if (Lang.Equals("VN"))
                    {
                        label5.Text = "đặt hàng không!";
                        label4.Text = "";
                    }
                    PLAYMP3();
                }
                else
                {
                    tran.Commit();      //執行交易  
                    if (Lang.Equals("CH"))
                    {
                        label5.Text = "訂餐成功!";
                        label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();

                    }
                    else if (Lang.Equals("VN"))
                    {
                        label5.Text = "thành công đặt phòng!";
                        label4.Text = Name.ToString() + " bạn đặt: " + OrderBoxed.ToString();

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

        public void OrderCanel(string Meal, string Dish, string OrderBoxed)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD COPTC

                if (Meal.Equals("10+20"))
                {
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') AND [DISH]='{1}' ", EmployeeID, Dish);   
                }
                else
                {
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' AND [DISH]='{2}' ", EmployeeID, Meal,Dish);                   
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
                        label5.Text = "取消訂餐失敗!";
                        label4.Text = "";
                    }
                    else if (Lang.Equals("VN"))
                    {
                        label5.Text = "Hủy bỏ Đặt Không!";
                        label4.Text = "";
                    }
                    PLAYMP3();
                }
                else
                {
                    tran.Commit();      //執行交易  
                    if (Lang.Equals("CH"))
                    {
                        label5.Text = "取消訂餐成功!";
                        label4.Text = Name.ToString() + " 您取消了: " + OrderBoxed.ToString();
                    }
                    else if (Lang.Equals("VN"))
                    {
                        label5.Text = "Hủy bỏ thành công Reservation!";
                        label4.Text = Name.ToString() + " bạn đã huỷ: " + OrderBoxed.ToString();
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

            if (DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //ADD COPTC
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE  CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND [ID]='{0}' ", EmployeeID);
                    sbSql.Append(" INSERT INTO [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDNO],[DATE],[MEAL],[DISH],[NUM],[EATNUM]) SELECT [ID],[NAME],[CARDNO],GETDATE()AS  [DATE],[MEAL],[DISH],[NUM],0 AS [EATNUM] FROM [TKBOXEDMEAL].[dbo].[EMPORDER]");
                    sbSql.AppendFormat(" WHERE CONVERT(varchar(20),[DATE],112) IN (SELECT TOP 1 CONVERT(varchar(20),[DATE],112) FROM [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE [ID]='{0}' ORDER BY SERNO DESC) ", EmployeeID);


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
                            label5.Text = "訂餐失敗!";
                            label4.Text = "";
                        }
                        else if (Lang.Equals("VN"))
                        {
                            label5.Text = "đặt hàng không!";
                            label4.Text = "";
                        }
                        PLAYMP3();
                    }
                    else
                    {
                        tran.Commit();      //訂餐成功  
                        OrderBoxed=SearchMeal();
                        if(!string.IsNullOrEmpty(OrderBoxed))
                        {                          
                            if (Lang.Equals("CH"))
                            {
                                label5.Text = "訂餐成功!";
                                label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();

                            }
                            else if (Lang.Equals("VN"))
                            {
                                label5.Text = "thành công đặt phòng!";
                                label4.Text = Name.ToString() + " bạn đặt: " + OrderBoxed.ToString();

                            }
                        }
                        
                    }

                    sqlConn.Close();
                    Search();

                    textBox1.Text = null;
                }
                catch
                {

                }
                finally
                {

                }

            }
            else
            {
                
                if (Lang.Equals("CH"))
                {
                    label5.Text = "超過可點餐時間!";
                    label4.Text = "";
                }
                else if (Lang.Equals("VN"))
                {
                    label5.Text = "Vượt quá thời gian bữa ăn!";
                    label4.Text = "";
                }
                
                PLAYMP3();
            }
            textBox1.Select();
        }

        public string SearchMeal()
        {
            QueryMeal = null;
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"SELECT [MEAL],[DISH] FROM [{0}].[dbo].[EMPORDER] WHERE CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND [ID]='{1}'", sqlConn.Database.ToString(), EmployeeID);

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
                    for(int i=0; i< ds2.Tables["TEMPds2"].Rows.Count; i++)
                    {
                        if(ds2.Tables["TEMPds2"].Rows[i][0].ToString().Equals("10"))
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
        public void SetLang()
        {
            if (Lang.Equals("CH"))
            {
                label6.Text = "請刷卡";
                button1.Text = "複製上次訂餐";
                button2.Text = "重新訂餐";
                button3.Text = "中餐-葷";
                button4.Text = "晚餐-葷";
                button5.Text = "中/晚餐-葷";
                button6.Text = "中餐-素";
                button7.Text = "晚餐-素";
                button8.Text = "中/晚餐-素";
                button9.Text = "取消訂餐";
                button10.Text = "取消本次操作";
            }
            else if (Lang.Equals("VN"))
            {
                label6.Text = "Hãy swipe";
                button1.Text = "Sao chép các đặt hàng cuối cùng";
                button2.Text = "Sắp xếp lại";
                button3.Text = "Ăn trưa - thịt";
                button4.Text = "Ăn tối - C";
                button5.Text = "Ăn trưa / Ăn tối - C";
                button6.Text = "Thực phẩm Trung Quốc - Surat";
                button7.Text = "Thực phẩm Trung Quốc - Surat";
                button8.Text = "Ăn trưa / Ăn tối - Surat";
                button9.Text = "Hủy bỏ bữa ăn";
                button10.Text = "Hủy bỏ hoạt động này";
            }
        }

        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
            {
                InputID = textBox1.Text.ToString();
                SearchEmplyee();
                if (!string.IsNullOrEmpty(Name))
                {
                    OrderLast();
                }
            }
        }
                                   
           
        
        private void button2_Click(object sender, EventArgs e)
        {
            OrderCancel = "Order";
            SetOrderButton();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OrderCancel = "Cancel";
            SetCancelButton();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SetCancel();
        }

        private void button3_Click(object sender, EventArgs e)
        {
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
            else if(OrderCancel.Equals("Cancel"))
            {
                OrderCanel(Meal, Dish, OrderBoxed);
            }



            button1.Visible = true;
            button2.Visible = true;
            button9.Visible = true;

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
        }

        private void button4_Click(object sender, EventArgs e)
        {
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


            button1.Visible = true;
            button2.Visible = true;
            button9.Visible = true;

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
        }

        private void button5_Click(object sender, EventArgs e)
        {
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

            button1.Visible = true;
            button2.Visible = true;
            button9.Visible = true;

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
        }

        private void button6_Click(object sender, EventArgs e)
        {
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

            button1.Visible = true;
            button2.Visible = true;
            button9.Visible = true;

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
        }

        private void button7_Click(object sender, EventArgs e)
        {
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

            button1.Visible = true;
            button2.Visible = true;
            button9.Visible = true;

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
        }

        private void button8_Click(object sender, EventArgs e)
        {
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

            button1.Visible = true;
            button2.Visible = true;
            button9.Visible = true;

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
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Lang = "CH";
            ds.Tables["TEMPds"].Columns.Clear();
            Search();
            SetLang();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Lang = "VN";
            ds.Tables["TEMPds"].Columns.Clear();
            Search();
            SetLang();
        }




        #endregion


    }
}
