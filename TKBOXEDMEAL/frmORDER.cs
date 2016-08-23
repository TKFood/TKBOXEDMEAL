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
        DateTime comdt;
        string InputID;
        string CardID;
        string ID;
        string Name;
        string Meal;
        string Dish;
        string OrderCancel;
        string QueryMeal;



        public frmORDER()
        {
            InitializeComponent();

        }

        #region FUNCTION
        private void frmORDER_Load(object sender, EventArgs e)
        {
            Search();
            textBox1.Select();
            //comdt = DateTime.Now;
            comdt = Convert.ToDateTime("09:10");
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


       

        public void SetOrderButton()
        {

            if (DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0)
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
                label5.Text = "超過可點餐時間!";
                //label4.Text = "";
            }

        }

        public void SetCancelButton()
        {     
            if (DateTime.Compare(startdt, comdt) < 0 && DateTime.Compare(enddt, comdt) > 0)
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
                label5.Text = "超過可取消點餐時間!";
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

                sbSql.AppendFormat(@"SELECT TOP 1 [ID],[NAME],[CARDID] FROM [{0}].[dbo].[EMPLOYEE] WHERE [CARDID]='{1}' OR [ID]='{1}'", sqlConn.Database.ToString(), InputID);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    label5.Text = "沒有此員工!";
                    label4.Text = "";

                    textBox1.Text = "";
                    ID = null;
                    Name = null;
                    CardID = null;
                    Meal = null;
                }
                else
                {
                    ID = ds1.Tables["TEMPds1"].Rows[0][0].ToString();
                    Name= ds1.Tables["TEMPds1"].Rows[0][1].ToString();
                    CardID= ds1.Tables["TEMPds1"].Rows[0][2].ToString();

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
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') ", ID, Meal);
                    Meal = "10";
                    sbSql.Append(" ");                    
                    sbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDID],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}',GETDATE(),'{3}','{4}',1) ", ID, Name, CardID, Meal,Dish);

                    Meal = "20";
                    sbSql.Append(" ");                 
                    sbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDID],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}',GETDATE(),'{3}','{4}',1) ", ID, Name, CardID, Meal, Dish);

                }              
                else
                {
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' ", ID, Meal);
                    sbSql.AppendFormat(" INSERT INTO  [TKBOXEDMEAL].[dbo].[EMPORDER] ([ID],[NAME],[CARDID],[DATE],[MEAL],[DISH],[NUM]) VALUES ('{0}','{1}','{2}',GETDATE(),'{3}','{4}',1) ", ID, Name, CardID, Meal, Dish);
                }

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();
                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    label5.Text = "訂餐失敗!";
                    label4.Text = "";
                }
                else
                {
                    tran.Commit();      //執行交易  
                    label5.Text = "訂餐成功!";
                    label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();
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
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  ([MEAL]='10' OR [MEAL]='20') AND [DISH]='{1}' ", ID, Dish);   
                }
                else
                {
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE], 112)=CONVERT(varchar(100),GETDATE(), 112) AND [ID]='{0}' AND  [MEAL]='{1}' AND [DISH]='{2}' ", ID, Meal,Dish);                   
                }

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();
                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    label5.Text = "取消訂餐失敗!";
                    label4.Text = "";
                }
                else
                {
                    tran.Commit();      //執行交易  
                    label5.Text = "取消訂餐成功!";
                    label4.Text = Name.ToString() + " 您取消了: " + OrderBoxed.ToString();
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
                    sbSql.AppendFormat(" DELETE [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE  CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND [ID]='{0}' ", ID);
                    sbSql.Append(" INSERT INTO [TKBOXEDMEAL].[dbo].[EMPORDER] SELECT [ID],[NAME],[CARDID],GETDATE()AS  [DATE],[MEAL],[DISH],[NUM] FROM [TKBOXEDMEAL].[dbo].[EMPORDER]");
                    sbSql.AppendFormat(" WHERE CONVERT(varchar(20),[DATE],112) IN (SELECT TOP 1 CONVERT(varchar(20),[DATE],112) FROM [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE [ID]='{0}' ORDER BY SERNO DESC) ", ID);


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();
                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        label5.Text = "訂餐失敗!";
                        label4.Text = "";
                    }
                    else
                    {
                        tran.Commit();      //訂餐成功  
                        OrderBoxed=SearchMeal();
                        if(!string.IsNullOrEmpty(OrderBoxed))
                        {
                            label5.Text = "訂餐成功!";
                            label4.Text = Name.ToString() + " 您訂了: " + OrderBoxed.ToString();
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
                label5.Text = "超過可點餐時間!";
                label4.Text = "";
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

                sbSql.AppendFormat(@"SELECT [MEAL],[DISH] FROM [{0}].[dbo].[EMPORDER] WHERE CONVERT(varchar(20),[DATE],112)=CONVERT(varchar(20),GETDATE(),112) AND [ID]='{1}'", sqlConn.Database.ToString(), ID);

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
                            QueryMeal = QueryMeal + "中餐-";
                        }
                        if (ds2.Tables["TEMPds2"].Rows[i][0].ToString().Equals("20"))
                        {
                            QueryMeal = QueryMeal + "晚餐-";
                        }
                        if (ds2.Tables["TEMPds2"].Rows[i][1].ToString().Equals("1"))
                        {
                            QueryMeal = QueryMeal + "葷 ";
                        }
                        if (ds2.Tables["TEMPds2"].Rows[i][1].ToString().Equals("2"))
                        {
                            QueryMeal = QueryMeal + "素 ";
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
            OrderBoxed = "中餐-葷";

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
            ID = null;
            Name = null;
            CardID = null;
            Meal = null;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Meal = "20";
            Dish = "1";
            OrderBoxed = "晚餐-葷";

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
            ID = null;
            Name = null;
            CardID = null;
            Meal = null;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Meal = "10+20";
            Dish = "1";
            OrderBoxed = "中/晚餐-葷";

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
            ID = null;
            Name = null;
            CardID = null;
            Meal = null;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Meal = "10";
            Dish = "2";
            OrderBoxed = "中餐-素";

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
            ID = null;
            Name = null;
            CardID = null;
            Meal = null;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Meal = "20";
            Dish = "2";
            OrderBoxed = "晚餐-素";

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
            ID = null;
            Name = null;
            CardID = null;
            Meal = null;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Meal = "10+20";
            Dish = "2";
            OrderBoxed = "中/晚餐-素";

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
            ID = null;
            Name = null;
            CardID = null;
            Meal = null;
        }




        #endregion


    }
}
