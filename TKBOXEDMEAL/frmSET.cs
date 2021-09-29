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
using TKITDLL;

namespace TKBOXEDMEAL
{
    public partial class frmSET : Form
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
        string NowDay;
        int rownum = 0;

        public frmSET()
        {
            InitializeComponent();
            
        }

        #region FUNCTION       
        private void frmSET_Load(object sender, EventArgs e)
        {
            Search();
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            var curRow = dataGridView1.CurrentRow;
            if (curRow != null)
            {
                textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                comboBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[2].Value.ToString());
                dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[3].Value.ToString());
                dateTimePicker3.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[4].Value.ToString());
                dateTimePicker4.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[5].Value.ToString());

            }
        }
        public void Search()
        {
            try
            {

                if (!string.IsNullOrEmpty(dateTimePicker1.Text.ToString()) && !string.IsNullOrEmpty(dateTimePicker2.Text.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                        textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString();
                        //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADDtoDB()
        {


            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                
                sbSql.AppendFormat(" INSERT INTO  [{0}].[dbo].[BOXEDMEALSET] ([NAME],[STARTORDERTIME],[ENDORDERTIME],[STARTEATTIME],[ENDEATTIME])  VALUES ('{1}','{2}','{3}','{4}','{5}') ", sqlConn.Database.ToString(), comboBox1.Text.ToString(),dateTimePicker1.Value.ToString("HH:mm"), dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), dateTimePicker4.Value.ToString("HH:mm"));
                //sbSql.AppendFormat("  UPDATE Member SET Cname='{1}',Mobile1='{2}' WHERE ID='{0}' ", list_Member[0].ID.ToString(), list_Member[0].Cname.ToString(), list_Member[0].Mobile1.ToString());

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
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

                rownum = dataGridView1.RowCount;

                Search();

               
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

        }

        public void UpdateDB()
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("是否真的要更新", "UPDATE?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //sbSql.Append("UPDATE Member SET Cname='009999',Mobile1='009999',Telphone='',Email='',Address='',Sex='',Birthday='' WHERE ID='009999'");

                    sbSql.AppendFormat("UPDATE [{0}].dbo.[BOXEDMEALSET]   SET [NAME]='{2}',[STARTORDERTIME]='{3}',[ENDORDERTIME]='{4}',[STARTEATTIME]='{5}',[ENDEATTIME]='{6}' WHERE [ID]='{1}' ", sqlConn.Database.ToString(), textBox1.Text.ToString(), comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("HH:mm"), dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), dateTimePicker4.Value.ToString("HH:mm"));
                    //sbSql.AppendFormat("  UPDATE Member SET Cname='{1}',Mobile1='{2}' WHERE ID='{0}' ", list_Member[0].ID.ToString(), list_Member[0].Cname.ToString(), list_Member[0].Mobile1.ToString());

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
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

                    Search();
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ClearText()
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;
            dateTimePicker4.Value = DateTime.Now;
        }
        public void DelDB()
        {
            try
            {
                textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString();
                DialogResult dialogResult = MessageBox.Show("是否真的要刪除", "del?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //sbSql.Append("UPDATE Member SET Cname='009999',Mobile1='009999',Telphone='',Email='',Address='',Sex='',Birthday='' WHERE ID='009999'");

                    sbSql.AppendFormat("DELETE [{0}].dbo.[BOXEDMEALSET] WHERE ID='{1}' ", sqlConn.Database.ToString(), textBox1.Text.ToString());
                    //sbSql.AppendFormat("  UPDATE Member SET Cname='{1}',Mobile1='{2}' WHERE ID='{0}' ", list_Member[0].ID.ToString(), list_Member[0].Cname.ToString(), list_Member[0].Mobile1.ToString());

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
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

                    Search();
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDtoDB();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UpdateDB();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DelDB();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            ClearText();
        }



        #endregion

       
    }
}
