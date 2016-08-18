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
                        rownum = ds.Tables["TEMPds"].Rows.Count - 1;
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

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //sbSql.Append("UPDATE Member SET Cname='009999',Mobile1='009999',Telphone='',Email='',Address='',Sex='',Birthday='' WHERE ID='009999'");

                //sbSql.AppendFormat(" INSERT INTO  [{0}].[dbo].[CALLRECORD] ([CallDate],[CallTime],[TypeID],[CallName],[CallPhone],[OrderID],[ShipID],[InvoiceNo],[CallText],[CallTextRe])  VALUES ('{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}') ", sqlConn.Database.ToString(), list_CallsRecord[0].CallDate.ToString(), list_CallsRecord[0].CallTime.ToString(), list_CallsRecord[0].TypeID.ToString(), list_CallsRecord[0].CallName.ToString(), list_CallsRecord[0].CallPhone.ToString(), list_CallsRecord[0].OrderID.ToString(), list_CallsRecord[0].ShipID.ToString(), list_CallsRecord[0].InvoiceNo.ToString(), list_CallsRecord[0].CallText.ToString(), list_CallsRecord[0].CallTextRe.ToString());
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
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //sbSql.Append("UPDATE Member SET Cname='009999',Mobile1='009999',Telphone='',Email='',Address='',Sex='',Birthday='' WHERE ID='009999'");

                    //sbSql.AppendFormat("UPDATE [{0}].dbo.[CALLRECORD]   SET [TypeID]='{2}',[CallName]='{3}',[CallPhone]='{4}',[CallText]='{5}',[CallTextRe]='{6}',[OrderID]='{7}',[ShipID]='{8}',[InvoiceNo]='{9}' WHERE [ID]='{1}' ", sqlConn.Database.ToString(), textBox5.Text.ToString(), comboBox1.SelectedValue.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString());
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

        }
        public void DelDB()
        {
            try
            {
                textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString();
                DialogResult dialogResult = MessageBox.Show("是否真的要刪除", "del?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //sbSql.Append("UPDATE Member SET Cname='009999',Mobile1='009999',Telphone='',Email='',Address='',Sex='',Birthday='' WHERE ID='009999'");

                    sbSql.AppendFormat("DELETE [{0}].dbo.[CALLRECORD] WHERE ID='{1}' ", sqlConn.Database.ToString(), textBox1.Text.ToString());
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
