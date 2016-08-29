using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace TKBOXEDMEAL
{
    
    public partial class frmREPORT : Form
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
        string tablename = null;
        int rownum = 0;

        public frmREPORT()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);



                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    label1.Text = "資料筆數:" + ds.Tables[tablename].Rows.Count.ToString();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

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

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder(); 

            if (comboBox1.Text.ToString().Equals("便當統計"))
            {

                STR.Append(@" SELECT CONVERT(varchar(10),[DATE],112) AS '日期',[MEAL].[MEALNAME] AS '午晚餐',[MEALDISH].[DISHNAME] AS '葷素',SUM([NUM]) AS '訂餐量'");
                STR.Append(@" FROM [TKBOXEDMEAL].[dbo].[EMPORDER],[TKBOXEDMEAL].[dbo].[MEAL],[TKBOXEDMEAL].[dbo].[MEALDISH] ");
                STR.Append(@" WHERE [EMPORDER].[MEAL]=[MEAL].[MEAL] AND [EMPORDER].[DISH]=[MEALDISH].[DISH]");
                STR.AppendFormat(@"  AND CONVERT(varchar(10),[DATE],112)='{0}' ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                STR.Append(@" GROUP BY CONVERT(varchar(10),[DATE],112),[MEAL].[MEALNAME],[MEALDISH].[DISHNAME]");
                tablename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals("各人訂單及用餐查詢"))
            {
                STR.Append(@" SELECT [ID] AS '工號' ,[NAME] AS '姓名',[CARDNO] AS '卡號',CONVERT(varchar(10),[DATE],112) AS '日期',[MEAL].[MEALNAME] AS '午晚餐',[MEALDISH].[DISHNAME] AS '葷素',[NUM] AS '訂餐量',[EATNUM] AS '用餐量'");
                STR.Append(@" FROM [TKBOXEDMEAL].[dbo].[EMPORDER],[TKBOXEDMEAL].[dbo].[MEAL],[TKBOXEDMEAL].[dbo].[MEALDISH]");
                STR.Append(@" WHERE [EMPORDER].[MEAL]=[MEAL].[MEAL] AND [EMPORDER].[DISH]=[MEALDISH].[DISH]");
                STR.AppendFormat(@"  AND CONVERT(varchar(10),[DATE],112)='{0}' ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                tablename = "TEMPds2";
            }
            else if (comboBox1.Text.ToString().Equals("有訂未用餐查詢"))
            {
                STR.Append(@" SELECT [ID] AS '工號' ,[NAME] AS '姓名',[CARDNO] AS '卡號',CONVERT(varchar(10),[DATE],112) AS '日期',[MEAL].[MEALNAME] AS '午晚餐',[MEALDISH].[DISHNAME] AS '葷素',[NUM] AS '訂餐量',[EATNUM] AS '用餐量'");
                STR.Append(@" FROM [TKBOXEDMEAL].[dbo].[EMPORDER],[TKBOXEDMEAL].[dbo].[MEAL],[TKBOXEDMEAL].[dbo].[MEALDISH]");
                STR.Append(@" WHERE [EMPORDER].[MEAL]=[MEAL].[MEAL] AND [EMPORDER].[DISH]=[MEALDISH].[DISH]");
                STR.AppendFormat(@"  AND CONVERT(varchar(10),[DATE],112)='{0}'  AND [NUM]<>[EATNUM]", dateTimePicker1.Value.ToString("yyyyMMdd"));
                tablename = "TEMPds2";
            }




            return STR;
        }

        public void ExcelExport()
        {
            Search();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables[tablename];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }


            int j = 0;
            if (tablename.Equals("TEMPds1"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                   
                    j++;
                }
            }
            if (tablename.Equals("TEMPds2"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));

                    j++;
                }

            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\便當{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
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
            ExcelExport();
        }
        #endregion


    }
}
