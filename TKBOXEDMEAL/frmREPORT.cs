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
                tablename = "TEMPds3";
            }
            else if (comboBox1.Text.ToString().Equals("部門訂餐統計"))
            {
                STR.AppendFormat(@"  SELECT [Department].[Name],[ID] AS '工號' ,[EMPORDER].[NAME] AS '姓名',[CARDNO] AS '卡號'");
                STR.AppendFormat(@"  ,CONVERT(varchar(10),[EMPORDER].[DATE],112) AS '日期',[MEAL].[MEALNAME] AS '午晚餐'");
                STR.AppendFormat(@"  ,[MEALDISH].[DISHNAME] AS '葷素',[NUM] AS '訂餐量',[EATNUM] AS '用餐量'");
                STR.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[EMPORDER],[TKBOXEDMEAL].[dbo].[MEAL],[TKBOXEDMEAL].[dbo].[MEALDISH],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] ");
                STR.AppendFormat(@"  WHERE [Employee].[Code]=[EMPORDER].[ID] COLLATE Chinese_Taiwan_Stroke_BIN");
                STR.AppendFormat(@"  AND [Employee].DepartmentId=[Department] .DepartmentId");
                STR.AppendFormat(@"  AND [EMPORDER].[MEAL]=[MEAL].[MEAL] AND [EMPORDER].[DISH]=[MEALDISH].[DISH]  ");
                STR.AppendFormat(@"  AND CONVERT(varchar(10),[EMPORDER].[DATE],112)='{0}' ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  ORDER BY  CONVERT(varchar(10),[EMPORDER].[DATE],112),[Department].[Name],[MEAL].[MEALNAME],[MEALDISH].[DISHNAME],[ID]");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds4";
            }
            else if (comboBox1.Text.ToString().Equals("部門訂餐統計-午餐"))
            {
                STR.AppendFormat(@"  SELECT [Department].[Name],[ID] AS '工號' ,[EMPORDER].[NAME] AS '姓名',[CARDNO] AS '卡號'");
                STR.AppendFormat(@"  ,CONVERT(varchar(10),[EMPORDER].[DATE],112) AS '日期',[MEAL].[MEALNAME] AS '午晚餐'");
                STR.AppendFormat(@"  ,[MEALDISH].[DISHNAME] AS '葷素',[NUM] AS '訂餐量',[EATNUM] AS '用餐量'");
                STR.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[EMPORDER],[TKBOXEDMEAL].[dbo].[MEAL],[TKBOXEDMEAL].[dbo].[MEALDISH],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] ");
                STR.AppendFormat(@"  WHERE [Employee].[Code]=[EMPORDER].[ID] COLLATE Chinese_Taiwan_Stroke_BIN");
                STR.AppendFormat(@"  AND [Employee].DepartmentId=[Department] .DepartmentId");
                STR.AppendFormat(@"  AND [EMPORDER].[MEAL]=[MEAL].[MEAL] AND [EMPORDER].[DISH]=[MEALDISH].[DISH]  ");
                STR.AppendFormat(@"  AND CONVERT(varchar(10),[EMPORDER].[DATE],112)='{0}' ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  AND [EMPORDER].[MEAL]='10' ");
                STR.AppendFormat(@"  ORDER BY  CONVERT(varchar(10),[EMPORDER].[DATE],112),[Department].[Name],[MEAL].[MEALNAME],[MEALDISH].[DISHNAME],[ID]");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds5";
            }

            else if (comboBox1.Text.ToString().Equals("部門訂餐統計-晚餐"))
            {
                STR.AppendFormat(@"  SELECT [Department].[Name],[ID] AS '工號' ,[EMPORDER].[NAME] AS '姓名',[CARDNO] AS '卡號'");
                STR.AppendFormat(@"  ,CONVERT(varchar(10),[EMPORDER].[DATE],112) AS '日期',[MEAL].[MEALNAME] AS '午晚餐'");
                STR.AppendFormat(@"  ,[MEALDISH].[DISHNAME] AS '葷素',[NUM] AS '訂餐量',[EATNUM] AS '用餐量'");
                STR.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[EMPORDER],[TKBOXEDMEAL].[dbo].[MEAL],[TKBOXEDMEAL].[dbo].[MEALDISH],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] ");
                STR.AppendFormat(@"  WHERE [Employee].[Code]=[EMPORDER].[ID] COLLATE Chinese_Taiwan_Stroke_BIN");
                STR.AppendFormat(@"  AND [Employee].DepartmentId=[Department] .DepartmentId");
                STR.AppendFormat(@"  AND [EMPORDER].[MEAL]=[MEAL].[MEAL] AND [EMPORDER].[DISH]=[MEALDISH].[DISH]  ");
                STR.AppendFormat(@"  AND CONVERT(varchar(10),[EMPORDER].[DATE],112)='{0}' ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@"  AND [EMPORDER].[MEAL]='20' ");
                STR.AppendFormat(@"  ORDER BY  CONVERT(varchar(10),[EMPORDER].[DATE],112),[Department].[Name],[MEAL].[MEALNAME],[MEALDISH].[DISHNAME],[ID]");
                STR.AppendFormat(@"  ");

                tablename = "TEMPds6";
            }
            else if (comboBox1.Text.ToString().Equals("有上班未訂餐查詢"))
            {
                STR.AppendFormat(@"  SELECT  DISTINCT CONVERT(varchar(100),[AttendanceRollcall].[Date],112) AS '日期', [Department].[Name] AS '部門',[Employee].[Code] AS '工號',[Employee].[CnName] AS '姓名'");
                STR.AppendFormat(@"  FROM [SQL102].[Chiyu].[dbo].[DoorLog],[SQL102].[Chiyu].[dbo].[Person],[HRMDB].[dbo].[AttendanceRollcall],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department]");
                STR.AppendFormat(@"  WHERE [DoorLog].[UserID]=[Person].[UserID]");
                STR.AppendFormat(@"  AND [Person].[EmployeeID]=[Employee].[Code]");
                STR.AppendFormat(@"  AND CONVERT(varchar(100),[DoorLog].[DateTime],112) LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.AppendFormat(@"  AND [AttendanceRollcall].EmployeeId=[Employee].EmployeeId");
                STR.AppendFormat(@"  AND [Department].[DepartmentId]=[Employee].[DepartmentId]");
                STR.AppendFormat(@"  AND CONVERT(varchar(100),[AttendanceRollcall].[Date],112) LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.AppendFormat(@"  AND CONVERT(varchar(100),[DoorLog].[DateTime],112)= CONVERT(varchar(100),[AttendanceRollcall].[Date],112)");
                STR.AppendFormat(@"  AND NOT  EXISTS (SELECT ID FROM [TKBOXEDMEAL].[dbo].[EMPORDER] WHERE CONVERT(varchar(100),[DATE],112)=CONVERT(varchar(100),[AttendanceRollcall].[Date],112) AND ID=[Employee].[Code] COLLATE Chinese_PRC_CI_AS) ");
                STR.AppendFormat(@"  ORDER BY CONVERT(varchar(100),[AttendanceRollcall].[Date],112), [Department].[Name]");
                STR.AppendFormat(@"  AND NOT  EXISTS ");
                STR.AppendFormat(@"  (SELECT ID FROM [TKBOXEDMEAL].[dbo].[EMPCARDNOTEAT]");
                STR.AppendFormat(@"  WHERE CONVERT(varchar(100),[DATE],112)=CONVERT(varchar(100),[AttendanceRollcall].[Date],112) ");
                STR.AppendFormat(@"  AND ID=[Employee].[Code] COLLATE Chinese_PRC_CI_AS) ");
                STR.AppendFormat(@"  ");

                STR.AppendFormat(@"  ");

                tablename = "TEMPds7";
            }
            else if (comboBox1.Text.ToString().Equals("部門員工查詢"))
            {
                STR.AppendFormat(@"  SELECT  DISTINCT [Department].[Name] AS '部門',[Person].CardNo AS '卡號',[Employee].[Code] AS '工號',[Employee].[CnName] AS '姓名'  ");
                STR.AppendFormat(@"  FROM [SQL102].[Chiyu].[dbo].[Person],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department]  ");
                STR.AppendFormat(@"  WHERE [Person].[EmployeeID]=[Employee].[Code] ");
                STR.AppendFormat(@"  AND [Department].[DepartmentId]=[Employee].[DepartmentId]  ");
                STR.AppendFormat(@"  ORDER BY [Department].[Name],[Employee].[Code],[Employee].[CnName]");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds8";
            }
            else if (comboBox1.Text.ToString().Equals("每月統計"))
            {
                DateTime dt1 = dateTimePicker1.Value;
                dt1.AddMonths(-1);
                DateTime dt2 = dateTimePicker1.Value;
                string sdt1 = dt1.ToString("yyyyMM");
                string sdt2 = dt2.ToString("yyyyMM");



                STR.AppendFormat(@"  SELECT DISTINCT [ID],[NAME]");
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'26%' AND [MEAL]='10'),0) AS '26-中餐'",sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'27%' AND [MEAL]='10'),0) AS '27-中餐'", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'28%' AND [MEAL]='10'),0) AS '28-中餐'", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'29%' AND [MEAL]='10'),0) AS '29-中餐' ", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'30%' AND [MEAL]='10'),0) AS '30-中餐'", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'31%' AND [MEAL]='10'),0) AS '31-中餐'", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'01%' AND [MEAL]='10'),0) AS '01-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'02%' AND [MEAL]='10'),0) AS '02-中餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'03%' AND [MEAL]='10'),0) AS '03-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'04%' AND [MEAL]='10'),0) AS '04-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'05%' AND [MEAL]='10'),0) AS '05-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'06%' AND [MEAL]='10'),0) AS '06-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'07%' AND [MEAL]='10'),0) AS '07-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'08%' AND [MEAL]='10'),0) AS '08-中餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'09%' AND [MEAL]='10'),0) AS '09-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'10%' AND [MEAL]='10'),0) AS '10-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'11%' AND [MEAL]='10'),0) AS '11-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'12%' AND [MEAL]='10'),0) AS '12-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'13%' AND [MEAL]='10'),0) AS '13-中餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'14%' AND [MEAL]='10'),0) AS '14-中餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'15%' AND [MEAL]='10'),0) AS '15-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'16%' AND [MEAL]='10'),0) AS '16-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'17%' AND [MEAL]='10'),0) AS '17-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'18%' AND [MEAL]='10'),0) AS '18-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'19%' AND [MEAL]='10'),0) AS '19-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'20%' AND [MEAL]='10'),0) AS '20-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'21%' AND [MEAL]='10'),0) AS '21-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'22%' AND [MEAL]='10'),0) AS '22-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'23%' AND [MEAL]='10'),0) AS '23-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'24%' AND [MEAL]='10'),0) AS '24-中餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'25%' AND [MEAL]='10'),0) AS '25-中餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'26%' AND [MEAL]='20'),0) AS '26-晚餐' ", sdt1);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'27%' AND [MEAL]='20'),0) AS '27-晚餐' ", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'28%' AND [MEAL]='20'),0) AS '28-晚餐'", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'29%' AND [MEAL]='20'),0) AS '29-晚餐'", sdt1);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'30%' AND [MEAL]='20'),0) AS '30-晚餐'", sdt1);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'31%' AND [MEAL]='20'),0) AS '31-晚餐' ", sdt1);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'01%' AND [MEAL]='20'),0) AS '01-晚餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'02%' AND [MEAL]='20'),0) AS '02-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'03%' AND [MEAL]='20'),0) AS '03-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'04%' AND [MEAL]='20'),0) AS '04-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'05%' AND [MEAL]='20'),0) AS '05-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'06%' AND [MEAL]='20'),0) AS '06-晚餐'", sdt2);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'07%' AND [MEAL]='20'),0) AS '07-晚餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'08%' AND [MEAL]='20'),0) AS '08-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'09%' AND [MEAL]='20'),0) AS '09-晚餐'", sdt2);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'10%' AND [MEAL]='20'),0) AS '10-晚餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'12%' AND [MEAL]='20'),0) AS '12-晚餐'", sdt2);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'13%' AND [MEAL]='20'),0) AS '13-晚餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'14%' AND [MEAL]='20'),0) AS '14-晚餐'", sdt2);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'15%' AND [MEAL]='20'),0) AS '15-晚餐' ", sdt2);
                STR.AppendFormat(@" ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'16%' AND [MEAL]='20'),0) AS '16-晚餐' ", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'17%' AND [MEAL]='20'),0) AS '17-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'18%' AND [MEAL]='20'),0) AS '18-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'19%' AND [MEAL]='20'),0) AS '19-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'20%' AND [MEAL]='20'),0) AS '20-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'21%' AND [MEAL]='20'),0) AS '21-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'22%' AND [MEAL]='20'),0) AS '22-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'23%' AND [MEAL]='20'),0) AS '23-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'24%' AND [MEAL]='20'),0) AS '24-晚餐'", sdt2);
                STR.AppendFormat(@"  ,ISNULL((SELECT 1 FROM [TKBOXEDMEAL].[dbo].[EMPORDER] EMP WHERE EMP.[ID]=[EMPORDER].[ID] AND CONVERT(varchar(100),EMP.[DATE],112) LIKE '{0}'+'25%' AND [MEAL]='20'),0) AS '25-晚餐'", sdt2);
                STR.AppendFormat(@"  FROM [TKBOXEDMEAL].[dbo].[EMPORDER]");
                STR.AppendFormat(@"  WHERE CONVERT(varchar(100),[DATE],112)>= '{0}26' AND  CONVERT(varchar(100),[DATE],112)>= '{1}25'",sdt1,sdt2);
                STR.AppendFormat(@" ORDER BY [ID],[NAME] ");
                STR.AppendFormat(@"  ");


                tablename = "TEMPds9";
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
            if (tablename.Equals("TEMPds4"))
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
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));

                    j++;
                }

            }
            if (tablename.Equals("TEMPds5"))
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
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));

                    j++;
                }

            }
            if (tablename.Equals("TEMPds6"))
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
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));

                    j++;
                }

            }
            if (tablename.Equals("TEMPds7"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                
                    j++;
                }

            }
            if (tablename.Equals("TEMPds8"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());

                    j++;
                }

            }
            if(tablename.Equals("TEMPds9"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                    ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                    ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));
                    ws.GetRow(j + 1).CreateCell(14).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString()));
                    ws.GetRow(j + 1).CreateCell(15).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString()));
                    ws.GetRow(j + 1).CreateCell(16).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[16].ToString()));
                    ws.GetRow(j + 1).CreateCell(17).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[17].ToString()));
                    ws.GetRow(j + 1).CreateCell(18).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[18].ToString()));
                    ws.GetRow(j + 1).CreateCell(19).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[19].ToString()));
                    ws.GetRow(j + 1).CreateCell(20).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[20].ToString()));
                    ws.GetRow(j + 1).CreateCell(21).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[21].ToString()));
                    ws.GetRow(j + 1).CreateCell(22).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[22].ToString()));
                    ws.GetRow(j + 1).CreateCell(23).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[23].ToString()));
                    ws.GetRow(j + 1).CreateCell(24).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[24].ToString()));
                    ws.GetRow(j + 1).CreateCell(25).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[25].ToString()));
                    ws.GetRow(j + 1).CreateCell(26).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[26].ToString()));
                    ws.GetRow(j + 1).CreateCell(27).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[27].ToString()));
                    ws.GetRow(j + 1).CreateCell(28).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[28].ToString()));
                    ws.GetRow(j + 1).CreateCell(29).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[29].ToString()));
                    ws.GetRow(j + 1).CreateCell(30).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[30].ToString()));
                    ws.GetRow(j + 1).CreateCell(31).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[31].ToString()));
                    ws.GetRow(j + 1).CreateCell(32).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[32].ToString()));
                    ws.GetRow(j + 1).CreateCell(33).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[33].ToString()));
                    ws.GetRow(j + 1).CreateCell(34).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[34].ToString()));
                    ws.GetRow(j + 1).CreateCell(35).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[35].ToString()));
                    ws.GetRow(j + 1).CreateCell(36).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[36].ToString()));
                    ws.GetRow(j + 1).CreateCell(37).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[37].ToString()));
                    ws.GetRow(j + 1).CreateCell(38).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[38].ToString()));
                    ws.GetRow(j + 1).CreateCell(39).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[39].ToString()));
                    ws.GetRow(j + 1).CreateCell(40).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[40].ToString()));
                    ws.GetRow(j + 1).CreateCell(41).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[41].ToString()));
                    ws.GetRow(j + 1).CreateCell(42).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[42].ToString()));
                    ws.GetRow(j + 1).CreateCell(43).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[43].ToString()));
                    ws.GetRow(j + 1).CreateCell(44).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[44].ToString()));
                    ws.GetRow(j + 1).CreateCell(45).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[45].ToString()));
                    ws.GetRow(j + 1).CreateCell(46).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[46].ToString()));
                    ws.GetRow(j + 1).CreateCell(47).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[47].ToString()));
                    ws.GetRow(j + 1).CreateCell(48).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[48].ToString()));
                    ws.GetRow(j + 1).CreateCell(49).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[49].ToString()));
                    ws.GetRow(j + 1).CreateCell(50).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[50].ToString()));
                    ws.GetRow(j + 1).CreateCell(51).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[51].ToString()));
                    ws.GetRow(j + 1).CreateCell(52).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[52].ToString()));
                    ws.GetRow(j + 1).CreateCell(53).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[53].ToString()));
                    ws.GetRow(j + 1).CreateCell(54).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[54].ToString()));
                    ws.GetRow(j + 1).CreateCell(55).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[55].ToString()));
                    ws.GetRow(j + 1).CreateCell(56).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[56].ToString()));
                    ws.GetRow(j + 1).CreateCell(57).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[57].ToString()));
                    ws.GetRow(j + 1).CreateCell(58).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[58].ToString()));
                    ws.GetRow(j + 1).CreateCell(59).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[59].ToString()));
                    ws.GetRow(j + 1).CreateCell(60).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[60].ToString()));
                    ws.GetRow(j + 1).CreateCell(61).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[61].ToString()));
                    ws.GetRow(j + 1).CreateCell(62).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[62].ToString()));
                    //ws.GetRow(j + 1).CreateCell(63).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[63].ToString()));
                    //ws.GetRow(j + 1).CreateCell(64).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[64].ToString()));
                    //ws.GetRow(j + 1).CreateCell(65).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[65].ToString()));
                    //ws.GetRow(j + 1).CreateCell(66).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[66].ToString()));
                    //ws.GetRow(j + 1).CreateCell(67).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[67].ToString()));
                    //ws.GetRow(j + 1).CreateCell(68).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[68].ToString()));
                   // ws.GetRow(j + 1).CreateCell(69).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArra[2].ToString()));
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
