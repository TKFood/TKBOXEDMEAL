using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Reflection;
using System;
using System.Diagnostics;
using System.Drawing;

namespace TKBOXEDMEAL
{
    public partial class FrmParent : Form
    {
        SqlConnection conn;
        MenuStrip MnuStrip;
        ToolStripMenuItem MnuStripItem;
        string UserName;

        public FrmParent()
        {
            InitializeComponent();
        }
        public FrmParent(string txt_UserName)
        {
            InitializeComponent();
            UserName = txt_UserName;
        }

        //private void InitializeComponent()
        //{
        //    throw new NotImplementedException();
        //}

        private void FrmParent_Load(object sender, EventArgs e)
        {

            // To make this Form the Parent Form
            this.IsMdiContainer = true;

            //Creating object of MenuStrip class
            MnuStrip = new MenuStrip();

            //Placing the control to the Form
            this.Controls.Add(MnuStrip);

            String connectionString;
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            conn = new SqlConnection(connectionString);
            String Sequel = "SELECT MAINMNU,MENUPARVAL,STATUS FROM MNU_PARENT";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, conn);
            DataTable dt = new DataTable();
            conn.Open();
            da.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                MnuStripItem = new ToolStripMenuItem(dr["MAINMNU"].ToString());
                SubMenu(MnuStripItem, dr["MENUPARVAL"].ToString());
                MnuStrip.Items.Add(MnuStripItem);
            }
            // The Form.MainMenuStrip property determines the merge target.
            this.MainMenuStrip = MnuStrip;

            timer1.Enabled = true;
            timer1.Interval = 1000;
            timer1.Start();

            textBox1.Select();
        }


        public void SubMenu(ToolStripMenuItem mnu, string submenu)
        {
            StringBuilder Seqchild = new StringBuilder();
            Seqchild.AppendFormat("SELECT FRM_NAME FROM MNU_SUBMENU ,MNU_SUBMENULogin WHERE MNU_SUBMENU.FRM_CODE=MNU_SUBMENULogin.FRM_CODE AND  MNU_SUBMENULogin.UserName='{0}' AND MENUPARVAL='{1}'", UserName.ToString(), submenu.ToString());
            //Seqchild.AppendFormat( "SELECT FRM_NAME FROM MNU_SUBMENU ,MNU_SUBMENULogin WHERE MNU_SUBMENU.FRM_CODE=MNU_SUBMENULogin.FRM_CODE AND  MNU_SUBMENULogin.UserName='1' AND MENUPARVAL='1'");
            SqlDataAdapter dachildmnu = new SqlDataAdapter(Seqchild.ToString(), conn);
            DataTable dtchild = new DataTable();
            dachildmnu.Fill(dtchild);

            foreach (DataRow dr in dtchild.Rows)
            {
                ToolStripMenuItem SSMenu = new ToolStripMenuItem(dr["FRM_NAME"].ToString(), null, new EventHandler(ChildClick));
                mnu.DropDownItems.Add(SSMenu);
            }
        }

        private void ChildClick(object sender, EventArgs e)
        {

        }

        private void FrmParent_FormClosed(object sender, FormClosedEventArgs e)
        {

            //=====偵測執行中的外部程式並關閉=====
            Process[] MyProcess = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);
            if (MyProcess.Length > 0)
                MyProcess[0].Kill(); //關閉執行中的程式

        }

        #region FUNCTION
        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox1.Text.Equals("1"))
                {
                    label4.Text = "1";
                    textBox2.Select();
                }
                else if (textBox1.Text.Equals("2"))
                {
                    label4.Text = "2";
                    textBox2.Select();
                }
                else if (textBox1.Text.Equals("3"))
                {
                    label4.Text = "3";
                    textBox2.Select();
                }
                else
                {
                    textBox1.Text = "";
                }

            }

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter&&!string.IsNullOrEmpty(textBox2.Text.ToString()))
            {
                if (textBox1.Text.Equals("1"))
                {
                    label11.Text = "~用餐愉快~";
                    SETLABEL();
                    label4.Text = "";
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox1.Select();

                }
                else if (textBox1.Text.Equals("2"))
                {
                    textBox3.Select();
                }
                else if (textBox1.Text.Equals("3"))
                {
                    textBox4.Select();
                }
            }

        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox3.Text.Equals("1"))
                {
                    label11.Text = "~訂餐 午餐-葷 完成~";                    
                }
                else if (textBox3.Text.Equals("2"))
                {
                    label11.Text = "~訂餐 晚餐-葷 完成~";
                }
                else if (textBox3.Text.Equals("3"))
                {
                    label11.Text = "~訂餐 午餐-葷/晚餐-葷 完成~";
                }
                else if (textBox3.Text.Equals("4"))
                {
                    label11.Text = "~訂餐 午餐-素 完成~";
                }
                else if (textBox3.Text.Equals("5"))
                {
                    label11.Text = "~訂餐 晚餐-素 完成~";
                }
                else if (textBox3.Text.Equals("6"))
                {
                    label11.Text = "~訂餐 午餐-素/晚餐-素 完成~";
                }
                else 
                {
                    label11.Text = "~取消~";
                }

                SETLABEL();
                label4.Text = "";
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox1.Select();
            }
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (textBox4.Text.Equals("1"))
                    {
                        label11.Text = "~取消 午餐 完成~";
                    }
                    else if (textBox4.Text.Equals("2"))
                    {
                        label11.Text = "~取消 晚餐 完成~";
                    }
                    else if (textBox4.Text.Equals("3"))
                    {
                        label11.Text = "~取消 午餐/晚餐 完成~";
                    }
                    else
                    {
                        label11.Text = "~取消~";
                    }


                    SETLABEL();
                    label4.Text = "";
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox1.Select();
                }
            }
        }

        public void SETLABEL()
        {
            label11.Font = new Font("Arial", 24, FontStyle.Bold);
            label11.ForeColor = Color.Blue;
        }
        #endregion


    }
}