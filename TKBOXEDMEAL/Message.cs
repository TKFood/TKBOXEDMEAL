using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace TKBOXEDMEAL
{
    public partial class Message : Form
    {
        int messagetime = 3000;


        public Message(String text)
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
            tbxMessage.Text = text;
            btnOK.Focus();
        }

        private void Message_Load(object sender, EventArgs e)
        {
            Thread a = new Thread(new ThreadStart(CloseAfter5Sce));
            a.Start();
        }

        private void CloseAfter5Sce()
	    { 
	        Thread.Sleep(messagetime);
            btnOK.PerformClick();

        }

         private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
       

        
    }
}
