using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoScoreFiller
{
    public partial class MessageForm : Form
    {
        private MainForm parent = null;
        public MessageForm(MainForm owner)
        {
            InitializeComponent();
            parent = owner;
        }

        public void ShowMessageDialog(string msg, string title = "")
        {
            this.Text = title;
            this.lblMsg.Text = msg;
            this.ShowDialog(parent);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
