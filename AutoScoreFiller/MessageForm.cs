using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoScoreFiller {
    public partial class MessageForm : Form {
        private MainForm parent = null;
        private bool confirm = false;
        public MessageForm(MainForm owner) {
            InitializeComponent();
            parent = owner;
        }

        public void ShowMessageDialog(string msg, string title = "", bool cancel = false) {
            this.Text = title;
            this.lblMsg.Text = msg;
            if (cancel) {
                this.btnCancel.Visible = true;
            }
            else {
                this.btnCancel.Visible = false;
            }
            this.ShowDialog(parent);
        }

        public bool Confirm() {
            return this.confirm;
        }

        private void btnOK_Click(object sender, EventArgs e) {
            this.confirm = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e) {
            this.confirm = false;
            this.Close();
        }
    }
}
