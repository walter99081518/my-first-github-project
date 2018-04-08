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
            this.parent = owner;
            this.btnCancel.Text = ResourceCulture.GetString("BUTTON_TEXT_CANCEL");
            this.btnOK.Text = ResourceCulture.GetString("BUTTON_TEXT_OK");
        }

        public void ShowMessageDialog(string msg, string title = "", bool cancel = false) {
            if (title == "error") {
                this.Text = ResourceCulture.GetString("MSGBOX_TITLE_ERROR");
            }
            else if (title == "warning") {
                this.Text = ResourceCulture.GetString("MSGBOX_TITLE_WARNING");
            }
            else {
                this.Text = ResourceCulture.GetString("MSGBOX_TITLE_CONFIRM");
            }
            if (cancel) {
                this.btnCancel.Visible = true;
                this.btnCancel.Select();
            }
            else {
                this.btnCancel.Visible = false;
            }
            this.lblMsg.Text = msg;
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
