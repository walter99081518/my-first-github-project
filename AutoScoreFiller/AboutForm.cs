using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoScoreFiller {
    public partial class AboutForm : Form {
        private MainForm parent = null;
        public AboutForm(MainForm owner) {
            InitializeComponent();
            this.parent = owner;
            this.Text = ResourceCulture.GetString("ABOUT_DIALOG_TITLE");
            this.txbAboutApp.Text = ResourceCulture.GetString("ABOUT_DIALOG_TXT");
        }

        public void ShowAboutDialog() {
            this.ShowDialog(parent);
        }
    }
}
