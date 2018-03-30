﻿namespace AutoScoreFiller
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.cmbUrl = new System.Windows.Forms.ComboBox();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.btnClearTable = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.lvTableShow = new AutoScoreFiller.ListViewNF();
            this.colId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colDemension = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.MenuItemTool = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemFindTable = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemImportFromCsv = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemExportToCsv = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemClearEditorSelector = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.MenuItemAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // webBrowser
            // 
            this.webBrowser.AllowWebBrowserDrop = false;
            this.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser.IsWebBrowserContextMenuEnabled = false;
            this.webBrowser.Location = new System.Drawing.Point(3, 38);
            this.webBrowser.Margin = new System.Windows.Forms.Padding(3, 0, 3, 4);
            this.webBrowser.MinimumSize = new System.Drawing.Size(23, 28);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.ScriptErrorsSuppressed = true;
            this.webBrowser.Size = new System.Drawing.Size(1060, 626);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_DocumentCompleted);
            this.webBrowser.NewWindow += new System.ComponentModel.CancelEventHandler(this.webBrowser_NewWindow);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 206F));
            this.tableLayoutPanel1.Controls.Add(this.webBrowser, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.lvTableShow, 2, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 25);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1272, 668);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Controls.Add(this.cmbUrl, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 3);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(1066, 32);
            this.tableLayoutPanel2.TabIndex = 3;
            // 
            // cmbUrl
            // 
            this.cmbUrl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbUrl.FormattingEnabled = true;
            this.cmbUrl.Location = new System.Drawing.Point(3, 3);
            this.cmbUrl.Name = "cmbUrl";
            this.cmbUrl.Size = new System.Drawing.Size(1060, 25);
            this.cmbUrl.TabIndex = 4;
            this.cmbUrl.SelectedIndexChanged += new System.EventHandler(this.cmbUrl_SelectedIndexChanged);
            this.cmbUrl.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbUrl_KeyPress);
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 3;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel3.Controls.Add(this.btnClearTable, 2, 0);
            this.tableLayoutPanel3.Controls.Add(this.btnImport, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.btnExport, 1, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(1066, 3);
            this.tableLayoutPanel3.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(206, 32);
            this.tableLayoutPanel3.TabIndex = 4;
            // 
            // btnClearTable
            // 
            this.btnClearTable.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnClearTable.Location = new System.Drawing.Point(139, 3);
            this.btnClearTable.Name = "btnClearTable";
            this.btnClearTable.Size = new System.Drawing.Size(64, 26);
            this.btnClearTable.TabIndex = 3;
            this.btnClearTable.Text = "清空";
            this.btnClearTable.UseVisualStyleBackColor = true;
            this.btnClearTable.Click += new System.EventHandler(this.btnClearTable_Click);
            // 
            // btnImport
            // 
            this.btnImport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnImport.Location = new System.Drawing.Point(3, 3);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(62, 26);
            this.btnImport.TabIndex = 2;
            this.btnImport.Text = "导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnExport
            // 
            this.btnExport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnExport.Location = new System.Drawing.Point(71, 3);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(62, 26);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // lvTableShow
            // 
            this.lvTableShow.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lvTableShow.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colId,
            this.colDemension});
            this.lvTableShow.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvTableShow.Location = new System.Drawing.Point(1069, 38);
            this.lvTableShow.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lvTableShow.Name = "lvTableShow";
            this.lvTableShow.Size = new System.Drawing.Size(200, 627);
            this.lvTableShow.TabIndex = 5;
            this.lvTableShow.UseCompatibleStateImageBehavior = false;
            this.lvTableShow.View = System.Windows.Forms.View.Details;
            this.lvTableShow.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lvTableShow_ColumnClick);
            this.lvTableShow.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.lvTableShow_ItemSelectionChanged);
            this.lvTableShow.MouseMove += new System.Windows.Forms.MouseEventHandler(this.lvTableShow_MouseMove);
            // 
            // colId
            // 
            this.colId.Text = "表格ID";
            this.colId.Width = 130;
            // 
            // colDemension
            // 
            this.colDemension.Text = "表格维度";
            this.colDemension.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.colDemension.Width = 70;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuItemTool});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(0, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1272, 25);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // MenuItemTool
            // 
            this.MenuItemTool.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuItemFindTable,
            this.MenuItemImportFromCsv,
            this.MenuItemExportToCsv,
            this.MenuItemClearEditorSelector,
            this.toolStripSeparator1,
            this.MenuItemAbout});
            this.MenuItemTool.Font = new System.Drawing.Font("Microsoft YaHei", 9F);
            this.MenuItemTool.Name = "MenuItemTool";
            this.MenuItemTool.Size = new System.Drawing.Size(44, 21);
            this.MenuItemTool.Text = "工具";
            this.MenuItemTool.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // MenuItemFindTable
            // 
            this.MenuItemFindTable.Name = "MenuItemFindTable";
            this.MenuItemFindTable.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.F)));
            this.MenuItemFindTable.Size = new System.Drawing.Size(263, 22);
            this.MenuItemFindTable.Text = "查找表格";
            this.MenuItemFindTable.Click += new System.EventHandler(this.MenuItemFindTable_Click);
            // 
            // MenuItemImportFromCsv
            // 
            this.MenuItemImportFromCsv.Name = "MenuItemImportFromCsv";
            this.MenuItemImportFromCsv.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.I)));
            this.MenuItemImportFromCsv.Size = new System.Drawing.Size(263, 22);
            this.MenuItemImportFromCsv.Text = "从CSV文件导入数据到网页";
            this.MenuItemImportFromCsv.Click += new System.EventHandler(this.MenuItemImportFromCsv_Click);
            // 
            // MenuItemExportToCsv
            // 
            this.MenuItemExportToCsv.Name = "MenuItemExportToCsv";
            this.MenuItemExportToCsv.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.MenuItemExportToCsv.Size = new System.Drawing.Size(263, 22);
            this.MenuItemExportToCsv.Text = "将网页数据导出到CSV文件";
            this.MenuItemExportToCsv.Click += new System.EventHandler(this.MenuItemExportToCsv_Click);
            // 
            // MenuItemClearEditorSelector
            // 
            this.MenuItemClearEditorSelector.Name = "MenuItemClearEditorSelector";
            this.MenuItemClearEditorSelector.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D)));
            this.MenuItemClearEditorSelector.Size = new System.Drawing.Size(263, 22);
            this.MenuItemClearEditorSelector.Text = "清空编辑框和选择框";
            this.MenuItemClearEditorSelector.Click += new System.EventHandler(this.MenuItemClearEditorSelector_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(260, 6);
            // 
            // MenuItemAbout
            // 
            this.MenuItemAbout.Name = "MenuItemAbout";
            this.MenuItemAbout.Size = new System.Drawing.Size(263, 22);
            this.MenuItemAbout.Text = "关于";
            this.MenuItemAbout.Click += new System.EventHandler(this.MenuItemAbout_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1272, 693);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "MainForm";
            this.Text = "网页表格数据导入导出助手";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnClearTable;
        private System.Windows.Forms.ComboBox cmbUrl;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ColumnHeader colId;
        private System.Windows.Forms.ColumnHeader colDemension;
        private ListViewNF lvTableShow;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem MenuItemTool;
        private System.Windows.Forms.ToolStripMenuItem MenuItemImportFromCsv;
        private System.Windows.Forms.ToolStripMenuItem MenuItemExportToCsv;
        private System.Windows.Forms.ToolStripMenuItem MenuItemClearEditorSelector;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem MenuItemAbout;
        private System.Windows.Forms.ToolStripMenuItem MenuItemFindTable;
    }
}

