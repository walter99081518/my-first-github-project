using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using System.IO;
using System.Xml;
using Microsoft.Win32;

namespace AutoScoreFiller {
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class MainForm : Form {
        private string initdir = "C:\\";    // 打开数据源时的初始目录
        private WriteLogThread thr = null;  // 程序运行状态记录线程
        private MessageForm msgbox = null;  // 自定义的MessageBox，用于向用户提示出错信息
        private AboutForm aboutform = null; // 软件说明对话框
        private XmlDocument xmldoc = null;  // 配置文件的Xml模型
        private XmlElement root = null;     // 配置文件的根节点
        private string tablename = null;    // 待查找表格的name属性的值
        private List<string> classnames;    // 配置文件中会指定具有哪些class属性的input元素在导入数据时被忽略
        private int log = 0;

        public MainForm() {
            InitializeComponent();
            try {
                classnames = new List<string>();
                xmldoc = new XmlDocument();
                msgbox = new MessageForm(this);
                aboutform = new AboutForm(this);
                ParseConfigXml();
                if (this.log != 0) {
                    thr = new WriteLogThread();
                    WriteLogThread.EmitMessage += new EmitCustomizedMessage(ReceiveMessage);
                    thr.Start();
                }
            }
            catch (Exception ex) {
                if (thr != null) thr.Append(ex.StackTrace);
                msgbox.ShowMessageDialog("程序启动错误\n" + ex.Message, "警告");
            }
        }

        private void ReceiveMessage(string msg) {
            msgbox.ShowMessageDialog(msg, "错误");
            Application.Exit();
        }

        private bool IsNodeIsExist(string xmlpath, string nodename) {
            using (XmlReader reader = XmlReader.Create(xmlpath)) {
                while (reader.Read()) {
                    if (reader.Name == nodename && reader.NodeType == XmlNodeType.Element) {
                        return true;
                    }
                }
            }
            return false;
        }

        private void ParseConfigXml() {
            this.xmldoc.Load("./config.xml");
            root = xmldoc.DocumentElement;

            XmlNode table = root.SelectSingleNode("table");
            this.tablename = table.InnerText;

            XmlNode skip = root.SelectSingleNode("skip");
            XmlNodeList nodelist = skip.SelectNodes("classname");
            for (int i = 0; i < nodelist.Count; i++) {
                this.classnames.Add(nodelist[i].InnerText.ToLower());
            }

            if (IsNodeIsExist("./config.xml", "favorites") && IsNodeIsExist("./config.xml", "url")) {
                XmlNode favorites = root.SelectSingleNode("favorites");
                XmlNodeList urllist = favorites.SelectNodes("url");
                for (int i = 0; i < urllist.Count; i++) {
                    string innertext = urllist[i].InnerText.Trim();
                    if (innertext != "") {
                        this.cmbUrl.Items.Add(innertext);
                    }
                }
            }

            XmlNode log = root.SelectSingleNode("log");
            this.log = int.Parse(log.InnerText);
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e) {
            try {
                //将所有的链接的目标，指向本窗体
                foreach (HtmlElement archor in this.webBrowser.Document.Links) {
                    archor.SetAttribute("target", "_self");
                }

                //将所有的FORM的提交目标，指向本窗体
                foreach (HtmlElement form in this.webBrowser.Document.Forms) {
                    form.SetAttribute("target", "_self");
                }
            }
            catch (System.Exception ex) {
                this.msgbox.ShowMessageDialog("加载网页时发生错误", "错误");
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void webBrowser_NewWindow(object sender, CancelEventArgs e) {
            e.Cancel = true;
        }

        private void btnClearTable_Click(object sender, EventArgs e) {
            try {

                HtmlElement table = this.webBrowser.Document.GetElementById(this.tablename);
                HtmlElementCollection hec = table.GetElementsByTagName("input");
                if (hec != null && hec.Count > 0) {
                    this.clearVisibleElements(hec, table);
                }

                hec = table.GetElementsByTagName("select");
                if (hec != null && hec.Count > 0) {
                    this.clearVisibleElements(hec, table);
                }
            }
            catch (System.Exception ex) {
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void clearVisibleElements(HtmlElementCollection hec, HtmlElement table) {
            HtmlElement script = this.webBrowser.Document.GetElementById("_#@#_embedded_stript_id_check_if_element_visible");
            if (script == null) {
                script = this.webBrowser.Document.CreateElement("script");
                script.SetAttribute("type", "text/javascript");
                script.SetAttribute("id", "_#@#_embedded_stript_id_check_if_element_visible");
                script.SetAttribute("text", "\n" +
                    @"function __is_visible(id) {
                            var element = document.getElementById(id);
                            var display = element.currentStyle.display + '';
                            if (id.match(/_#@#_temp_id_123456$/) != null) {
                                element.id = id.replace(/_#@#_temp_id_123456$/, '');
                            }
                            else {
                                element.id = null;
                            }
                            if (display.toLowerCase().search('none') != -1) {
                                return false;
                            }
                            return true;
                        }" + "\n"
                );
                this.webBrowser.Document.Body.AppendChild(script);
            }
            object[] args = new object[1];

            foreach (HtmlElement he in hec) {
                try {
                    if (!he.Enabled) {
                        continue;
                    }

                    string tag_name = he.TagName.ToLower();
                    if (tag_name != "input" && tag_name != "select" && tag_name != "textarea") {
                        continue;
                    }
                    if (tag_name == "input" && he.GetAttribute("type").ToLower() != "text") {
                        continue;
                    }
                    if (!he.Parent.Parent.Parent.Parent.Equals(table)) {
                        continue;
                    }

                    string id = he.GetAttribute("id");
                    if (id == null) {
                        id = "#@#_temp_id_123456";
                    }
                    else {
                        id += "_#@#_temp_id_123456";
                    }
                    he.SetAttribute("id", id);
                    args[0] = id;
                    object jsret = this.webBrowser.Document.InvokeScript("__is_visible", args);
                    if (jsret == null || jsret.ToString().ToLower() == "false") {
                        continue;
                    }

                    if (tag_name == "input" || tag_name == "textarea") {
                        he.InnerText = "";
                    }
                    else {
                        HtmlElementCollection options = he.GetElementsByTagName("option");
                        for (int i = 0; i < options.Count; i++) {
                            if (options[i].InnerHtml.Trim() == "") {
                                options[i].SetAttribute("selected", "selected");
                                break;
                            }
                        }
                    }
                }
                catch (System.Exception ex) {
                    if (thr != null) thr.Append(ex.StackTrace);
                }
            }
        }

        private void cmbUrl_KeyPress(object sender, KeyPressEventArgs e) {
            try {
                if (e.KeyChar == (char)Keys.Enter) {
                    this.webBrowser.Navigate(this.cmbUrl.Text);

                    if (IsNodeIsExist("./config.xml", "favorites") && IsNodeIsExist("./config.xml", "url")) {
                        XmlNode favorites = this.root.SelectSingleNode("favorites");
                        XmlNodeList urllist = favorites.SelectNodes("url");
                        int i;
                        string newurl = this.cmbUrl.Text.Trim().Replace("\\", "/");
                        for (i = 0; i < urllist.Count; i++) {
                            string innertext = urllist[i].InnerText.Trim().Replace("\\", "/").ToLower();
                            if (innertext == newurl.ToLower()) {
                                break;
                            }
                        }
                        if (i == urllist.Count) {
                            XmlNode node = this.xmldoc.CreateNode(XmlNodeType.Element, "url", null);
                            node.InnerText = newurl;
                            favorites.AppendChild(node);
                            this.xmldoc.Save("./config.xml");
                        }
                    }
                }
            }
            catch (System.Exception ex) {
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void btnPrevious_Click(object sender, EventArgs e) {
            this.webBrowser.GoBack();
        }

        private void btnNext_Click(object sender, EventArgs e) {
            this.webBrowser.GoForward();
        }

        private void cmbUrl_SelectedIndexChanged(object sender, EventArgs e) {
            try {
                if (this.cmbUrl.Text != null && this.cmbUrl.Text.Trim() != "") {
                    this.webBrowser.Navigate(this.cmbUrl.Text);
                }
            }
            catch (System.Exception ex) {
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void MenuItemAbout_Click(object sender, EventArgs e) {
            this.aboutform.ShowDialog();
        }

        private void MenuItemImportFromExcel_Click(object sender, EventArgs e) {
            try {
                HtmlElement table = this.webBrowser.Document.GetElementById(this.tablename);

                OpenFileDialog ofdlg = new OpenFileDialog();
                ofdlg.InitialDirectory = this.initdir;
                ofdlg.Filter = "csv file|*.csv";
                ofdlg.RestoreDirectory = false;
                ofdlg.FilterIndex = 1;
                ofdlg.Title = "打开数据源";
                if (ofdlg.ShowDialog() != DialogResult.OK) {
                    return;
                }
                string filename = ofdlg.FileName;
                initdir = filename;

                HtmlElement script = this.webBrowser.Document.GetElementById("_#@#_embedded_stript_id_check_if_element_visible");
                if (script == null) {
                    script = this.webBrowser.Document.CreateElement("script");
                    script.SetAttribute("type", "text/javascript");
                    script.SetAttribute("id", "_#@#_embedded_stript_id_check_if_element_visible");
                    script.SetAttribute("text", "\n" +
                        @"function __is_visible(id) {
                            var element = document.getElementById(id);
                            var display = element.currentStyle.display + '';
                            if (id.match(/_#@#_temp_id_123456$/) != null) {
                                element.id = id.replace(/_#@#_temp_id_123456$/, '');
                            }
                            else {
                                element.id = null;
                            }
                            if (display.toLowerCase().search('none') != -1) {
                                return false;
                            }
                            return true;
                        }" + "\n"
                    );
                    this.webBrowser.Document.Body.AppendChild(script);
                }

                object[] args = new object[1];

                HtmlElementCollection trs = table.GetElementsByTagName("tr");
                CsvStreamReader csr = new CsvStreamReader(filename);

                for (int i = 0, m = 0; i < trs.Count && m < csr.RowCount; i++) {
                    if (!trs[i].Parent.Parent.Equals(table)) { // trs[i].Parent.TagName == "TBODY";
                        continue;
                    }
                    HtmlElementCollection tds = trs[i].GetElementsByTagName("td");
                    for (int j = 0, n = 0; j < tds.Count && n < csr.ColCount; j++) {
                        HtmlElementCollection hec = tds[j].Children;
                        if (hec == null) {
                            continue;
                        }
                        for (int k = 0; k < hec.Count; k++) {
                            // Only input, select and textarea can be filled with value.
                            string tagname = hec[k].TagName.ToLower();
                            if (tagname != "input" && tagname != "select" && tagname != "textarea") {
                                continue;
                            }
                            // If the element is a input, then the type of it must be text.
                            if (tagname == "input" && hec[k].GetAttribute("type").ToLower() != "text") {
                                continue;
                            }
                            // Check if the object is visible.
                            string id = hec[k].GetAttribute("id");
                            if (id == null) {
                                id = "#@#_temp_id_123456";
                            }
                            else {
                                id += "_#@#_temp_id_123456";
                            }
                            args[0] = id;
                            hec[k].SetAttribute("id", id);
                            object jsret = this.webBrowser.Document.InvokeScript("__is_visible", args);
                            if (jsret == null || jsret.ToString().ToLower() == "false") {
                                continue;
                            }
                            // Elements must be editable.
                            if (!hec[k].Enabled) {
                                ++n;
                                continue;
                            }
                            // Copy text to elements' inner area.
                            if (tagname == "input" || tagname == "textarea") {
                                hec[k].InnerText = csr[m, ++n].ToString();
                            }
                            else // select
                            {
                                foreach (HtmlElement opt in hec[k].GetElementsByTagName("option")) {
                                    string text = csr[m, n + 1].ToString();
                                    if (opt.InnerText == text) {
                                        opt.SetAttribute("selected", "selected");
                                        break;
                                    }
                                }
                                ++n;
                            }
                        }
                    }
                    ++m;
                }
                //csr.WriteTextFile("d:/debug.txt", this.webBrowser.Document.Body.InnerHtml);
            }
            catch (System.Exception ex) {
                this.msgbox.ShowMessageDialog("导入数据失败", "错误");
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void MenuItemExportToExcel_Click(object sender, EventArgs e) {
            try {
                HtmlElement table = this.webBrowser.Document.GetElementById(this.tablename);

                SaveFileDialog dlg = new SaveFileDialog();
                dlg.InitialDirectory = this.initdir;
                dlg.Filter = "csv file|*.csv";
                dlg.RestoreDirectory = false;
                dlg.FilterIndex = 1;
                dlg.Title = "导出网页数据";
                if (dlg.ShowDialog() != DialogResult.OK) {
                    return;
                }

                string filename = initdir = dlg.FileName;
                if (!Regex.IsMatch(dlg.FileName.ToLower(), @".csv$")) {
                    filename = initdir += ".csv";
                }

                HtmlElement script = this.webBrowser.Document.GetElementById("_#@#_embedded_stript_id_check_if_element_visible");
                if (script == null) {
                    script = this.webBrowser.Document.CreateElement("script");
                    script.SetAttribute("type", "text/javascript");
                    script.SetAttribute("id", "_#@#_embedded_stript_id_check_if_element_visible");
                    script.SetAttribute("text", "\n" +
                        @"function __is_visible(id) {
                            var element = document.getElementById(id);
                            var display = element.currentStyle.display + '';
                            if (id.match(/_#@#_temp_id_123456$/) != null) {
                                element.id = id.replace(/_#@#_temp_id_123456$/, '');
                            }
                            else {
                                element.id = null;
                            }
                            if (display.toLowerCase().search('none') != -1) {
                                return false;
                            }
                            return true;
                        }" + "\n"
                    );
                    this.webBrowser.Document.Body.AppendChild(script);
                }

                HtmlElementCollection trs = table.GetElementsByTagName("tr");

                CsvStreamReader csv = new CsvStreamReader();
                csv.OpenCsvFile(filename);

                object[] args = new object[1];
                string[] tagnames = { "a", "b", "p", "font", "h1", "h2", "h3", "h4", "h5", "h6", "input", "select", "span", "textarea" };

                foreach (HtmlElement tr in trs) {
                    if (!tr.Parent.Parent.Equals(table)) {
                        continue;
                    }
                    HtmlElementCollection tds = tr.GetElementsByTagName("td");
                    LinkedList<string> lls = new LinkedList<string>();
                    foreach (HtmlElement td in tds) {
                        if (!td.Parent.Equals(tr)) {
                            continue;
                        }

                        string innerhtml = "";
                        HtmlElementCollection elements = td.Children;
                        if (elements != null && elements.Count > 0) {
                            foreach (HtmlElement he in elements) {
                                string id = he.GetAttribute("id");
                                if (id == null) {
                                    id = "#@#_temp_id_123456";
                                }
                                else {
                                    id += "_#@#_temp_id_123456";
                                }
                                args[0] = id;
                                he.SetAttribute("id", id);

                                object jsret = this.webBrowser.Document.InvokeScript("__is_visible", args);
                                if (jsret == null || jsret.ToString().ToLower() == "false") {
                                    continue;
                                }

                                string tagname = he.TagName.ToLower();
                                if (!tagnames.Contains(tagname)) {
                                    continue;
                                }

                                string text = "";
                                if (tagname == "input" && he.GetAttribute("type").ToLower() == "text") {
                                    text = he.GetAttribute("value");
                                }
                                else if (tagname == "select") {
                                    foreach (HtmlElement opt in he.GetElementsByTagName("option")) {
                                        string is_selected = opt.GetAttribute("selected");
                                        if (is_selected == null) {
                                            continue;
                                        }
                                        if (is_selected.ToLower() == "true") {
                                            text = opt.InnerText;
                                            break;
                                        }
                                    }
                                }
                                else {
                                    text = he.InnerText;
                                }
                                innerhtml = (text == null ? "" : text);
                                lls.AddLast(innerhtml);
                            }
                        }
                        else {
                            innerhtml = td.InnerText == null ? "" : td.InnerText;
                            lls.AddLast(innerhtml);
                        }
                    }
                    csv.WriteCsvFile(lls);
                }
                csv.CloseCsvFile();
            }
            catch (System.Exception ex) {
                this.msgbox.ShowMessageDialog("导出Excel文件失败", "错误");
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }
    }
}
