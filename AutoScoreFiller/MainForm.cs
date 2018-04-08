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
        private string language = "zh";
        private int log = 0;
        private string embededScript =
            @"function __getInnerText(id) {
                try {
                    var nodes = document.getElementById(id).childNodes;
                    if (nodes == null || nodes == undefined) {
                        return '';
                    }
                    var arr = [];
                    var tagnames = ['a', 'b', 'p', 'font', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'input', 'select', 'span', 'textarea'];
                    for (var i = 0; i < nodes.length; i++) {
                        var node = nodes[i];
                        if (node.nodeType == 1) { // Node.ELEMENT_NODE
                            if (__getComputedStyle(node, 'display') == 'none') {
                                continue;
                            }
                            if (tagnames.indexOf(node.nodeName.toLowerCase()) == -1) {
                                continue;
                            }
                            if (node.nodeName.toLowerCase() == 'input' && node.type == 'text') {
                                arr[arr.length] = node.value;
                            }
                            else if (node.nodeName.toLowerCase() == 'textarea') {
                                arr[arr.length] = node.textContent;
                            }
                            else if (node.nodeName.toLowerCase() == 'select') {
                                var options = node.children;
                                for (var i = 0; i < options.length; i++) {
                                    if (options[i].selected) {
                                        arr[arr.length] = options[i].textContent;
                                        break;
                                    }
                                }
                            }
                            else {
                                arr[arr.length] = node.innerHTML;
                            }
                        }
                        else if (node.nodeType == 3) { // Node.TEXT_NODE(3)
                            arr[arr.length] = node.nodeValue;
                        }
                        else {
                            arr[arr.length] = node.innerHTML;
                        }
                    }
                    return arr.join('\t');
                }
                catch (ex) {
                    return '';
                }
            }

            function __getComputedStyle(ele, attr){
                if (window.getComputedStyle) {
                    return window.getComputedStyle(ele, null)[attr];
                }
                return ele.currentStyle[attr];
            }

            function __isVisible(id) {
                var element = document.getElementById(id);
                var display = __getComputedStyle(element, 'display') + '';
                return display.toLowerCase();
            }";

        public MainForm() {
            InitializeComponent();
            try {
                this.xmldoc = new XmlDocument();
                ParseConfigXml();

                if (this.language.Equals("en")) {
                    ResourceCulture.SetCurrentCulture("en-US");
                }
                else {
                    ResourceCulture.SetCurrentCulture("zh-CN");
                }

                this.Text = ResourceCulture.GetString("MAINFORM_TITLE");
                this.MenuItemTool.Text = ResourceCulture.GetString("MENU_TOOL");
                this.MenuItemImportFromCsv.Text = ResourceCulture.GetString("MENU_TOOL_IMPORT_FROM_CSV");
                this.MenuItemExportToCsv.Text = ResourceCulture.GetString("MENU_TOOL_EXPORT_TO_CSV");
                this.MenuItemClearBrowsingRecord.Text = ResourceCulture.GetString("MENU_TOOL_CLEAR_BROWSING_RECORD");
                this.MenuItemAbout.Text = ResourceCulture.GetString("MENU_TOOL_ABOUT");
                this.btnPrevious.Text = ResourceCulture.GetString("BUTTON_TEXT_BACK");
                this.btnNext.Text = ResourceCulture.GetString("BUTTON_TEXT_FRONT");
                this.btnClearTable.Text = ResourceCulture.GetString("BUTTON_TEXT_CLEAR");
                this.msgbox = new MessageForm(this);
                this.aboutform = new AboutForm(this);

                if (this.log != 0) {
                    thr = new WriteLogThread();
                    WriteLogThread.EmitMessage += new EmitCustomizedMessage(ReceiveMessage);
                    thr.Start();
                }
            }
            catch (Exception ex) {
                if (thr != null) thr.Append(ex.StackTrace);
                string txt = ResourceCulture.GetString("MSGBOX_TXT_FAILS_TO_START_APP");
                msgbox.ShowMessageDialog(txt, "warning");
            }
        }

        private void ReceiveMessage(string msg) {
            msgbox.ShowMessageDialog(msg, "error");
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

            XmlNode lang = root.SelectSingleNode("language");
            this.language = lang.InnerText.ToLower();

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
                string txt = ResourceCulture.GetString("MSGBOX_TXT_FAILS_TO_LOAD_WEB_PAGE");
                this.msgbox.ShowMessageDialog(txt, "error");
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
                script.SetAttribute("text", this.embededScript);
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
                    string newId = "#@#_temp_id_123456";
                    if (id != null) {
                        newId = id + "_#@#_temp_id_123456";
                    }
                    he.SetAttribute("id", newId);
                    args[0] = newId;
                    object obj = this.webBrowser.Document.InvokeScript("__isVisible", args);
                    he.SetAttribute("id", id);
                    if (obj == null || obj.ToString() == "none") {
                        continue;
                    }

                    if (tag_name == "input" || tag_name == "textarea") {
                        he.InnerText = "";
                    }
                    else {
                        HtmlElementCollection options = he.GetElementsByTagName("option");
                        if (options != null && options.Count > 0) {
                            options[0].SetAttribute("selected", "selected");
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
                    string newurl = this.cmbUrl.Text.Trim().Replace("\\", "/");

                    if (IsNodeIsExist("./config.xml", "favorites")) {
                        XmlNode favorites = this.root.SelectSingleNode("favorites");
                        if (IsNodeIsExist("./config.xml", "url")) {
                            XmlNodeList urllist = favorites.SelectNodes("url");
                            int i = 0;
                            for (; i < urllist.Count; i++) {
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
                        else {
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
            this.aboutform.ShowAboutDialog();
        }

        private void MenuItemImportFromCsv_Click(object sender, EventArgs e) {
            try {
                HtmlElement table = null;
                try {
                    table = this.webBrowser.Document.GetElementById(this.tablename);
                    if (table == null) {
                        throw new Exception();
                    }
                }
                catch (System.Exception ex) {
                    string fmt = ResourceCulture.GetString("MSGBOX_TXT_TABLE_NOT_FOUND");
                    string txt = string.Format(fmt, this.tablename);
                    this.msgbox.ShowMessageDialog(txt, "error");
                    if (thr != null) thr.Append(ex.StackTrace);
                    return;
                }

                OpenFileDialog ofdlg = new OpenFileDialog();
                ofdlg.InitialDirectory = this.initdir;
                ofdlg.Filter = "csv file|*.csv";
                ofdlg.RestoreDirectory = false;
                ofdlg.FilterIndex = 1;
                ofdlg.Title = ResourceCulture.GetString("MENU_TOOL_IMPORT_FROM_CSV");
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
                    script.SetAttribute("text", this.embededScript);
                    this.webBrowser.Document.Body.AppendChild(script);
                }

                object[] args = new object[1];

                HtmlElementCollection trs = table.GetElementsByTagName("tr");
                CsvStreamReader csr = new CsvStreamReader(filename);

                for (int i = 0, m = 0; i < trs.Count && m < csr.RowCount; i++) {
                    if (!trs[i].Parent.Parent.Equals(table)) { // trs[i].Parent.TagName == "TBODY";
                        continue;
                    }
                    bool validRow = false;
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
                            string newId = "#@#_temp_id_123456";
                            if (id != null) {
                                newId = id + "_#@#_temp_id_123456";
                            }
                            args[0] = newId;
                            hec[k].SetAttribute("id", newId);
                            object obj = this.webBrowser.Document.InvokeScript("__isVisible", args);
                            hec[k].SetAttribute("id", id);

                            if (obj == null || obj.ToString() == "none") {
                                continue;
                            }

                            // Elements must be editable.
                            if (!hec[k].Enabled) {
                                ++n;
                                continue;
                            }
                            // Copy text to elements' inner area.
                            if (tagname == "input" || tagname == "textarea") {
                                hec[k].InnerText = csr[m + 1, n + 1].ToString();
                                validRow = true;
                            }
                            else {// select
                                foreach (HtmlElement opt in hec[k].GetElementsByTagName("option")) {
                                    string text = csr[m + 1, n + 1].ToString();
                                    validRow = true;
                                    if (opt.InnerText == text) {
                                        opt.SetAttribute("selected", "selected");
                                        break;
                                    }
                                }
                            }
                            ++n;
                        }
                    }
                    if (validRow) {
                        ++m;
                    }
                }
            }
            catch (System.Exception ex) {
                string txt = ResourceCulture.GetString("MSGBOX_TXT_IMPORT_FAIL");
                this.msgbox.ShowMessageDialog(txt, "error");
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void MenuItemExportToCsv_Click(object sender, EventArgs e) {
            try {
                HtmlElement table = null;
                try {
                    table = this.webBrowser.Document.GetElementById(this.tablename);
                    if (table == null) {
                        throw new Exception();
                    }
                }
                catch (System.Exception ex) {
                    string fmt = ResourceCulture.GetString("MSGBOX_TXT_TABLE_NOT_FOUND");
                    string txt = string.Format(fmt, this.tablename);
                    this.msgbox.ShowMessageDialog(txt, "error");
                    if (thr != null) thr.Append(ex.StackTrace);
                    return;
                }

                SaveFileDialog dlg = new SaveFileDialog();
                dlg.InitialDirectory = this.initdir;
                dlg.Filter = "csv file|*.csv";
                dlg.RestoreDirectory = false;
                dlg.FilterIndex = 1;
                dlg.Title = ResourceCulture.GetString("MENU_TOOL_EXPORT_TO_CSV");
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
                    script.SetAttribute("text", this.embededScript);
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

                        string id = td.GetAttribute("id");
                        string newId = "#@#_temp_id_123456";
                        if (id != null) {
                            newId = id + "_#@#_temp_id_123456";
                        }
                        td.SetAttribute("id", newId);
                        args[0] = newId;
                        object obj = this.webBrowser.Document.InvokeScript("__getInnerText", args);
                        td.SetAttribute("id", id);

                        string innerText = (obj == null ? "" : obj.ToString());
                        string[] texts = innerText.Split('\t');
                        for (int j = 0; j < texts.Length; j++) {
                            lls.AddLast(texts[j]);
                        }
                    }
                    csv.WriteCsvFile(lls);
                }
                csv.CloseCsvFile();
            }
            catch (System.Exception ex) {
                string txt = ResourceCulture.GetString("MSGBOX_TXT_EXPORT_FAIL");
                this.msgbox.ShowMessageDialog(txt, "error");
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void MenuItemClearBrowsingRecord_Click(object sender, EventArgs e) {
            try {
                string txt = ResourceCulture.GetString("MSGBOX_TXT_CONFIRM_DELETE_BROWSING_RECORD");
                this.msgbox.ShowMessageDialog(txt, "confirm", true);
                if (this.msgbox.Confirm()) {
                    if (IsNodeIsExist("./config.xml", "favorites") && IsNodeIsExist("./config.xml", "url")) {
                        XmlNode favorites = this.root.SelectSingleNode("favorites");
                        favorites.RemoveAll();
                        this.xmldoc.Save("./config.xml");
                        this.cmbUrl.Items.Clear();
                    }
                }
            }
            catch (System.Exception ex) {
                string txt = ResourceCulture.GetString("MSGBOX_TXT_CLEAR_BROWSING_DATA_FAIL");
                this.msgbox.ShowMessageDialog(txt, "error");
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }

        private void cmbUrl_DropDown(object sender, EventArgs e) {
            try {
                if (IsNodeIsExist("./config.xml", "favorites") && IsNodeIsExist("./config.xml", "url")) {
                    XmlNode favorites = root.SelectSingleNode("favorites");
                    XmlNodeList urllist = favorites.SelectNodes("url");
                    this.cmbUrl.Items.Clear();
                    for (int i = 0; i < urllist.Count; i++) {
                        string innertext = urllist[i].InnerText.Trim();
                        if (innertext != "") {
                            this.cmbUrl.Items.Add(innertext);
                        }
                    }
                }
            }
            catch (System.Exception ex) {
                if (thr != null) thr.Append(ex.StackTrace);
            }
        }
    }
}
