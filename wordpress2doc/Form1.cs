using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Xml;
using Pechkin;
using MetroFramework.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Text;
using MetroFramework.Components;
using System.Reflection;
using System.Net;

namespace Wordpress2Doc
{
    public partial class Form1 : MetroForm
    {
        private PrivateFontCollection fontColl;
        private Font handwrittenFont;
        private Localization loc;
        private int version = 4;

        public Form1()
        {
            InitializeComponent();


            if (this.StyleManager == null)
                this.StyleManager = new MetroStyleManager();
            this.StyleManager.Owner = this;

            this.ShadowType = MetroFormShadowType.Flat;
            this.BorderStyle = MetroFramework.Drawing.MetroBorderStyle.None;

            LoadLanguage();
            LoadLanguageSettings();
            UpdateLanguage();

            LoadHandwrittenFont();
            labelStartDescription.UseCompatibleTextRendering = true;
            labelStartDescription.Font = handwrittenFont;
            richTextBoxSettingsCredits.Font = this.StyleManager.GetThemeFont(MetroFramework.Drawing.MetroFontSize.Medium, MetroFramework.Drawing.MetroFontWeight.Regular);          

            LoadSettingsStyleList();
            RenderLoadFilePlease();
            metroTabControlContainer.TabPages.Remove(metroTabPageSettings);

            metroToolTipChoose.SetToolTip(metroButtonChooseTip, loc.C_ttPreviewTip);
            
            LoadSettingsTab();            
        }

        private void LoadSettingsTab()
        {
            if (SettingsHelper.GetAppSetting("style") == null)
            {
                SettingsHelper.SetAppSetting("style", "blue");
            }
            metroComboBoxSettingsStyle.Text = SettingsHelper.GetAppSetting("style");


            if (SettingsHelper.GetAppSetting("formatPdf") == null)
            {
                SettingsHelper.SetAppSetting("formatPdf", "FALSE");
                SettingsHelper.SetAppSetting("formatDocx", "TRUE");
            }
            metroToggleFormatDocx.Checked = Convert.ToBoolean(SettingsHelper.GetAppSetting("formatDocx"));
            metroToggleFormatPdf.Checked = Convert.ToBoolean(SettingsHelper.GetAppSetting("formatPdf"));

            if (SettingsHelper.GetAppSetting("AIO") == null)
            {
                SettingsHelper.SetAppSetting("AIO", "FALSE");
            }
            metroToggleConvertAIO.Checked = Convert.ToBoolean(SettingsHelper.GetAppSetting("AIO"));

            if (SettingsHelper.GetAppSetting("proxy-use") == null)
            {
                SettingsHelper.SetAppSetting("proxy-use", "FALSE");
                SettingsHelper.SetAppSetting("proxy-server", "proxy.example.com");
                SettingsHelper.SetAppSetting("proxy-port", "8080");
            }
            metroToggleSettingsProxy.Checked = Convert.ToBoolean(SettingsHelper.GetAppSetting("proxy-use"));
            metroTextBoxSettingsProxyPort.Text = SettingsHelper.GetAppSetting("proxy-port");
            metroTextBoxSettingsProxy.Text = SettingsHelper.GetAppSetting("proxy-server");            
        }

        private void UpdateLanguage()
        {
            /*
            bool rtl = (loc.GetInfo.LanguageShort == "fa");
            this.RightToLeft = (rtl ? System.Windows.Forms.RightToLeft.Yes : System.Windows.Forms.RightToLeft.No);
            this.RightToLeftLayout = rtl;
             */

            metroTabPageLoad.Text = loc.C_tpLoad;
            labelStartDescription.Text = loc.C_lblStartDescription;
            metroTileHelp.Text = loc.C_mtHelp;
            metroTileSettings.Text = loc.C_mtSettings;
            metroTileLoadExportXML.Text = loc.C_mtLoadEportFile;
            metroTabPageChoose.Text = loc.C_tpChoose;
            metroButtonChooseTip.Text = loc.C_btnChooseTip;
            metroButtonChooseDeselectAll.Text = loc.C_btnDeselectAll;
            metroButtonChooseSelectAll.Text = loc.C_btnSelectAll;
            metroTabPagePreviewHtml.Text = loc.C_tpPreviewWeb;
            metroTabPagePreviewText.Text = loc.C_tpPreviewText;
            ColumnExport.HeaderText = loc.C_clmnExport;
            ColumnTitle.HeaderText = loc.C_clmnTitle;
            metroTabPageExport.Text = loc.C_tpExport;
            metroLabelConvertFormat.Text = loc.C_lblConvertFormat;
            metroLabelConvertAIO.Text = loc.C_lblConvertAIO;
            metroTileConvert.Text = loc.C_mtConvertStart;
            metroTabPageSettings.Text = loc.C_tpSettings;
            metroLabelSettingsStyle.Text = loc.C_lblSettingsStyle;
            metroLabelSettingsLanguage.Text = loc.C_lblSettingsLanguage;
            metroButtonSettingsClose.Text = loc.C_btnSettingsClose;
            metroLabelSettingsCredits.Text = loc.C_lblSettingsCredits + " (Version: " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + ")";
            metroLabelSettingsProxyPort.Text = loc.C_lblProxyPort;
            metroLabelSettingsProxyServer.Text = loc.C_lblProxyServer;
            metroLabelSettingsUseProxy.Text = loc.C_lblUseProxy;
            metroLabelSettingsAuthor.Text = loc.C_lblSettingsTranslated;
            metroLinkSettingsAuthor.Text = loc.GetInfo.Author;
            metroLinkSettingsAuthor.Tag = loc.GetInfo.AuthorLink;
            metroToolTipChoose.SetToolTip(metroButtonChooseTip, loc.C_ttPreviewTip);
            RenderLoadFilePlease();
        }

        private void LoadLanguageSettings()
        {            
            if (SettingsHelper.GetAppSetting("lang") == null)
                SettingsHelper.SetAppSetting("lang", "en");
            metroComboBoxSettingsLanguage.Items.Clear();
            var availableLangs = loc.GetAvailableLanguages();
            metroComboBoxSettingsLanguage.Items.AddRange(availableLangs.Select(x => x.Value).ToArray());
            metroComboBoxSettingsLanguage.Text = availableLangs[SettingsHelper.GetAppSetting("lang")];
        }

        private void LoadLanguage()
        {
            if (SettingsHelper.GetAppSetting("lang") == null)
                SettingsHelper.SetAppSetting("lang", "en");
            loc = new Localization(Application.StartupPath + "\\local", SettingsHelper.GetAppSetting("lang"));            
        }

        private void LoadSettingsStyleList()
        {
            metroComboBoxSettingsStyle.Items.AddRange(MetroStyleManager.Styles.Styles.Keys.Where(x => x != "White").ToArray());           
        }
     
        private void LoadHandwrittenFont()
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Wordpress2Doc.FGVirgil.ttf"))
            {
                handwrittenFont = new Font(LoadFontFamily(s, out fontColl), 16f);                
            }
        }

        public static FontFamily LoadFontFamily(Stream stream, out PrivateFontCollection fontColl)
        {
            var buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);
            var handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            try
            {
                var ptr = Marshal.UnsafeAddrOfPinnedArrayElement(buffer, 0);
                fontColl = new PrivateFontCollection();
                fontColl.AddMemoryFont(ptr, buffer.Length);
                return fontColl.Families[0];
            }
            finally
            {
                handle.Free();
            }
        }
    
        XDocument xDoc;
        XNamespace nsContent = "http://purl.org/rss/1.0/modules/content/";

        private void RenderArticlePreview()
        {
            if (dataGridViewArticles.SelectedRows.Count > 0)
            {
                var url = dataGridViewArticles.SelectedRows[0].Tag.ToString();
                var article = xDoc.Descendants("item").Where(x => x.Descendants("link").First().Value == url);
                var title = article.Descendants("title").First().Value;
                var contentNode = article.Descendants(nsContent + "encoded").First();
                var content = "<html><head></head><body>" + contentNode.Value.Replace("\n", "<br />") + "</body></html>";
                richTextBoxPreview.Text = content;
                webBrowserHtml.DocumentText = content;
                metroLabelPreviewTitleHtml.Text = title;
                metroLabelPreviewTitleText.Text = title;
            }
        }

        private void RenderArticleCount()
        {
            var count = dataGridViewArticles.Rows.Cast<DataGridViewRow>().Where(x => (bool)x.Cells[0].Value).Count();
            metroLabelCount.Text = string.Concat(xDoc.Root.Descendants("item").Count(), 
                                                 loc.C_lblConvertStatus, 
                                                 count, 
                                                 loc.C_lblConvertStatus2);
            metroTileConvert.TileCount = count;
        }

        private void RenderLoadFilePlease()
        {
            Control[] tabPages = new Control[] { metroTabPageExport, metroTabPageChoose };
            foreach (var tabPage in tabPages)
            {
                Panel p = new Panel();
                p.Dock = DockStyle.Fill;
                p.Name = "blockpanel";
                Label lbl = new Label();
                lbl.Text = loc.C_hwLoadExportFile;
                lbl.UseCompatibleTextRendering = true;
                lbl.Font = handwrittenFont;
                lbl.Location = new Point(100, 140);
                lbl.AutoSize = true;
                p.Controls.Add(lbl);
                PictureBox pb = new PictureBox();
                pb.BackColor = Color.Transparent;
                pb.Image = Wordpress2Doc.Properties.Resources.arrow_left_up;
                pb.Size = new Size(164, 117);
                pb.Location = new Point(20, 10);
                p.Controls.Add(pb);
                tabPage.Controls.Add(p);
                p.BringToFront();
            }
        }

        private void HideLoadFilePlease()
        {
            var tabPages = new Control[] { metroTabPageExport, metroTabPageChoose };
            foreach (var tabPage in tabPages)
            {
                while (tabPage.Controls.Cast<Control>().Where(x => x.Name == "blockpanel").Count() > 0)
                {
                    tabPage.Controls.RemoveByKey("blockpanel");
                }
            }            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridViewArticles.DoubleClick += delegate(object s, EventArgs ev)
            {
                RenderArticlePreview();
            };

            metroTabControlContainer.SelectedTab = metroTabPageLoad;
            metroTabControlPreview.SelectedTab = metroTabPagePreviewHtml;

            WebClient wc = new WebClient();
            wc.DownloadStringCompleted += delegate(object s, DownloadStringCompletedEventArgs ev)
            {
                try
                {
                    if (Convert.ToInt32(ev.Result) > version)
                        ShowMessage("There is an update available! Please visit http://code-bude.net/update/wordpress2doc");
                }
                catch { }
            };
            try
            {
                wc.DownloadStringAsync(new Uri("http://code-bude.net/downloads/wordpress2doc/update.txt"));
            }
            catch { }

            metroTabControlContainer.TabPages.Clear();
            metroTabControlContainer.TabPages.AddRange(new TabPage[] { metroTabPageLoad, metroTabPageChoose, metroTabPageExport });
            webBrowserHtml.Size = new Size(428, 288);
            richTextBoxPreview.Size = new Size(428, 288);
            dataGridViewArticles.Size = new Size(306, 294);
        }

        private static string SanitizeXMLString(string inp)
        {
            return Regex.Replace(inp, @"[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u10000-\u10FFFF]", "987654321");
        }

        private void LoadExportXml(string file)
        {
            try
            {
                string xmlRaw = File.ReadAllText(file);
                xDoc = XDocument.Parse(SanitizeXMLString(xmlRaw));
                dataGridViewArticles.Rows.Clear();                
                metroLabelCount.Text = loc.C_lblArticleCountZero;
                XNamespace nsWp = "http://wordpress.org/export/1.2/";
                var articleList = xDoc.Root.Descendants("item").Where(x => x.Descendants(nsWp + "post_type").Count() == 0 || x.Descendants(nsWp + "post_type").First().Value == "post" || x.Descendants(nsWp + "post_type").First().Value == "page").ToList();
                if (articleList.Count == 0)
                {
                    richTextBoxPreview.Clear();
                    webBrowserHtml.DocumentText = string.Empty;
                    ShowMessage(loc.C_dlgNoArticles);
                    return;
                }
                articleList.ForEach(item =>
                {
                    DataGridViewRow dgrv = new DataGridViewRow();
                    dgrv.Tag = item.Descendants("link").First().Value;
                    dgrv.CreateCells(dataGridViewArticles, true, item.Descendants("title").First().Value);
                    dataGridViewArticles.Rows.Add(dgrv);
                });

                if (dataGridViewArticles.Rows.Count > 0)
                {
                    dataGridViewArticles.Rows[0].Selected = true;
                    RenderArticlePreview();
                    dataGridViewArticles.Rows[0].Selected = false;
                }

                RenderArticleCount();
                HideLoadFilePlease();
                metroTabControlContainer.SelectedTab = metroTabPageChoose;
            }
            catch (Exception e)
            {
                ShowMessage(e.Message);
            }
        }

        public static string GetRandomTempPath()
        {
            var tempDir = System.IO.Path.GetTempPath();

            string fullPath;
            do
            {
                var randomName = System.IO.Path.GetRandomFileName();
                fullPath = System.IO.Path.Combine(tempDir, randomName);
            }
            while (Directory.Exists(fullPath));

            return fullPath;
        }

        private static void SaveHtmlAsDocx(string fName, string htmlBody)
        {
            FileInfo fi = new FileInfo(fName);
            var htmlByteContent = Encoding.UTF8.GetBytes(String.Concat("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\"><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"><title></title></head><body>", htmlBody, "</body></html>"));

            var namespaceDocx = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var namespaceRs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var namespaceRsOfficeDoc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            var namespaceRsOfficeAltCh = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";

            var uriDoc = new Uri("/word/document.xml", UriKind.Relative);
            var uriBase = new Uri("/word/document.xml", UriKind.Relative);
            var uri = new Uri("/word/" + Regex.Replace(fi.Name.Substring(0, fi.Name.LastIndexOf('.')), @"[^0-9a-zA-Z]", "") + ".html", UriKind.Relative);
            var uriRelative = PackUriHelper.GetRelativeUri(uriBase, uri);
            
            Package packageDocx = Package.Open(fi.FullName, FileMode.Create, FileAccess.ReadWrite);

            XmlDocument documentRoot = new XmlDocument();
            XmlElement documentTag = documentRoot.CreateElement("w:document", namespaceDocx);
            documentRoot.AppendChild(documentTag);
            XmlElement documentBody = documentRoot.CreateElement("w:body", namespaceDocx);
            documentTag.AppendChild(documentBody);
           
            XmlElement chunkAlt = documentRoot.CreateElement("w:altChunk", namespaceDocx);
            XmlAttribute relationId = chunkAlt.Attributes.Append(documentRoot.CreateAttribute("r:id", namespaceRs));
            relationId.Value = "rId2";
            documentBody.AppendChild(chunkAlt);
                                    
            PackagePart packageDocxXml = packageDocx.CreatePart(uriDoc, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            StreamWriter streamStartPart = new StreamWriter(packageDocxXml.GetStream(FileMode.Create, FileAccess.Write));
            documentRoot.Save(streamStartPart);
            streamStartPart.Close();
            packageDocx.Flush();

            packageDocx.CreateRelationship(uriDoc, TargetMode.Internal, namespaceRsOfficeDoc, "rId1");
            packageDocx.Flush();
            
            PackagePart packageDocxXmlDocument = packageDocx.GetPart(uriBase);            
            PackagePart packageDocxXmlAltchunk = packageDocx.CreatePart(uri, "text/html");
            using (Stream targetStream = packageDocxXmlAltchunk.GetStream())
            {
                targetStream.Write(htmlByteContent, 0, htmlByteContent.Length);
            }
            packageDocxXmlDocument.CreateRelationship(uriRelative, TargetMode.Internal, namespaceRsOfficeAltCh, "rId2");
            packageDocx.Close();
        }

        private void SaveHtmlAsPdf(string fName, string html)
        {
            GlobalConfig cfg = new GlobalConfig();
            SimplePechkin sp = new SimplePechkin(cfg);
            ObjectConfig oc = new ObjectConfig();
            oc.SetCreateExternalLinks(true)
                .SetFallbackEncoding(Encoding.ASCII)
                .SetLoadImages(true)
                .SetAllowLocalContent(true)
                .SetPrintBackground(true);
            
            byte[] pdfBuf = sp.Convert(oc, "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"><title></title></head><body>" + html + "</body></html>");
            FileStream fs = new FileStream(fName, FileMode.Create, FileAccess.ReadWrite);
            foreach (var byteSymbol in pdfBuf)
                fs.WriteByte(byteSymbol);
            fs.Close();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            
       
            
        }

        public string CleanFileName(string strSource)
        {
            var regexSearch = new string(Path.GetInvalidFileNameChars());
            var regEx = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            strSource = regEx.Replace(strSource, "");
            //return Regex.Replace(strSource, @"[^0-9a-zA-Z \-\.#]", "");
            return strSource;
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(Application.StartupPath + "\\manual.pdf");
            }
            catch (Exception ee)
            {
                ShowMessage(string.Concat(loc.C_dlgErrorHelp, ee.Message));
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            metroTabControlContainer.TabPages.Clear();
            metroTabControlContainer.TabPages.Add(metroTabPageSettings);
        }

        private void dataGridViewArticles_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dataGridViewArticles.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = !(bool)dataGridViewArticles.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                RenderArticleCount();
            }
        }

        private void metroTileLoadExportXML_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;
            ofd.Filter = string.Concat(loc.V_ofdFilter," (*.xml)|*.xml|",loc.V_ofdFilter2," (*.*)|*.*");
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (File.Exists(ofd.FileName))
                {
                    StreamReader sr = new StreamReader(ofd.FileName);
                    var content = sr.ReadToEnd();
                    sr.Close();
                    if (content.IndexOf("This is a WordPress eXtended RSS file generated by WordPress as an export of your site") != -1)
                        LoadExportXml(ofd.FileName);
                    else
                        ShowMessage(loc.C_dlgInvalidExportFile);
                }
                
            }
        }

        private void ShowMessage(string message)
        {
            using (MetroMessageDialog mmd = new MetroMessageDialog(message, this.StyleManager))
            {
                mmd.ShowDialog();
            }
        }

        private void metroComboBoxSettingsStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            var style = (sender as MetroFramework.Controls.MetroComboBox).SelectedItem.ToString();
            this.StyleManager.Style = style;
            SettingsHelper.SetAppSetting("style", style);
        }

        private void metroButtonSettingsClose_Click(object sender, EventArgs e)
        {
            metroTabControlContainer.TabPages.Clear();
            metroTabControlContainer.TabPages.AddRange(new TabPage[]{metroTabPageLoad, metroTabPageChoose, metroTabPageExport });
        }

        private void metroButtonChooseSelectAll_Click(object sender, EventArgs e)
        {
            SetAllCheckboxes(dataGridViewArticles, 0, true);
        }

        private void metroButtonChooseDeselectAll_Click(object sender, EventArgs e)
        {
            SetAllCheckboxes(dataGridViewArticles, 0, false);            
        }

        private void SetAllCheckboxes(DataGridView dgv, int checkboxColumnIndex, bool isChecked)
        {
            dgv.Enabled = false;
            dgv.SuspendLayout();
            dgv.Rows.Cast<DataGridViewRow>().ToList().ForEach(row =>
            {
                row.Cells[checkboxColumnIndex].Value = isChecked;
            });
            dgv.ResumeLayout();
            RenderArticleCount();
            dgv.Enabled = true;
        }

        private void metroTileConvert_Click(object sender, EventArgs e)
        {
            if (metroTileConvert.TileCount == 0)
            {
                ShowMessage(loc.C_dlgNoArticleChosen);
                return;
            }
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowNewFolderButton = true;
            fbd.Description = loc.C_fbdDescription;
            if (SettingsHelper.GetAppSetting("lastSavePath") != null && Directory.Exists(SettingsHelper.GetAppSetting("lastSavePath")))
            {
                fbd.SelectedPath = SettingsHelper.GetAppSetting("lastSavePath");
            }
            if (fbd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            metroToggleFormatDocx.Enabled = false;
            metroToggleFormatPdf.Enabled = false;
            metroToggleConvertAIO.Enabled = false;
            metroTileConvert.Enabled = false;
           
            metroTabControlContainer.TabPages.Remove(metroTabPageLoad);
            metroTabControlContainer.TabPages.Remove(metroTabPageChoose);

            metroProgressSpinner1.Visible = true;
            progressBarStatus.Visible = true;
            metroLabelConvertStatus.Text = "";
            metroLabelConvertStatus.Visible = true;

            if (!Directory.Exists(fbd.SelectedPath))
                Directory.CreateDirectory(fbd.SelectedPath);
            SettingsHelper.SetAppSetting("lastSavePath", fbd.SelectedPath);

            var choosenArticles = dataGridViewArticles.Rows.Cast<DataGridViewRow>().Where(x => (bool)x.Cells[0].Value).Select(x => x.Tag.ToString()).ToList();
            var articles = xDoc.Root.Descendants("item").Where(x => choosenArticles.Contains(x.Descendants("link").First().Value)).ToList();

            progressBarStatus.Maximum = articles.Count;
            progressBarStatus.Value = 0;

            BackgroundWorker bgw = new BackgroundWorker();
            bgw.DoWork += bgw_DoWork;
            bgw.ProgressChanged += bgw_ProgressChanged;
            bgw.RunWorkerCompleted += bgw_RunWorkerCompleted;
            bgw.WorkerReportsProgress = true;
            bgw.RunWorkerAsync(new object[]{articles, fbd.SelectedPath, metroToggleFormatDocx.Checked, metroToggleFormatPdf.Checked, metroToggleConvertAIO.Checked });

            

            
        }

        void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            metroToggleFormatDocx.Enabled = true;
            metroToggleFormatPdf.Enabled = true;
            metroTileConvert.Enabled = true;
            metroToggleConvertAIO.Enabled = true;
            metroProgressSpinner1.Visible = false;
            metroLabelConvertStatus.Text = loc.C_lblConvertStatusReady;
            metroTabControlContainer.TabPages.Clear();
            metroTabControlContainer.TabPages.AddRange(new TabPage[] { metroTabPageLoad, metroTabPageChoose, metroTabPageExport });
            metroTabControlContainer.SelectedTab = metroTabPageExport;
        }

        void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarStatus.Value++;
            metroLabelConvertStatus.Text = string.Concat(loc.C_lblConvertStatusProgress, " ", progressBarStatus.Value, " ", loc.C_lblConvertStatusProgress2, " ", progressBarStatus.Maximum);
        }

        void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string masterBody = string.Empty;

                (((object[])e.Argument)[0] as List<XElement>).ForEach(item =>
                {
                    var contentBody = item.Descendants(nsContent + "encoded").First().Value.Replace("\n", "<br />");

                    //Clean BB-Code images
                    Regex re = new Regex(@"\[imag.*?src=""(.*?)"".*?\]\[\/image\]", RegexOptions.Multiline | RegexOptions.Singleline);
                    MatchCollection mc = re.Matches(contentBody);
                    foreach (Match m in mc)
                    {
                        try
                        {
                            contentBody = contentBody.Replace(m.Value, "<img src=\"" + m.Groups[1].Value + "\">");
                        }
                        catch { }
                    }

                    if (!(bool)((object[])e.Argument)[4]) // != AllInOne
                    {
                        var fNameBase = string.Concat((string)((object[])e.Argument)[1], "\\", CleanFileName(item.Descendants("title").First().Value));
                        var fNameDocx = string.Concat(fNameBase, ".docx");
                        var fNamePdf = string.Concat(fNameBase, ".pdf");

                        if ((bool)((object[])e.Argument)[2])
                            SaveHtmlAsDocx(fNameDocx, contentBody);

                        if ((bool)((object[])e.Argument)[3])
                            SaveHtmlAsPdf(fNamePdf, contentBody);
                    }
                    else
                    {
                        masterBody += "<article><h1>" + item.Descendants("title").First().Value + "</h1><br/>" + contentBody + "</article><br/><br/><br/>";
                    }

                    (sender as BackgroundWorker).ReportProgress(0);
                });

                if ((bool)((object[])e.Argument)[4]) // == AllInOne
                {
                    if ((bool)((object[])e.Argument)[2])
                        SaveHtmlAsDocx(((object[])e.Argument)[1] + "\\wordpress2doc-export_"+DateTime.Now.ToFileTimeUtc().ToString() +".docx", masterBody);

                    if ((bool)((object[])e.Argument)[3])
                        SaveHtmlAsPdf(((object[])e.Argument)[1] + "\\wordpress2doc-export_"+DateTime.Now.ToFileTimeUtc().ToString() +".pdf", masterBody);
                }
            }
            catch (Exception ee)
            {
                ShowMessage(string.Concat(loc.C_dlgErrorInExport,"\n\n", ee.Message));
            }
        }

        private void metroToggleFormatDocx_CheckedChanged(object sender, EventArgs e)
        {
            MetroFramework.Controls.MetroToggle cbThis, cbOther;
            if ((sender as Control).Name.Contains("Pdf"))
            {
                cbThis = metroToggleFormatPdf;
                cbOther = metroToggleFormatDocx;
            }
            else
            {
                cbThis = metroToggleFormatDocx;
                cbOther = metroToggleFormatPdf;
            }

            if (!cbThis.Checked && !cbOther.Checked)
                cbOther.Checked = true;
            
            SettingsHelper.SetAppSetting("formatPdf", metroToggleFormatPdf.Checked.ToString());
            SettingsHelper.SetAppSetting("formatDocx", metroToggleFormatDocx.Checked.ToString());
        }

        private void metroComboBoxSettingsLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            var language = (sender as MetroFramework.Controls.MetroComboBox).SelectedItem.ToString();
            var languageShort = loc.GetAvailableLanguages().Where(x => x.Value == language).First().Key;
            loc.ChangeLang(languageShort);
            UpdateLanguage();
            SettingsHelper.SetAppSetting("lang", languageShort);
        }

        private void metroToggleSettingsProxy_CheckedChanged(object sender, EventArgs e)
        {
            var proxyActive = (sender as MetroFramework.Controls.MetroToggle).Checked;
            if (proxyActive)
            {
                metroTextBoxSettingsProxy.Enabled = true;
                metroTextBoxSettingsProxyPort.Enabled = true;
            }
            else
            {
                metroTextBoxSettingsProxy.Enabled = false;
                metroTextBoxSettingsProxyPort.Enabled = false;
            }
            SettingsHelper.SetAppSetting("proxy-use", proxyActive.ToString());
        }

        private void metroTextBoxSettingsProxy_TextChanged(object sender, EventArgs e)
        {
            var con = (sender as Control);
            var settingsKey = con.Name.Contains("Port") ? "proxy-port" : "proxy-server";
            SettingsHelper.SetAppSetting(settingsKey, con.Text);
        }

        private void metroLinkSettingsAuthor_Click(object sender, EventArgs e)
        {
            string link = (sender as Control).Tag.ToString();
            if (link.Length > 0 && (link.StartsWith("http://") || link.StartsWith("https://") || link.StartsWith("www.")))
            {
                try
                {
                    System.Diagnostics.Process.Start(link);
                }
                catch { }
            }
        }

        private void metroToggleConvertAIO_CheckedChanged(object sender, EventArgs e)
        {
            SettingsHelper.SetAppSetting("AIO", metroToggleConvertAIO.Checked.ToString());
        }
        
    }
}
