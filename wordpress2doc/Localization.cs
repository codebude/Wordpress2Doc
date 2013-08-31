using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
namespace Wordpress2Doc
{
    class Localization
    {
        private XDocument xdoc;
        private string workingDir;
        private TranslationInfo info;

        public Localization(string langFileFolder, string lang = "en")
        {
            workingDir = langFileFolder;
            ChangeLang(lang);
        }

        public void ChangeLang(string lang)
        {
            try
            {                
                Directory.GetFiles(workingDir).ToList().ForEach(file =>
                {
                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        if (fi.Extension == ".xml")
                        {
                            XDocument xdocT = XDocument.Load(fi.FullName);
                            if (xdocT.Descendants("info").First().Attribute("language").Value == lang)
                            {
                                xdoc = XDocument.Load(fi.FullName);

                                var infoSource = xdoc.Descendants("info").Select(x => new
                                {
                                    Language = x.Attribute("language").Value,
                                    LanguageLong = x.Attribute("language-long").Value,
                                    Application = x.Attribute("application").Value,
                                    LastEdit = x.Attribute("lastedit").Value,
                                    Author = x.Attribute("author") == null ? string.Empty : x.Attribute("author").Value,
                                    AuthorLink = x.Attribute("author-link") == null ? string.Empty : x.Attribute("author-link").Value
                                }).First();

                                info = new TranslationInfo()
                                {
                                    Author = infoSource.Author,
                                    AuthorLink = infoSource.AuthorLink,
                                    Language = infoSource.LanguageLong,
                                    LanguageShort = infoSource.Language,
                                    LastEdit = infoSource.LastEdit
                                };                                   

                                return;
                            }
                        }
                    }
                    catch { }
                });
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Couldn't change language file. Please reinstall the software!\n\nError: \n" + e.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private string GetText(string label)
        {
            return xdoc.Descendants("ressource").Where(x => x.Attribute("label").Value == label).First().Value;                
        }

        public Dictionary<string, string> GetAvailableLanguages()
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            Directory.GetFiles(workingDir).ToList().ForEach(file =>
            {
                try
                {
                    FileInfo fi = new FileInfo(file);
                    if (fi.Extension == ".xml")
                    {
                        XDocument xdocT = XDocument.Load(fi.FullName);
                        if (xdocT.Descendants("info").First().Attribute("application").Value == System.Reflection.Assembly.GetExecutingAssembly().GetName().Name)
                        {
                            dict.Add(xdocT.Descendants("info").First().Attribute("language").Value, xdocT.Descendants("info").First().Attribute("language-long").Value);                            
                        }
                    }
                }
                catch { }               
            });
            return dict;
        }

        public TranslationInfo GetInfo
        {
            get { return info; }
        }

        public string C_ttPreviewTip { get { return GetText("C_ttPreviewTip"); } }
        public string C_lblConvertStatus { get { return GetText("C_lblConvertStatus"); } }
        public string C_lblConvertStatus2 { get { return GetText("C_lblConvertStatus2"); } }
        public string C_hwLoadExportFile { get { return GetText("C_hwLoadExportFile"); } }
        public string C_lblArticleCountZero { get { return GetText("C_lblArticleCountZero"); } }
        public string C_dlgNoArticles { get { return GetText("C_dlgNoArticles"); } }
        public string V_ofdFilter { get { return GetText("V_ofdFilter"); } }
        public string V_ofdFilter2 { get { return GetText("V_ofdFilter2"); } }
        public string C_dlgInvalidExportFile { get { return GetText("C_dlgInvalidExportFile"); } }
        public string C_dlgNoArticleChosen { get { return GetText("C_dlgNoArticleChosen"); } }
        public string C_fbdDescription { get { return GetText("C_fbdDescription"); } }
        public string C_lblConvertStatusReady { get { return GetText("C_lblConvertStatusReady"); } }
        public string C_lblConvertStatusProgress { get { return GetText("C_lblConvertStatusProgress"); } }
        public string C_lblConvertStatusProgress2 { get { return GetText("C_lblConvertStatusProgress2"); } }
        public string C_dlgErrorInExport { get { return GetText("C_dlgErrorInExport"); } }
        public string C_tpLoad { get { return GetText("C_tpLoad"); } }
        public string C_lblStartDescription { get { return GetText("C_lblStartDescription"); } }
        public string C_mtHelp { get { return GetText("C_mtHelp"); } }
        public string C_mtSettings { get { return GetText("C_mtSettings"); } }
        public string C_mtLoadEportFile { get { return GetText("C_mtLoadEportFile"); } }
        public string C_tpChoose { get { return GetText("C_tpChoose"); } }
        public string C_btnChooseTip { get { return GetText("C_btnChooseTip"); } }
        public string C_btnDeselectAll { get { return GetText("C_btnDeselectAll"); } }
        public string C_btnSelectAll { get { return GetText("C_btnSelectAll"); } }
        public string C_tpPreviewWeb { get { return GetText("C_tpPreviewWeb"); } }
        public string C_tpPreviewText { get { return GetText("C_tpPreviewText"); } }
        public string C_clmnExport { get { return GetText("C_clmnExport"); } }
        public string C_clmnTitle { get { return GetText("C_clmnTitle"); } }
        public string C_tpExport { get { return GetText("C_tpExport"); } }
        public string C_lblConvertFormat { get { return GetText("C_lblConvertFormat"); } }
        public string C_mtConvertStart { get { return GetText("C_mtConvertStart"); } }
        public string C_tpSettings { get { return GetText("C_tpSettings"); } }
        public string C_lblSettingsStyle { get { return GetText("C_lblSettingsStyle"); } }
        public string C_lblSettingsLanguage { get { return GetText("C_lblSettingsLanguage"); } }
        public string C_btnSettingsClose { get { return GetText("C_btnSettingsClose"); } }
        public string C_lblSettingsCredits { get { return GetText("C_lblSettingsCredits"); } }
        public string C_lblUseProxy { get { return GetText("C_lblUseProxy"); } }
        public string C_lblProxyServer { get { return GetText("C_lblProxyServer"); } }
        public string C_lblProxyPort { get { return GetText("C_lblProxyPort"); } }
        public string C_lblSettingsTranslated { get { return GetText("C_lblSettingsTranslated"); } }
        public string C_dlgErrorHelp { get { return GetText("C_dlgErrorHelp"); } }

        public struct TranslationInfo
        {
            public string Author;
            public string AuthorLink;
            public string Language;
            public string LanguageShort;
            public string LastEdit;
        }
    }
}
