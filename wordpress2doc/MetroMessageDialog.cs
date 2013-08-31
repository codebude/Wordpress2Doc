using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MetroFramework.Forms;
using MetroFramework.Components;


namespace Wordpress2Doc
{
    public partial class MetroMessageDialog : MetroForm
    {
        public MetroMessageDialog()
        {
            InitializeComponent();
        }

        public MetroMessageDialog(string message, MetroStyleManager ms)
        {
            InitializeComponent();
            this.StyleManager = new MetroStyleManager();
            this.StyleManager.Owner = this;
            this.StyleManager.Theme = ms.Theme;
            this.StyleManager.Style = ms.Style;
            this.ShadowType = MetroFormShadowType.SystemDropShadow;
            this.BorderStyle = MetroFramework.Drawing.MetroBorderStyle.None;
            richTextBoxMessage.Font = this.StyleManager.GetThemeFont(MetroFramework.Drawing.MetroFontSize.Medium, MetroFramework.Drawing.MetroFontWeight.Regular);
            richTextBoxMessage.Text = message;
            richTextBoxMessage.Enter += delegate(object s, EventArgs eA){
                richTextBoxMessage.DeselectAll();
                metroButtonOk.Focus();
            };
        }      

        private void metroButtonOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
