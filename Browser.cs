using System;
using System.Windows.Forms;


namespace PluginSalas
{
    public partial class Browser : Form
    {
        public Browser(string url)
        {
            InitializeComponent();
            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.Navigate(url);
        }

        private void Browser_Load(object sender, EventArgs e)
        {
        }

        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
        }
    }
             
}
