using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using MarkdownSharp;
using System.IO;

namespace WordAddIn1
{
    public partial class SimpleBrowser : UserControl
    {
        string FDscendHome = Ribbon1.FDscendHome;
        string online_url = Ribbon1.readme_path;
        bool mobileMod = false;


        public SimpleBrowser()
        {
            InitializeComponent();

            InitializeWebView2Async();

            this.Resize += new System.EventHandler(this.Form_Resize);

            addressBar.Text = online_url;
            addressBar.KeyUp += new KeyEventHandler(textBox1_KeyUp);
            toolTip1.SetToolTip(switchMod, "当前桌面模式");
        }

        private async void InitializeWebView2Async()
        {
            var env = await CoreWebView2Environment.CreateAsync(null, FDscendHome);
            await webView21.EnsureCoreWebView2Async(env);
            webView21.CoreWebView2.Navigate(online_url);
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            webView21.Size = this.ClientSize - new System.Drawing.Size(webView21.Location);
            addressBar.Width = webView21.Width - 270 - 20 - 4;
            gourl.Location = new Point(24 + switchMod.Width + addressBar.Width, 22);
            openFile.Location = new Point(gourl.Location.X + 94, 22);
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                gourl_Click(sender, e);
            }
        }

        private void gourl_Click(object sender, EventArgs e)
        {
            if (addressBar.Text.StartsWith("https://")) webView21.CoreWebView2.Navigate(addressBar.Text);
            else
            {
                webView21.CoreWebView2.Navigate("https://" + addressBar.Text);
                addressBar.Text = "https://" + addressBar.Text;
            }
        }

        private void switchMod_Click(object sender, EventArgs e)
        {
            if (!mobileMod)
            {
                webView21.CoreWebView2.Settings.UserAgent = "Mozilla/5.0 (iPhone; CPU iPhone OS 15_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148";
                mobileMod = true;
                toolTip1.SetToolTip(switchMod, "当前手机模式");
            }
            else
            {
                webView21.CoreWebView2.Settings.UserAgent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36";
                mobileMod = false;
                toolTip1.SetToolTip(switchMod, "当前桌面模式");
            }

            webView21.Reload();
        }

        private void openFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "打开文件";
            openFileDialog1.Filter= "HTML|*.htm;*.html|PDF|*.pdf|Markdown|*.md";
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                addressBar.Text = file;

                string fileExtension = Path.GetExtension(file);

                //markdown
                if (fileExtension.Equals(".md", StringComparison.OrdinalIgnoreCase))
                {
                    string markdownContent = File.ReadAllText(file);

                    Markdown markdown = new Markdown();
                    string htmlContent = markdown.Transform(markdownContent);

                    webView21.NavigateToString(htmlContent);
                }
                else
                {
                    webView21.CoreWebView2.Navigate(file);
                }
                
            }
        }
    }
}
