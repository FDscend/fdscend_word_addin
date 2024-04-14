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
using Microsoft.Office.Interop.Word;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    public partial class HighlightForm : UserControl
    {
        string FDscendHome = Ribbon1.FDscendHome;
        string highlight_index = Ribbon1.highlight_index;
        string online_url = @"https://highlightjs.org/demo";
        bool mobileMod = false;

        public static JObject ImportJSON(string jsonfile)
        {
            StreamReader reader = File.OpenText(jsonfile);
            JsonTextReader jsonTextReader = new JsonTextReader(reader);
            JObject jsonObject = (JObject)JToken.ReadFrom(jsonTextReader);
            reader.Close();
            return jsonObject;
        }

        public HighlightForm()
        {
            InitializeComponent();

            InitializeWebView2Async();

            JObject js = ImportJSON(Ribbon1.PresetHighlight);
            online_url = js["onlineMod"]["url"].ToString();
            mobileMod = (bool)js["onlineMod"]["mobileMod"];

#if DEBUG
            webView21.Location = new System.Drawing.Point(0, 80);
#endif
#if !DEBUG
            webView21.Location = new System.Drawing.Point(0, 0);
#endif

            insertCode.Width = 0;
            onlineMod.Left = insertCode.Left;
            onlineMod.Text = "切换离线模式";
            this.Resize += new System.EventHandler(this.Form_Resize);
        }

        private async void InitializeWebView2Async()
        {
            var env = await CoreWebView2Environment.CreateAsync(null, FDscendHome);
            await webView21.EnsureCoreWebView2Async(env);
            if (mobileMod)
            {
                webView21.CoreWebView2.Settings.UserAgent = "Mozilla/5.0 (iPhone; CPU iPhone OS 15_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148";
            }
            webView21.CoreWebView2.Navigate(online_url);
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            webView21.Size = this.ClientSize - new System.Drawing.Size(webView21.Location);
            onlineMod.Width = this.ClientSize.Width - insertCode.Width + insertCode.Left;
        }

        private void insertCode_Click(object sender, EventArgs e)
        {
            webView21.CoreWebView2.ExecuteScriptAsync($"document.getElementById(\"codeContent\").innerHTML=\"\"");
            foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                string codeContent = paragraph.Range.Text;
                codeContent = codeContent.Substring(0, codeContent.Length - 1);

                webView21.CoreWebView2.ExecuteScriptAsync($"document.getElementById(\"codeContent\").innerHTML+=\"{codeContent}\"\n");
            }

            //string codeContent = Globals.ThisAddIn.Application.Selection.Text;
            //codeContent = codeContent.Substring(0, codeContent.Length - 1);

            //MessageBox.Show(codeContent);
            //webView21.CoreWebView2.ExecuteScriptAsync($"document.getElementById(\"codeContent\").textContent=unescape(\"{codeContent}\")");

            //webView21.CoreWebView2.ExecuteScriptAsync("hljs.highlightAll()");
        }

        private void onlineMod_Click(object sender, EventArgs e)
        {
            if (insertCode.Width != 0)
            {
                insertCode.Width = 0;
                onlineMod.Width = this.ClientSize.Width;
                onlineMod.Left = insertCode.Left;
                onlineMod.Text = "切换离线模式";
                webView21.CoreWebView2.Navigate(online_url);
            }
            else
            {
                insertCode.Width = 122;
                onlineMod.Left = insertCode.Left + insertCode.Width;
                onlineMod.Text = "切换在线模式";
                webView21.CoreWebView2.Navigate(highlight_index);
            }
            
        }
    }
}
