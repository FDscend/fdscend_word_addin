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
using System.IO.Compression;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading;


namespace WordAddIn1
{
    public partial class HighlightForm : UserControl
    {
        string FDscendHome = Ribbon1.FDscendHome;
        string highlight_index = Ribbon1.highlight_index;
        string tempFilePath = Ribbon1.temp_websource_htm;
        bool mobileMod = false;

        string srcHighlight = Ribbon1.srcHighlight;

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

            ExtractSrc(srcHighlight);

            JObject js = ImportJSON(Ribbon1.PresetHighlight);
            mobileMod = (bool)js["onlineMod"]["mobileMod"];

            this.Resize += new System.EventHandler(this.Form_Resize);

            toolTip1.SetToolTip(insertCode, "选中 WORD 中的代码，一键复制到输入框");
            toolTip1.SetToolTip(codeFormat, "选中粘贴的高亮代码，一键格式化");
        }

        public void ExtractSrc(string srczip)
        {
            string src_dir = Path.GetDirectoryName(srczip) + "\\highlight";

            if (!Directory.Exists(src_dir))
            {
                ZipFile.ExtractToDirectory(srczip, src_dir);
            }
        }

        private async void InitializeWebView2Async()
        {
            var env = await CoreWebView2Environment.CreateAsync(null, FDscendHome);
            await webView21.EnsureCoreWebView2Async(env);
            if (mobileMod)
            {
                webView21.CoreWebView2.Settings.UserAgent = "Mozilla/5.0 (iPhone; CPU iPhone OS 15_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148";
            }
            webView21.CoreWebView2.Navigate(highlight_index);

            webView21.CoreWebView2.NavigationCompleted += (sender, args) =>
            {
                // 在页面加载完成后执行 JavaScript 修改按钮文本
                webView21.CoreWebView2.ExecuteScriptAsync("document.getElementById('copyButton').innerText = '将高亮代码插入 WORD';");
                webView21.CoreWebView2.ExecuteScriptAsync("document.getElementById('title').style.display = 'none';");
            };

            webView21.CoreWebView2.WebMessageReceived += WebView_WebMessageReceived;
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            webView21.Size = this.ClientSize - new System.Drawing.Size(webView21.Location);
            insertCode.Width = this.ClientSize.Width / 2;
            codeFormat.Width = this.ClientSize.Width / 2;
            codeFormat.Location = new System.Drawing.Point(insertCode.Location.X + insertCode.Width, 3);
        }

        private void insertCode_Click(object sender, EventArgs e)
        {
            CopyWordSelectionToWebView();
        }

        /// <summary>
        /// 将 Word 选中的文本发送到 WebView2
        /// </summary>
        private async void CopyWordSelectionToWebView()
        {
            if (Globals.ThisAddIn.Application.Selection != null && Globals.ThisAddIn.Application.Selection.Text.Trim().Length > 0)
            {
                // string selectedText = Globals.ThisAddIn.Application.Selection.Text.Replace("\r", "\\n").Replace("\"", "\\\""); // 处理换行符和引号
                string selectedText = JsonConvert.SerializeObject(Globals.ThisAddIn.Application.Selection.Text);
                string script = $"document.getElementById('codeInput').value = {selectedText};";
                await webView21.CoreWebView2.ExecuteScriptAsync(script);
            }
        }

        /// <summary>
        /// 监听剪贴板内容并粘贴到 Word 中
        /// </summary>
        private void PasteFromClipboardToWord()
        {
            if (Globals.ThisAddIn.Application.Selection != null)
            {
                // 粘贴剪贴板内容到当前选定位置
                Globals.ThisAddIn.Application.Selection.Paste();

                // 等待 Word 处理粘贴
                // System.Windows.Forms.Application.DoEvents();

                // MessageBox.Show("finish copy");
            }
        }


        public void ApplyParagraphShading()
        {
            Range range = Globals.ThisAddIn.Application.Selection.Range;

            if (range == null)
            {
                MessageBox.Show("请选择一些文本。");
                return;
            }

            Range firstCharRange = range.Duplicate;
            firstCharRange.End = firstCharRange.Start + 1; // 仅获取第一个字符

            // 获取第一个字符的背景颜色
            WdColor firstCharShadingColor = firstCharRange.Shading.BackgroundPatternColor;
            Globals.ThisAddIn.Application.Selection.Font.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.Selection.ParagraphFormat.Shading.BackgroundPatternColor = firstCharShadingColor;
        }

        private void WebView_WebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            // 使用 TryGetWebMessageAsString 来获取消息
            string message = e.TryGetWebMessageAsString();
            if (message == "codeCopiedToClipboard")
            {
                PasteFromClipboardToWord();
            }
            else
            {
                MessageBox.Show("先点击“高亮代码”");
            }
        }

        private void codeFormat_Click(object sender, EventArgs e)
        {
            ApplyParagraphShading();
        }
    }
}
