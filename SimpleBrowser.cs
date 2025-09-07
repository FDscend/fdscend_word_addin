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
using System.IO.Compression;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    public partial class SimpleBrowser : UserControl
    {
        string FDscendHome = Ribbon1.FDscendHome;
        string online_url = Ribbon1.readme_path;
        string mdIndex = Ribbon1.mdIndex;
        string srcMarkdown = Ribbon1.srcMarkdown;
        string srcHighlight = Ribbon1.srcHighlight;

        bool mobileMod = false;


        public SimpleBrowser()
        {
            InitializeComponent();

            InitializeWebView2Async();

            this.Resize += new System.EventHandler(this.Form_Resize);

            ExtractSrc(srcMarkdown);
            ExtractSrc(srcHighlight, "highlight");

            addressBar.Text = online_url;
            addressBar.KeyUp += new KeyEventHandler(textBox1_KeyUp);
            toolTip1.SetToolTip(switchMod, "当前桌面模式");
        }

        public void ExtractSrc(string srczip, string dir="src")
        {
            string src_dir = Path.GetDirectoryName(srczip) + "\\" + dir;

            if (!Directory.Exists(src_dir))
            {
                ZipFile.ExtractToDirectory(srczip, src_dir);
            }
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
                    LoadMarkdownFile(file);


                    string markdownDirectory = Path.GetDirectoryName(file);  // 获取Markdown文件所在目录

                    webView21.CoreWebView2.WebMessageReceived += (sender_2, e_2) =>
                    {
                        try
                        {
                            var message = JsonConvert.DeserializeObject<dynamic>(e_2.WebMessageAsJson);
                            if (message.type == "loadMarkdown")
                            {
                                string mdPath = message.path;
                                // MessageBox.Show(mdPath);
                                // 处理路径，可能是相对路径
                                string fullPath = new Uri(mdPath).LocalPath;
                                LoadMarkdownFile(fullPath);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error processing web message: " + ex.Message);
                        }
                    };
                }
                else
                {
                    webView21.CoreWebView2.Navigate(file);
                }
                
            }
        }

        private void LoadMarkdownFile(string path)
        {
            string markdownContent = File.ReadAllText(path);
            string markdownDirectory = Path.GetDirectoryName(path);  // 获取Markdown文件所在目录
            string processedContent = PreprocessImagePaths(markdownContent, markdownDirectory);
            processedContent = ProcessHtmlLinks(processedContent, markdownDirectory);

            string escapedContent = JsonConvert.SerializeObject(processedContent);


            webView21.CoreWebView2.Navigate(mdIndex);

            webView21.CoreWebView2.NavigationCompleted += (sender_, args) =>
            {
                // 在页面加载完成后执行 JavaScrip
                string script = $@"document.getElementById(""text"").value = {escapedContent};
                                   renderAll();";
                webView21.CoreWebView2.ExecuteScriptAsync(script);
            };
        }



        public string PreprocessImagePaths(string markdownContent, string baseDirectory)
        {
            // 转换Windows路径为URL格式
            string basePath = new Uri(baseDirectory).AbsoluteUri;

            // 使用正则表达式匹配Markdown图片语法
            string pattern = @"\[(.*?)\]\((?!http|data:|file:///|#)(.*?)\)";
            string replacement = $"[$1](file:///{basePath}/$2)";

            // 替换所有相对路径
            return Regex.Replace(markdownContent, pattern, match =>
            {
                // 处理相对路径
                string relativePath = match.Groups[2].Value;
                string fullPath = Path.GetFullPath(Path.Combine(baseDirectory, relativePath));
                return $"[{match.Groups[1].Value}](file:///{fullPath.Replace("\\", "/")})";
            });
        }

        public string ProcessHtmlLinks(string htmlContent, string baseDirectory)
        {
            // 确保基础目录格式正确
            if (string.IsNullOrWhiteSpace(baseDirectory))
            {
                throw new ArgumentException("基础目录不能为空");
            }

            // 规范化基础目录路径
            baseDirectory = Path.GetFullPath(baseDirectory);
            if (!baseDirectory.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                baseDirectory += Path.DirectorySeparatorChar;
            }

            string reg = @"<(?<tag>a|img|link)\s+(?<pre>[^>]*?)(?:href|src)\s*=\s*(?<quote>[""'])(?<url>.+?)(\k<quote>)(?<post>[^>]*)>";


            // 匹配HTML中的href属性
            return Regex.Replace(htmlContent,
                reg,
                match =>
                {
                    string tag = match.Groups["tag"].Value;           // a / img / link
                    string preAttributes = match.Groups["pre"].Value; // href/src 前的属性
                    string quoteChar = match.Groups["quote"].Value;   // 引号
                    string originalUrl = match.Groups["url"].Value;   // 链接
                    string postAttributes = match.Groups["post"].Value; // href/src 后的属性

                    // 跳过 http、data、file://、# 等
                    if (ShouldSkipUrl(originalUrl))
                    {
                        return match.Value;
                    }

                    // 转换为绝对路径
                    string processedUrl = ProcessSingleUrl(originalUrl, baseDirectory);

                    // 重建标签
                    return $"<{tag} {preAttributes}href={quoteChar}{processedUrl}{quoteChar}{postAttributes}>";
                },
                RegexOptions.IgnoreCase);
        }

        private bool ShouldSkipUrl(string url)
        {
            // 跳过以下类型的URL：
            return url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("#") ||
                   string.IsNullOrWhiteSpace(url);
        }

        private string ProcessSingleUrl(string originalUrl, string baseDirectory)
        {
            try
            {
                // 分离主路径和片段标识（如#page=19）
                int fragmentIndex = originalUrl.IndexOf('#');
                string mainPath = fragmentIndex >= 0 ? originalUrl.Substring(0, fragmentIndex) : originalUrl;
                string fragment = fragmentIndex >= 0 ? originalUrl.Substring(fragmentIndex) : "";

                // 如果是相对路径
                if (!Path.IsPathRooted(mainPath))
                {
                    // 获取完整路径并转换为URI格式
                    string fullPath = Path.GetFullPath(Path.Combine(baseDirectory, mainPath));

                    // 转换为file:///格式
                    Uri fileUri = new Uri(fullPath);

                    // 重新组合URL
                    return fileUri.AbsoluteUri + fragment;
                }

                // 已经是绝对路径的情况
                return originalUrl;
            }
            catch (Exception ex)
            {
                // 记录错误并返回原始URL
                System.Diagnostics.Debug.WriteLine($"URL处理失败: {originalUrl}, 错误: {ex.Message}");
                return originalUrl;
            }
        }
    }
}
