using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Net.NetworkInformation;


namespace WordAddIn1
{
    public partial class AboutForm : Form
    {
        //全局路径
        string latest_info = Ribbon1.latest_info;
        string url = Properties.Resources.latest_info_url;
        string doc_path = Ribbon1.pdf_path;


        public AboutForm()
        {
            InitializeComponent();

            //version
            string mod = "";
#if DEBUG
            mod = "DEBUG";
#endif
#if !DEBUG
            mod = "RELEASE";
#endif
            //label2.Text = "当前版本：" + Properties.Resources.current_ver + "\r\n内部版本：" + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            label2.Text = "当前版本：" + Properties.Resources.current_ver + " " + mod;
        }

        private void AboutForm_Load(object sender, EventArgs e)
        {

        }

        private void help_doc_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(doc_path);
        }

        private void git_web_Click(object sender, EventArgs e)
        {
            if (GetNetStatus()) System.Diagnostics.Process.Start(Properties.Resources.github_code_url);
            else MessageBox.Show("network error");
        }

        private void check_ver_Click(object sender, EventArgs e)
        {
            if (GetNetStatus())
            {
                if (File.Exists(latest_info))
                {
                    File.Delete(latest_info);
                }


                WebClient webClient = new WebClient();
                webClient.Headers.Add("user-agent", "Mozilla/4.0 ((compatible; MSIE 8.0; Windows NT 6.1;.NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729;)");

                ServicePointManager.SecurityProtocol = (SecurityProtocolType)48
                                                | (SecurityProtocolType)192
                                                | (SecurityProtocolType)768
                                                | (SecurityProtocolType)3072;

                webClient.DownloadFile(url, latest_info);


                JObject js = ImportJSON(latest_info);

                Version ver_cur = new Version(Properties.Resources.current_ver);
                Version ver_latest = new Version(js["tag_name"].ToString().Substring(1));

                if (ver_cur == ver_latest) MessageBox.Show("已是最新版啦！");
                if (ver_cur < ver_latest) MessageBox.Show("有新版可用！");
                if (ver_cur > ver_latest) MessageBox.Show("你怎么会比网站版本还要新？");
            }
            else MessageBox.Show("network error");
            
        }

        public static JObject ImportJSON(string jsonfile)
        {
            StreamReader reader = File.OpenText(jsonfile);
            JsonTextReader jsonTextReader = new JsonTextReader(reader);
            JObject jsonObject = (JObject)JToken.ReadFrom(jsonTextReader);
            reader.Close();
            return jsonObject;
        }

        private void download_latest_Click(object sender, EventArgs e)
        {
            if (GetNetStatus())
            {
                if (File.Exists(latest_info))
                {
                    File.Delete(latest_info);
                }


                WebClient webClient = new WebClient();
                webClient.Headers.Add("user-agent", "Mozilla/4.0 ((compatible; MSIE 8.0; Windows NT 6.1;.NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729;)");

                ServicePointManager.SecurityProtocol = (SecurityProtocolType)48
                                                | (SecurityProtocolType)192
                                                | (SecurityProtocolType)768
                                                | (SecurityProtocolType)3072;

                webClient.DownloadFile(url, latest_info);


                JObject js = ImportJSON(latest_info);
                //MessageBox.Show(js["assets"][0]["browser_download_url"].ToString());
                System.Diagnostics.Process.Start(js["assets"][0]["browser_download_url"].ToString());
            }
            else MessageBox.Show("network error");

        }

        bool GetNetStatus()
        {
            Ping ping = new Ping();
            try
            {
                PingReply pingReply = ping.Send("api.github.com");
                if (pingReply.Status == IPStatus.Success)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            } 
        }
    }
}
