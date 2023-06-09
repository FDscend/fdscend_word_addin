﻿using System;
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


namespace WordAddIn1
{
    public partial class AboutForm : Form
    {
        //全局路径
        string latest_info = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\分点作答\\FDscend\\latest.json";
        string url = Properties.Resources.latest_info_url;
#if DEBUG
        string doc_path = "D:\\code\\WordAddIn1\\Resources\\说明文档.pdf";
#endif
#if !DEBUG
        string doc_path = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\分点作答\\FDscend\\说明文档.pdf";
#endif
        public AboutForm()
        {
            InitializeComponent();

            //version
            //label2.Text = "当前版本：" + Properties.Resources.current_ver + "\r\n内部版本：" + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            label2.Text = "当前版本：" + Properties.Resources.current_ver;
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
            System.Diagnostics.Process.Start(Properties.Resources.github_code_url);
        }

        private void check_ver_Click(object sender, EventArgs e)
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
    }
}
