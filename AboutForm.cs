using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace WordAddIn1
{
    public partial class AboutForm : Form
    {
        //全局路径
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
            this.Text = "关于（ver: " + Assembly.GetExecutingAssembly().GetName().Version.ToString() + "）";
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
            System.Diagnostics.Process.Start("https://github.com/FDscend/fdscend_word_addin");
        }
    }
}
