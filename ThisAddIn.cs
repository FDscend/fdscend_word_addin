using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.IO;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            File.Delete(Ribbon1.latest_info);
            DeleteFolder(Ribbon1.tempFile);
            DeleteFolder(Ribbon1.FDscendHome + "\\EBWebView");
        }

        void DeleteFolder(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                try
                {
                    // 删除文件夹内的所有文件
                    string[] files = Directory.GetFiles(folderPath);
                    foreach (string file in files)
                    {
                        File.Delete(file);
                    }

                    // 递归删除子文件夹
                    string[] subFolders = Directory.GetDirectories(folderPath);
                    foreach (string subFolder in subFolders)
                    {
                        DeleteFolder(subFolder);
                    }

                    // 删除空文件夹
                    Directory.Delete(folderPath);
                }
                catch (Exception ex)
                {
#if DEBUG
                    MessageBox.Show($"删除文件夹失败: {ex.Message}");
#endif
                }
            }     
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
