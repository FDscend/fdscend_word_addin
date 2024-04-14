using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ChangeCharForm : UserControl
    {
        //全局路径
        string PresetFile = Ribbon1.ChangeCharList;

        public ChangeCharForm()
        {
            InitializeComponent();

            JObject js = ImportJSON(PresetFile);
            foreach (JObject jsob in js["data"])
            {
                //MessageBox.Show(jsob.ToString());

                bool ck;
                if (jsob["checked"].ToString() == "1") ck = true;
                else ck = false;
                checkedListBox1.Items.Add(jsob["from"] + " -> " + jsob["to"], ck);
            }

            this.Resize += new System.EventHandler(this.Form_Resize);
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            checkedListBox1.Width = this.ClientSize.Width - 8;
            button1.Width = checkedListBox1.Width;
            groupBox1.Width = checkedListBox1.Width;
            addFrom.Width = groupBox1.Width - 89 - 9;
            addTo.Width = addFrom.Width;
            addListChar.Width = groupBox1.Width - 12;
        }

        public static JObject ImportJSON(string jsonfile)
        {
            StreamReader reader = File.OpenText(jsonfile);
            JsonTextReader jsonTextReader = new JsonTextReader(reader);
            JObject jsonObject = (JObject)JToken.ReadFrom(jsonTextReader);
            reader.Close();
            return jsonObject;
        }

        public static void SetjsonFun(string jsonfile, JObject jsonObject)
        {
            string output = Newtonsoft.Json.JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(jsonfile, output);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //替换
            JObject js = ImportJSON(PresetFile);

            int selectCount = checkedListBox1.Items.Count;
            for (int i = 1; i <= selectCount; i++)
            {

                if (checkedListBox1.GetItemChecked(i - 1))
                {
                    //要替换的字符
                    object Replace_String = js["data"][i - 1]["from"];
                    js["data"][i - 1]["checked"] = 1;


                    //最终替换成的字符
                    object ReplaceWith = js["data"][i - 1]["to"].ToString();


                    object ms = System.Type.Missing;
                    object Replace = Word.WdReplace.wdReplaceAll;//设置替换方式:一，全部替换；二，只替换一个；三，一个都不替换。

                    //执行Word自带的查找/替换功能函数
                    Globals.ThisAddIn.Application.Selection.Find.Execute(ref Replace_String, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ReplaceWith, ref Replace, ref ms, ref ms, ref ms, ref ms);
                    
                }
                else
                    js["data"][i - 1]["checked"] = 0;
            
            }

            SetjsonFun(PresetFile, js);
        }

        private void addListChar_Click(object sender, EventArgs e)
        {
            string textFrom = addFrom.Text;
            string textTo = addTo.Text;

            if (textFrom.Length == 0) MessageBox.Show("From 字符串不能为空");
            else
            {
                JObject js = ImportJSON(PresetFile);
                ((JArray)js["data"]).Add(
                    new JObject()
                    {
                        { "from", textFrom},
                        { "to", textTo },
                        { "checked", 1 }
                    }
                    );

                SetjsonFun(PresetFile, js);

                checkedListBox1.Items.Add(textFrom + " -> " + textTo, true);

                addFrom.Text = "";
                addTo.Text = "";
            }

        }
    }
}
