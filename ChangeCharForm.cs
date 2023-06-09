﻿using System;
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
#if DEBUG
        const string PresetFile = "D:\\code\\WordAddIn1\\Resources\\ChangeCharList";
#endif
#if !DEBUG
        string PresetFile = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\分点作答\\FDscend\\Presets\\ChangeCharList";
#endif

        public ChangeCharForm()
        {
            InitializeComponent();

            JObject js = ImportJSON(PresetFile);
            int PresetNum = js.Count - 2;
            for (int i = 1; i <= PresetNum; i++)
            {
                bool ck;
                if (js[i.ToString()]["checked"].ToString() == "1") ck = true;
                else ck = false;
                checkedListBox1.Items.Add(js[i.ToString()]["from"] + " -> " + js[i.ToString()]["to"]["to1"], ck);
            }
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
                    object Replace_String = js[i.ToString()]["from"];
                    js[i.ToString()]["checked"] = 1;


                    //最终替换成的字符
                    JObject js_to = (JObject)js[i.ToString()]["to"];
                    int tocount = js_to.Count;

                    if(tocount == 1)
                    {
                        object ReplaceWith = js_to["to" + 1].ToString();


                        object ms = System.Type.Missing;
                        object Replace = Word.WdReplace.wdReplaceAll;//设置替换方式:一，全部替换；二，只替换一个；三，一个都不替换。

                        //执行Word自带的查找/替换功能函数
                        Globals.ThisAddIn.Application.ActiveDocument.Content.Find.Execute(ref Replace_String, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ReplaceWith, ref Replace, ref ms, ref ms, ref ms, ref ms);

                    }
                    
                }
                else
                    js[i.ToString()]["checked"] = 0;
            
            }

            SetjsonFun(PresetFile, js);
        }
    }
}
