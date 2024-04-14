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

namespace WordAddIn1
{
    public partial class SettingForm : Form
    {
        
        string ControlKey = Ribbon1.ControlKey;
        string SettingsFile = Ribbon1.SettingsFile;


        //全局常量
        const string KeyDefault = Ribbon1.KeyDefault;
        const string KeyXMT = Ribbon1.KeyXMT;
        const string KeyCode = Ribbon1.KeyCode;
        const string KeyCode2 = Ribbon1.KeyCode2;
        const string KeyCodeLatex = Ribbon1.KeyCodeLatex;
        const string KeyCode4 = Ribbon1.KeyCode4;
        const string KeyToolsBox = Ribbon1.KeyToolsBox;
        const string KeyRunCode = Ribbon1.KeyRunCode;
        const string KeyBrowser = Ribbon1.KeyBrowser;
        const string KeyAdmin = Ribbon1.KeyAdmin;

        string[] keyArray = new string[] { KeyXMT, KeyCode, KeyCode2, KeyCodeLatex, KeyCode4, KeyToolsBox, KeyRunCode , KeyBrowser };
        string[] nameArray = new string[] { "文案", "代码排版", "代码排版3", "代码排版2", "代码排版4", "工具箱", "运行代码" ,"简单浏览器"};

        public bool boolKeyAllTrue;

        public SettingForm()
        {
            InitializeComponent();

            loadList();

            boolKeyAllTrue = false;

            textBox1.KeyUp += new KeyEventHandler(textBox1_KeyUp);

            JObject js_settings = ImportJSON(SettingsFile);
            usePyScripts.Checked = (bool)js_settings["usePyScripts"];
            toolTip1.SetToolTip(usePyScripts, "如勾选，则使用 FDscend/scripts 中的 Python 脚本;\r\n否则，使用 FDscend/scripts 中的同名exe文件");
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //MessageBox.Show("回车");
                button1_Click(sender, e);
            }
        }

        void loadList()
        {
            JObject js = ImportJSON(ControlKey);
            for (int i = 3; i < js.Count; i++)
            {
                if (js.Properties().ToList()[i].Name.ToString().StartsWith("key"))
                {
                    //MessageBox.Show(js.Properties().ToList()[i].Value.ToString());
                    checkedListBox1.Items.Add(GetControlName(js.Properties().ToList()[i].Value.ToString()));
                }
            }

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (js["state_" + GetControlStateCode(checkedListBox1.Items[i].ToString())].ToString() == "1")
                {
                    checkedListBox1.SetItemChecked(i, true);
                }

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

        string GetControlName(string input)
        {
            string output = "";

            if(keyArray.Contains(input))
            {
                int index_1 = Array.IndexOf(keyArray, input);
                output = nameArray[index_1];
            }
            else
            {
                MessageBox.Show("加载错误name", "操作码");
            }

            return output;
        }

        string GetControlStateCode(string input)
        {
            string output = "";

            if (nameArray.Contains(input))
            {
                int index_1 = Array.IndexOf(nameArray, input);
                output = keyArray[index_1];
            }
            else
            {
                MessageBox.Show("加载错误state", "操作码");
            }

            return output;
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            JObject js = ImportJSON(ControlKey);

            if (e.NewValue == CheckState.Checked)
            {
                js["state_" + GetControlStateCode(checkedListBox1.Items[e.Index].ToString())] = "1";
            }
            else
            {
                js["state_" + GetControlStateCode(checkedListBox1.Items[e.Index].ToString())] = "0";
            }

            SetjsonFun(ControlKey, js);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string key = textBox1.Text;

            if (key == KeyAdmin)
            {
                boolKeyAllTrue = true;
                this.Close();
            }
            else if (BoolKey(key) == true)
            {
                JObject Jtemp = ImportJSON(ControlKey);
                Jtemp["key_" + key] = key;
                Jtemp["state_" + key] = "1";
                SetjsonFun(ControlKey, Jtemp);

                checkedListBox1.Items.Clear();
                loadList();

                textBox1.Text = "";
            }
        }

        private bool BoolKey(string key)
        {
            switch (key)
            {
                case KeyDefault:
                    break;
                case KeyXMT:
                    break;
                case KeyCode:
                    break;
                case KeyCode2:
                    break;
                case KeyCodeLatex:
                    break;
                case KeyCode4:
                    break;
                case KeyToolsBox:
                    break;
                case KeyRunCode:
                    break;
                case KeyBrowser:
                    break;
                default:
                    return false;
            }
            return true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void usePyScripts_CheckedChanged(object sender, EventArgs e)
        {
            JObject js = ImportJSON(SettingsFile);
            js["usePyScripts"] = usePyScripts.Checked;
            SetjsonFun(SettingsFile, js);
        }
    }
}
