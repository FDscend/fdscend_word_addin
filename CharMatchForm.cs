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
    public partial class CharMatchForm : UserControl
    {
        //全局路径
        string PresetFile = Ribbon1.CharMatchList;



        //全局变量
        List<string> charLeft = new List<string>();
        List<string> charRight = new List<string>();

        public CharMatchForm()
        {
            InitializeComponent();

            JObject js = ImportJSON(PresetFile);
            int PresetNum = js.Count - 2;
            for (int i = 1; i <= PresetNum; i++)
            {
                bool ck;
                if (js[i.ToString()]["checked"].ToString() == "1") ck = true;
                else ck = false;
                checkedListBox1.Items.Add(js[i.ToString()]["name"], ck);
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

        // 符号匹配栈
        static bool IsBracketsMatch(string matchStr, List<string> LeftList, List<string> RightList)
        {
            Stack<string> stack = new Stack<string>();
            for (int i = 0; i < matchStr.Length; i++)
            {
                string current = matchStr[i].ToString();

                if(LeftList.Contains(current))
                {
                    stack.Push(current);
                }
                if(RightList.Contains(current))
                {
                    if (stack.Count <= 0)
                    {
                        return false;
                    }
                    else
                    {
                        string top = stack.Peek();
                        if (IsCouple(top, current, LeftList, RightList))
                        {
                            stack.Pop();
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                
            }
            if (stack.Count <= 0) return true;
            return false;
        }

        static bool IsCouple(string left, string right, List<string> LeftList, List<string> RightList)
        {
            for (int i = 0; i < LeftList.Count; i++)
            {
                if (left == LeftList[i] && right == RightList[i])
                {
                    return true;
                }
            }
                       
            return false;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            JObject js = ImportJSON(PresetFile);

            // 初始化左右列表
            charLeft.Clear();
            charRight.Clear();

            int selectCount = checkedListBox1.Items.Count;
            for (int i = 1; i <= selectCount; i++)
            {
                if (checkedListBox1.GetItemChecked(i - 1))
                {                   
                    charLeft.Add(js[i.ToString()]["left"].ToString());
                    charRight.Add(js[i.ToString()]["right"].ToString());
                }
                else
                    js[i.ToString()]["checked"] = 0;
            }

            //MessageBox.Show("charLeft: " + charLeft.Count.ToString());
            //MessageBox.Show("charRight: " + charRight.Contains("）").ToString());


            //检测符号匹配

            //https://blog.csdn.net/wucdsg/article/details/99009205

            string text = Globals.ThisAddIn.Application.Selection.Text;

            //MessageBox.Show(text);
            
            bool IsMatch = IsBracketsMatch(text, charLeft, charRight);

            if (IsMatch == true)
            {
                MessageBox.Show("符号成功配对");
            }
            else
            {
                MessageBox.Show("警告！存在未配对符号");
            }
            //MessageBox.Show(IsMatch.ToString());

            SetjsonFun(PresetFile, js);
        }
    }
}
