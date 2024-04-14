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
    public partial class CodeControlForm : Form
    {
        //全局路径
        //#if DEBUG
        //        const string PresetCodeFile = "D:\\code\\WordAddIn1\\Resources\\Preset_Code";
        //#endif
        //#if !DEBUG
        //        string PresetCodeFile = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + "\\分点作答\\FDscend\\Presets\\Preset_Code";
        //#endif

        string PresetCodeFile;

        public CodeControlForm(string preset)
        {
            InitializeComponent();

            PresetCodeFile = preset;

            //设置代码左侧行号的按钮
            JObject js = ImportJSON(PresetCodeFile);
            string CodeListNumStatus = js["CodeListNum"].ToString();
            if (CodeListNumStatus == "yes")
            {
                CodeListYN.Checked = true;
            }
            else
            {
                CodeListYN.Checked = false;
            }

            //代码预设列表
            int PresetNum = int.Parse(js["num"].ToString());
            for (int i = 1; i <= PresetNum; i++)
            {
                CodePresetList.Items.Add(js["preset_" + i + "_name"].ToString());
            }
            DefaultPreset.Text = "默认预设：" + js["preset_" + int.Parse(js["DefaultPreset"].ToString()) + "_name"].ToString();

            //默认字体
            DefaultFont.Text = js["DefaultFont"].ToString();

            //设置代码左侧行号的竖线
            string CodeBorder = js["CodeBorder"].ToString();
            if (CodeBorder == "yes")
            {
                checkBoxBorderLine.Checked = true;
            }
            else
            {
                checkBoxBorderLine.Checked = false;
            }

            //设置代码块是否平铺
            string FullPageWidth = js["FullPageWidth"].ToString();
            if (FullPageWidth == "yes")
            {
                CodeFullPageSet.Checked = true;
            }
            else
            {
                CodeFullPageSet.Checked = false;
            }

            //隐藏导入预设控件组
            groupBoxFont.Top = 305;
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

        private void CodeListYN_CheckedChanged(object sender, EventArgs e)
        {
            //默认加载代码行号

            JObject js = ImportJSON(PresetCodeFile);
            if (CodeListYN.Checked == true)
            {
                js["CodeListNum"] = "yes";
                //MessageBox.Show(CodeListYN.Checked.ToString() + "1");
            }
            else
            {
                js["CodeListNum"] = "no";
                //MessageBox.Show(CodeListYN.Checked.ToString() + "2");
            }
            SetjsonFun(PresetCodeFile, js);
        }

        private void checkBoxBorderLine_CheckedChanged(object sender, EventArgs e)
        {
            //默认加载代码竖线

            JObject js = ImportJSON(PresetCodeFile);
            if (checkBoxBorderLine.Checked == true)
            {
                js["CodeBorder"] = "yes";
            }
            else
            {
                js["CodeBorder"] = "no";
            }
            SetjsonFun(PresetCodeFile, js);
        }

        private void CodeFullPageSet_CheckedChanged(object sender, EventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);
            if (CodeFullPageSet.Checked == true)
            {
                js["FullPageWidth"] = "yes";
            }
            else
            {
                js["FullPageWidth"] = "no";
            }
            SetjsonFun(PresetCodeFile, js);
        }

        private void DefaultPresetOK_Click(object sender, EventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);

            int PresetSelected = CodePresetList.SelectedIndex + 1;

            if (PresetSelected != 0)
            {
                js["DefaultPreset"] = (CodePresetList.SelectedIndex + 1).ToString();
                SetjsonFun(PresetCodeFile, js);
                DefaultPreset.Text = "默认预设：" + js["preset_" + (CodePresetList.SelectedIndex + 1) + "_name"].ToString();
            }         
        }

        private void DeletePreset_Click(object sender, EventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);
            
            int PresetSelected = CodePresetList.SelectedIndex + 1;
            int PresetCount = int.Parse(js["num"].ToString());

            //MessageBox.Show(PresetSelected.ToString());

            if(PresetSelected!=0)
            {
                for (int i = PresetSelected; i < PresetCount; i++)
                {
                    int j = i + 1;
                    js["preset_" + i + "_name"] = js["preset_" + j + "_name"];
                    js["preset_" + i + "_CodeBackcolor1_r"] = js["preset_" + j + "_CodeBackcolor1_r"];
                    js["preset_" + i + "_CodeBackcolor1_g"] = js["preset_" + j + "_CodeBackcolor1_g"];
                    js["preset_" + i + "_CodeBackcolor1_b"] = js["preset_" + j + "_CodeBackcolor1_b"];
                    js["preset_" + i + "_CodeBackcolor2_r"] = js["preset_" + j + "_CodeBackcolor2_r"];
                    js["preset_" + i + "_CodeBackcolor2_g"] = js["preset_" + j + "_CodeBackcolor2_g"];
                    js["preset_" + i + "_CodeBackcolor2_b"] = js["preset_" + j + "_CodeBackcolor2_b"];
                    js["preset_" + i + "_CodeBorderLine_r"] = js["preset_" + j + "_CodeBorderLine_r"];
                    js["preset_" + i + "_CodeBorderLine_g"] = js["preset_" + j + "_CodeBorderLine_g"];
                    js["preset_" + i + "_CodeBorderLine_b"] = js["preset_" + j + "_CodeBorderLine_b"];
                }

                //delete json opt
                js["num"] = (PresetCount - 1).ToString();

                js["preset_" + PresetCount + "_name"] = "";
                js["preset_" + PresetCount + "_CodeBackcolor1_r"] = "";
                js["preset_" + PresetCount + "_CodeBackcolor1_g"] = "";
                js["preset_" + PresetCount + "_CodeBackcolor1_b"] = "";
                js["preset_" + PresetCount + "_CodeBackcolor2_r"] = "";
                js["preset_" + PresetCount + "_CodeBackcolor2_g"] = "";
                js["preset_" + PresetCount + "_CodeBackcolor2_b"] = "";
                js["preset_" + PresetCount + "_CodeBorderLine_r"] = "";
                js["preset_" + PresetCount + "_CodeBorderLine_g"] = "";
                js["preset_" + PresetCount + "_CodeBorderLine_b"] = "";

                if (PresetSelected == int.Parse(js["DefaultPreset"].ToString()))
                {
                    js["DefaultPreset"] = "1";
                }
                if (PresetSelected < int.Parse(js["DefaultPreset"].ToString()))
                {
                    js["DefaultPreset"] = (PresetSelected).ToString();
                }

                SetjsonFun(PresetCodeFile, js);

                //refresh preset list
                CodePresetList.Items.Clear();
                for (int i = 1; i < PresetCount; i++)
                {
                    CodePresetList.Items.Add(js["preset_" + i + "_name"].ToString());
                }
                DefaultPreset.Text = "默认预设：" + js["preset_" + int.Parse(js["DefaultPreset"].ToString()) + "_name"].ToString();
            }          
        }

        private void DefaultFontOK_Click(object sender, EventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);

            DialogResult dr = fontDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                DefaultFont.Text = fontDialog1.Font.Name;
                DefaultFont.Font = fontDialog1.Font;

                js["DefaultFont"] = fontDialog1.Font.Name;
                js["DefaultFontSize"] = fontDialog1.Font.Size;
                SetjsonFun(PresetCodeFile, js);
            }
        }

        private void ExportPreset_Click(object sender, EventArgs e)
        {
            //导出预设            
            string Path;
            saveFileDialogExPreset.Filter = "FDscend代码预设(*.fdcp)|*.fdcp";
            DialogResult dr = saveFileDialogExPreset.ShowDialog();
            if (dr == DialogResult.OK)
            {
                //记录选中的目录  
                Path = saveFileDialogExPreset.FileName;
                //MessageBox.Show(Path);
                if (File.Exists(PresetCodeFile))//必须判断要复制的文件是否存在
                {
                    File.Copy(PresetCodeFile, Path, true);//三个参数分别是源文件路径，存储路径，若存储路径有相同文件是否替换
                }
            }
        }

        string InportFile = "";//导入预设文件
        private void InportPreset_Click(object sender, EventArgs e)
        {
            //导入预设
            openFileDialogINPreset.Title = "选择预设文件";
            openFileDialogINPreset.Filter = "FDscend代码预设(*.fdcp)|*.fdcp|Preset_Code(*.*)|*.*";
            DialogResult dr = openFileDialogINPreset.ShowDialog();
            if (dr == DialogResult.OK)
            {
                InportFile = openFileDialogINPreset.FileName;
                //MessageBox.Show(InportFile);
                
                //显示导入预设控件组，其他控件让位
                InportPresetGroup.Visible = true;
                groupBoxFont.Top = 434;

                //导入预设列表                
                InportPresetList.Items.Clear();//刷新导入预设列表
                JObject js = ImportJSON(InportFile);                
                int PresetNum = int.Parse(js["num"].ToString());
                for (int i = 1; i <= PresetNum; i++)
                {
                    InportPresetList.Items.Add(js["preset_" + i + "_name"].ToString());
                }
            }
        }

        private void InportOK_Click(object sender, EventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);//原先预设
            JObject jsIn = ImportJSON(InportFile);//要导入的预设

            int InportSelected = InportPresetList.SelectedIndex + 1;
            int StartPresetCount = int.Parse(js["num"].ToString());

            if (InportSelected != 0)
            {
                js["preset_" + (StartPresetCount + 1) + "_name"] = jsIn["preset_" + InportSelected + "_name"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBackcolor1_r"] = jsIn["preset_" + InportSelected + "_CodeBackcolor1_r"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBackcolor1_g"] = jsIn["preset_" + InportSelected + "_CodeBackcolor1_g"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBackcolor1_b"] = jsIn["preset_" + InportSelected + "_CodeBackcolor1_b"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBackcolor2_r"] = jsIn["preset_" + InportSelected + "_CodeBackcolor2_r"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBackcolor2_g"] = jsIn["preset_" + InportSelected + "_CodeBackcolor2_g"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBackcolor2_b"] = jsIn["preset_" + InportSelected + "_CodeBackcolor2_b"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBorderLine_r"] = jsIn["preset_" + InportSelected + "_CodeBorderLine_r"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBorderLine_g"] = jsIn["preset_" + InportSelected + "_CodeBorderLine_g"];
                js["preset_" + (StartPresetCount + 1) + "_CodeBorderLine_b"] = jsIn["preset_" + InportSelected + "_CodeBorderLine_b"];
                js["num"] = (StartPresetCount + 1).ToString();

                SetjsonFun(PresetCodeFile, js);

                //刷新本来预设列表
                CodePresetList.Items.Clear();
                for (int i = 1; i <= StartPresetCount + 1; i++)
                {
                    CodePresetList.Items.Add(js["preset_" + i + "_name"].ToString());
                }

                //刷新导入预设列表
                //InportPresetList.Items.Clear();
                InportPresetList.Items.Remove(InportPresetList.SelectedItem);
                InportPresetList.Refresh();
            }            
        }
    }
}
