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
using Microsoft.VisualBasic;//inputbox

namespace WordAddIn1
{
    public partial class TableColoringForm : UserControl
    {
        //全局路径
        // #if DEBUG
        //         const string PresetFile = "D:\\code\\WordAddIn1\\Resources\\ToolsBox_TablePreset";
        // #endif
        // #if !DEBUG
        //         string PresetFile = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + "\\分点作答\\FDscend\\Presets\\ToolsBox_TablePreset";
        // #endif

        string PresetFile;
        Color TableBackcolor1;
        Color TableBackcolor2;
        Color TableBackcolorH;

        public TableColoringForm(string preset)
        {
            InitializeComponent();

            PresetFile = preset;

            //表格预设列表
            JObject js = ImportJSON(PresetFile);
            int PresetNum = int.Parse(js["num"].ToString());
            for (int i = 1; i <= PresetNum; i++)
            {
                TableColorPresetList.Items.Add(js["preset_" + i + "_name"].ToString());
            }
            TableColorPresetList.SelectedIndex = 0;


            //表格效果展示
            TableBackcolor1 = ReadTHINGcolor(1, "TableBackcolor1");
            TableBackcolor2 = ReadTHINGcolor(1, "TableBackcolor2");
            TableBackcolorH = ReadTHINGcolor(1, "TableBackcolorH");
            SetTableLineColor();
            SetTableLineTxt();


            this.Resize += new System.EventHandler(this.Form_Resize);
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            groupTableColor.Width = this.ClientSize.Width - 8;
            TableColorPresetList.Width = groupTableColor.Width - 14;

            groupColorEffect.Width = groupTableColor.Width;
            TableLine1.Width = TableColorPresetList.Width;
            TableLine2.Width = TableLine1.Width;
            TableLine3.Width = TableLine1.Width;
            TableLine4.Width = TableLine1.Width;
            TableLine5.Width = TableLine1.Width;
            TableLine6.Width = TableLine1.Width;


            int margin = (this.ClientSize.Width - 309) / 2;

            ColorOK.Location = new Point(4 + margin, 486);
            ExchangeColor.Location = new Point(137 + margin, 486);
            SavePreset.Location = new Point(247 + margin, 486);

            NewColor.Location = new Point(4 + margin, 531);
            NewColor2.Location = new Point(82 + margin, 531);
            NewColorH.Location = new Point(161 + margin, 531);
            DeletePreset.Location = new Point(247 + margin, 531);

            checkBoxFirstLine.Location = new Point(4 + margin, 576);
            Export.Location = new Point(126 + margin, 576);
            checkBoxInport.Location = new Point(220 + margin, 576);


            InportPresetGroup.Width = groupTableColor.Width;
            InportPresetList.Width = InportPresetGroup.Width - 101;
            InportOK.Location = new Point(InportPresetGroup.Location.X + InportPresetGroup.Width - 92, 35);
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

        private Word.WdColor GetColor(Color c)
        {
            //color
            UInt32 R = 0x1, G = 0x100, B = 0x10000;
            return (Word.WdColor)(R * c.R + G * c.G + B * c.B);
        }

        public Color ReadTHINGcolor(int choice, string colorseries)
        {
            //读取代码预设颜色

            JObject js = ImportJSON(PresetFile);

            int r = int.Parse(js["preset_" + choice + "_" + colorseries + "_r"].ToString());
            int g = int.Parse(js["preset_" + choice + "_" + colorseries + "_g"].ToString());
            int b = int.Parse(js["preset_" + choice + "_" + colorseries + "_b"].ToString());

            return Color.FromArgb(r, g, b);
        }

        
        private void TableColorPresetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            //更新颜色预览
            int PresetChoice;
            if (TableColorPresetList.SelectedIndex == -1)
                PresetChoice = 1;
            else
                PresetChoice = TableColorPresetList.SelectedIndex + 1;

            TableBackcolor1 = ReadTHINGcolor(PresetChoice, "TableBackcolor1");
            TableBackcolor2 = ReadTHINGcolor(PresetChoice, "TableBackcolor2");
            TableBackcolorH = ReadTHINGcolor(PresetChoice, "TableBackcolorH");

            SetTableLineColor();
        }

        void SetTableLineColor()
        {
            //刷新预览的颜色

            if (checkBoxFirstLine.Checked == false)
            {
                TableLine1.BackColor = TableBackcolor1;
                TableLine2.BackColor = TableBackcolor2;
                TableLine3.BackColor = TableBackcolor1;
                TableLine4.BackColor = TableBackcolor2;
                TableLine5.BackColor = TableBackcolor1;
                TableLine6.BackColor = TableBackcolor2;
            }
            else
            {
                TableLine1.BackColor = TableBackcolorH;
                TableLine2.BackColor = TableBackcolor2;
                TableLine3.BackColor = TableBackcolor1;
                TableLine4.BackColor = TableBackcolor2;
                TableLine5.BackColor = TableBackcolor1;
                TableLine6.BackColor = TableBackcolor2;
            }
        }

        void SetTableLineTxt()
        {
            //刷新显示内容
            if (checkBoxFirstLine.Checked == false)
            {
                TableLine1.Text = "第1行\t(๑•̀ㅂ•́)و✧";
                TableLine2.Text = "第2行\t(>▽<)";
                TableLine3.Text = "第3行\t(*ฅ́ˇฅ̀*)";
                TableLine4.Text = "第4行\t(′▽`〃)";
                TableLine5.Text = "第5行\tヾ(^▽^*)))";
                TableLine6.Text = "第6行\tヾ(≧▽≦*)o";
            }
            else
            {
                TableLine1.Text = "首行\t表格标题";
                TableLine2.Text = "第1行\t(๑•̀ㅂ•́)و✧";
                TableLine3.Text = "第2行\t(>▽<)";
                TableLine4.Text = "第3行\t(*ฅ́ˇฅ̀*)";
                TableLine5.Text = "第4行\t(′▽`〃)";
                TableLine6.Text = "第5行\tヾ(^▽^*)))";
            }
        }

        private void ColorOK_Click(object sender, EventArgs e)
        {
            //确认着色

            int table_select_end = Globals.ThisAddIn.Application.Selection.End;
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
            int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
            Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

            int row = tables_select.Range.Rows.Count;

            if (checkBoxFirstLine.Checked == false)
            {
                for (int i = 1; i <= row; i++)
                {
                    if (i % 2 == 1)
                    {
                        tables_select.Rows[i].Shading.BackgroundPatternColor = GetColor(TableBackcolor1);
                    }
                    else
                    {
                        tables_select.Rows[i].Shading.BackgroundPatternColor = GetColor(TableBackcolor2);
                    }
                }
            }
            else
            {
                tables_select.Rows[1].Shading.BackgroundPatternColor = GetColor(TableBackcolorH);

                for (int i = 2; i <= row; i++)  //排除首行
                {
                    if (i % 2 == 1)
                    {
                        tables_select.Rows[i].Shading.BackgroundPatternColor = GetColor(TableBackcolor1);
                    }
                    else
                    {
                        tables_select.Rows[i].Shading.BackgroundPatternColor = GetColor(TableBackcolor2);
                    }
                }
            }
                
        }

        private void ExchangeColor_Click(object sender, EventArgs e)
        {
            //交换奇数行颜色&偶数行颜色
            
            Color temp;
            temp = TableBackcolor1;
            TableBackcolor1 = TableBackcolor2;
            TableBackcolor2 = temp;

            SetTableLineColor();
        }

        private void NewColor_Click(object sender, EventArgs e)
        {
            //自定义颜色1
            DialogResult dr = colorDialog1.ShowDialog();
            if (dr == DialogResult.OK) TableBackcolor1 = colorDialog1.Color;

            SetTableLineColor();
        }

        private void NewColor2_Click(object sender, EventArgs e)
        {
            //自定义颜色2
            DialogResult dr = colorDialog1.ShowDialog();
            if (dr == DialogResult.OK) TableBackcolor2 = colorDialog1.Color;

            SetTableLineColor();
        }

        private void SavePreset_Click(object sender, EventArgs e)
        {
            //将颜色保存到预设文件
            JObject js = ImportJSON(PresetFile);

            string presetName = Interaction.InputBox("预设名称", "保存预设").ToString();

            string c1r = TableBackcolor1.R.ToString();
            string c1g = TableBackcolor1.G.ToString();
            string c1b = TableBackcolor1.B.ToString();
            string c2r = TableBackcolor2.R.ToString();
            string c2g = TableBackcolor2.G.ToString();
            string c2b = TableBackcolor2.B.ToString();
            string cHr = TableBackcolorH.R.ToString();
            string cHg = TableBackcolorH.G.ToString();
            string cHb = TableBackcolorH.B.ToString();

            if (presetName != "")
            {
                int num = int.Parse(js["num"].ToString());
                num = num + 1;
                js["num"] = num.ToString();
                js["preset_" + num + "_name"] = presetName;
                js["preset_" + num + "_TableBackcolor1_r"] = c1r;
                js["preset_" + num + "_TableBackcolor1_g"] = c1g;
                js["preset_" + num + "_TableBackcolor1_b"] = c1b;
                js["preset_" + num + "_TableBackcolor2_r"] = c2r;
                js["preset_" + num + "_TableBackcolor2_g"] = c2g;
                js["preset_" + num + "_TableBackcolor2_b"] = c2b;
                js["preset_" + num + "_TableBackcolorH_r"] = cHr;
                js["preset_" + num + "_TableBackcolorH_g"] = cHg;
                js["preset_" + num + "_TableBackcolorH_b"] = cHb;

                SetjsonFun(PresetFile, js);

                TableColorPresetList.Items.Clear();//刷新导入预设列表
                for (int i = 1; i <= num; i++)
                {
                    TableColorPresetList.Items.Add(js["preset_" + i + "_name"].ToString());
                }
            }
        }

        private void DeletePreset_Click(object sender, EventArgs e)
        {
            //删除预设
            JObject js = ImportJSON(PresetFile);

            int PresetSelected = TableColorPresetList.SelectedIndex + 1;
            int PresetCount = int.Parse(js["num"].ToString());

            if (PresetSelected != 0)
            {
                for (int i = PresetSelected; i < PresetCount; i++)
                {
                    int j = i + 1;
                    js["preset_" + i + "_name"] = js["preset_" + j + "_name"];
                    js["preset_" + i + "_TableBackcolor1_r"] = js["preset_" + j + "_TableBackcolor1_r"];
                    js["preset_" + i + "_TableBackcolor1_g"] = js["preset_" + j + "_TableBackcolor1_g"];
                    js["preset_" + i + "_TableBackcolor1_b"] = js["preset_" + j + "_TableBackcolor1_b"];
                    js["preset_" + i + "_TableBackcolor2_r"] = js["preset_" + j + "_TableBackcolor2_r"];
                    js["preset_" + i + "_TableBackcolor2_g"] = js["preset_" + j + "_TableBackcolor2_g"];
                    js["preset_" + i + "_TableBackcolor2_b"] = js["preset_" + j + "_TableBackcolor2_b"];
                    js["preset_" + i + "_TableBackcolorH_r"] = js["preset_" + j + "_TableBackcolorH_r"];
                    js["preset_" + i + "_TableBackcolorH_g"] = js["preset_" + j + "_TableBackcolorH_g"];
                    js["preset_" + i + "_TableBackcolorH_b"] = js["preset_" + j + "_TableBackcolorH_b"];
                }

                //delete json opt
                js["num"] = (PresetCount - 1).ToString();

                js["preset_" + PresetCount + "_name"] = "";
                js["preset_" + PresetCount + "_TableBackcolor1_r"] = "";
                js["preset_" + PresetCount + "_TableBackcolor1_g"] = "";
                js["preset_" + PresetCount + "_TableBackcolor1_b"] = "";
                js["preset_" + PresetCount + "_TableBackcolor2_r"] = "";
                js["preset_" + PresetCount + "_TableBackcolor2_g"] = "";
                js["preset_" + PresetCount + "_TableBackcolor2_b"] = "";
                js["preset_" + PresetCount + "_TableBackcolorH_r"] = "";
                js["preset_" + PresetCount + "_TableBackcolorH_g"] = "";
                js["preset_" + PresetCount + "_TableBackcolorH_b"] = "";

                SetjsonFun(PresetFile, js);

                //refresh preset list
                TableColorPresetList.Items.Clear();
                for (int i = 1; i < PresetCount; i++)
                {
                    TableColorPresetList.Items.Add(js["preset_" + i + "_name"].ToString());
                }

                //refresh choice and preview
                //TableColorPresetList.SelectedIndex = 0;             
                //TableBackcolor1 = ReadTHINGcolor(1, "TableBackcolor1");
                //TableBackcolor2 = ReadTHINGcolor(1, "TableBackcolor2");
                //SetTableLineColor();
            }
        }

        private void checkBoxFirstLine_CheckedChanged(object sender, EventArgs e)
        {
            // 排除首行
            SetTableLineTxt();

            SetTableLineColor();
        }

        private void Export_Click(object sender, EventArgs e)
        {
            //导出预设            
            string Path;
            saveFileDialogExPreset.Filter = "FDscend表格预设(*.fdtp)|*.fdtp";
            DialogResult dr = saveFileDialogExPreset.ShowDialog();
            if (dr == DialogResult.OK)
            {
                //记录选中的目录  
                Path = saveFileDialogExPreset.FileName;
                //MessageBox.Show(Path);
                if (File.Exists(PresetFile))//必须判断要复制的文件是否存在
                {
                    File.Copy(PresetFile, Path, true);//三个参数分别是源文件路径，存储路径，若存储路径有相同文件是否替换
                }
            }
        }

        string InportFile = "";//导入预设文件
        private void checkBoxInport_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxInport.Checked == true)
            {
                //导入预设
                openFileDialogINPreset.Title = "选择预设文件";
                openFileDialogINPreset.Filter = "FDscend预设(*.fdtp)|*.fdtp|ToolsBox_TablePreset(*.*)|*.*";
                DialogResult dr = openFileDialogINPreset.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    InportFile = openFileDialogINPreset.FileName;
                    //MessageBox.Show(InportFile);

                    //显示导入预设控件组
                    InportPresetGroup.Visible = true;

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
            else
            {
                InportPresetGroup.Visible = false;
            }
        }

        private void InportOK_Click(object sender, EventArgs e)
        {
            JObject js = ImportJSON(PresetFile);//原先预设
            JObject jsIn = ImportJSON(InportFile);//要导入的预设

            int InportSelected = InportPresetList.SelectedIndex + 1;
            int StartPresetCount = int.Parse(js["num"].ToString());

            if (InportSelected != 0)
            {
                js["preset_" + (StartPresetCount + 1) + "_name"] = jsIn["preset_" + InportSelected + "_name"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolor1_r"] = jsIn["preset_" + InportSelected + "_TableBackcolor1_r"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolor1_g"] = jsIn["preset_" + InportSelected + "_TableBackcolor1_g"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolor1_b"] = jsIn["preset_" + InportSelected + "_TableBackcolor1_b"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolor2_r"] = jsIn["preset_" + InportSelected + "_TableBackcolor2_r"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolor2_g"] = jsIn["preset_" + InportSelected + "_TableBackcolor2_g"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolor2_b"] = jsIn["preset_" + InportSelected + "_TableBackcolor2_b"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolorH_r"] = jsIn["preset_" + InportSelected + "_TableBackcolorH_r"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolorH_g"] = jsIn["preset_" + InportSelected + "_TableBackcolorH_g"];
                js["preset_" + (StartPresetCount + 1) + "_TableBackcolorH_b"] = jsIn["preset_" + InportSelected + "_TableBackcolorH_b"];
                js["num"] = (StartPresetCount + 1).ToString();

                SetjsonFun(PresetFile, js);

                //刷新本来预设列表
                TableColorPresetList.Items.Clear();
                for (int i = 1; i <= StartPresetCount + 1; i++)
                {
                    TableColorPresetList.Items.Add(js["preset_" + i + "_name"].ToString());
                }

                //刷新导入预设列表
                //InportPresetList.Items.Clear();
                InportPresetList.Items.Remove(InportPresetList.SelectedItem);
                InportPresetList.Refresh();
            }
        }

        private void NewColorH_Click(object sender, EventArgs e)
        {
            //自定义颜色3
            DialogResult dr = colorDialog1.ShowDialog();
            if (dr == DialogResult.OK) TableBackcolorH = colorDialog1.Color;

            SetTableLineColor();
        }
    }
}
