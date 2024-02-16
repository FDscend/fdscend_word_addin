using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;//MessageBox,ColorDialog
using Microsoft.VisualBasic;//inputbox
using Microsoft.Win32;//registry
using System.Drawing;//color
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;  //正则表达式
using System.Diagnostics;
using System.Threading.Tasks;
using System.Net.NetworkInformation;


namespace WordAddIn1
{
    public partial class Ribbon1
    {
        Word.Application app;


        //全局路径
#if DEBUG
        public static string FDscendHome = "D:\\code\\WordAddIn1\\FDscend";
#endif
#if !DEBUG
        public static string FDscendHome = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\分点作答\\FDscend";
#endif
        public static string ControlKey = FDscendHome + Properties.Resources.ControlKey;
        public static string CheckUpdateFile = FDscendHome + Properties.Resources.CheckUpdateFile;
        public static string SettingsFile = FDscendHome + Properties.Resources.SettingsFile;

        public static string ChangeCharList = FDscendHome + Properties.Resources.PresetChangeCharList;
        public static string CharMatchList = FDscendHome + Properties.Resources.PresetCharMatchList;
        public static string PresetCodeFile = FDscendHome + Properties.Resources.PresetCodeFile;
        public static string PresetCodeLatexFile = FDscendHome + Properties.Resources.PresetCodeLatexFile;
        public static string PresetCodeMDFile = FDscendHome + Properties.Resources.PresetCodeMDFile;
        public static string PresetCode4File = FDscendHome + Properties.Resources.PresetCode4File;
        public static string PresetToolsBoxShadeColor = FDscendHome + Properties.Resources.PresetToolsBoxShadeColor;
        public static string PresetToolsBoxTable = FDscendHome + Properties.Resources.PresetToolsBoxTable;
        public static string XMTsetting = FDscendHome + Properties.Resources.PresetXMTsetting;
        public static string XMTstyle = FDscendHome + Properties.Resources.PresetXMTstyle;
        public static string highlight_index = FDscendHome + Properties.Resources.highlight_index;

        public static string tempFile = FDscendHome + "\\temp";
        public static string scriptsDic = FDscendHome + "\\scripts";

        public static string latest_info = FDscendHome + "\\latest.json";
        public static string temp_websource_htm = FDscendHome + "\\temp\\temp.htm";
        public static string temp_pic_path = FDscendHome + "\\temp\\temp_pic_path";

        public static string pdf_path = FDscendHome + "\\说明文档.pdf";



        //全局常量
        public const string KeyDefault = "default";
        public const string KeyXMT = "xmt";
        public const string KeyCode = "code";
        public const string KeyCode2 = "codeB";
        public const string KeyCodeLatex = "codeL";
        public const string KeyCode4 = "code4";
        public const string KeyToolsBox = "tools";
        public const string KeyRunCode = "runCode";
        public const string KeyAdmin = "admin";


        //全局变量        
        Color CodeBackcolor1 = Color.FromArgb(226, 239, 217);//代码背景颜色
        Color CodeBackcolor2 = Color.White;//代码背景颜色
        Color CodeBorderLine = Color.FromArgb(112, 173, 71);//代码竖线颜色
        Color CodeTxtColor = Color.FromArgb(165, 165, 165);
        //string CodeFontName = "新宋体";
        //double CodeFontSize = 9.5;
        Word.WdColor codeMD_CurrentColor;
        Word.WdColor code4_CurrentColor;

        ColorDialog MyColorDialog = new ColorDialog();

        string codeChoice = "python";
        static int codeOutputCount = 0;
        static int codeOutputTotal = 0;

        static int gbtOutputCount = 1;


        //窗体句柄字典
        private Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> HwndPaneDic_tab = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> { };
        private Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> HwndPaneDic_tableColor = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> { };
        private Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> HwndPaneDic_ChangeChar = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> { };
        private Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> HwndPaneDic_CharMatch = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> { };
        private Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> HwndPaneDic_highlight = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> { };

        //侧面板
        Microsoft.Office.Tools.CustomTaskPane FileTabPane;
        Microsoft.Office.Tools.CustomTaskPane tableColoringPane;
        Microsoft.Office.Tools.CustomTaskPane changCharPane;
        Microsoft.Office.Tools.CustomTaskPane charMatchPane;
        Microsoft.Office.Tools.CustomTaskPane highlightPane;

        // 样式是否存在 int 默认 0
        int maintitle_bool;
        int title_1_bool;
        int title_2_bool;
        int maintext_bool;
        int codeMD_bool;

        int paraShandingChoice = 1;
        int styleShadingChoice = 1;

        //标签栏
        static int fileTabPaneHeight = 83;
        public List<string> DocNamesList = new List<string>();


        public static JObject ImportJSON(string jsonfile)
        {
            StreamReader reader = File.OpenText(jsonfile);
            JsonTextReader jsonTextReader = new JsonTextReader(reader);
            JObject jsonObject = (JObject)JToken.ReadFrom(jsonTextReader);
            reader.Close();
            return jsonObject;
        }

        public static string ReadjsonFun(string jsonfile, string key)
        {
            //StreamReader reader = File.OpenText(jsonfile);
            //JsonTextReader jsonTextReader = new JsonTextReader(reader);
            //JObject jsonObject = (JObject)JToken.ReadFrom(jsonTextReader);
            //reader.Close();
            JObject jsonObject = ImportJSON(jsonfile);
            string value = jsonObject[key].ToString();           
            return value;
        }

        public static void SetjsonFun(string jsonfile, JObject jsonObject)
        {
            /*StreamReader reader = File.OpenText(jsonfile);
            JsonTextReader jsonTextReader = new JsonTextReader(reader);
            JObject jsonObject = (JObject)JToken.ReadFrom(jsonTextReader);
            jsonObject[key] = value;
            reader.Close();*/
            string output = Newtonsoft.Json.JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(jsonfile, output);
        }

        //暂时用不到
        public static void AddjsonFun(string jsonfile, JObject JStemp)
        {
            //string JsonString = File.ReadAllText(jsonfile);
            StreamReader reader = File.OpenText(jsonfile);
            JsonTextReader jsonTextReader = new JsonTextReader(reader);
            JObject jobject = (JObject)JToken.ReadFrom(jsonTextReader);
            //JObject jobject = JObject.Parse(JsonString);
            //JObject newkey = new JObject(new JProperty(key, value));
            jobject.Last.AddAfterSelf(JStemp);
            //jobject.Add(newkey);
            string convertString = Convert.ToString(jobject);
            File.WriteAllText(jsonfile, convertString);
        }

        private bool BoolKey(string key)
        {
            switch (key)
            {              
                case KeyDefault:
                    group_tuisong.Visible = false;
                    code.Visible = false;
                    break;
                case KeyXMT:
                    group_tuisong.Visible = true;
                    break;
                case KeyCode:
                    code.Visible = true;
                    break;
                case KeyCode2:
                    Code2.Visible = true;
                    break;
                case KeyCodeLatex:
                    CodeLatex.Visible = true;
                    break;
                case KeyCode4:
                    CodeGroup4.Visible = true;
                    break;
                case KeyToolsBox:
                    ToolsBox.Visible = true;
                    break;
                case KeyRunCode:
                    runCodeGroup.Visible = true;
                    break;
                case KeyAdmin:
                    KeyAllTrue();
                    break;
                default:
                    return false;
            }
            return true;
        }

        private bool BoolKeyState(string key, string jsonfile)
        {
            switch (key)
            {
                case KeyDefault:
                    //group_tuisong.Visible = false;
                    //code.Visible = false;
                    break;
                case KeyXMT:
                    if (ReadjsonFun(jsonfile, "state_"+ KeyXMT) == "1")
                    { 
                        group_tuisong.Visible = true;
                        return true;
                    }
                    else
                        group_tuisong.Visible = false;
                    break;
                case KeyCode:
                    if (ReadjsonFun(jsonfile, "state_"+ KeyCode) == "1")
                    {
                        code.Visible = true;
                        return true;
                    }
                    else
                        code.Visible = false;
                    break;
                case KeyCode2:
                    if (ReadjsonFun(jsonfile, "state_" + KeyCode2) == "1")
                    {
                        Code2.Visible = true;
                        return true;
                    }
                    else
                        Code2.Visible = false;
                    break;
                case KeyCodeLatex:
                    if (ReadjsonFun(jsonfile, "state_" + KeyCodeLatex) == "1")
                    {
                        CodeLatex.Visible = true;
                        return true;
                    }
                    else
                        CodeLatex.Visible = false;
                    break;
                case KeyCode4:
                    if (ReadjsonFun(jsonfile, "state_" + KeyCode4) == "1")
                    {
                        CodeGroup4.Visible = true;
                        return true;
                    }
                    else
                        CodeGroup4.Visible = false;
                    break;
                case KeyToolsBox:
                    if (ReadjsonFun(jsonfile, "state_" + KeyToolsBox) == "1")
                    {
                        ToolsBox.Visible = true;
                        return true;
                    }
                    else
                        ToolsBox.Visible = false;
                    break;
                case KeyRunCode:
                    if (ReadjsonFun(jsonfile, "state_" + KeyToolsBox) == "1")
                    {
                        runCodeGroup.Visible = true;
                        return true;
                    }
                    else
                        runCodeGroup.Visible = false;
                    break;
                case KeyAdmin:
                    if (ReadjsonFun(jsonfile, "state_"+KeyAdmin) == "1")
                    {
                        KeyAllTrue();
                        return true;
                    }
                    break;
                default:
                    return false;
            }
            return true;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;

            //重置临时文件夹
            if (Directory.Exists(tempFile))
            {
                Directory.Delete(tempFile, true);
                Directory.CreateDirectory(tempFile);
            }
            else Directory.CreateDirectory(tempFile);



            //颜色对话框自定义颜色集
            MyColorDialog.CustomColors = new int[] { 14282722, 13684944, 13298939, 14869500 };

            //
#if DEBUG
            KeyAllTrue();
            //KeyStateLoad();
#endif
#if !DEBUG
            KeyStateLoad();
            button_tuisong.Visible = false;
            //FileTabOnOff.Visible = false;
#endif

            //代码相关颜色初始化
            JObject js = ImportJSON(PresetCodeFile);
            CodeBackcolor1 = ReadCodeTHINGcolor(int.Parse(js["DefaultPreset"].ToString()), "CodeBackcolor1");
            CodeBackcolor2 = ReadCodeTHINGcolor(int.Parse(js["DefaultPreset"].ToString()), "CodeBackcolor2");
            CodeBorderLine = ReadCodeTHINGcolor(int.Parse(js["DefaultPreset"].ToString()), "CodeBorderLine");

            JObject js_codeMD = ImportJSON(PresetCodeMDFile);
            int r = int.Parse(js_codeMD["ParaShadingColor_r"].ToString());
            int g = int.Parse(js_codeMD["ParaShadingColor_g"].ToString());
            int b = int.Parse(js_codeMD["ParaShadingColor_b"].ToString());
            codeMD_CurrentColor = GetColor(Color.FromArgb(r, g, b));

            JObject js_code4 = ImportJSON(PresetCode4File);
            r = int.Parse(js_code4["DefaultShadingColor_r"].ToString());
            g = int.Parse(js_code4["DefaultShadingColor_g"].ToString());
            b = int.Parse(js_code4["DefaultShadingColor_b"].ToString());
            code4_CurrentColor = GetColor(Color.FromArgb(r, g, b));


            // 文案
            JObject js_xmt = ImportJSON(XMTsetting);
            string authorChoice = js_xmt["author"].ToString();
            if (authorChoice == "1") author.Image = Properties.Resources.office;
            if (authorChoice == "2") author.Image = Properties.Resources.wps;
            if (authorChoice == "3") author.Image = Properties.Resources.document;
            if (authorChoice == "4") author.Image = Properties.Resources.computer;
            if (authorChoice == "5")
            {
                if (js_xmt["author_new"].ToString() == "") author.Image = Properties.Resources.add;
                else
                {
                    author.Image = Properties.Resources.署名;
                    author_new.Image = Properties.Resources.署名;
                    author_new.Label = js_xmt["author_new"].ToString();
                }
            }


            //check update
            JObject js_update = ImportJSON(CheckUpdateFile);
            if (js_update["auto_check"].ToString() == "yes")
            {
                if (GetNetStatus("api.github.com")) AutoUpdate();
            }

        }

        bool GetNetStatus(string ping_url)
        {
            Ping ping = new Ping();
            try
            {
                PingReply pingReply = ping.Send(ping_url);
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


        void AutoUpdate()
        {
            //check update

            JObject js = ImportJSON(CheckUpdateFile);
            string strTime = js["latest_date"].ToString();
            double fre_day = double.Parse(js["delay_day"].ToString());

            DateTime dt_latest = Convert.ToDateTime(strTime);
            DateTime dt_check = dt_latest.AddDays(fre_day);

            DateTime dt_now = DateTime.Now;

            if (dt_check < dt_now)  //检查上次更新时间
            {
                //检查版本
                if (File.Exists(latest_info))
                {
                    File.Delete(latest_info);
                }

                string url = Properties.Resources.latest_info_url;
                WebClient webClient = new WebClient();
                webClient.Headers.Add("user-agent", "Mozilla/4.0 ((compatible; MSIE 8.0; Windows NT 6.1;.NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729;)");

                ServicePointManager.SecurityProtocol = (SecurityProtocolType)48
                                                | (SecurityProtocolType)192
                                                | (SecurityProtocolType)768
                                                | (SecurityProtocolType)3072;

                webClient.DownloadFile(url, latest_info);

                JObject js_latest = ImportJSON(latest_info);
                Version ver_cur = new Version(Properties.Resources.current_ver);
                string latest_version = js_latest["tag_name"].ToString().Substring(1);
                Version ver_latest = new Version(latest_version);

                if (ver_cur < ver_latest)
                {
                    DialogResult dr = MessageBox.Show("插件更新啦，去看看吧！\r\n\r\n当前版本：" + Properties.Resources.current_ver + "\r\n最新版本：" + latest_version, "分点作答", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign, false);

                    if (dr == DialogResult.OK)
                    {
                        System.Diagnostics.Process.Start(Properties.Resources.github_code_url);
                    }
                }
                js["latest_date"] = dt_now.ToString("d");
                SetjsonFun(CheckUpdateFile, js);
            }
        }


        void KeyAllTrue()
        {
            group_tuisong.Visible = true;
            code.Visible = true;
            Code2.Visible = true;
            CodeLatex.Visible = true;
            CodeGroup4.Visible = true;
            ToolsBox.Visible = true;
            runCodeGroup.Visible = true;
        }

        public void KeyStateLoad()
        {
            //string key;
            JObject jstemp = ImportJSON(ControlKey);

            foreach (JProperty jp in jstemp.Properties())
            {
                string key_temp = jp.Value.ToString();
                BoolKeyState(key_temp, ControlKey);
            }
        }

        private void button_tuisong_Click(object sender, RibbonControlEventArgs e)
        {
            //插入文本
            //Globals.ThisAddIn.Application.Selection; //获取光标位置
            //int insert_start = Globals.ThisAddIn.Application.Selection.Start;
            //int insert_end = Globals.ThisAddIn.Application.Selection.End;
            //Word.Range rng = app.Application.ActiveDocument.Range(insert_start, insert_end);
            //rng.Text = "分点作答，一键排版！";
            //rng.Select();


            button_header_Click(sender, e);//设置页眉
            button_footer_Click(sender, e);//设置页脚
            button_font_Click(sender, e);//替换全文字体


            //清除全文部分格式属性
            app.Application.ActiveDocument.Content.Bold = 0;
            //app.Application.ActiveDocument.Content.HighlightColorIndex = 0;
            app.Application.ActiveDocument.Content.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            app.Application.ActiveDocument.Content.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;//单倍行间距
            //app.Application.ActiveDocument.Content.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
            //app.Application.ActiveDocument.Content.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
            //app.Application.ActiveDocument.Content.ParagraphFormat.LineSpacing = (float)1.25;//行间距
            app.Application.ActiveDocument.Content.ParagraphFormat.SpaceBefore= float.Parse("0");//段前间距
            app.Application.ActiveDocument.Content.ParagraphFormat.SpaceAfter= float.Parse("0");//段后间距           
            
            

            Word.Document myDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs paragraphs = myDoc.Paragraphs;
            int numberOfParagraphs = paragraphs.Count;            
            for(int i=1;i<=numberOfParagraphs;i++)
            {
                if((int)paragraphs[i].OutlineLevel==1)
                {
                    //换竖线
                    object Replace_String = "|";       //要替换的字符
                    object ms = System.Type.Missing;
                    object Replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;//设置替换方式:一，全部替换；二，只替换一个；三，一个都不替换。
                    object ReplaceWith = "｜";             //最终替换成的字符
                                                          //执行Word自带的查找/替换功能函数
                    app.Application.ActiveDocument.Content.Find.Execute(ref Replace_String, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ReplaceWith, ref Replace, ref ms, ref ms, ref ms, ref ms);


                    paragraphs[i].Range.Font.Bold = 1;
                    paragraphs[i].Range.Font.Color = Word.WdColor.wdColorBlack;
                    paragraphs[i].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraphs[i].Range.ParagraphFormat.SpaceAfter = float.Parse("5");
                }
                if ((int)paragraphs[i].OutlineLevel == 2)
                {
                    paragraphs[i].Range.Font.Bold = 1;
                    paragraphs[i].Range.Font.Color = Word.WdColor.wdColorBlack;
                    paragraphs[i].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraphs[i].Range.ParagraphFormat.SpaceBefore = float.Parse("5");
                }
                if ((int)paragraphs[i].OutlineLevel == 3)
                {
                    paragraphs[i].Range.Font.Bold = 1;
                    paragraphs[i].Range.Font.Color = Word.WdColor.wdColorBlack;
                    paragraphs[i].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                if ((int)paragraphs[i].OutlineLevel == 10)
                {
                    if (paragraphs[i].Range.ParagraphFormat.CharacterUnitFirstLineIndent == float.Parse("2") || paragraphs[i].Range.ParagraphFormat.FirstLineIndent == paragraphs[i].Range.Font.Size * 2)
                    {
                        ;//如果已经缩进了就不管了
                    }
                    else
                    {
                        //paragraphs[i].Range.ParagraphFormat.FirstLineIndent = app.CentimetersToPoints(float.Parse("0"));
                        paragraphs[i].Range.ParagraphFormat.IndentFirstLineCharWidth(2);
                    }
                }
            }

            author_doc_Click(sender, e);

        }

        private void button_header_Click(object sender, RibbonControlEventArgs e)
        {
            //设置页眉
            JObject js = ImportJSON(XMTsetting);

            int nullkey = 1;
            string header_text = js["header_text"].ToString();

            if (header_text == "")
            {
                string key = Interaction.InputBox("输入自定义页眉", "页眉设置").ToString();
                if (key == "")
                {
                    nullkey = 0;
                }
                else
                {
                    js["header_text"] = key;
                    header_text = key;
                    SetjsonFun(XMTsetting, js);
                }
            }


            if (nullkey == 1)
            {
                foreach (Word.Section section in app.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.Name = js["heade_font"].ToString();
                    headerRange.Text = header_text;
                }
            }
            
        }

        private void button_footer_Click(object sender, RibbonControlEventArgs e)
        {
            //设置页脚
            JObject js = ImportJSON(XMTsetting);

            int nullkey = 1;
            string footer_text = js["footer_text"].ToString();

            if (footer_text == "")
            {
                string key = Interaction.InputBox("输入自定义页脚", "页脚设置").ToString();
                if (key == "")
                {
                    nullkey = 0;
                }
                else
                {
                    js["footer_text"] = key;
                    footer_text = key;
                    SetjsonFun(XMTsetting, js);
                }
            }


            if (nullkey == 1)
            {
                foreach (Word.Section wordSection in app.Application.ActiveDocument.Sections)
                {
                    Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    //footerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    //footerRange.Font.Size = 11;
                    footerRange.Font.Name = js["footer_font"].ToString();
                    footerRange.Text = footer_text;
                }
            }

        }

        private void button_font_Click(object sender, RibbonControlEventArgs e)
        {
            //替换全文字体
            // Set the Range to the first paragraph. 
            //Word.Document document = app.Application.ActiveDocument;
            //Word.Range rng = document.Paragraphs[1].Range;
            Word.Range rng = app.Application.ActiveDocument.Content;

            rng.Font.Size = 12;
            rng.Font.Name = "宋体";

            rng.Select();
            //MessageBox.Show("Formatted Range");

        }

        private void button_MainTitle_Click(object sender, RibbonControlEventArgs e)
        {
            //大标题
            Globals.ThisAddIn.Application.Selection.ClearFormatting();

            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
            {
                //MessageBox.Show(Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal);
                if(Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal== "标题_XMT")
                {
                    maintitle_bool = 1;
                    break;
                }
            }

            if (maintitle_bool == 0)
            {
                // CreatMainTitleStyle();
                CopyStyle(XMTstyle, "标题_XMT");

                maintitle_bool = 1;
            }
            
            Globals.ThisAddIn.Application.Selection.set_Style("标题_XMT");

            //换竖线
            object Replace_String = "|";       //要替换的字符
            object ms = System.Type.Missing;
            object Replace = Word.WdReplace.wdReplaceAll;//设置替换方式:一，全部替换；二，只替换一个；三，一个都不替换。
            object ReplaceWith = "｜";             //最终替换成的字符
            //执行Word自带的查找/替换功能函数
            Globals.ThisAddIn.Application.Selection.Find.Execute(ref Replace_String, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ReplaceWith, ref Replace, ref ms, ref ms, ref ms, ref ms);

            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Microsoft.Office.Core.DocumentProperties)app.Application.ActiveDocument.BuiltInDocumentProperties;
            //string title = properties["Title"].Value;

            properties["Title"].Value = Globals.ThisAddIn.Application.Selection.Paragraphs[1].Range.Text;
            //MessageBox.Show(Globals.ThisAddIn.Application.Selection.Paragraphs[1].Range.Text);
        }

        void CopyStyle(string source, string styleName)
        {
            string activeDocPath = app.ActiveDocument.Path + "\\" + app.ActiveDocument.Name;
            Globals.ThisAddIn.Application.OrganizerCopy(source, activeDocPath, styleName, Word.WdOrganizerObject.wdOrganizerObjectStyles);
        }

        void CreatMainTitleStyle()
        {
            Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("标题_XMT", Word.WdStyleType.wdStyleTypeParagraph);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].AutomaticallyUpdate = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.NameFarEast = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.NameAscii = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.NameOther = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Name = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Size = 14;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Bold = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Italic = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.StrikeThrough = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Outline = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Emboss = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Shadow = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Hidden = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.SmallCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.AllCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Color = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Engrave = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Superscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Subscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Scaling = 100;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Kerning = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Animation = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.DisableCharacterSpaceGrid = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.EmphasisMark = Word.WdEmphasisMark.wdEmphasisMarkNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.Ligatures = Word.WdLigatures.wdLigaturesNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.NumberSpacing = Word.WdNumberSpacing.wdNumberSpacingDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.NumberForm = Word.WdNumberForm.wdNumberFormDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.StylisticSet = Word.WdStylisticSet.wdStylisticSetDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Font.ContextualAlternates = 0;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.LeftIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.RightIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.SpaceBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.SpaceBeforeAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.SpaceAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.SpaceAfterAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.WidowControl = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.KeepWithNext = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.KeepTogether = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.PageBreakBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.NoLineNumber = 0;
            //Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Hyphenation = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.FirstLineIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.CharacterUnitLeftIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.CharacterUnitRightIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.LineUnitBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.LineUnitAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.MirrorIndents = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.CollapsedByDefault = 0;
            //Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.AutoAdjustRightIndent = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.DisableLineHeightGrid = 0;
            //Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.FarEastLineBreakControl = 1;
            //Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.WordWrap = 1;
            //Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.HangingPunctuation = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.HalfWidthPunctuationOnTopOfLine = 0;
            //Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.AddSpaceBetweenFarEastAndAlpha = 1;
            //Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.AddSpaceBetweenFarEastAndDigit = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.TabStops.ClearAll();

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders.DistanceFromTop = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders.DistanceFromRight = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders.DistanceFromLeft = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].ParagraphFormat.Borders.DistanceFromBottom = 1;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].set_BaseStyle("");

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题_XMT"].Frame.Delete();
        }

        void CreatTitle1Style()
        {
            Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("标题1_XMT", Word.WdStyleType.wdStyleTypeParagraph);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].AutomaticallyUpdate = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.NameFarEast = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.NameAscii = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.NameOther = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Name = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Size = (float)10.5;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Bold = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Italic = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.StrikeThrough = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Outline = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Emboss = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Shadow = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Hidden = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.SmallCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.AllCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Color = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Engrave = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Superscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Subscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Scaling = 100;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Kerning = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Animation = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.DisableCharacterSpaceGrid = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.EmphasisMark = Word.WdEmphasisMark.wdEmphasisMarkNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.Ligatures = Word.WdLigatures.wdLigaturesNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.NumberSpacing = Word.WdNumberSpacing.wdNumberSpacingDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.NumberForm = Word.WdNumberForm.wdNumberFormDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.StylisticSet = Word.WdStylisticSet.wdStylisticSetDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Font.ContextualAlternates = 0;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.LeftIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.RightIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.SpaceBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.SpaceBeforeAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.SpaceAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.SpaceAfterAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.WidowControl = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.KeepWithNext = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.KeepTogether = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.PageBreakBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.NoLineNumber = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.FirstLineIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.CharacterUnitLeftIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.CharacterUnitRightIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.LineUnitBefore = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.LineUnitAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.MirrorIndents = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.CollapsedByDefault = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.DisableLineHeightGrid = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.HalfWidthPunctuationOnTopOfLine = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.TabStops.ClearAll();

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders.DistanceFromTop = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders.DistanceFromRight = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders.DistanceFromLeft = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].ParagraphFormat.Borders.DistanceFromBottom = 1;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].set_BaseStyle("");

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题1_XMT"].Frame.Delete();
        }

        void CreatTitle2Style()
        {
            Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("标题2_XMT", Word.WdStyleType.wdStyleTypeParagraph);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].AutomaticallyUpdate = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.NameFarEast = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.NameAscii = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.NameOther = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Name = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Size = (float)10.5;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Bold = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Italic = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.StrikeThrough = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Outline = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Emboss = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Shadow = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Hidden = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.SmallCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.AllCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Color = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Engrave = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Superscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Subscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Scaling = 100;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Kerning = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Animation = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.DisableCharacterSpaceGrid = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.EmphasisMark = Word.WdEmphasisMark.wdEmphasisMarkNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.Ligatures = Word.WdLigatures.wdLigaturesNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.NumberSpacing = Word.WdNumberSpacing.wdNumberSpacingDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.NumberForm = Word.WdNumberForm.wdNumberFormDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.StylisticSet = Word.WdStylisticSet.wdStylisticSetDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Font.ContextualAlternates = 0;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.LeftIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.RightIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.SpaceBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.SpaceBeforeAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.SpaceAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.SpaceAfterAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.WidowControl = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.KeepWithNext = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.KeepTogether = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.PageBreakBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.NoLineNumber = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.FirstLineIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel2;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.CharacterUnitLeftIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.CharacterUnitRightIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.LineUnitBefore = (float)0.5;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.LineUnitAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.MirrorIndents = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.CollapsedByDefault = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.DisableLineHeightGrid = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.HalfWidthPunctuationOnTopOfLine = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.TabStops.ClearAll();

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders.DistanceFromTop = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders.DistanceFromRight = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders.DistanceFromLeft = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].ParagraphFormat.Borders.DistanceFromBottom = 1;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].set_BaseStyle("");

            Globals.ThisAddIn.Application.ActiveDocument.Styles["标题2_XMT"].Frame.Delete();
        }

        void CreatMainTextStyle()
        {
            Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("正文_XMT", Word.WdStyleType.wdStyleTypeParagraph);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].AutomaticallyUpdate = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.NameFarEast = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.NameAscii = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.NameOther = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Name = "宋体";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Size = (float)10.5;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Bold = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Italic = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.StrikeThrough = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Outline = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Emboss = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Shadow = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Hidden = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.SmallCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.AllCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Color = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Engrave = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Superscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Subscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Scaling = 100;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Kerning = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Animation = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.DisableCharacterSpaceGrid = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.EmphasisMark = Word.WdEmphasisMark.wdEmphasisMarkNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Ligatures = Word.WdLigatures.wdLigaturesNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.NumberSpacing = Word.WdNumberSpacing.wdNumberSpacingDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.NumberForm = Word.WdNumberForm.wdNumberFormDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.StylisticSet = Word.WdStylisticSet.wdStylisticSetDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.ContextualAlternates = 0;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.LeftIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.RightIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.SpaceBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.SpaceBeforeAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.SpaceAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.SpaceAfterAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.WidowControl = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.KeepWithNext = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.KeepTogether = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.PageBreakBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.NoLineNumber = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.FirstLineIndent = Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Font.Size * 2;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.CharacterUnitLeftIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.CharacterUnitRightIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.LineUnitBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.LineUnitAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.MirrorIndents = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.CollapsedByDefault = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.DisableLineHeightGrid = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.HalfWidthPunctuationOnTopOfLine = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.TabStops.ClearAll();

            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders.DistanceFromTop = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders.DistanceFromRight = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders.DistanceFromLeft = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].ParagraphFormat.Borders.DistanceFromBottom = 1;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].set_BaseStyle("");

            Globals.ThisAddIn.Application.ActiveDocument.Styles["正文_XMT"].Frame.Delete();
        }

        private void button_title_1_Click(object sender, RibbonControlEventArgs e)
        {
            //标题1
            Globals.ThisAddIn.Application.Selection.ClearFormatting();

            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
            {
                //MessageBox.Show(Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal);
                if (Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal == "标题1_XMT")
                {
                    title_1_bool = 1;
                    break;
                }
            }

            if (title_1_bool == 0)
            {
                // CreatTitle1Style();
                CopyStyle(XMTstyle, "标题1_XMT");

                title_1_bool = 1;
            }

            Globals.ThisAddIn.Application.Selection.set_Style("标题1_XMT");
        }

        private void button_title_2_Click(object sender, RibbonControlEventArgs e)
        {
            //标题2
            Globals.ThisAddIn.Application.Selection.ClearFormatting();

            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
            {
                //MessageBox.Show(Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal);
                if (Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal == "标题2_XMT")
                {
                    title_2_bool = 1;
                    break;
                }
            }

            if (title_2_bool == 0)
            {
                //CreatTitle2Style();
                CopyStyle(XMTstyle, "标题2_XMT");

                title_2_bool = 1;
            }

            Globals.ThisAddIn.Application.Selection.set_Style("标题2_XMT");
        }

        private void button_MainText_Click(object sender, RibbonControlEventArgs e)
        {
            //正文
            Globals.ThisAddIn.Application.Selection.ClearFormatting();

            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
            {
                //MessageBox.Show(Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal);
                if (Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal == "正文_XMT")
                {
                    maintext_bool = 1;
                    break;
                }
            }

            if (maintext_bool == 0)
            {
                // CreatMainTextStyle();
                CopyStyle(XMTstyle, "正文_XMT");

                maintext_bool = 1;
            }

            Globals.ThisAddIn.Application.Selection.set_Style("正文_XMT");

            // 如果在列表中创建样式，无法缩进
            //Globals.ThisAddIn.Application.Selection.ParagraphFormat.FirstLineIndent = Globals.ThisAddIn.Application.Selection.Font.Size * 2;
        }

        private void button_clearFont_Click(object sender, RibbonControlEventArgs e)
        {
            //清除格式
            
            Globals.ThisAddIn.Application.Selection.ClearFormatting();
        }

        
        public string CodeFontFunN()
        {
            JObject js = ImportJSON(PresetCodeFile);
            return js["DefaultFont"].ToString();           
        }
        public double CodeFontFunS()
        {
            JObject js = ImportJSON(PresetCodeFile);
            return double.Parse(js["DefaultFontSize"].ToString());
        }

        public bool CodeBorderYN()
        {
            JObject js = ImportJSON(PresetCodeFile);
            if (js["CodeBorder"].ToString() == "yes")
                return true;
            else
                return false;
        }

        public float CodeNoWidth()
        {
            JObject js = ImportJSON(PresetCodeFile);
            return float.Parse(js["DefaultNoWidth"].ToString());
        }

        public bool CodeFullPageWidth()
        {
            JObject js = ImportJSON(PresetCodeFile);
            if (js["FullPageWidth"].ToString() == "yes")
                return true;
            else
                return false;
        }

        private void CodeFormat_Click(object sender, RibbonControlEventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);
            string CodeListNumStatus = js["CodeListNum"].ToString();

            Globals.ThisAddIn.Application.Selection.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;


            wordReplace("^l", "^p");

            //Globals.ThisAddIn.Application.Selection.ConvertToTable(Separator: Word.WdTableFieldSeparator.wdSeparateByParagraphs, NumColumns: 1, NumRows: 10, AutoFitBehavior: Word.WdAutoFitBehavior.wdAutoFitFixed);            
            Word.Table table = Globals.ThisAddIn.Application.Selection.ConvertToTable(Separator: Word.WdTableFieldSeparator.wdSeparateByParagraphs);
            //table.Borders.Enable = 5;
            table.Range.Font.Name = CodeFontFunN();
            table.Range.Font.Size = (float)CodeFontFunS();


            float width1 = CodeNoWidth();
            float width2 = Globals.ThisAddIn.Application.Selection.PageSetup.PageWidth - Globals.ThisAddIn.Application.Selection.PageSetup.LeftMargin - Globals.ThisAddIn.Application.Selection.PageSetup.RightMargin;
            

            int row = table.Range.Rows.Count;

            if (CodeListNumStatus == "yes")
            {
                table.Columns.Add(table.Columns[1]);
                //table.Columns[1].Width = app.CentimetersToPoints(float.Parse("0.8"));
                table.Columns[1].Width = width1;
                if(CodeFullPageWidth()==false)
                    table.Columns[2].AutoFit();
                else
                    table.Columns[2].Width = width2 - width1;


                for (int i = 1; i <= row; i++)
                {
                    table.Cell(i, 1).Range.Text = "" + i;
                    table.Cell(i, 1).Range.Font.Color = GetColor(CodeTxtColor);
                }

                //代码竖线
                if(CodeBorderYN()==true)
                {
                    table.Columns[1].Borders[Word.WdBorderType.wdBorderRight].Visible = true;
                    table.Columns[1].Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    table.Columns[1].Borders[Word.WdBorderType.wdBorderRight].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    table.Columns[1].Borders[Word.WdBorderType.wdBorderRight].Color = GetColor(CodeBorderLine);                    
                }                
            }
            if (CodeBorderYN() == false)
                table.Columns[1].Borders[Word.WdBorderType.wdBorderRight].Visible = false;


            for (int i=1;i<=row;i++)
            {                               
                if (i % 2 == 1)
                {
                    //table.Rows[i].Shading.BackgroundPatternColor = Word.WdColor.wdColorViolet;
                    //Color c1 = Color.FromArgb(226, 239, 217);
                    Color c1 = CodeBackcolor1;
                    table.Rows[i].Shading.BackgroundPatternColor = GetColor(c1);
                }
                else
                {
                    //table.Rows[i].Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    Color c2 = CodeBackcolor2;
                    table.Rows[i].Shading.BackgroundPatternColor = GetColor(c2);
                }
            }            
        }
        
        private Word.WdColor GetColor(Color c)
        {
            //color
            UInt32 R = 0x1, G = 0x100, B = 0x10000;
            return (Word.WdColor)(R * c.R + G * c.G + B * c.B);
        }

       
        private bool Registry_bool(string sPath, string sFile)
        {
            //判断注册表项是否存在
            RegistryKey subKey = Registry.CurrentUser.OpenSubKey(name: sPath);
            string[] keyNames = subKey.GetSubKeyNames();
            subKey.Close();
            foreach (string keyName in keyNames)
            {
                if (keyName == sFile)
                    return true;               
            }
            return false;
        }
        

        private void author_o_Click(object sender, RibbonControlEventArgs e)
        {
            if (Registry_bool("SOFTWARE\\Microsoft", "Office") == true)
            {   //office用户

                JObject js = ImportJSON(XMTsetting);
                js["author"] = 1;
                SetjsonFun(XMTsetting, js);
                author.Image = Properties.Resources.office;

                string userName_office;
                RegistryKey security;
                security = Registry.CurrentUser.OpenSubKey(name: "SOFTWARE\\Microsoft\\Office\\Common\\UserInfo", writable: true);
                userName_office = security.GetValue("UserName").ToString();
                //MessageBox.Show(userName);
                app.Application.ActiveDocument.Content.InsertAfter("\r\n\r\n文案：" + userName_office);

                Word.Document myDoc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Paragraphs paragraphs = myDoc.Paragraphs;
                paragraphs[paragraphs.Count].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }
        }

        private void author_wps_Click(object sender, RibbonControlEventArgs e)
        {
            if (Registry_bool("SOFTWARE", "Kingsoft") == true)
            {
                if (Registry_bool("SOFTWARE\\Kingsoft", "Office") == true)
                {   //wps user

                    JObject js = ImportJSON(XMTsetting);
                    js["author"] = 2;
                    SetjsonFun(XMTsetting, js);
                    author.Image = Properties.Resources.wps;

                    string userName_wps;
                    RegistryKey security_wps;
                    security_wps = Registry.CurrentUser.OpenSubKey(name: "SOFTWARE\\Kingsoft\\Office\\6.0\\Common\\UserInfo", writable: true);
                    userName_wps = security_wps.GetValue("UserName").ToString();
                    app.Application.ActiveDocument.Content.InsertAfter("\r\n\r\n文案：" + userName_wps);

                    Word.Document myDoc = Globals.ThisAddIn.Application.ActiveDocument;
                    Word.Paragraphs paragraphs = myDoc.Paragraphs;
                    paragraphs[paragraphs.Count].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }
            }
        }

        private void author_doc_Click(object sender, RibbonControlEventArgs e)
        {
            //文档创建者

            JObject js = ImportJSON(XMTsetting);
            js["author"] = 3;
            SetjsonFun(XMTsetting, js);
            author.Image = Properties.Resources.document;

            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Microsoft.Office.Core.DocumentProperties)app.Application.ActiveDocument.BuiltInDocumentProperties;
            //Microsoft.Office.Core.DocumentProperty author;
            string author_doc;
            author_doc = properties["Author"].Value;
            //MessageBox.Show(author);
            app.Application.ActiveDocument.Content.InsertAfter("\r\n\r\n文案：" + author_doc);

            Word.Document myDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs paragraphs = myDoc.Paragraphs;
            paragraphs[paragraphs.Count].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }

        private void author_c_Click(object sender, RibbonControlEventArgs e)
        {
            //计算机用户

            JObject js = ImportJSON(XMTsetting);
            js["author"] = 4;
            SetjsonFun(XMTsetting, js);
            author.Image = Properties.Resources.computer;

            //string userName_computer = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string userName_computer = Environment.UserName;
            //MessageBox.Show(userName_computer);
            app.Application.ActiveDocument.Content.InsertAfter("\r\n\r\n文案：" + userName_computer);

            Word.Document myDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs paragraphs = myDoc.Paragraphs;
            paragraphs[paragraphs.Count].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }
        

        public Color ReadCodeTHINGcolor(int choice,string colorseries)
        {
            //读取代码预设颜色

            JObject js = ImportJSON(PresetCodeFile);

            int r = int.Parse(js["preset_" + choice + "_" + colorseries + "_r"].ToString());
            int g = int.Parse(js["preset_" + choice + "_" + colorseries + "_g"].ToString());
            int b = int.Parse(js["preset_" + choice + "_" + colorseries + "_b"].ToString());

            return Color.FromArgb(r, g, b);
        }

        
        private void CodeTabSetting_Click(object sender, RibbonControlEventArgs e)
        {
            //设置颜色1
            MyColorDialog.Color = CodeBackcolor1;
            MyColorDialog.FullOpen = true;

            DialogResult dr = MyColorDialog.ShowDialog();
            if (dr == DialogResult.OK) CodeBackcolor1 = MyColorDialog.Color;
        }

        private void About_Click(object sender, RibbonControlEventArgs e)
        {
            AboutForm abForm = new AboutForm();
            abForm.ShowDialog();
        }

        private void CodeTabSetting2_Click(object sender, RibbonControlEventArgs e)
        {
            //设置颜色2
            MyColorDialog.Color = CodeBackcolor2;
            MyColorDialog.FullOpen = true;

            DialogResult dr = MyColorDialog.ShowDialog();
            if (dr == DialogResult.OK) CodeBackcolor2 = MyColorDialog.Color;
        }

        private void ResetCodeBGcolor_Click(object sender, RibbonControlEventArgs e)
        {
            //着色

            int table_select_end = Globals.ThisAddIn.Application.Selection.End;
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
            int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
            //MessageBox.Show(LevelOfTables.ToString());
            Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

            int row = tables_select.Range.Rows.Count;
            for (int i = 1; i <= row; i++)
            {
                if (i % 2 == 1)
                {
                    Color c1 = CodeBackcolor1;
                    tables_select.Rows[i].Shading.BackgroundPatternColor = GetColor(c1);
                }
                else
                {
                    Color c2 = CodeBackcolor2;
                    tables_select.Rows[i].Shading.BackgroundPatternColor = GetColor(c2);
                }
            }
            if (CodeBorderYN() == true)
                tables_select.Columns[1].Borders[Word.WdBorderType.wdBorderRight].Color = GetColor(CodeBorderLine);
        }

        private void BorderLine_Click(object sender, RibbonControlEventArgs e)
        {
            //代码竖线颜色
            MyColorDialog.Color = CodeBorderLine;
            MyColorDialog.FullOpen = true;

            DialogResult dr = MyColorDialog.ShowDialog();
            if (dr == DialogResult.OK)
            {
                CodeBorderLine = MyColorDialog.Color;

                int table_select_end = Globals.ThisAddIn.Application.Selection.End;
                Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
                int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
                                                 //MessageBox.Show(LevelOfTables.ToString());
                Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

                if (tables_select.Columns.Count > 1 && CodeBorderYN() == true)
                {
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Visible = true;
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Color = GetColor(CodeBorderLine);
                }
                if (CodeBorderYN() == false)
                {
                    tables_select.Columns[1].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
                    tables_select.Columns[1].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tables_select.Columns[1].Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    tables_select.Columns[1].Borders[Word.WdBorderType.wdBorderLeft].Color = GetColor(CodeBorderLine);
                }
            }           
        }

        private void TableWidth_Click(object sender, RibbonControlEventArgs e)
        {
            //统一表宽度，以最宽的为准
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Tables;
            int NumOfTables = tables.Count;

            float width;
            float widthT;
            if(tables[1].Columns.Count==1)
            {
                width = tables[1].Columns[1].Width;
            }
            else
            {
                width = tables[1].Columns[1].Width + tables[1].Columns[2].Width;
            }
            for (int i = 2; i <= NumOfTables; i++)
            {
                if(tables[i].Columns.Count == 1)
                {
                    widthT = tables[i].Columns[1].Width;
                }
                else
                {
                    widthT = tables[i].Columns[1].Width + tables[i].Columns[2].Width;
                }
                if(widthT>width)
                {
                    width = widthT;
                }
            }
            for (int i = 1; i <= NumOfTables; i++)
            {
                if (tables[i].Columns.Count == 1)
                {
                    tables[i].Columns[1].Width = width;
                }
                else
                {
                    tables[i].Columns[2].Width = width - tables[i].Columns[1].Width;
                }
            }

            //tables[2].Columns[2].Width = width;
            //MessageBox.Show(tables[1].Columns[2].Width.ToString());
        }


        private void CodeOpenPreset_Click(object sender, RibbonControlEventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);
            int num = int.Parse(js["num"].ToString());
            string sentence = "请选择预设的编号：\r\n编号  名称";
            /*for (int i = 1; i < js.Count; i++)
            {
                sentence = sentence + "\r\n  " + i + "   " + js.Properties().ToList()[i].Name.ToString();               
            }*/
            for(int i=1;i<=num;i++)
            {
                sentence = sentence + "\r\n  " + i + "   " + js["preset_" + i + "_name"].ToString();
            }
            string choice = Interaction.InputBox(sentence, "选择预设").ToString();
            if(choice!="")
            {
                int choice_int = int.Parse(choice);
                
                CodeBackcolor1 = ReadCodeTHINGcolor(choice_int, "CodeBackcolor1");
                CodeBackcolor2 = ReadCodeTHINGcolor(choice_int, "CodeBackcolor2");
                CodeBorderLine = ReadCodeTHINGcolor(choice_int, "CodeBorderLine");

            }               
        }

        private void CodeSavePreset_Click(object sender, RibbonControlEventArgs e)
        {
            JObject js = ImportJSON(PresetCodeFile);
            string presetName = Interaction.InputBox("预设名称", "保存预设").ToString();            
            
            string c1r = CodeBackcolor1.R.ToString();
            string c1g = CodeBackcolor1.G.ToString();
            string c1b = CodeBackcolor1.B.ToString();
            string c2r = CodeBackcolor2.R.ToString();
            string c2g = CodeBackcolor2.G.ToString();
            string c2b = CodeBackcolor2.B.ToString();
            string cbr = CodeBorderLine.R.ToString();
            string cbg = CodeBorderLine.G.ToString();
            string cbb = CodeBorderLine.B.ToString();

            if(presetName!="")
            {
                int num = int.Parse(js["num"].ToString());
                num = num + 1;
                js["num"] = num.ToString();
                js["preset_" + num + "_name"] = presetName;
                js["preset_" + num + "_CodeBackcolor1_r"] = c1r;
                js["preset_" + num + "_CodeBackcolor1_g"] = c1g;
                js["preset_" + num + "_CodeBackcolor1_b"] = c1b;
                js["preset_" + num + "_CodeBackcolor2_r"] = c2r;
                js["preset_" + num + "_CodeBackcolor2_g"] = c2g;
                js["preset_" + num + "_CodeBackcolor2_b"] = c2b;
                js["preset_" + num + "_CodeBorderLine_r"] = cbr;
                js["preset_" + num + "_CodeBorderLine_g"] = cbg;
                js["preset_" + num + "_CodeBorderLine_b"] = cbb;

                SetjsonFun(PresetCodeFile, js);
            }            
        }

        

        private void CodeListNum_Click(object sender, RibbonControlEventArgs e)
        {
            int table_select_end = Globals.ThisAddIn.Application.Selection.End;
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
            int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
            Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

            if (tables_select.Columns.Count == 1)
            {
                //如果只有一列
                tables_select.Range.Font.Name = CodeFontFunN();
                tables_select.Range.Font.Size = (float)CodeFontFunS();
                
                tables_select.Columns.Add(tables_select.Columns[1]);
                tables_select.Columns[1].Width = CodeNoWidth();
                tables_select.Columns[2].Width = tables_select.Columns[2].Width - tables_select.Columns[1].Width;

                int row = tables_select.Range.Rows.Count;
                for (int i = 1; i <= row; i++)
                {
                    tables_select.Cell(i, 1).Range.Text = "" + i;
                    tables_select.Cell(i, 1).Shading.BackgroundPatternColor = tables_select.Cell(i, 2).Shading.BackgroundPatternColor;
                    tables_select.Cell(i, 1).Range.Font.Color = GetColor(CodeTxtColor);
                }
                tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Visible = true;
                tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Color = GetColor(CodeBorderLine);
            }
            else if (tables_select.Columns.Count == 2)
            {
                //如果有两列，也就是有编号
                tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
                tables_select.Columns[2].Width = tables_select.Columns[2].Width + tables_select.Columns[1].Width;
                tables_select.Columns[1].Delete();               
            }
            else
            {
                MessageBox.Show("这不是代码块！");
            }
            tables_select.Columns[1].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
        }

        //private void CodeControl_Click(object sender, RibbonControlEventArgs e)
        //{
        //    CodeControlForm codeControlForm = new CodeControlForm();
        //    codeControlForm.Show();
        //}

        private void CodeBoderLineFun_Click(object sender, RibbonControlEventArgs e)
        {
            int table_select_end = Globals.ThisAddIn.Application.Selection.End;
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
            int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
            Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

            if (tables_select.Columns.Count == 2)
            {
                if(tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Visible==false)
                {
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Visible = true;
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Color = GetColor(CodeBorderLine);
                }
                else
                {
                    tables_select.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
                }
            }
        }

        private void CodeTableFullWidth_Click(object sender, RibbonControlEventArgs e)
        {
            //将选中的代码块左右拉到最大
            int table_select_end = Globals.ThisAddIn.Application.Selection.End;
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
            int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
            Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

            if (tables_select.Columns.Count == 1)
            {
                tables_select.Columns[1].Width = CodeNoWidth();
            }
            if (tables_select.Columns.Count == 2)
            {
                tables_select.Columns[1].Width = CodeNoWidth();
                float width2 = Globals.ThisAddIn.Application.Selection.PageSetup.PageWidth - Globals.ThisAddIn.Application.Selection.PageSetup.LeftMargin - Globals.ThisAddIn.Application.Selection.PageSetup.RightMargin;
                tables_select.Columns[2].Width = width2 - tables_select.Columns[1].Width;
            }
        }

        private void CodeNoWidthSet_Click(object sender, RibbonControlEventArgs e)
        {
            //设置选中的代码块的行号宽度为行号宽度的默认值
            int table_select_end = Globals.ThisAddIn.Application.Selection.End;
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
            int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
            Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];
            if (tables_select.Columns.Count == 2)
            {
                JObject js = ImportJSON(PresetCodeFile);
                js["DefaultNoWidth"] = tables_select.Columns[1].Width.ToString();
                SetjsonFun(PresetCodeFile, js);
            }
        }
       
        private void ParagraphShading_Click(object sender, RibbonControlEventArgs e)
        {
            //段落底纹颜色
            JObject js = ImportJSON(PresetToolsBoxShadeColor);
            int r = int.Parse(js["ParaShadingColor_r"].ToString());
            int g = int.Parse(js["ParaShadingColor_g"].ToString());
            int b = int.Parse(js["ParaShadingColor_b"].ToString());
            Color c = Color.FromArgb(r, g, b);
            Globals.ThisAddIn.Application.Selection.ParagraphFormat.Shading.BackgroundPatternColor = GetColor(c);

            paraShandingChoice = 1;
            ParaShadeSplit.Image = Properties.Resources.底纹;
        }

        private void CodeControl_Click(object sender, RibbonControlEventArgs e)
        {
            CodeControlForm codeControlForm = new CodeControlForm(PresetCodeFile);
            codeControlForm.Show();
        }

        private void ParaShadingColorSet_Click(object sender, RibbonControlEventArgs e)
        {
            //修改段落底纹颜色
            JObject js = ImportJSON(PresetToolsBoxShadeColor);
            int r = int.Parse(js["ParaShadingColor_r"].ToString());
            int g = int.Parse(js["ParaShadingColor_g"].ToString());
            int b = int.Parse(js["ParaShadingColor_b"].ToString());
            Color c = Color.FromArgb(r, g, b);

            MyColorDialog.Color = c;
            MyColorDialog.FullOpen = true;

            DialogResult dr = MyColorDialog.ShowDialog();

            if (dr == DialogResult.OK)
            {
                c = MyColorDialog.Color;
                //MessageBox.Show(c.ToString());

                js["ParaShadingColor_r"] = c.R.ToString();
                js["ParaShadingColor_g"] = c.G.ToString();
                js["ParaShadingColor_b"] = c.B.ToString();
                SetjsonFun(PresetToolsBoxShadeColor, js);
            }

            paraShandingChoice = 2;
            ParaShadeSplit.Image = Properties.Resources.底纹颜色;
        }

        private void CodeFormat2_Click(object sender, RibbonControlEventArgs e)
        {
            //用段落底纹排版, markdown style

            int SelectStart = Globals.ThisAddIn.Application.Selection.Start;
            int Selectend = Globals.ThisAddIn.Application.Selection.End;

            // 标记替换
            object Replace_String = "^p";       //要替换的字符
            object ms = System.Type.Missing;
            object Replace = Word.WdReplace.wdReplaceAll;//设置替换方式:一，全部替换；二，只替换一个；三，一个都不替换。
            object ReplaceWith = "^l";             //最终替换成的字符
            //执行Word自带的查找/替换功能函数
            Globals.ThisAddIn.Application.ActiveDocument.Range(SelectStart, Selectend - 1).Find.Execute(ref Replace_String, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ReplaceWith, ref Replace, ref ms, ref ms, ref ms, ref ms);


            // 底纹样式
            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
            {
                //MessageBox.Show(Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal);
                if (Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal == "code_MD")
                {
                    codeMD_bool = 1;
                    break;
                }
            }

            if (codeMD_bool == 0)
            {
                CreatCodeMDStyle(codeMD_CurrentColor);
                codeMD_bool = 1;
            }

            Globals.ThisAddIn.Application.Selection.set_Style("code_MD");
        }

        void CreatCodeMDStyle(Word.WdColor BGcolor)
        {
            JObject Js = ImportJSON(PresetCodeMDFile);

            Globals.ThisAddIn.Application.ActiveDocument.Styles.Add("code_MD", Word.WdStyleType.wdStyleTypeParagraph);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].AutomaticallyUpdate = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.NameFarEast = Js["DefaultFontC"].ToString();
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.NameAscii = Js["DefaultFontE"].ToString();
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.NameOther = Js["DefaultFontE"].ToString();
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Name = "";
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Size = float.Parse(Js["DefaultFontSize"].ToString());
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Bold = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Italic = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Underline = Word.WdUnderline.wdUnderlineNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.StrikeThrough = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Outline = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Emboss = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Shadow = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Hidden = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.SmallCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.AllCaps = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Color = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Engrave = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Superscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Subscript = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Scaling = 100;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Kerning = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Animation = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.DisableCharacterSpaceGrid = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.EmphasisMark = Word.WdEmphasisMark.wdEmphasisMarkNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.Ligatures = Word.WdLigatures.wdLigaturesNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.NumberSpacing = Word.WdNumberSpacing.wdNumberSpacingDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.NumberForm = Word.WdNumberForm.wdNumberFormDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.StylisticSet = Word.WdStylisticSet.wdStylisticSetDefault;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Font.ContextualAlternates = 0;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.LeftIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.RightIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.SpaceBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.SpaceBeforeAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.SpaceAfter = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.SpaceAfterAuto = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.WidowControl = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.KeepWithNext = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.KeepTogether = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.PageBreakBefore = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.NoLineNumber = 0;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.FirstLineIndent = app.CentimetersToPoints(0);
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.CharacterUnitLeftIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.CharacterUnitRightIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.LineUnitBefore = (float)0.5;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.LineUnitAfter = (float)0.5;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.MirrorIndents = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.CollapsedByDefault = 0;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.DisableLineHeightGrid = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.HalfWidthPunctuationOnTopOfLine = 0;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.TabStops.ClearAll();

            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Shading.BackgroundPatternColor = BGcolor;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders.DistanceFromTop = 1;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders.DistanceFromRight = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders.DistanceFromLeft = 4;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Borders.DistanceFromBottom = 1;

            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].NoSpaceBetweenParagraphsOfSameStyle = false;
            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].set_BaseStyle("");

            Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].Frame.Delete();
        }


        private void DeleteShading_Click(object sender, RibbonControlEventArgs e)
        {
            //清除底纹
            Globals.ThisAddIn.Application.Selection.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;

            paraShandingChoice = 3;
            ParaShadeSplit.Image = Properties.Resources.删除底纹;
        }


        private void ThreeLine_Click(object sender, RibbonControlEventArgs e)
        {
            //表格转为三线表

            int table_select_end = Globals.ThisAddIn.Application.Selection.End;
            Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
            int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
            Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

            tables_select.Borders.Enable = 0;  //delete all borders

            tables_select.Borders[Word.WdBorderType.wdBorderBottom].Visible = true;
            tables_select.Borders[Word.WdBorderType.wdBorderTop].Visible = true;
            tables_select.Rows[1].Borders[Word.WdBorderType.wdBorderBottom].Visible = true;

            tables_select.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
            tables_select.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
            tables_select.Rows[1].Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
        }

        private void CodeFormatLatex_Click(object sender, RibbonControlEventArgs e)
        {
            //latex 样式排版
            Globals.ThisAddIn.Application.Selection.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;

            wordReplace("^l", "^p");

            Word.Table table = Globals.ThisAddIn.Application.Selection.ConvertToTable(Separator: Word.WdTableFieldSeparator.wdSeparateByParagraphs);

            JObject js = ImportJSON(PresetCodeLatexFile);

            table.Range.Font.NameFarEast = js["DefaultFontC"].ToString();
            table.Range.Font.NameAscii = js["DefaultFontE"].ToString();
            table.Range.Font.NameOther = js["DefaultFontE"].ToString();
            table.Range.Font.Name = "";
            table.Range.Font.Size = float.Parse(js["DefaultFontSize"].ToString());
            float NoSize = float.Parse(js["DefaultNoFontSize"].ToString());

            float width1 = float.Parse(js["DefaultNoWidth"].ToString());
            float width2 = Globals.ThisAddIn.Application.Selection.PageSetup.PageWidth - Globals.ThisAddIn.Application.Selection.PageSetup.LeftMargin - Globals.ThisAddIn.Application.Selection.PageSetup.RightMargin;
            width2 = width2 - width1;

            table.Columns.Add(table.Columns[1]);
            table.Columns[1].Width = width1;
            table.Columns[2].Width = width2;

            int row = table.Range.Rows.Count;
            for (int i = 1; i <= row; i++)
            {
                table.Cell(i, 1).Range.Text = "" + i;
                table.Cell(i, 1).Range.Font.Color = Word.WdColor.wdColorBlack;
                table.Cell(i, 1).Range.Font.Size = NoSize;
            }

            table.Columns[2].Borders[Word.WdBorderType.wdBorderTop].Visible = true;
            table.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].Visible = true;
            table.Columns[2].Borders[Word.WdBorderType.wdBorderRight].Visible = true;
            table.Columns[2].Borders[Word.WdBorderType.wdBorderBottom].Visible = true;

            table.Columns[2].Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Columns[2].Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Columns[2].Borders[Word.WdBorderType.wdBorderRight].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Columns[2].Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
        }


        private void FileTabOnOff_Click(object sender, RibbonControlEventArgs e)
        {
            //面板开关

            if (FileTabOnOff.Checked == false)
            {
                FileTabOnOff.Label = "标签栏";

                foreach (var tempPane in HwndPaneDic_tab)
                {
                    HwndPaneDic_tab[tempPane.Key].Visible = false;
                    HwndPaneDic_tab[tempPane.Key].Dispose();
                    Globals.ThisAddIn.CustomTaskPanes.Remove(HwndPaneDic_tab[tempPane.Key]);
                }
            }
            else
            {
                FileTabOnOff.Label = "关标签栏";

                //标签栏
                int TempInt = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
                RefreshFileTabPane(TempInt);
            }

            //设置面板可见性
            FileTabPane.Visible = FileTabOnOff.Checked;

        }

        private void RefreshFileTabPane(int HwndInt)
        {
            //重新绑定面板
            //MessageBox.Show(app.Documents[1].Name);


            //保持标签顺序
            DocNamesList = TabForm.DocNamesList_pane;
            Doc2List();


            if (HwndPaneDic_tab.ContainsKey(HwndInt))
            {
                //FileTabPane = HwndPaneDic[HwndInt];

                //HwndPaneDic[HwndInt].Visible = false;
                Globals.ThisAddIn.CustomTaskPanes.Remove(HwndPaneDic_tab[HwndInt]);
                //HwndPaneDic.Remove(HwndInt);

                //创建控件
                TabForm FileTabPane2 = new TabForm(DocNamesList);
                //添加控件
                FileTabPane = Globals.ThisAddIn.CustomTaskPanes.Add(FileTabPane2, "分点标签栏");               
                //设置在上方显示
                FileTabPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                //禁止用户修改位置
                FileTabPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                //设置高度
                //MessageBox.Show("2:"+Globals.ThisAddIn.Application.ActiveWindow.Width.ToString());
                if (GetWinScaling() * Globals.ThisAddIn.Application.ActiveWindow.Width > 129* app.Documents.Count)//放大倍数，窗口尺寸和实际不符
                {
                    FileTabPane.Height = fileTabPaneHeight;
                    //MessageBox.Show("FileTabPane.Width> 129* app.Documents.Count");
                }                   
                else
                {
                    FileTabPane.Height = 103;
                    //MessageBox.Show("FileTabPane.Width< 129* app.Documents.Count");
                }
                //事件
                //FileTabPane.Visible = true;
                //FileTabPane.VisibleChanged -= new EventHandler(CustomPane_VisibleChanged);
                //添加到字典
                //HwndPaneDic.Add(HwndInt, FileTabPane);
                HwndPaneDic_tab[HwndInt] = FileTabPane;
            }
            else
            {
                //创建控件
                TabForm FileTabPane2 = new TabForm(DocNamesList);
                //添加控件
                FileTabPane = Globals.ThisAddIn.CustomTaskPanes.Add(FileTabPane2, "分点标签栏");
                //设置在上方显示
                FileTabPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                //禁止用户修改位置
                FileTabPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                //设置高度
                if (GetWinScaling() * Globals.ThisAddIn.Application.ActiveWindow.Width > 129 * app.Documents.Count)
                {
                    FileTabPane.Height = fileTabPaneHeight;
                    //MessageBox.Show("FileTabPane.Width> 129* app.Documents.Count");
                }
                else
                {
                    FileTabPane.Height = 103;
                    //MessageBox.Show("FileTabPane.Width< 129* app.Documents.Count");
                }
                //事件
                //FileTabPane.Visible = false;
                FileTabPane.VisibleChanged += new EventHandler(CustomPane_VisibleChanged);
                //添加到字典
                HwndPaneDic_tab.Add(HwndInt, FileTabPane);
            }
        }

        private void CustomPane_VisibleChanged(object sender, System.EventArgs e)
        {
            //保持按钮与面板状态一致

            int TempInt = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
            FileTabOnOff.Checked = FileTabPane.Visible;

            if (!FileTabOnOff.Checked)
            {
                app.WindowActivate -= new Word.ApplicationEvents4_WindowActivateEventHandler(CustomPane_WindowActivate);
                app.WindowDeactivate -= new Word.ApplicationEvents4_WindowDeactivateEventHandler(CustomPane_WindowDeactivate);
            }
            else
            {
                app.WindowActivate += new Word.ApplicationEvents4_WindowActivateEventHandler(CustomPane_WindowActivate);
                app.WindowDeactivate += new Word.ApplicationEvents4_WindowDeactivateEventHandler(CustomPane_WindowDeactivate);
            }

        }

        private void CustomPane_WindowActivate(Word.Document Doc, Word.Window WD)
        {
            //窗体激活事件
            
            app.WindowActivate -= new Word.ApplicationEvents4_WindowActivateEventHandler(CustomPane_WindowActivate);
            //app = Globals.ThisAddIn.Application;
            //string DocName = Doc.Name;


            //保持标签栏高度
            fileTabPaneHeight = FileTabPane.Height;


            int TempHwnd = WD.Hwnd;
            RefreshFileTabPane(TempHwnd);
            //设置面板可见性
            FileTabPane.Visible = true;
        }

        private void CustomPane_WindowDeactivate(Word.Document Doc, Word.Window WD)
        {
            //窗体取消激活事件

            int TempHwnd = WD.Hwnd;

            if (HwndPaneDic_tab.ContainsKey(TempHwnd))
            {
                HwndPaneDic_tab[TempHwnd].Visible = false;
                app.WindowActivate += new Word.ApplicationEvents4_WindowActivateEventHandler(CustomPane_WindowActivate);
                //app.WindowDeactivate += new Word.ApplicationEvents4_WindowDeactivateEventHandler(CustomPane_WindowDeactivate);
            }
        }
        

        private void TableColoring_Click(object sender, RibbonControlEventArgs e)
        {
            int TempInt = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;

            if (HwndPaneDic_tableColor.ContainsKey(TempInt))
            {
                HwndPaneDic_tableColor[TempInt].Visible = false;
                HwndPaneDic_tableColor[TempInt].Dispose();
                Globals.ThisAddIn.CustomTaskPanes.Remove(HwndPaneDic_tableColor[TempInt]); 
                HwndPaneDic_tableColor.Remove(TempInt);
            }
            else
            {
                TableColoringForm tableColoringForm = new TableColoringForm(PresetToolsBoxTable);

                tableColoringPane = Globals.ThisAddIn.CustomTaskPanes.Add(tableColoringForm, "设置表格颜色");
                tableColoringPane.Width = 333;
                tableColoringPane.Visible = true;
                HwndPaneDic_tableColor.Add(TempInt, tableColoringPane);
            }
        }

        private void TimeHeader_Click(object sender, RibbonControlEventArgs e)
        {
            //设置页眉
            foreach (Word.Section section in app.Application.ActiveDocument.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                headerRange.Font.Name = "宋体";
                headerRange.Text = System.DateTime.Now.ToString("d");
            }
        }

        private void author_Click(object sender, RibbonControlEventArgs e)
        {
            //署名

            JObject js = ImportJSON(XMTsetting);

            //office用户
            if (js["author"].ToString() == "1")
                author_o_Click(sender, e);

            //wps user
            if (js["author"].ToString() == "2")
                author_wps_Click(sender, e);

            //文档创建者
            if (js["author"].ToString() == "3")
                author_doc_Click(sender, e);

            //计算机用户
            if (js["author"].ToString() == "4")
                author_c_Click(sender, e);

            //自定义署名
            if (js["author"].ToString() == "5")
                author_new_Click(sender, e);
        }

        private void author_new_Click(object sender, RibbonControlEventArgs e)
        {
            // 自定义署名

            JObject js = ImportJSON(XMTsetting);
            int nullkey = 1;

            if(js["author_new"].ToString()=="")
            {
                string key = Interaction.InputBox("输入自定义署名", "署名设置").ToString();
                if (key == "")
                {                    
                    nullkey = 0;
                }
                else
                {
                    js["author_new"] = key;
                    author_new.Image = Properties.Resources.署名;
                    author_new.Label = key;
                }
            }

            
            if (nullkey == 1)
            {
                app.Application.ActiveDocument.Content.InsertAfter("\r\n\r\n文案：" + js["author_new"].ToString());

                Word.Document myDoc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Paragraphs paragraphs = myDoc.Paragraphs;
                paragraphs[paragraphs.Count].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                js["author"] = 5;
                SetjsonFun(XMTsetting, js);
                author.Image = Properties.Resources.署名;
            }            
        }


        private void changecharCE_Click(object sender, RibbonControlEventArgs e)
        {
            //将英文标点换成中文标点

            int TempInt = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;

            if (HwndPaneDic_ChangeChar.ContainsKey(TempInt))
            {
                HwndPaneDic_ChangeChar[TempInt].Visible = false;
                HwndPaneDic_ChangeChar[TempInt].Dispose();
                Globals.ThisAddIn.CustomTaskPanes.Remove(HwndPaneDic_ChangeChar[TempInt]);
                HwndPaneDic_ChangeChar.Remove(TempInt);
            }
            else
            {
                ChangeCharForm changeCharForm = new ChangeCharForm();

                changCharPane = Globals.ThisAddIn.CustomTaskPanes.Add(changeCharForm, "设置替换字符");
                changCharPane.Width = 215;
                changCharPane.Visible = true;
                HwndPaneDic_ChangeChar.Add(TempInt, changCharPane);
            }
        }



        private void checkCharMatch_Click(object sender, RibbonControlEventArgs e)
        {
            //检查符号是否配对

            int TempInt = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;

            if (HwndPaneDic_CharMatch.ContainsKey(TempInt))
            {
                HwndPaneDic_CharMatch[TempInt].Visible = false;
                HwndPaneDic_CharMatch[TempInt].Dispose();
                Globals.ThisAddIn.CustomTaskPanes.Remove(HwndPaneDic_CharMatch[TempInt]);
                HwndPaneDic_CharMatch.Remove(TempInt);
            }
            else
            {
                CharMatchForm charMatchForm = new CharMatchForm();

                charMatchPane = Globals.ThisAddIn.CustomTaskPanes.Add(charMatchForm, "设置匹配字符");
                charMatchPane.Width = 335;
                charMatchPane.Visible = true;
                HwndPaneDic_CharMatch.Add(TempInt, charMatchPane);
            }
        }

        private void del_header_line_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (Word.Section section in app.Application.ActiveDocument.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            }
        }

        private void SetCode3CurrentColor_Click(object sender, RibbonControlEventArgs e)
        {
            //设置代码3底纹颜色

            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
            {
                if (Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal == "code_MD")
                {
                    codeMD_bool = 1;
                    codeMD_CurrentColor = Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Shading.BackgroundPatternColor;
                    break;
                }
            }


            // 响应

            //Wdcolor2Color(codeMD_CurrentColor);
            //MessageBox.Show(Wdcolor2Color(codeMD_CurrentColor).ToString());

            MyColorDialog.Color = Wdcolor2Color(codeMD_CurrentColor);
            MyColorDialog.FullOpen = true;

            DialogResult dr = MyColorDialog.ShowDialog();
            if (dr == DialogResult.OK) codeMD_CurrentColor = GetColor(MyColorDialog.Color);

            if (codeMD_bool == 1)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Styles["code_MD"].ParagraphFormat.Shading.BackgroundPatternColor = codeMD_CurrentColor;
            }
        }

        public Color Wdcolor2Color(Word.WdColor color_wd)
        {
            //convert WdColor to color class

            string color_wd16 = Convert.ToString(((int)color_wd), 16);
            string b_str16 = color_wd16.Substring(0, 2);
            string g_str16 = color_wd16.Substring(2, 2);
            string r_str16 = color_wd16.Substring(4, 2);

            int r = Convert.ToInt32(r_str16, 16);
            int g = Convert.ToInt32(g_str16, 16);
            int b = Convert.ToInt32(b_str16, 16);

            return Color.FromArgb(r, g, b);
        }

        private void saveCode3Color_Click(object sender, RibbonControlEventArgs e)
        {
            JObject js_codeMD = ImportJSON(PresetCodeMDFile);
            js_codeMD["ParaShadingColor_r"] = Wdcolor2Color(codeMD_CurrentColor).R.ToString();
            js_codeMD["ParaShadingColor_g"] = Wdcolor2Color(codeMD_CurrentColor).G.ToString();
            js_codeMD["ParaShadingColor_b"] = Wdcolor2Color(codeMD_CurrentColor).B.ToString();

            SetjsonFun(PresetCodeMDFile, js_codeMD);
        }

        private void ParaShadeSplit_Click(object sender, RibbonControlEventArgs e)
        {
            //段落底纹工具集

            switch(paraShandingChoice)
            {
                case 1:
                    ParagraphShading_Click(sender, e);
                    break;
                case 2:
                    ParaShadingColorSet_Click(sender, e);  
                    break;
                case 3:
                    DeleteShading_Click(sender, e);
                    break;
                default:
                    MessageBox.Show("段落底纹错误！", "段落底纹");
                    break;
            }
        }

        private void styleShading_Click(object sender, RibbonControlEventArgs e)
        {
            //设置样式底纹颜色

            PatternSelectForm patterns = new PatternSelectForm(app.Documents);
            patterns.ShowDialog();
            string style_name = patterns.selectedName;

            // string style_name = Interaction.InputBox("输入样式名称", "样式底纹").ToString();
            int style_exsit = 0;
            Word.WdColor style_color;

            if (style_name != "")
            {
                for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
                {
                    if (Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal == style_name)
                    {
                        style_exsit = 1;
                        break;
                    }
                }

                if (style_exsit == 1)
                {
                    if (Globals.ThisAddIn.Application.ActiveDocument.Styles[style_name].ParagraphFormat.Shading.BackgroundPatternColor != Word.WdColor.wdColorAutomatic)
                    {
                        MyColorDialog.Color = Wdcolor2Color(Globals.ThisAddIn.Application.ActiveDocument.Styles[style_name].ParagraphFormat.Shading.BackgroundPatternColor);
                        MyColorDialog.FullOpen = true;
                    }

                    DialogResult dr = MyColorDialog.ShowDialog();

                    if (dr == DialogResult.OK)
                    {
                        style_color = GetColor(MyColorDialog.Color);
                        Globals.ThisAddIn.Application.ActiveDocument.Styles[style_name].ParagraphFormat.Shading.BackgroundPatternColor = style_color;

                        styleShadingChoice = 1;
                        StyleShadeSplit.Image = Properties.Resources.样式底纹;
                    }
                }
                else
                {
                    MessageBox.Show("样式输入错误！", "样式底纹");
                }
            }

        }

        private void styleShadeClear_Click(object sender, RibbonControlEventArgs e)
        {
            //删除样式底纹颜色

            PatternSelectForm patterns = new PatternSelectForm(app.Documents);
            patterns.ShowDialog();
            string style_name = patterns.selectedName;

            //string style_name = Interaction.InputBox("输入样式名称", "样式底纹").ToString();
            int style_exsit = 0;

            if (style_name != "")
            {
                for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Styles.Count; i++)
                {
                    if (Globals.ThisAddIn.Application.ActiveDocument.Styles[i].NameLocal == style_name)
                    {
                        style_exsit = 1;
                        break;
                    }
                }

                if (style_exsit == 1)
                {
                    Globals.ThisAddIn.Application.ActiveDocument.Styles[style_name].ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;

                    styleShadingChoice = 2;
                    StyleShadeSplit.Image = Properties.Resources.样式底纹清除;
                }
                else
                {
                    MessageBox.Show("样式输入错误！", "样式底纹");
                }
            }
           
        }

        private void StyleShadeSplit_Click(object sender, RibbonControlEventArgs e)
        {
            //样式底纹工具集

            switch (styleShadingChoice)
            {
                case 1:
                    styleShading_Click(sender, e);
                    break;
                case 2:
                    styleShadeClear_Click(sender, e);
                    break;
                default:
                    MessageBox.Show("样式底纹错误！", "样式底纹");
                    break;
            }
        }

        private void SettingBt_Click(object sender, RibbonControlEventArgs e)
        {
            SettingForm settingForm = new SettingForm();
            settingForm.ShowDialog();
            bool boolKeyAllTrue = settingForm.boolKeyAllTrue;
            if(boolKeyAllTrue==true)
            {
                KeyAllTrue();
            }
            else
            {
                KeyStateLoad();
            }

        }

        /// <summary>
        /// 文档列表转列表，以及更新列表内容
        /// </summary>
        private void Doc2List()
        {
            int DocNums = app.Documents.Count;
            for (int i = 1; i <= DocNums; i++)
            {
                if (!DocNamesList.Contains(app.Documents[i].Path + "\\" + app.Documents[i].Name))
                {
                    DocNamesList.Add(app.Documents[i].Path + "\\" + app.Documents[i].Name);
                } 
            }

            foreach (string docname in DocNamesList)
            {
                if (!boolDocsContain(docname))
                {
                    DocNamesList.Remove(docname);
                }
            }
        }

        /// <summary>
        /// 判断是否含有某个文档
        /// </summary>
        /// <param name="doc_name">文档名称</param>
        /// <returns></returns>
        private bool boolDocsContain(string doc_name)
        {
            int DocNums = app.Documents.Count;
            for (int i = 1; i <= DocNums; i++)
            {
                if (doc_name == (app.Documents[i].Path + "\\" + app.Documents[i].Name))
                {
                    return true;
                }
            }

            return false;
        }


        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr ptr);

        [DllImport("gdi32.dll")]
        public static extern int GetDeviceCaps(
            IntPtr hdc, // handle to DC
            int nIndex // index of capability
        );

        [DllImport("user32.dll", EntryPoint = "ReleaseDC")]
        public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();

        /// <summary>
        /// 获取屏幕缩放系数
        /// </summary>
        public float GetWinScaling()
        {
            // 获取屏幕缩放
            var hdc = GetDC(GetDesktopWindow());
            int nWidth = GetDeviceCaps(hdc, 117);
            ReleaseDC(IntPtr.Zero, hdc);
            float f_Scale = (float)nWidth / (float)Screen.PrimaryScreen.Bounds.Width;
            return 1 / f_Scale;
        }

        private void WeixinPic_Click(object sender, RibbonControlEventArgs e)
        {
            if (GetNetStatus("mp.weixin.qq.com"))
            {
                if (File.Exists(temp_websource_htm))
                {
                    File.Delete(temp_websource_htm);
                }

                string url = Interaction.InputBox("输入微信推送链接", "获取封面").ToString();

                if (url != "")
                {
                    if (url.Contains("mp.weixin.qq.com"))
                    {
                        WebClient webClient = new WebClient();
                        webClient.Headers.Add("user-agent", "Mozilla/4.0 ((compatible; MSIE 8.0; Windows NT 6.1;.NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729;)");

                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)48
                                                        | (SecurityProtocolType)192
                                                        | (SecurityProtocolType)768
                                                        | (SecurityProtocolType)3072;

                        webClient.Encoding = Encoding.UTF8;

                        webClient.DownloadFile(url, temp_websource_htm);
                        //string html_string = webClient.DownloadString(url);
                    }
                    else
                    {
                        MessageBox.Show("不是微信公众号链接！");
                    }
                }
            }
            else MessageBox.Show("network error");

            if (File.Exists(temp_websource_htm))
            {
                StreamReader reader = File.OpenText(temp_websource_htm);
                string strhtml = reader.ReadToEnd();

                //提取图片连接

                int index_1 = strhtml.IndexOf("var msg_cdn_url = ");
                int index_2 = strhtml.IndexOf("var cdn_url_1_1 = ");

                strhtml = strhtml.Remove(index_2);
                strhtml = strhtml.Substring(index_1);
                // var msg_cdn_url = "https...jpeg"; // 首图idx=0时2.35:1 ， 次图idx!=0时1:1

                index_1 = strhtml.IndexOf("\"");
                index_2 = strhtml.IndexOf(";");

                strhtml = strhtml.Remove(index_2);
                strhtml = strhtml.Substring(index_1);
                // "https...jpeg"

                strhtml = strhtml.Substring(1, strhtml.Length - 2);
                // https...jpeg

                index_1 = strhtml.IndexOf("=");
                string pic_format = strhtml.Substring(index_1 + 1);
                string pic_name = temp_pic_path + "." + pic_format;

                //MessageBox.Show(pic_format);


                //下载图片

                WebClient webClient = new WebClient();
                webClient.Headers.Add("user-agent", "Mozilla/4.0 ((compatible; MSIE 8.0; Windows NT 6.1;.NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729;)");

                ServicePointManager.SecurityProtocol = (SecurityProtocolType)48
                                                | (SecurityProtocolType)192
                                                | (SecurityProtocolType)768
                                                | (SecurityProtocolType)3072;

                webClient.Encoding = Encoding.UTF8;

                webClient.DownloadFile(strhtml, pic_name);


                //图片插入至文档
                Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pic_name, false, true);

                //删除临时文件
                File.Delete(temp_websource_htm);
                File.Delete(pic_name);
            }

        }

        private void bilibiliPic_Click(object sender, RibbonControlEventArgs e)
        {
            if (GetNetStatus("bilibili.com"))
            {
                string pic_path = tempFile + "\\temp_pic_path";

                if (File.Exists(temp_websource_htm))
                {
                    File.Delete(temp_websource_htm);
                }

                if (File.Exists(pic_path))
                {
                    File.Delete(pic_path);
                }

                string url = Interaction.InputBox("输入B站视频or直播链接", "获取封面").ToString();

                if (url != "")
                {

                    if (url.Contains("bilibili.com"))
                    {
                        if (url.Contains("bilibili.com/video"))  // 视频封面
                        {
                            int index = url.IndexOf('?');
                            url = url.Remove(index);

                            if (url.Substring(0, 8) == "https://")
                            {
                                runPythonScript_bilibili(url, tempFile, "bilibili_vid_cover.py");

                                InsertPicAsync(pic_path);
                            }
                            else MessageBox.Show("请正确输入https链接");
                        }
                        else if (url.Contains("live.bilibili.com"))
                        {
                            int index = url.IndexOf('?');
                            url = url.Remove(index);
                            
                            if (url.Substring(0, 8) == "https://")
                            {
                                runPythonScript_bilibili(url, tempFile, "bilibili_live_cover.py");

                                InsertPicAsync(pic_path);
                            }
                            else MessageBox.Show("请正确输入https链接");
                        }
                        else
                        {
                            MessageBox.Show("不是b站视频or直播链接");
                        }
                    }
                    else
                    {
                        MessageBox.Show("不是B站链接！");
                    }
                }
            }
            else MessageBox.Show("network error");

        }

        public static Task WhenFileCreated(string path)
        {
            //https://www.coder.work/article/959277

            if (File.Exists(path))
                return Task.FromResult(true);

            var tcs = new TaskCompletionSource<bool>();
            FileSystemWatcher watcher = new FileSystemWatcher(Path.GetDirectoryName(path));

            FileSystemEventHandler createdHandler = null;
            RenamedEventHandler renamedHandler = null;
            createdHandler = (s, e) =>
            {
                if (e.Name == Path.GetFileName(path))
                {
                    tcs.TrySetResult(true);
                    watcher.Created -= createdHandler;
                    watcher.Dispose();
                }
            };

            renamedHandler = (s, e) =>
            {
                if (e.Name == Path.GetFileName(path))
                {
                    tcs.TrySetResult(true);
                    watcher.Renamed -= renamedHandler;
                    watcher.Dispose();
                }
            };

            watcher.Created += createdHandler;
            watcher.Renamed += renamedHandler;

            watcher.EnableRaisingEvents = true;

            return tcs.Task;
        }

        /// <summary>
        /// 当图片存在时插入图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static async Task InsertPicAsync(string filePath)
        {
            await WhenFileCreated(filePath);
            //MessageBox.Show("It's aliiiiiive!!!");
            Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(filePath, false, true);
        }


        /// <summary>
        /// 运行scripts文件夹的python脚本-bilibili
        /// </summary>
        /// <param name="url">b站资源链接</param>
        /// <param name="DicPath">下载文件的存放文件夹</param>
        /// <param name="pyscript">要运行的脚本</param>
        private void runPythonScript_bilibili(string url, string DicPath, string pyscript)
        {
            string startFile = "";
            string py_file = scriptsDic + "\\" + pyscript;
            string sArg = "";
            string catchError = "";
            JObject js_settings = ImportJSON(SettingsFile);
            if ((bool)js_settings["usePyScripts"])
            {
                startFile = @"python.exe";
                sArg = py_file + " ";
                catchError = "Python Environment Error";
            }
            else
            {
                startFile = scriptsDic + $"\\{pyscript.Substring(0, pyscript.Length - 3)}.exe";
                catchError = $"{startFile}.exe Error";
            }


            try
            {

                if (File.Exists(py_file))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = startFile;

                    string sArguments = sArg + url + " " + DicPath;

                    p.StartInfo.Arguments = sArguments;
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardInput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.StartInfo.CreateNoWindow = true;

                    p.Start();
                    //p.BeginOutputReadLine();
                    //p.OutputDataReceived += new DataReceivedEventHandler(p_OutputDataReceived);
                }
                else MessageBox.Show("python file not exist");

            }
            catch
            { 
                MessageBox.Show(catchError);
            }
        }


        private void CodeFormat4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;

            wordReplace("^l", "^p");

            Word.Table table = Globals.ThisAddIn.Application.Selection.ConvertToTable(Separator: Word.WdTableFieldSeparator.wdSeparateByParagraphs);

            JObject js = ImportJSON(PresetCode4File);

            table.Range.Font.NameFarEast = js["DefaultFontC"].ToString();
            table.Range.Font.NameAscii = js["DefaultFontE"].ToString();
            table.Range.Font.NameOther = js["DefaultFontE"].ToString();
            table.Range.Font.Name = "";
            table.Range.Font.Size = float.Parse(js["DefaultFontSize"].ToString());
            float NoSize = float.Parse(js["DefaultNoFontSize"].ToString());

            float width1 = float.Parse(js["DefaultNoWidth"].ToString());
            float width2 = Globals.ThisAddIn.Application.Selection.PageSetup.PageWidth - Globals.ThisAddIn.Application.Selection.PageSetup.LeftMargin - Globals.ThisAddIn.Application.Selection.PageSetup.RightMargin;
            width2 = width2 - width1;

            table.Columns.Add(table.Columns[1]);
            table.Columns[1].Width = width1;
            table.Columns[2].Width = width2;

            int row = table.Range.Rows.Count;
            for (int i = 1; i <= row; i++)
            {
                table.Cell(i, 1).Range.Text = "" + i;
                table.Cell(i, 1).Range.Font.Color = Word.WdColor.wdColorBlack;
                table.Cell(i, 1).Range.Font.Size = NoSize;
            }

            table.Range.Shading.BackgroundPatternColor = code4_CurrentColor;
        }

        private void SetCode4CurrentColor_Click(object sender, RibbonControlEventArgs e)
        {
            MyColorDialog.Color = Wdcolor2Color(code4_CurrentColor);
            MyColorDialog.FullOpen = true;

            DialogResult dr = MyColorDialog.ShowDialog();
            if (dr == DialogResult.OK) 
            {
                code4_CurrentColor = GetColor(MyColorDialog.Color);

                int table_select_end = Globals.ThisAddIn.Application.Selection.End;
                Word.Tables tables = Globals.ThisAddIn.Application.ActiveDocument.Range(0, table_select_end).Tables;
                int LevelOfTables = tables.Count;//从文档开始数表格数，相当于表格排序等级
                                                 //MessageBox.Show(LevelOfTables.ToString());
                Word.Table tables_select = Globals.ThisAddIn.Application.ActiveDocument.Tables[LevelOfTables];

                tables_select.Range.Shading.BackgroundPatternColor = code4_CurrentColor;
            }

        }

        private void saveCode4Color_Click(object sender, RibbonControlEventArgs e)
        {
            JObject js_code4 = ImportJSON(PresetCode4File);
            js_code4["DefaultShadingColor_r"] = Wdcolor2Color(code4_CurrentColor).R.ToString();
            js_code4["DefaultShadingColor_g"] = Wdcolor2Color(code4_CurrentColor).G.ToString();
            js_code4["DefaultShadingColor_b"] = Wdcolor2Color(code4_CurrentColor).B.ToString();

            SetjsonFun(PresetCode4File, js_code4);
        }

        private void inlineCode_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Font.Shading.BackgroundPatternColor = GetColor(Color.FromArgb(242, 242, 242));
            
            Globals.ThisAddIn.Application.Selection.Font.NameFarEast = "宋体";
            Globals.ThisAddIn.Application.Selection.Font.NameAscii = "Consolas";
            Globals.ThisAddIn.Application.Selection.Font.NameOther = "Consolas";
            Globals.ThisAddIn.Application.Selection.Font.Name = "";

            Globals.ThisAddIn.Application.Selection.Font.Shrink();
            Globals.ThisAddIn.Application.Selection.Font.Shrink();
        }

        private void runCode_Click(object sender, RibbonControlEventArgs e)
        {
            // 文字处理
            wordReplace("^l", "^p");
            wordReplace("^t", "    ");


            foreach (Word.Paragraph para in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                if (para.Format.FirstLineIndent != 0)
                {
                    para.Format.FirstLineIndent = 0;
                    //string beforePara = new string(' ', 2 * Convert.ToInt32(para.FirstLineIndent));
                    string beforePara = "    ";
                    para.Range.Text = beforePara + para.Range.Text;
                }
            }


            
            string outputComment = ">>>  lang=" + codeChoice + "  " + System.DateTime.Now.ToString("G");

            
            // 代码语言
            string code_file = "";
            if (codeChoice == "python") code_file = tempFile + "\\runCode.py";
            else if (codeChoice == "R") code_file = tempFile + "\\runCode.R";
            else if (codeChoice == "java") code_file = tempFile + "\\runCode.java";
            else MessageBox.Show("代码语言错误");


            // 转代码运行
            string codetxt = Globals.ThisAddIn.Application.Selection.Text.Replace('\r', '\n');
            

            if ((codetxt != "") && (code_file != ""))
            {
                if(!appendCodeMod.Checked)
                {
                    if (File.Exists(code_file)) File.Delete(code_file);

                    FileStream fs = new FileStream(code_file, FileMode.Create, FileAccess.ReadWrite);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.WriteLine(codetxt);
                    sw.Flush();
                    sw.Dispose();
                    sw.Close();
                    fs.Close();
                }
                else
                {
                    if (File.Exists(code_file))
                    {
                        StreamWriter sw = File.AppendText(code_file);
                        sw.WriteLine("\n" + codetxt);
                        sw.Close();
                    }
                    else
                    {
                        FileStream fs = new FileStream(code_file, FileMode.Create, FileAccess.ReadWrite);
                        StreamWriter sw = new StreamWriter(fs);
                        sw.WriteLine(codetxt);
                        sw.Flush();
                        sw.Dispose();
                        sw.Close();
                        fs.Close();
                    }
                }


                codeOutputTotal = codeOutputCount;
                codeOutputCount = 0;


                Globals.ThisAddIn.Application.Selection.InsertAfter("\r\n" + outputComment + "\r\n");

                if (codeChoice == "python") runCode_python();
                else if (codeChoice == "R") runCode_R();
                else if (codeChoice == "java") runCode_java();
                else MessageBox.Show("代码语言错误");

                Globals.ThisAddIn.Application.Selection.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                Globals.ThisAddIn.Application.Selection.Find.ClearFormatting();
                Globals.ThisAddIn.Application.Selection.Find.Replacement.ClearFormatting();
                Globals.ThisAddIn.Application.Selection.Find.Replacement.Highlight = 1;
                //Globals.ThisAddIn.Application.Selection.Find.Replacement.Font.Color = Word.WdColor.wdColorRed;
                Globals.ThisAddIn.Application.Selection.Find.Text = outputComment;
                Globals.ThisAddIn.Application.Selection.Find.Replacement.Text = "^&";
                Globals.ThisAddIn.Application.Selection.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }

        }

        /// <summary>
        /// 运行python代码
        /// </summary>
        private void runCode_python()
        {

            string py_file = tempFile + "\\runCode.py";

            try
            {
                string python_path = @"python.exe";

                if (File.Exists(py_file))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = python_path;

                    string sArguments = py_file;

                    p.StartInfo.Arguments = sArguments;
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardInput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.StartInfo.CreateNoWindow = true;

                    p.Start();
                    p.BeginOutputReadLine();
                    p.OutputDataReceived += new DataReceivedEventHandler(p_OutputDataReceived);
                }
                else MessageBox.Show("python file not exist");

            }
            catch
            {
                MessageBox.Show("python environment error");
            }
        }

        //输出打印的信息
        static void p_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Data))
            {
                AppendText(e.Data);
            }
        }
        public delegate void AppendTextCallback(string text);
        public static void AppendText(string text)
        {
            if (codeOutputCount >= codeOutputTotal) Globals.ThisAddIn.Application.Selection.InsertAfter("> " + text + "\r\n");
            codeOutputCount++;
            //MessageBox.Show("codeOutputTotal " + codeOutputTotal.ToString() + "\r\ncodeOutputCount " + codeOutputCount.ToString());
        }


        /// <summary>
        /// 选中部分进行替换
        /// </summary>
        /// <param name="ReplaceRrom">要替换的字符</param>
        /// <param name="ReplaceTo">最终替换成的字符</param>
        private void wordReplace(string ReplaceRrom, string ReplaceTo)
        {
            object Replace_String = ReplaceRrom;       //要替换的字符
            object ms = System.Type.Missing;
            object Replace = Word.WdReplace.wdReplaceAll;//设置替换方式:一，全部替换；二，只替换一个；三，一个都不替换。
            object ReplaceWith = ReplaceTo;             //最终替换成的字符
            //执行Word自带的查找/替换功能函数
            Globals.ThisAddIn.Application.Selection.Find.Execute(ref Replace_String, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ms, ref ReplaceWith, ref Replace, ref ms, ref ms, ref ms, ref ms);
        }

        private void choosePython_Click(object sender, RibbonControlEventArgs e)
        {
            codeChoice = "python";
            chooseCode.Image = Properties.Resources.python;
        }

        /// <summary>
        /// 运行R代码
        /// </summary>
        private void runCode_R()
        {
            string code_file = tempFile + "\\runCode.R";

            try
            {

                if (File.Exists(code_file))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = @"Rscript.exe";

                    string sArguments = code_file;

                    p.StartInfo.Arguments = sArguments;
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardInput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.StartInfo.CreateNoWindow = true;

                    p.Start();
                    p.BeginOutputReadLine();
                    p.OutputDataReceived += new DataReceivedEventHandler(p_OutputDataReceived);
                }
                else MessageBox.Show("R file not exist");

            }
            catch
            {
                MessageBox.Show("R environment error");
            }
        }

        private void chooseR_Click(object sender, RibbonControlEventArgs e)
        {
            codeChoice = "R";
            chooseCode.Image = Properties.Resources.R;
        }

        private void appendCodeMod_Click(object sender, RibbonControlEventArgs e)
        {
            codeOutputCount = 0;
            codeOutputTotal = 0;

            if (appendCodeMod.Checked)
            {
                string code_file = "";
                if (codeChoice == "python") code_file = tempFile + "\\runCode.py";
                else if (codeChoice == "R") code_file = tempFile + "\\runCode.R";
                else if (codeChoice == "java") code_file = tempFile + "\\runCode.java";
                else MessageBox.Show("代码语言错误");
                if (File.Exists(code_file)) File.Delete(code_file);
            }
        }

        private void bib2gbt_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择BibTeX文件";
            openFileDialog.Filter = "BibTeX文件(*.bib,*.txt)|*.bib;*.txt|BibTeX文件(*.bib)|*.bib|全部文件(*.*)|*.*";
            openFileDialog.Multiselect = true;


            DialogResult dr = openFileDialog.ShowDialog();
            if (dr == DialogResult.OK)
            {
                string startFile = "";
                string py_file = scriptsDic + "\\bib2GBT.py";
                string sArg = "";
                string catchError = "";
                JObject js_settings = ImportJSON(SettingsFile);
                if ((bool)js_settings["usePyScripts"])
                {
                    startFile = @"python.exe";
                    sArg = py_file + " ";
                    catchError = "Python Environment Error";
                }
                else
                {
                    startFile = scriptsDic + "\\bib2GBT.exe";
                    catchError = "bib2GBT.exe Error";
                }


                foreach (string InportFile in openFileDialog.FileNames)
                {
                    //MessageBox.Show(InportFile);

                    try
                    {
                    
                        if (File.Exists(py_file))
                        {
                            Process p = new Process();

                            p.StartInfo.FileName = startFile;

                            string sArguments = sArg + InportFile;
                    
                            p.StartInfo.Arguments = sArguments;
                            p.StartInfo.UseShellExecute = false;
                            p.StartInfo.RedirectStandardOutput = true;
                            p.StartInfo.RedirectStandardInput = true;
                            p.StartInfo.RedirectStandardError = true;
                            p.StartInfo.CreateNoWindow = true;
                    
                            p.Start();
                            p.BeginOutputReadLine();
                            p.OutputDataReceived += new DataReceivedEventHandler(p_OutputGBT);
                        }
                        else MessageBox.Show("python file not exist");
                    
                    }
                    catch
                    {
                        MessageBox.Show(catchError);
                    }
                }


                gbtOutputCount = 1;
            }
        }

        static void p_OutputGBT(object sender, DataReceivedEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Data))
            {
                string output = Regex.Replace(e.Data, @"(?<=\[)(\d+)(?=\])", "");

                if (output == e.Data)
                {
                    output = $"[{gbtOutputCount}] " + output;
                }
                else
                {
                    output = Regex.Replace(e.Data, @"(?<=\[)(\d+)(?=\])", gbtOutputCount.ToString());
                }

                gbtOutputCount += 1;


                Globals.ThisAddIn.Application.Selection.InsertAfter(output + "\r\n");
            }
        }

        private void chooseJava_Click(object sender, RibbonControlEventArgs e)
        {
            codeChoice = "java";
            chooseCode.Image = Properties.Resources.java;
        }

        /// <summary>
        /// 运行java代码
        /// </summary>
        private void runCode_java()
        {
            string code_file = tempFile + "\\runCode.java";

            try
            {

                if (File.Exists(code_file))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = @"java.exe";

                    string sArguments = code_file;

                    p.StartInfo.Arguments = sArguments;
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardInput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.StartInfo.CreateNoWindow = true;

                    p.Start();
                    p.BeginOutputReadLine();
                    p.OutputDataReceived += new DataReceivedEventHandler(p_OutputDataReceived);
                }
                else MessageBox.Show("java file not exist");

            }
            catch
            {
                MessageBox.Show("Java environment error");
            }
        }

        private void highlight_Click(object sender, RibbonControlEventArgs e)
        {
            int TempInt = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
            
            if (HwndPaneDic_highlight.ContainsKey(TempInt))
            {
                HwndPaneDic_highlight[TempInt].Visible = false;
                HwndPaneDic_highlight[TempInt].Dispose();
                Globals.ThisAddIn.CustomTaskPanes.Remove(HwndPaneDic_highlight[TempInt]);
                HwndPaneDic_highlight.Remove(TempInt);
            }
            else
            {
                HighlightForm highlightForm = new HighlightForm();
            
                highlightPane = Globals.ThisAddIn.CustomTaskPanes.Add(highlightForm, "Highlight");
                highlightPane.Width = 500;
                highlightPane.Visible = true;
                HwndPaneDic_highlight.Add(TempInt, highlightPane);
            }
        }
    }
}

