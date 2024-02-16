
namespace WordAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group_tuisong = this.Factory.CreateRibbonGroup();
            this.code = this.Factory.CreateRibbonGroup();
            this.CodeLatex = this.Factory.CreateRibbonGroup();
            this.Code2 = this.Factory.CreateRibbonGroup();
            this.CodeGroup4 = this.Factory.CreateRibbonGroup();
            this.ToolsBox = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.runCodeGroup = this.Factory.CreateRibbonGroup();
            this.appendCodeMod = this.Factory.CreateRibbonCheckBox();
            this.control = this.Factory.CreateRibbonGroup();
            this.button_tuisong = this.Factory.CreateRibbonButton();
            this.button_MainTitle = this.Factory.CreateRibbonButton();
            this.button_title_1 = this.Factory.CreateRibbonButton();
            this.button_title_2 = this.Factory.CreateRibbonButton();
            this.button_MainText = this.Factory.CreateRibbonButton();
            this.button_header = this.Factory.CreateRibbonButton();
            this.button_footer = this.Factory.CreateRibbonButton();
            this.button_font = this.Factory.CreateRibbonButton();
            this.button_clearFont = this.Factory.CreateRibbonButton();
            this.TimeHeader = this.Factory.CreateRibbonButton();
            this.author = this.Factory.CreateRibbonSplitButton();
            this.author_o = this.Factory.CreateRibbonButton();
            this.author_wps = this.Factory.CreateRibbonButton();
            this.author_doc = this.Factory.CreateRibbonButton();
            this.author_c = this.Factory.CreateRibbonButton();
            this.author_new = this.Factory.CreateRibbonButton();
            this.WeixinPic = this.Factory.CreateRibbonButton();
            this.bilibiliPic = this.Factory.CreateRibbonButton();
            this.CodeFormat = this.Factory.CreateRibbonButton();
            this.CodeTabSetting = this.Factory.CreateRibbonButton();
            this.CodeTabSetting2 = this.Factory.CreateRibbonButton();
            this.BorderLine = this.Factory.CreateRibbonButton();
            this.ResetCodeBGcolor = this.Factory.CreateRibbonButton();
            this.TableWidthSet = this.Factory.CreateRibbonMenu();
            this.TableWidth = this.Factory.CreateRibbonButton();
            this.CodeTableFullWidth = this.Factory.CreateRibbonButton();
            this.CodeNoWidthSet = this.Factory.CreateRibbonButton();
            this.PresetCode = this.Factory.CreateRibbonMenu();
            this.CodeOpenPreset = this.Factory.CreateRibbonButton();
            this.CodeSavePreset = this.Factory.CreateRibbonButton();
            this.CodeListNum = this.Factory.CreateRibbonButton();
            this.CodeBoderLineFun = this.Factory.CreateRibbonButton();
            this.CodeControl = this.Factory.CreateRibbonButton();
            this.CodeFormatLatex = this.Factory.CreateRibbonButton();
            this.CodeFormat2 = this.Factory.CreateRibbonButton();
            this.SetCode3CurrentColor = this.Factory.CreateRibbonButton();
            this.saveCode3Color = this.Factory.CreateRibbonButton();
            this.CodeFormat4 = this.Factory.CreateRibbonButton();
            this.SetCode4CurrentColor = this.Factory.CreateRibbonButton();
            this.saveCode4Color = this.Factory.CreateRibbonButton();
            this.ParaShadeSplit = this.Factory.CreateRibbonSplitButton();
            this.ParagraphShading = this.Factory.CreateRibbonButton();
            this.ParaShadingColorSet = this.Factory.CreateRibbonButton();
            this.DeleteShading = this.Factory.CreateRibbonButton();
            this.StyleShadeSplit = this.Factory.CreateRibbonSplitButton();
            this.styleShading = this.Factory.CreateRibbonButton();
            this.styleShadeClear = this.Factory.CreateRibbonButton();
            this.TableColoring = this.Factory.CreateRibbonButton();
            this.ThreeLine = this.Factory.CreateRibbonButton();
            this.changecharCE = this.Factory.CreateRibbonButton();
            this.checkCharMatch = this.Factory.CreateRibbonButton();
            this.del_header_line = this.Factory.CreateRibbonButton();
            this.inlineCode = this.Factory.CreateRibbonButton();
            this.FileTabOnOff = this.Factory.CreateRibbonToggleButton();
            this.bib2gbt = this.Factory.CreateRibbonButton();
            this.highlight = this.Factory.CreateRibbonButton();
            this.runCode = this.Factory.CreateRibbonButton();
            this.chooseCode = this.Factory.CreateRibbonMenu();
            this.choosePython = this.Factory.CreateRibbonButton();
            this.chooseR = this.Factory.CreateRibbonButton();
            this.chooseJava = this.Factory.CreateRibbonButton();
            this.SettingBt = this.Factory.CreateRibbonButton();
            this.About = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group_tuisong.SuspendLayout();
            this.code.SuspendLayout();
            this.CodeLatex.SuspendLayout();
            this.Code2.SuspendLayout();
            this.CodeGroup4.SuspendLayout();
            this.ToolsBox.SuspendLayout();
            this.runCodeGroup.SuspendLayout();
            this.control.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group_tuisong);
            this.tab1.Groups.Add(this.code);
            this.tab1.Groups.Add(this.CodeLatex);
            this.tab1.Groups.Add(this.Code2);
            this.tab1.Groups.Add(this.CodeGroup4);
            this.tab1.Groups.Add(this.ToolsBox);
            this.tab1.Groups.Add(this.runCodeGroup);
            this.tab1.Groups.Add(this.control);
            this.tab1.Label = "分点";
            this.tab1.Name = "tab1";
            // 
            // group_tuisong
            // 
            this.group_tuisong.Items.Add(this.button_tuisong);
            this.group_tuisong.Items.Add(this.button_MainTitle);
            this.group_tuisong.Items.Add(this.button_title_1);
            this.group_tuisong.Items.Add(this.button_title_2);
            this.group_tuisong.Items.Add(this.button_MainText);
            this.group_tuisong.Items.Add(this.button_header);
            this.group_tuisong.Items.Add(this.button_footer);
            this.group_tuisong.Items.Add(this.button_font);
            this.group_tuisong.Items.Add(this.button_clearFont);
            this.group_tuisong.Items.Add(this.TimeHeader);
            this.group_tuisong.Items.Add(this.author);
            this.group_tuisong.Items.Add(this.WeixinPic);
            this.group_tuisong.Items.Add(this.bilibiliPic);
            this.group_tuisong.Label = "文案";
            this.group_tuisong.Name = "group_tuisong";
            this.group_tuisong.Visible = false;
            // 
            // code
            // 
            this.code.Items.Add(this.CodeFormat);
            this.code.Items.Add(this.CodeTabSetting);
            this.code.Items.Add(this.CodeTabSetting2);
            this.code.Items.Add(this.BorderLine);
            this.code.Items.Add(this.ResetCodeBGcolor);
            this.code.Items.Add(this.TableWidthSet);
            this.code.Items.Add(this.PresetCode);
            this.code.Items.Add(this.CodeListNum);
            this.code.Items.Add(this.CodeBoderLineFun);
            this.code.Items.Add(this.CodeControl);
            this.code.Label = "代码排版";
            this.code.Name = "code";
            this.code.Tag = "";
            this.code.Visible = false;
            // 
            // CodeLatex
            // 
            this.CodeLatex.Items.Add(this.CodeFormatLatex);
            this.CodeLatex.Label = "代码排版2";
            this.CodeLatex.Name = "CodeLatex";
            this.CodeLatex.Visible = false;
            // 
            // Code2
            // 
            this.Code2.Items.Add(this.CodeFormat2);
            this.Code2.Items.Add(this.SetCode3CurrentColor);
            this.Code2.Items.Add(this.saveCode3Color);
            this.Code2.Label = "代码排版3";
            this.Code2.Name = "Code2";
            this.Code2.Visible = false;
            // 
            // CodeGroup4
            // 
            this.CodeGroup4.Items.Add(this.CodeFormat4);
            this.CodeGroup4.Items.Add(this.SetCode4CurrentColor);
            this.CodeGroup4.Items.Add(this.saveCode4Color);
            this.CodeGroup4.Label = "代码排版4";
            this.CodeGroup4.Name = "CodeGroup4";
            this.CodeGroup4.Visible = false;
            // 
            // ToolsBox
            // 
            this.ToolsBox.Items.Add(this.ParaShadeSplit);
            this.ToolsBox.Items.Add(this.StyleShadeSplit);
            this.ToolsBox.Items.Add(this.TableColoring);
            this.ToolsBox.Items.Add(this.ThreeLine);
            this.ToolsBox.Items.Add(this.changecharCE);
            this.ToolsBox.Items.Add(this.checkCharMatch);
            this.ToolsBox.Items.Add(this.del_header_line);
            this.ToolsBox.Items.Add(this.inlineCode);
            this.ToolsBox.Items.Add(this.bib2gbt);
            this.ToolsBox.Items.Add(this.separator1);
            this.ToolsBox.Items.Add(this.FileTabOnOff);
            this.ToolsBox.Items.Add(this.highlight);
            this.ToolsBox.Label = "工具箱";
            this.ToolsBox.Name = "ToolsBox";
            this.ToolsBox.Visible = false;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // runCodeGroup
            // 
            this.runCodeGroup.Items.Add(this.runCode);
            this.runCodeGroup.Items.Add(this.chooseCode);
            this.runCodeGroup.Items.Add(this.appendCodeMod);
            this.runCodeGroup.Label = "运行代码";
            this.runCodeGroup.Name = "runCodeGroup";
            this.runCodeGroup.Visible = false;
            // 
            // appendCodeMod
            // 
            this.appendCodeMod.Label = "追加模式";
            this.appendCodeMod.Name = "appendCodeMod";
            this.appendCodeMod.ScreenTip = "追加模式";
            this.appendCodeMod.SuperTip = "勾选后，代码将会追加到上次的运行内容之后";
            this.appendCodeMod.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.appendCodeMod_Click);
            // 
            // control
            // 
            this.control.Items.Add(this.SettingBt);
            this.control.Items.Add(this.About);
            this.control.Label = "控制";
            this.control.Name = "control";
            // 
            // button_tuisong
            // 
            this.button_tuisong.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_tuisong.Image = ((System.Drawing.Image)(resources.GetObject("button_tuisong.Image")));
            this.button_tuisong.Label = "一键排版";
            this.button_tuisong.Name = "button_tuisong";
            this.button_tuisong.ShowImage = true;
            this.button_tuisong.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_tuisong_Click);
            // 
            // button_MainTitle
            // 
            this.button_MainTitle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_MainTitle.Image = ((System.Drawing.Image)(resources.GetObject("button_MainTitle.Image")));
            this.button_MainTitle.Label = "文章标题";
            this.button_MainTitle.Name = "button_MainTitle";
            this.button_MainTitle.ScreenTip = "文章标题";
            this.button_MainTitle.ShowImage = true;
            this.button_MainTitle.SuperTip = "将段落设置为主标题，并保存至文档信息";
            this.button_MainTitle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_MainTitle_Click);
            // 
            // button_title_1
            // 
            this.button_title_1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_title_1.Image = ((System.Drawing.Image)(resources.GetObject("button_title_1.Image")));
            this.button_title_1.Label = "一级标题";
            this.button_title_1.Name = "button_title_1";
            this.button_title_1.ScreenTip = "一级标题";
            this.button_title_1.ShowImage = true;
            this.button_title_1.SuperTip = "将段落设置为一级标题";
            this.button_title_1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_title_1_Click);
            // 
            // button_title_2
            // 
            this.button_title_2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_title_2.Image = ((System.Drawing.Image)(resources.GetObject("button_title_2.Image")));
            this.button_title_2.Label = "二级标题";
            this.button_title_2.Name = "button_title_2";
            this.button_title_2.ScreenTip = "二级标题";
            this.button_title_2.ShowImage = true;
            this.button_title_2.SuperTip = "将段落设置为二级标题";
            this.button_title_2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_title_2_Click);
            // 
            // button_MainText
            // 
            this.button_MainText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_MainText.Image = ((System.Drawing.Image)(resources.GetObject("button_MainText.Image")));
            this.button_MainText.Label = "正文格式";
            this.button_MainText.Name = "button_MainText";
            this.button_MainText.ScreenTip = "正文格式";
            this.button_MainText.ShowImage = true;
            this.button_MainText.SuperTip = "将段落设置为正文";
            this.button_MainText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_MainText_Click);
            // 
            // button_header
            // 
            this.button_header.Label = "页眉";
            this.button_header.Name = "button_header";
            this.button_header.ScreenTip = "设置页眉";
            this.button_header.SuperTip = "中间页眉文字";
            this.button_header.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_header_Click);
            // 
            // button_footer
            // 
            this.button_footer.Label = "页脚";
            this.button_footer.Name = "button_footer";
            this.button_footer.ScreenTip = "设置页脚";
            this.button_footer.SuperTip = "右下角添加页脚文字";
            this.button_footer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_footer_Click);
            // 
            // button_font
            // 
            this.button_font.Label = "字体";
            this.button_font.Name = "button_font";
            this.button_font.ScreenTip = "全文字体";
            this.button_font.SuperTip = "设置全文为宋体";
            this.button_font.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_font_Click);
            // 
            // button_clearFont
            // 
            this.button_clearFont.Label = "清除格式";
            this.button_clearFont.Name = "button_clearFont";
            this.button_clearFont.ScreenTip = "清除格式";
            this.button_clearFont.SuperTip = "只留下无格式文本";
            this.button_clearFont.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_clearFont_Click);
            // 
            // TimeHeader
            // 
            this.TimeHeader.Label = "时间页眉";
            this.TimeHeader.Name = "TimeHeader";
            this.TimeHeader.ScreenTip = "设置页眉";
            this.TimeHeader.SuperTip = "右侧页眉显示上角当天日期（年/月/日）";
            this.TimeHeader.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TimeHeader_Click);
            // 
            // author
            // 
            this.author.Image = global::WordAddIn1.Properties.Resources.署名;
            this.author.Items.Add(this.author_o);
            this.author.Items.Add(this.author_wps);
            this.author.Items.Add(this.author_doc);
            this.author.Items.Add(this.author_c);
            this.author.Items.Add(this.author_new);
            this.author.Label = "署名";
            this.author.Name = "author";
            this.author.ScreenTip = "文章署名";
            this.author.SuperTip = "文末显示作者，默认使用上次的选择";
            this.author.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_Click);
            // 
            // author_o
            // 
            this.author_o.Image = ((System.Drawing.Image)(resources.GetObject("author_o.Image")));
            this.author_o.Label = "Office账户";
            this.author_o.Name = "author_o";
            this.author_o.ScreenTip = "使用Office账户名称";
            this.author_o.ShowImage = true;
            this.author_o.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_o_Click);
            // 
            // author_wps
            // 
            this.author_wps.Image = ((System.Drawing.Image)(resources.GetObject("author_wps.Image")));
            this.author_wps.Label = "WPS账户";
            this.author_wps.Name = "author_wps";
            this.author_wps.ScreenTip = "使用WPS账户名称";
            this.author_wps.ShowImage = true;
            this.author_wps.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_wps_Click);
            // 
            // author_doc
            // 
            this.author_doc.Image = ((System.Drawing.Image)(resources.GetObject("author_doc.Image")));
            this.author_doc.Label = "文档创建者";
            this.author_doc.Name = "author_doc";
            this.author_doc.ScreenTip = "使用文档创建者名称";
            this.author_doc.ShowImage = true;
            this.author_doc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_doc_Click);
            // 
            // author_c
            // 
            this.author_c.Image = ((System.Drawing.Image)(resources.GetObject("author_c.Image")));
            this.author_c.Label = "计算机账户";
            this.author_c.Name = "author_c";
            this.author_c.ScreenTip = "使用计算机账户名称";
            this.author_c.ShowImage = true;
            this.author_c.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_c_Click);
            // 
            // author_new
            // 
            this.author_new.Image = ((System.Drawing.Image)(resources.GetObject("author_new.Image")));
            this.author_new.Label = "自定义名称";
            this.author_new.Name = "author_new";
            this.author_new.ScreenTip = "使用自定义署名";
            this.author_new.ShowImage = true;
            this.author_new.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_new_Click);
            // 
            // WeixinPic
            // 
            this.WeixinPic.Image = ((System.Drawing.Image)(resources.GetObject("WeixinPic.Image")));
            this.WeixinPic.Label = "获取封面";
            this.WeixinPic.Name = "WeixinPic";
            this.WeixinPic.ScreenTip = "微信推送封面";
            this.WeixinPic.ShowImage = true;
            this.WeixinPic.SuperTip = "输入微信推送连接，插入封面图至文档选中处";
            this.WeixinPic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WeixinPic_Click);
            // 
            // bilibiliPic
            // 
            this.bilibiliPic.Image = ((System.Drawing.Image)(resources.GetObject("bilibiliPic.Image")));
            this.bilibiliPic.Label = "获取封面";
            this.bilibiliPic.Name = "bilibiliPic";
            this.bilibiliPic.ScreenTip = "B站视频or直播封面";
            this.bilibiliPic.ShowImage = true;
            this.bilibiliPic.SuperTip = "输入B站视频or直播连接，插入封面图至文档选中处";
            this.bilibiliPic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bilibiliPic_Click);
            // 
            // CodeFormat
            // 
            this.CodeFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CodeFormat.Image = ((System.Drawing.Image)(resources.GetObject("CodeFormat.Image")));
            this.CodeFormat.Label = "一键排版";
            this.CodeFormat.Name = "CodeFormat";
            this.CodeFormat.ScreenTip = "代码一键排版";
            this.CodeFormat.ShowImage = true;
            this.CodeFormat.SuperTip = "样式1";
            this.CodeFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeFormat_Click);
            // 
            // CodeTabSetting
            // 
            this.CodeTabSetting.Label = "颜色1";
            this.CodeTabSetting.Name = "CodeTabSetting";
            this.CodeTabSetting.ScreenTip = "代码背景颜色1";
            this.CodeTabSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeTabSetting_Click);
            // 
            // CodeTabSetting2
            // 
            this.CodeTabSetting2.Label = "颜色2";
            this.CodeTabSetting2.Name = "CodeTabSetting2";
            this.CodeTabSetting2.ScreenTip = "代码背景颜色2";
            this.CodeTabSetting2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeTabSetting2_Click);
            // 
            // BorderLine
            // 
            this.BorderLine.Label = "颜色3";
            this.BorderLine.Name = "BorderLine";
            this.BorderLine.ScreenTip = "竖线颜色";
            this.BorderLine.SuperTip = "一次选择一个表格";
            this.BorderLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BorderLine_Click);
            // 
            // ResetCodeBGcolor
            // 
            this.ResetCodeBGcolor.Label = "着色";
            this.ResetCodeBGcolor.Name = "ResetCodeBGcolor";
            this.ResetCodeBGcolor.ScreenTip = "设定代码背景颜色";
            this.ResetCodeBGcolor.SuperTip = "请一次选择一个表格";
            this.ResetCodeBGcolor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ResetCodeBGcolor_Click);
            // 
            // TableWidthSet
            // 
            this.TableWidthSet.Items.Add(this.TableWidth);
            this.TableWidthSet.Items.Add(this.CodeTableFullWidth);
            this.TableWidthSet.Items.Add(this.CodeNoWidthSet);
            this.TableWidthSet.Label = "宽度";
            this.TableWidthSet.Name = "TableWidthSet";
            // 
            // TableWidth
            // 
            this.TableWidth.Label = "统一宽度";
            this.TableWidth.Name = "TableWidth";
            this.TableWidth.ScreenTip = "统一所有代码块宽度，以最宽的为准";
            this.TableWidth.ShowImage = true;
            this.TableWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableWidth_Click);
            // 
            // CodeTableFullWidth
            // 
            this.CodeTableFullWidth.Label = "左右平铺";
            this.CodeTableFullWidth.Name = "CodeTableFullWidth";
            this.CodeTableFullWidth.ScreenTip = "将选中的代码块左右拉到最大";
            this.CodeTableFullWidth.ShowImage = true;
            this.CodeTableFullWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeTableFullWidth_Click);
            // 
            // CodeNoWidthSet
            // 
            this.CodeNoWidthSet.Label = "行号宽度";
            this.CodeNoWidthSet.Name = "CodeNoWidthSet";
            this.CodeNoWidthSet.ScreenTip = "设置选中的代码块的行号宽度(如果存在行号)为行号宽度的默认值";
            this.CodeNoWidthSet.ShowImage = true;
            this.CodeNoWidthSet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeNoWidthSet_Click);
            // 
            // PresetCode
            // 
            this.PresetCode.Items.Add(this.CodeOpenPreset);
            this.PresetCode.Items.Add(this.CodeSavePreset);
            this.PresetCode.Label = "预设";
            this.PresetCode.Name = "PresetCode";
            // 
            // CodeOpenPreset
            // 
            this.CodeOpenPreset.Label = "打开预设";
            this.CodeOpenPreset.Name = "CodeOpenPreset";
            this.CodeOpenPreset.ScreenTip = "打开一个预设";
            this.CodeOpenPreset.ShowImage = true;
            this.CodeOpenPreset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeOpenPreset_Click);
            // 
            // CodeSavePreset
            // 
            this.CodeSavePreset.Label = "保存预设";
            this.CodeSavePreset.Name = "CodeSavePreset";
            this.CodeSavePreset.ScreenTip = "将目前格式保存为预设";
            this.CodeSavePreset.ShowImage = true;
            this.CodeSavePreset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeSavePreset_Click);
            // 
            // CodeListNum
            // 
            this.CodeListNum.Label = "行号";
            this.CodeListNum.Name = "CodeListNum";
            this.CodeListNum.ScreenTip = "代码左侧增/减行号";
            this.CodeListNum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeListNum_Click);
            // 
            // CodeBoderLineFun
            // 
            this.CodeBoderLineFun.Label = "竖线";
            this.CodeBoderLineFun.Name = "CodeBoderLineFun";
            this.CodeBoderLineFun.ScreenTip = "增加/删除代码编号分隔线";
            this.CodeBoderLineFun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeBoderLineFun_Click);
            // 
            // CodeControl
            // 
            this.CodeControl.Label = "设置";
            this.CodeControl.Name = "CodeControl";
            this.CodeControl.ScreenTip = "代码排版的所有设置";
            this.CodeControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeControl_Click);
            // 
            // CodeFormatLatex
            // 
            this.CodeFormatLatex.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CodeFormatLatex.Image = ((System.Drawing.Image)(resources.GetObject("CodeFormatLatex.Image")));
            this.CodeFormatLatex.Label = "一键排版";
            this.CodeFormatLatex.Name = "CodeFormatLatex";
            this.CodeFormatLatex.ScreenTip = "代码一键排版";
            this.CodeFormatLatex.ShowImage = true;
            this.CodeFormatLatex.SuperTip = "样式2";
            this.CodeFormatLatex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeFormatLatex_Click);
            // 
            // CodeFormat2
            // 
            this.CodeFormat2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CodeFormat2.Image = ((System.Drawing.Image)(resources.GetObject("CodeFormat2.Image")));
            this.CodeFormat2.Label = "一键排版";
            this.CodeFormat2.Name = "CodeFormat2";
            this.CodeFormat2.ScreenTip = "代码一键排版";
            this.CodeFormat2.ShowImage = true;
            this.CodeFormat2.SuperTip = "样式3";
            this.CodeFormat2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeFormat2_Click);
            // 
            // SetCode3CurrentColor
            // 
            this.SetCode3CurrentColor.Label = "颜色";
            this.SetCode3CurrentColor.Name = "SetCode3CurrentColor";
            this.SetCode3CurrentColor.ScreenTip = "设置底纹颜色";
            this.SetCode3CurrentColor.SuperTip = "选择颜色，并且应用到本文档。注意：不会修改预设文件。";
            this.SetCode3CurrentColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetCode3CurrentColor_Click);
            // 
            // saveCode3Color
            // 
            this.saveCode3Color.Label = "保存";
            this.saveCode3Color.Name = "saveCode3Color";
            this.saveCode3Color.ScreenTip = "保存颜色";
            this.saveCode3Color.SuperTip = "将本文档使用的代码排版3的底纹颜色保存至预设文件";
            this.saveCode3Color.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveCode3Color_Click);
            // 
            // CodeFormat4
            // 
            this.CodeFormat4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CodeFormat4.Image = ((System.Drawing.Image)(resources.GetObject("CodeFormat4.Image")));
            this.CodeFormat4.Label = "一键排版";
            this.CodeFormat4.Name = "CodeFormat4";
            this.CodeFormat4.ScreenTip = "代码一键排版";
            this.CodeFormat4.ShowImage = true;
            this.CodeFormat4.SuperTip = "样式4";
            this.CodeFormat4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeFormat4_Click);
            // 
            // SetCode4CurrentColor
            // 
            this.SetCode4CurrentColor.Label = "颜色";
            this.SetCode4CurrentColor.Name = "SetCode4CurrentColor";
            this.SetCode4CurrentColor.ScreenTip = "设置底纹颜色";
            this.SetCode4CurrentColor.SuperTip = "先选中代码块。再选择颜色。更改只会应用到本文档，不会修改预设文件。";
            this.SetCode4CurrentColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetCode4CurrentColor_Click);
            // 
            // saveCode4Color
            // 
            this.saveCode4Color.Label = "保存";
            this.saveCode4Color.Name = "saveCode4Color";
            this.saveCode4Color.ScreenTip = "保存颜色";
            this.saveCode4Color.SuperTip = "将本文档使用的代码排版4的底纹颜色保存至预设文件";
            this.saveCode4Color.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveCode4Color_Click);
            // 
            // ParaShadeSplit
            // 
            this.ParaShadeSplit.Image = ((System.Drawing.Image)(resources.GetObject("ParaShadeSplit.Image")));
            this.ParaShadeSplit.Items.Add(this.ParagraphShading);
            this.ParaShadeSplit.Items.Add(this.ParaShadingColorSet);
            this.ParaShadeSplit.Items.Add(this.DeleteShading);
            this.ParaShadeSplit.Label = "段落底纹";
            this.ParaShadeSplit.Name = "ParaShadeSplit";
            this.ParaShadeSplit.ScreenTip = "段落底纹";
            this.ParaShadeSplit.SuperTip = "设置段落底纹的工具集";
            this.ParaShadeSplit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ParaShadeSplit_Click);
            // 
            // ParagraphShading
            // 
            this.ParagraphShading.Image = ((System.Drawing.Image)(resources.GetObject("ParagraphShading.Image")));
            this.ParagraphShading.Label = "段落着色";
            this.ParagraphShading.Name = "ParagraphShading";
            this.ParagraphShading.ScreenTip = "段落着色";
            this.ParagraphShading.ShowImage = true;
            this.ParagraphShading.SuperTip = "给段落加上底纹";
            this.ParagraphShading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ParagraphShading_Click);
            // 
            // ParaShadingColorSet
            // 
            this.ParaShadingColorSet.Image = ((System.Drawing.Image)(resources.GetObject("ParaShadingColorSet.Image")));
            this.ParaShadingColorSet.Label = "设置颜色";
            this.ParaShadingColorSet.Name = "ParaShadingColorSet";
            this.ParaShadingColorSet.ScreenTip = "设置颜色";
            this.ParaShadingColorSet.ShowImage = true;
            this.ParaShadingColorSet.SuperTip = "设置默认底纹颜色";
            this.ParaShadingColorSet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ParaShadingColorSet_Click);
            // 
            // DeleteShading
            // 
            this.DeleteShading.Image = ((System.Drawing.Image)(resources.GetObject("DeleteShading.Image")));
            this.DeleteShading.Label = "清除底纹";
            this.DeleteShading.Name = "DeleteShading";
            this.DeleteShading.ScreenTip = "清除底纹";
            this.DeleteShading.ShowImage = true;
            this.DeleteShading.SuperTip = "清除选择段落的底纹";
            this.DeleteShading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteShading_Click);
            // 
            // StyleShadeSplit
            // 
            this.StyleShadeSplit.Image = ((System.Drawing.Image)(resources.GetObject("StyleShadeSplit.Image")));
            this.StyleShadeSplit.Items.Add(this.styleShading);
            this.StyleShadeSplit.Items.Add(this.styleShadeClear);
            this.StyleShadeSplit.Label = "样式底纹";
            this.StyleShadeSplit.Name = "StyleShadeSplit";
            this.StyleShadeSplit.ScreenTip = "样式底纹";
            this.StyleShadeSplit.SuperTip = "设置样式底纹的工具集";
            this.StyleShadeSplit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StyleShadeSplit_Click);
            // 
            // styleShading
            // 
            this.styleShading.Image = ((System.Drawing.Image)(resources.GetObject("styleShading.Image")));
            this.styleShading.Label = "底纹颜色";
            this.styleShading.Name = "styleShading";
            this.styleShading.ScreenTip = "底纹颜色";
            this.styleShading.ShowImage = true;
            this.styleShading.SuperTip = "设置样式底纹颜色";
            this.styleShading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.styleShading_Click);
            // 
            // styleShadeClear
            // 
            this.styleShadeClear.Image = ((System.Drawing.Image)(resources.GetObject("styleShadeClear.Image")));
            this.styleShadeClear.Label = "清除底纹";
            this.styleShadeClear.Name = "styleShadeClear";
            this.styleShadeClear.ScreenTip = "清除底纹";
            this.styleShadeClear.ShowImage = true;
            this.styleShadeClear.SuperTip = "删除样式中的底纹";
            this.styleShadeClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.styleShadeClear_Click);
            // 
            // TableColoring
            // 
            this.TableColoring.Image = ((System.Drawing.Image)(resources.GetObject("TableColoring.Image")));
            this.TableColoring.Label = "表格着色";
            this.TableColoring.Name = "TableColoring";
            this.TableColoring.ScreenTip = "表格着色";
            this.TableColoring.ShowImage = true;
            this.TableColoring.SuperTip = "打开\\关闭表格颜色设置面板";
            this.TableColoring.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableColoring_Click);
            // 
            // ThreeLine
            // 
            this.ThreeLine.Image = ((System.Drawing.Image)(resources.GetObject("ThreeLine.Image")));
            this.ThreeLine.Label = "三线表";
            this.ThreeLine.Name = "ThreeLine";
            this.ThreeLine.ScreenTip = "将表格转为三线表";
            this.ThreeLine.ShowImage = true;
            this.ThreeLine.SuperTip = "不能有合并单元格";
            this.ThreeLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ThreeLine_Click);
            // 
            // changecharCE
            // 
            this.changecharCE.Image = ((System.Drawing.Image)(resources.GetObject("changecharCE.Image")));
            this.changecharCE.Label = "替换符号";
            this.changecharCE.Name = "changecharCE";
            this.changecharCE.ScreenTip = "替换符号";
            this.changecharCE.ShowImage = true;
            this.changecharCE.SuperTip = "打开\\关闭面板、将选中文字的英文标点换成中文标点";
            this.changecharCE.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changecharCE_Click);
            // 
            // checkCharMatch
            // 
            this.checkCharMatch.Image = ((System.Drawing.Image)(resources.GetObject("checkCharMatch.Image")));
            this.checkCharMatch.Label = "符号匹配";
            this.checkCharMatch.Name = "checkCharMatch";
            this.checkCharMatch.ScreenTip = "符号匹配";
            this.checkCharMatch.ShowImage = true;
            this.checkCharMatch.SuperTip = "打开\\关闭面板，检查符号是否配对";
            this.checkCharMatch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkCharMatch_Click);
            // 
            // del_header_line
            // 
            this.del_header_line.Image = ((System.Drawing.Image)(resources.GetObject("del_header_line.Image")));
            this.del_header_line.Label = "页眉横线";
            this.del_header_line.Name = "del_header_line";
            this.del_header_line.ScreenTip = "去除页眉横线";
            this.del_header_line.ShowImage = true;
            this.del_header_line.SuperTip = "去除页眉横线";
            this.del_header_line.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.del_header_line_Click);
            // 
            // inlineCode
            // 
            this.inlineCode.Image = ((System.Drawing.Image)(resources.GetObject("inlineCode.Image")));
            this.inlineCode.Label = "行内代码";
            this.inlineCode.Name = "inlineCode";
            this.inlineCode.ScreenTip = "行内代码";
            this.inlineCode.ShowImage = true;
            this.inlineCode.SuperTip = "选中文字，按照行内代码样式排版";
            this.inlineCode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.inlineCode_Click);
            // 
            // FileTabOnOff
            // 
            this.FileTabOnOff.Image = ((System.Drawing.Image)(resources.GetObject("FileTabOnOff.Image")));
            this.FileTabOnOff.Label = "标签栏";
            this.FileTabOnOff.Name = "FileTabOnOff";
            this.FileTabOnOff.ScreenTip = "【实验性】标签栏";
            this.FileTabOnOff.ShowImage = true;
            this.FileTabOnOff.SuperTip = "在多个文档间切换（请注意及时保存文档）";
            this.FileTabOnOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FileTabOnOff_Click);
            // 
            // bib2gbt
            // 
            this.bib2gbt.Image = global::WordAddIn1.Properties.Resources.bib2gbt;
            this.bib2gbt.Label = "BibTeX";
            this.bib2gbt.Name = "bib2gbt";
            this.bib2gbt.ScreenTip = "【实验性】将BibTeX转换成GBT格式";
            this.bib2gbt.ShowImage = true;
            this.bib2gbt.SuperTip = "将BibTeX转换成GBT7714-2015格式，并插入到选中位置.";
            this.bib2gbt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bib2gbt_Click);
            // 
            // highlight
            // 
            this.highlight.Image = global::WordAddIn1.Properties.Resources.highlight;
            this.highlight.Label = "代码高亮";
            this.highlight.Name = "highlight";
            this.highlight.ScreenTip = "【实验性】代码高亮";
            this.highlight.ShowImage = true;
            this.highlight.SuperTip = "打开代码高亮面板";
            this.highlight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.highlight_Click);
            // 
            // runCode
            // 
            this.runCode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runCode.Image = global::WordAddIn1.Properties.Resources.codeRun;
            this.runCode.Label = "运行代码";
            this.runCode.Name = "runCode";
            this.runCode.ScreenTip = "运行代码";
            this.runCode.ShowImage = true;
            this.runCode.SuperTip = "选中代码，点击运行。打印结果到选段下方。\n\n注意：需要相应编译环境";
            this.runCode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runCode_Click);
            // 
            // chooseCode
            // 
            this.chooseCode.Image = ((System.Drawing.Image)(resources.GetObject("chooseCode.Image")));
            this.chooseCode.Items.Add(this.choosePython);
            this.chooseCode.Items.Add(this.chooseR);
            this.chooseCode.Items.Add(this.chooseJava);
            this.chooseCode.Label = "当前语言";
            this.chooseCode.Name = "chooseCode";
            this.chooseCode.ScreenTip = "选择代码语言";
            this.chooseCode.ShowImage = true;
            this.chooseCode.SuperTip = "在下拉列表中选择需要的代码语言";
            // 
            // choosePython
            // 
            this.choosePython.Image = global::WordAddIn1.Properties.Resources.python;
            this.choosePython.Label = "python";
            this.choosePython.Name = "choosePython";
            this.choosePython.ScreenTip = "Python";
            this.choosePython.ShowImage = true;
            this.choosePython.SuperTip = "需要 Python 环境";
            this.choosePython.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.choosePython_Click);
            // 
            // chooseR
            // 
            this.chooseR.Image = global::WordAddIn1.Properties.Resources.R;
            this.chooseR.Label = "R";
            this.chooseR.Name = "chooseR";
            this.chooseR.ScreenTip = "R";
            this.chooseR.ShowImage = true;
            this.chooseR.SuperTip = "需要 R 环境";
            this.chooseR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chooseR_Click);
            // 
            // chooseJava
            // 
            this.chooseJava.Image = global::WordAddIn1.Properties.Resources.java;
            this.chooseJava.Label = "java";
            this.chooseJava.Name = "chooseJava";
            this.chooseJava.ScreenTip = "Java";
            this.chooseJava.ShowImage = true;
            this.chooseJava.SuperTip = "需要 Java 环境";
            this.chooseJava.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chooseJava_Click);
            // 
            // SettingBt
            // 
            this.SettingBt.Image = global::WordAddIn1.Properties.Resources.settings;
            this.SettingBt.Label = "设置";
            this.SettingBt.Name = "SettingBt";
            this.SettingBt.ScreenTip = "插件设置";
            this.SettingBt.ShowImage = true;
            this.SettingBt.SuperTip = "打开插件设置";
            this.SettingBt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SettingBt_Click);
            // 
            // About
            // 
            this.About.Image = ((System.Drawing.Image)(resources.GetObject("About.Image")));
            this.About.Label = "关于";
            this.About.Name = "About";
            this.About.ShowImage = true;
            this.About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.About_Click);
            // 
            // button1
            // 
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group_tuisong.ResumeLayout(false);
            this.group_tuisong.PerformLayout();
            this.code.ResumeLayout(false);
            this.code.PerformLayout();
            this.CodeLatex.ResumeLayout(false);
            this.CodeLatex.PerformLayout();
            this.Code2.ResumeLayout(false);
            this.Code2.PerformLayout();
            this.CodeGroup4.ResumeLayout(false);
            this.CodeGroup4.PerformLayout();
            this.ToolsBox.ResumeLayout(false);
            this.ToolsBox.PerformLayout();
            this.runCodeGroup.ResumeLayout(false);
            this.runCodeGroup.PerformLayout();
            this.control.ResumeLayout(false);
            this.control.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_tuisong;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_header;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_footer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_font;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_MainTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_title_1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_title_2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_MainText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_clearFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup code;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup control;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton author;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton author_o;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton author_wps;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton author_doc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton author_c;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeTabSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton About;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeTabSetting2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ResetCodeBGcolor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BorderLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TableWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu PresetCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeOpenPreset;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeSavePreset;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeListNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeBoderLineFun;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu TableWidthSet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeTableFullWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeNoWidthSet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ToolsBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ParagraphShading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ParaShadingColorSet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Code2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeFormat2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteShading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ThreeLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CodeLatex;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeFormatLatex;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton FileTabOnOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TableColoring;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TimeHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton author_new;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton changecharCE;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton checkCharMatch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton del_header_line;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetCode3CurrentColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveCode3Color;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton ParaShadeSplit;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton StyleShadeSplit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton styleShading;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton styleShadeClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SettingBt;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_tuisong;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WeixinPic;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bilibiliPic;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CodeGroup4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CodeFormat4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetCode4CurrentColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveCode4Color;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton inlineCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup runCodeGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu chooseCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton choosePython;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton chooseR;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox appendCodeMod;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bib2gbt;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton chooseJava;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton highlight;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
