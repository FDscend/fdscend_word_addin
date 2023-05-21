
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
            this.ToolsBox = this.Factory.CreateRibbonGroup();
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
            this.ParagraphShading = this.Factory.CreateRibbonButton();
            this.ParaShadingColorSet = this.Factory.CreateRibbonButton();
            this.DeleteShading = this.Factory.CreateRibbonButton();
            this.TableColoring = this.Factory.CreateRibbonToggleButton();
            this.ThreeLine = this.Factory.CreateRibbonButton();
            this.FileTabOnOff = this.Factory.CreateRibbonToggleButton();
            this.changecharCE = this.Factory.CreateRibbonToggleButton();
            this.checkCharMatch = this.Factory.CreateRibbonToggleButton();
            this.del_header_line = this.Factory.CreateRibbonButton();
            this.ControlActive = this.Factory.CreateRibbonButton();
            this.hide = this.Factory.CreateRibbonButton();
            this.showHidden = this.Factory.CreateRibbonButton();
            this.About = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group_tuisong.SuspendLayout();
            this.code.SuspendLayout();
            this.CodeLatex.SuspendLayout();
            this.Code2.SuspendLayout();
            this.ToolsBox.SuspendLayout();
            this.control.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group_tuisong);
            this.tab1.Groups.Add(this.code);
            this.tab1.Groups.Add(this.CodeLatex);
            this.tab1.Groups.Add(this.Code2);
            this.tab1.Groups.Add(this.ToolsBox);
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
            this.Code2.Label = "代码排版3";
            this.Code2.Name = "Code2";
            this.Code2.Visible = false;
            // 
            // ToolsBox
            // 
            this.ToolsBox.Items.Add(this.ParagraphShading);
            this.ToolsBox.Items.Add(this.ParaShadingColorSet);
            this.ToolsBox.Items.Add(this.DeleteShading);
            this.ToolsBox.Items.Add(this.TableColoring);
            this.ToolsBox.Items.Add(this.ThreeLine);
            this.ToolsBox.Items.Add(this.FileTabOnOff);
            this.ToolsBox.Items.Add(this.changecharCE);
            this.ToolsBox.Items.Add(this.checkCharMatch);
            this.ToolsBox.Items.Add(this.del_header_line);
            this.ToolsBox.Label = "工具箱";
            this.ToolsBox.Name = "ToolsBox";
            this.ToolsBox.Visible = false;
            // 
            // control
            // 
            this.control.Items.Add(this.ControlActive);
            this.control.Items.Add(this.hide);
            this.control.Items.Add(this.showHidden);
            this.control.Items.Add(this.About);
            this.control.Label = "控制";
            this.control.Name = "control";
            // 
            // button_tuisong
            // 
            this.button_tuisong.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_tuisong.Image = global::WordAddIn1.Properties.Resources._2;
            this.button_tuisong.Label = "一键排版";
            this.button_tuisong.Name = "button_tuisong";
            this.button_tuisong.ShowImage = true;
            this.button_tuisong.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_tuisong_Click);
            // 
            // button_MainTitle
            // 
            this.button_MainTitle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_MainTitle.Image = global::WordAddIn1.Properties.Resources.主标题;
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
            this.button_title_1.Image = global::WordAddIn1.Properties.Resources.一级标题;
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
            this.button_title_2.Image = global::WordAddIn1.Properties.Resources.二级标题;
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
            this.button_MainText.Image = global::WordAddIn1.Properties.Resources.正文;
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
            this.author_o.Image = global::WordAddIn1.Properties.Resources.office;
            this.author_o.Label = "Office账户";
            this.author_o.Name = "author_o";
            this.author_o.ScreenTip = "使用Office账户名称";
            this.author_o.ShowImage = true;
            this.author_o.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_o_Click);
            // 
            // author_wps
            // 
            this.author_wps.Image = global::WordAddIn1.Properties.Resources.wps;
            this.author_wps.Label = "WPS账户";
            this.author_wps.Name = "author_wps";
            this.author_wps.ScreenTip = "使用WPS账户名称";
            this.author_wps.ShowImage = true;
            this.author_wps.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_wps_Click);
            // 
            // author_doc
            // 
            this.author_doc.Image = global::WordAddIn1.Properties.Resources.document;
            this.author_doc.Label = "文档创建者";
            this.author_doc.Name = "author_doc";
            this.author_doc.ScreenTip = "使用文档创建者名称";
            this.author_doc.ShowImage = true;
            this.author_doc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_doc_Click);
            // 
            // author_c
            // 
            this.author_c.Image = global::WordAddIn1.Properties.Resources.computer;
            this.author_c.Label = "计算机账户";
            this.author_c.Name = "author_c";
            this.author_c.ScreenTip = "使用计算机账户名称";
            this.author_c.ShowImage = true;
            this.author_c.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_c_Click);
            // 
            // author_new
            // 
            this.author_new.Image = global::WordAddIn1.Properties.Resources.add;
            this.author_new.Label = "自定义名称";
            this.author_new.Name = "author_new";
            this.author_new.ScreenTip = "使用自定义署名";
            this.author_new.ShowImage = true;
            this.author_new.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.author_new_Click);
            // 
            // CodeFormat
            // 
            this.CodeFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CodeFormat.Image = global::WordAddIn1.Properties.Resources._2;
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
            this.CodeFormatLatex.Image = global::WordAddIn1.Properties.Resources._2;
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
            this.CodeFormat2.Image = global::WordAddIn1.Properties.Resources._2;
            this.CodeFormat2.Label = "一键排版";
            this.CodeFormat2.Name = "CodeFormat2";
            this.CodeFormat2.ScreenTip = "代码一键排版";
            this.CodeFormat2.ShowImage = true;
            this.CodeFormat2.SuperTip = "样式3";
            this.CodeFormat2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CodeFormat2_Click);
            // 
            // ParagraphShading
            // 
            this.ParagraphShading.Image = global::WordAddIn1.Properties.Resources.底纹;
            this.ParagraphShading.Label = "段落底纹";
            this.ParagraphShading.Name = "ParagraphShading";
            this.ParagraphShading.ScreenTip = "给段落加上底纹";
            this.ParagraphShading.ShowImage = true;
            this.ParagraphShading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ParagraphShading_Click);
            // 
            // ParaShadingColorSet
            // 
            this.ParaShadingColorSet.Image = ((System.Drawing.Image)(resources.GetObject("ParaShadingColorSet.Image")));
            this.ParaShadingColorSet.Label = "底纹颜色";
            this.ParaShadingColorSet.Name = "ParaShadingColorSet";
            this.ParaShadingColorSet.ScreenTip = "设置默认底纹颜色";
            this.ParaShadingColorSet.ShowImage = true;
            this.ParaShadingColorSet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ParaShadingColorSet_Click);
            // 
            // DeleteShading
            // 
            this.DeleteShading.Image = global::WordAddIn1.Properties.Resources.删除底纹;
            this.DeleteShading.Label = "清除底纹";
            this.DeleteShading.Name = "DeleteShading";
            this.DeleteShading.ScreenTip = "清除选择段落的底纹";
            this.DeleteShading.ShowImage = true;
            this.DeleteShading.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteShading_Click);
            // 
            // TableColoring
            // 
            this.TableColoring.Image = global::WordAddIn1.Properties.Resources.表格着色;
            this.TableColoring.Label = "表格着色";
            this.TableColoring.Name = "TableColoring";
            this.TableColoring.ScreenTip = "打开表格颜色设置面板";
            this.TableColoring.ShowImage = true;
            this.TableColoring.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableColoring_Click);
            // 
            // ThreeLine
            // 
            this.ThreeLine.Image = global::WordAddIn1.Properties.Resources.三线表;
            this.ThreeLine.Label = "三线表";
            this.ThreeLine.Name = "ThreeLine";
            this.ThreeLine.ScreenTip = "将表格转为三线表";
            this.ThreeLine.ShowImage = true;
            this.ThreeLine.SuperTip = "不能有合并单元格";
            this.ThreeLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ThreeLine_Click);
            // 
            // FileTabOnOff
            // 
            this.FileTabOnOff.Image = global::WordAddIn1.Properties.Resources.tab;
            this.FileTabOnOff.Label = "标签栏";
            this.FileTabOnOff.Name = "FileTabOnOff";
            this.FileTabOnOff.ScreenTip = "标签栏";
            this.FileTabOnOff.ShowImage = true;
            this.FileTabOnOff.SuperTip = "请注意及时保存文档";
            this.FileTabOnOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FileTabOnOff_Click);
            // 
            // changecharCE
            // 
            this.changecharCE.Image = global::WordAddIn1.Properties.Resources.替换;
            this.changecharCE.Label = "替换符号";
            this.changecharCE.Name = "changecharCE";
            this.changecharCE.ScreenTip = "替换符号";
            this.changecharCE.ShowImage = true;
            this.changecharCE.SuperTip = "将英文标点换成中文标点";
            this.changecharCE.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changecharCE_Click);
            // 
            // checkCharMatch
            // 
            this.checkCharMatch.Image = global::WordAddIn1.Properties.Resources.pair;
            this.checkCharMatch.Label = "符号匹配";
            this.checkCharMatch.Name = "checkCharMatch";
            this.checkCharMatch.ScreenTip = "符号匹配";
            this.checkCharMatch.ShowImage = true;
            this.checkCharMatch.SuperTip = "检查符号是否配对";
            this.checkCharMatch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkCharMatch_Click);
            // 
            // del_header_line
            // 
            this.del_header_line.Image = global::WordAddIn1.Properties.Resources.页眉横线;
            this.del_header_line.Label = "页眉横线";
            this.del_header_line.Name = "del_header_line";
            this.del_header_line.ScreenTip = "去除页眉横线";
            this.del_header_line.ShowImage = true;
            this.del_header_line.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.del_header_line_Click);
            // 
            // ControlActive
            // 
            this.ControlActive.Label = "增加";
            this.ControlActive.Name = "ControlActive";
            this.ControlActive.ScreenTip = "增加功能";
            this.ControlActive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ControlActive_Click);
            // 
            // hide
            // 
            this.hide.Label = "隐藏";
            this.hide.Name = "hide";
            this.hide.ScreenTip = "隐藏功能";
            this.hide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Hide_Click);
            // 
            // showHidden
            // 
            this.showHidden.Label = "展示";
            this.showHidden.Name = "showHidden";
            this.showHidden.ScreenTip = "显示隐藏的功能";
            this.showHidden.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowHidden_Click);
            // 
            // About
            // 
            this.About.Image = global::WordAddIn1.Properties.Resources.关于;
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
            this.ToolsBox.ResumeLayout(false);
            this.ToolsBox.PerformLayout();
            this.control.ResumeLayout(false);
            this.control.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_tuisong;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ControlActive;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton hide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showHidden;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton TableColoring;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TimeHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton author_new;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton changecharCE;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton checkCharMatch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton del_header_line;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
