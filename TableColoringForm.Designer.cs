
namespace WordAddIn1
{
    partial class TableColoringForm
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.groupTableColor = new System.Windows.Forms.GroupBox();
            this.TableColorPresetList = new System.Windows.Forms.ListBox();
            this.groupColorEffect = new System.Windows.Forms.GroupBox();
            this.TableLine6 = new System.Windows.Forms.TextBox();
            this.TableLine5 = new System.Windows.Forms.TextBox();
            this.TableLine4 = new System.Windows.Forms.TextBox();
            this.TableLine3 = new System.Windows.Forms.TextBox();
            this.TableLine2 = new System.Windows.Forms.TextBox();
            this.TableLine1 = new System.Windows.Forms.TextBox();
            this.ColorOK = new System.Windows.Forms.Button();
            this.ExchangeColor = new System.Windows.Forms.Button();
            this.NewColor = new System.Windows.Forms.Button();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.NewColor2 = new System.Windows.Forms.Button();
            this.SavePreset = new System.Windows.Forms.Button();
            this.DeletePreset = new System.Windows.Forms.Button();
            this.checkBoxFirstLine = new System.Windows.Forms.CheckBox();
            this.Export = new System.Windows.Forms.Button();
            this.saveFileDialogExPreset = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialogINPreset = new System.Windows.Forms.OpenFileDialog();
            this.InportPresetGroup = new System.Windows.Forms.GroupBox();
            this.InportOK = new System.Windows.Forms.Button();
            this.InportPresetList = new System.Windows.Forms.ListBox();
            this.checkBoxInport = new System.Windows.Forms.CheckBox();
            this.NewColorH = new System.Windows.Forms.Button();
            this.groupTableColor.SuspendLayout();
            this.groupColorEffect.SuspendLayout();
            this.InportPresetGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupTableColor
            // 
            this.groupTableColor.BackColor = System.Drawing.Color.OldLace;
            this.groupTableColor.Controls.Add(this.TableColorPresetList);
            this.groupTableColor.Location = new System.Drawing.Point(4, 4);
            this.groupTableColor.Margin = new System.Windows.Forms.Padding(4);
            this.groupTableColor.Name = "groupTableColor";
            this.groupTableColor.Padding = new System.Windows.Forms.Padding(4);
            this.groupTableColor.Size = new System.Drawing.Size(309, 262);
            this.groupTableColor.TabIndex = 0;
            this.groupTableColor.TabStop = false;
            this.groupTableColor.Text = "表格颜色预设";
            // 
            // TableColorPresetList
            // 
            this.TableColorPresetList.FormattingEnabled = true;
            this.TableColorPresetList.ItemHeight = 24;
            this.TableColorPresetList.Location = new System.Drawing.Point(7, 35);
            this.TableColorPresetList.Name = "TableColorPresetList";
            this.TableColorPresetList.Size = new System.Drawing.Size(295, 220);
            this.TableColorPresetList.TabIndex = 0;
            this.TableColorPresetList.SelectedIndexChanged += new System.EventHandler(this.TableColorPresetList_SelectedIndexChanged);
            // 
            // groupColorEffect
            // 
            this.groupColorEffect.BackColor = System.Drawing.Color.OldLace;
            this.groupColorEffect.Controls.Add(this.TableLine6);
            this.groupColorEffect.Controls.Add(this.TableLine5);
            this.groupColorEffect.Controls.Add(this.TableLine4);
            this.groupColorEffect.Controls.Add(this.TableLine3);
            this.groupColorEffect.Controls.Add(this.TableLine2);
            this.groupColorEffect.Controls.Add(this.TableLine1);
            this.groupColorEffect.Location = new System.Drawing.Point(4, 273);
            this.groupColorEffect.Name = "groupColorEffect";
            this.groupColorEffect.Size = new System.Drawing.Size(309, 206);
            this.groupColorEffect.TabIndex = 1;
            this.groupColorEffect.TabStop = false;
            this.groupColorEffect.Text = "效果预览";
            // 
            // TableLine6
            // 
            this.TableLine6.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TableLine6.Location = new System.Drawing.Point(7, 174);
            this.TableLine6.Multiline = true;
            this.TableLine6.Name = "TableLine6";
            this.TableLine6.Size = new System.Drawing.Size(295, 28);
            this.TableLine6.TabIndex = 5;
            this.TableLine6.Text = "ヾ(≧▽≦*)o ";
            // 
            // TableLine5
            // 
            this.TableLine5.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TableLine5.Location = new System.Drawing.Point(7, 146);
            this.TableLine5.Multiline = true;
            this.TableLine5.Name = "TableLine5";
            this.TableLine5.Size = new System.Drawing.Size(295, 28);
            this.TableLine5.TabIndex = 4;
            this.TableLine5.Text = "ヾ(^▽^*))) ";
            // 
            // TableLine4
            // 
            this.TableLine4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TableLine4.Location = new System.Drawing.Point(7, 118);
            this.TableLine4.Multiline = true;
            this.TableLine4.Name = "TableLine4";
            this.TableLine4.Size = new System.Drawing.Size(295, 28);
            this.TableLine4.TabIndex = 3;
            this.TableLine4.Text = "(′▽`〃)";
            // 
            // TableLine3
            // 
            this.TableLine3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TableLine3.Location = new System.Drawing.Point(7, 90);
            this.TableLine3.Multiline = true;
            this.TableLine3.Name = "TableLine3";
            this.TableLine3.Size = new System.Drawing.Size(295, 28);
            this.TableLine3.TabIndex = 2;
            this.TableLine3.Text = " (*ฅ́ˇฅ̀*) ";
            // 
            // TableLine2
            // 
            this.TableLine2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TableLine2.Location = new System.Drawing.Point(7, 62);
            this.TableLine2.Multiline = true;
            this.TableLine2.Name = "TableLine2";
            this.TableLine2.Size = new System.Drawing.Size(295, 28);
            this.TableLine2.TabIndex = 1;
            this.TableLine2.Text = "(>▽<)";
            // 
            // TableLine1
            // 
            this.TableLine1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TableLine1.Location = new System.Drawing.Point(7, 34);
            this.TableLine1.Multiline = true;
            this.TableLine1.Name = "TableLine1";
            this.TableLine1.Size = new System.Drawing.Size(295, 28);
            this.TableLine1.TabIndex = 0;
            this.TableLine1.Text = "(๑•̀ㅂ•́)و✧";
            // 
            // ColorOK
            // 
            this.ColorOK.BackColor = System.Drawing.Color.OldLace;
            this.ColorOK.Location = new System.Drawing.Point(4, 486);
            this.ColorOK.Name = "ColorOK";
            this.ColorOK.Size = new System.Drawing.Size(127, 39);
            this.ColorOK.TabIndex = 2;
            this.ColorOK.Text = "*使用*";
            this.ColorOK.UseVisualStyleBackColor = false;
            this.ColorOK.Click += new System.EventHandler(this.ColorOK_Click);
            // 
            // ExchangeColor
            // 
            this.ExchangeColor.BackColor = System.Drawing.Color.OldLace;
            this.ExchangeColor.Location = new System.Drawing.Point(137, 486);
            this.ExchangeColor.Name = "ExchangeColor";
            this.ExchangeColor.Size = new System.Drawing.Size(102, 39);
            this.ExchangeColor.TabIndex = 3;
            this.ExchangeColor.Text = "交换";
            this.ExchangeColor.UseVisualStyleBackColor = false;
            this.ExchangeColor.Click += new System.EventHandler(this.ExchangeColor_Click);
            // 
            // NewColor
            // 
            this.NewColor.BackColor = System.Drawing.Color.OldLace;
            this.NewColor.Location = new System.Drawing.Point(4, 531);
            this.NewColor.Name = "NewColor";
            this.NewColor.Size = new System.Drawing.Size(78, 39);
            this.NewColor.TabIndex = 4;
            this.NewColor.Text = "颜色1";
            this.NewColor.UseVisualStyleBackColor = false;
            this.NewColor.Click += new System.EventHandler(this.NewColor_Click);
            // 
            // NewColor2
            // 
            this.NewColor2.BackColor = System.Drawing.Color.OldLace;
            this.NewColor2.Location = new System.Drawing.Point(82, 531);
            this.NewColor2.Name = "NewColor2";
            this.NewColor2.Size = new System.Drawing.Size(78, 39);
            this.NewColor2.TabIndex = 5;
            this.NewColor2.Text = "颜色2";
            this.NewColor2.UseVisualStyleBackColor = false;
            this.NewColor2.Click += new System.EventHandler(this.NewColor2_Click);
            // 
            // SavePreset
            // 
            this.SavePreset.BackColor = System.Drawing.Color.BlanchedAlmond;
            this.SavePreset.Location = new System.Drawing.Point(247, 486);
            this.SavePreset.Name = "SavePreset";
            this.SavePreset.Size = new System.Drawing.Size(66, 39);
            this.SavePreset.TabIndex = 6;
            this.SavePreset.Text = "保存";
            this.SavePreset.UseVisualStyleBackColor = false;
            this.SavePreset.Click += new System.EventHandler(this.SavePreset_Click);
            // 
            // DeletePreset
            // 
            this.DeletePreset.BackColor = System.Drawing.Color.BlanchedAlmond;
            this.DeletePreset.Location = new System.Drawing.Point(247, 531);
            this.DeletePreset.Name = "DeletePreset";
            this.DeletePreset.Size = new System.Drawing.Size(66, 39);
            this.DeletePreset.TabIndex = 7;
            this.DeletePreset.Text = "删除";
            this.DeletePreset.UseVisualStyleBackColor = false;
            this.DeletePreset.Click += new System.EventHandler(this.DeletePreset_Click);
            // 
            // checkBoxFirstLine
            // 
            this.checkBoxFirstLine.Appearance = System.Windows.Forms.Appearance.Button;
            this.checkBoxFirstLine.BackColor = System.Drawing.Color.Cornsilk;
            this.checkBoxFirstLine.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.checkBoxFirstLine.Location = new System.Drawing.Point(4, 576);
            this.checkBoxFirstLine.Name = "checkBoxFirstLine";
            this.checkBoxFirstLine.Size = new System.Drawing.Size(116, 39);
            this.checkBoxFirstLine.TabIndex = 8;
            this.checkBoxFirstLine.Text = "排除首行";
            this.checkBoxFirstLine.UseVisualStyleBackColor = false;
            this.checkBoxFirstLine.CheckedChanged += new System.EventHandler(this.checkBoxFirstLine_CheckedChanged);
            // 
            // Export
            // 
            this.Export.BackColor = System.Drawing.Color.Moccasin;
            this.Export.Location = new System.Drawing.Point(126, 576);
            this.Export.Name = "Export";
            this.Export.Size = new System.Drawing.Size(93, 39);
            this.Export.TabIndex = 9;
            this.Export.Text = "导出";
            this.Export.UseVisualStyleBackColor = false;
            this.Export.Click += new System.EventHandler(this.Export_Click);
            // 
            // saveFileDialogExPreset
            // 
            this.saveFileDialogExPreset.FileName = "FDscend PresetTable";
            this.saveFileDialogExPreset.Title = "导出预设";
            // 
            // openFileDialogINPreset
            // 
            this.openFileDialogINPreset.FileName = "FDscend PresetTable";
            // 
            // InportPresetGroup
            // 
            this.InportPresetGroup.BackColor = System.Drawing.Color.Moccasin;
            this.InportPresetGroup.Controls.Add(this.InportOK);
            this.InportPresetGroup.Controls.Add(this.InportPresetList);
            this.InportPresetGroup.Location = new System.Drawing.Point(4, 625);
            this.InportPresetGroup.Name = "InportPresetGroup";
            this.InportPresetGroup.Size = new System.Drawing.Size(309, 146);
            this.InportPresetGroup.TabIndex = 11;
            this.InportPresetGroup.TabStop = false;
            this.InportPresetGroup.Text = "导入预设列表";
            this.InportPresetGroup.Visible = false;
            // 
            // InportOK
            // 
            this.InportOK.BackColor = System.Drawing.Color.OldLace;
            this.InportOK.Location = new System.Drawing.Point(221, 35);
            this.InportOK.Name = "InportOK";
            this.InportOK.Size = new System.Drawing.Size(82, 100);
            this.InportOK.TabIndex = 1;
            this.InportOK.Text = "确认";
            this.InportOK.UseVisualStyleBackColor = false;
            this.InportOK.Click += new System.EventHandler(this.InportOK_Click);
            // 
            // InportPresetList
            // 
            this.InportPresetList.FormattingEnabled = true;
            this.InportPresetList.ItemHeight = 24;
            this.InportPresetList.Location = new System.Drawing.Point(7, 35);
            this.InportPresetList.Name = "InportPresetList";
            this.InportPresetList.Size = new System.Drawing.Size(208, 100);
            this.InportPresetList.TabIndex = 0;
            // 
            // checkBoxInport
            // 
            this.checkBoxInport.Appearance = System.Windows.Forms.Appearance.Button;
            this.checkBoxInport.BackColor = System.Drawing.Color.Moccasin;
            this.checkBoxInport.Location = new System.Drawing.Point(220, 576);
            this.checkBoxInport.Name = "checkBoxInport";
            this.checkBoxInport.Size = new System.Drawing.Size(93, 39);
            this.checkBoxInport.TabIndex = 12;
            this.checkBoxInport.Text = "导入";
            this.checkBoxInport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.checkBoxInport.UseVisualStyleBackColor = false;
            this.checkBoxInport.CheckedChanged += new System.EventHandler(this.checkBoxInport_CheckedChanged);
            // 
            // NewColorH
            // 
            this.NewColorH.BackColor = System.Drawing.Color.OldLace;
            this.NewColorH.Location = new System.Drawing.Point(161, 531);
            this.NewColorH.Name = "NewColorH";
            this.NewColorH.Size = new System.Drawing.Size(78, 39);
            this.NewColorH.TabIndex = 13;
            this.NewColorH.Text = "颜色3";
            this.NewColorH.UseVisualStyleBackColor = false;
            this.NewColorH.Click += new System.EventHandler(this.NewColorH_Click);
            // 
            // TableColoringForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.NewColorH);
            this.Controls.Add(this.checkBoxInport);
            this.Controls.Add(this.InportPresetGroup);
            this.Controls.Add(this.Export);
            this.Controls.Add(this.checkBoxFirstLine);
            this.Controls.Add(this.DeletePreset);
            this.Controls.Add(this.SavePreset);
            this.Controls.Add(this.NewColor2);
            this.Controls.Add(this.NewColor);
            this.Controls.Add(this.ExchangeColor);
            this.Controls.Add(this.ColorOK);
            this.Controls.Add(this.groupColorEffect);
            this.Controls.Add(this.groupTableColor);
            this.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TableColoringForm";
            this.Size = new System.Drawing.Size(317, 963);
            this.groupTableColor.ResumeLayout(false);
            this.groupColorEffect.ResumeLayout(false);
            this.groupColorEffect.PerformLayout();
            this.InportPresetGroup.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupTableColor;
        private System.Windows.Forms.ListBox TableColorPresetList;
        private System.Windows.Forms.GroupBox groupColorEffect;
        private System.Windows.Forms.TextBox TableLine1;
        private System.Windows.Forms.TextBox TableLine2;
        private System.Windows.Forms.TextBox TableLine5;
        private System.Windows.Forms.TextBox TableLine4;
        private System.Windows.Forms.TextBox TableLine3;
        private System.Windows.Forms.TextBox TableLine6;
        private System.Windows.Forms.Button ColorOK;
        private System.Windows.Forms.Button ExchangeColor;
        private System.Windows.Forms.Button NewColor;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button NewColor2;
        private System.Windows.Forms.Button SavePreset;
        private System.Windows.Forms.Button DeletePreset;
        private System.Windows.Forms.CheckBox checkBoxFirstLine;
        private System.Windows.Forms.Button Export;
        private System.Windows.Forms.SaveFileDialog saveFileDialogExPreset;
        private System.Windows.Forms.OpenFileDialog openFileDialogINPreset;
        private System.Windows.Forms.GroupBox InportPresetGroup;
        private System.Windows.Forms.ListBox InportPresetList;
        private System.Windows.Forms.CheckBox checkBoxInport;
        private System.Windows.Forms.Button InportOK;
        private System.Windows.Forms.Button NewColorH;
    }
}
