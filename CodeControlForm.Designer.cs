
namespace WordAddIn1
{
    partial class CodeControlForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CodeControlForm));
            this.DefaultBehave = new System.Windows.Forms.GroupBox();
            this.CodeFullPageSet = new System.Windows.Forms.CheckBox();
            this.checkBoxBorderLine = new System.Windows.Forms.CheckBox();
            this.CodeListYN = new System.Windows.Forms.CheckBox();
            this.DefaultPreset = new System.Windows.Forms.GroupBox();
            this.ExportPreset = new System.Windows.Forms.Button();
            this.InportPreset = new System.Windows.Forms.Button();
            this.DeletePreset = new System.Windows.Forms.Button();
            this.DefaultPresetOK = new System.Windows.Forms.Button();
            this.CodePresetList = new System.Windows.Forms.ListBox();
            this.groupBoxFont = new System.Windows.Forms.GroupBox();
            this.DefaultFontOK = new System.Windows.Forms.Button();
            this.DefaultFont = new System.Windows.Forms.Label();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.openFileDialogINPreset = new System.Windows.Forms.OpenFileDialog();
            this.InportPresetGroup = new System.Windows.Forms.GroupBox();
            this.InportPresetList = new System.Windows.Forms.ListBox();
            this.InportOK = new System.Windows.Forms.Button();
            this.saveFileDialogExPreset = new System.Windows.Forms.SaveFileDialog();
            this.DefaultBehave.SuspendLayout();
            this.DefaultPreset.SuspendLayout();
            this.groupBoxFont.SuspendLayout();
            this.InportPresetGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // DefaultBehave
            // 
            this.DefaultBehave.Controls.Add(this.CodeFullPageSet);
            this.DefaultBehave.Controls.Add(this.checkBoxBorderLine);
            this.DefaultBehave.Controls.Add(this.CodeListYN);
            this.DefaultBehave.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DefaultBehave.Location = new System.Drawing.Point(12, 12);
            this.DefaultBehave.Name = "DefaultBehave";
            this.DefaultBehave.Size = new System.Drawing.Size(321, 137);
            this.DefaultBehave.TabIndex = 0;
            this.DefaultBehave.TabStop = false;
            this.DefaultBehave.Text = "默认行为";
            // 
            // CodeFullPageSet
            // 
            this.CodeFullPageSet.AutoSize = true;
            this.CodeFullPageSet.Location = new System.Drawing.Point(10, 90);
            this.CodeFullPageSet.Name = "CodeFullPageSet";
            this.CodeFullPageSet.Size = new System.Drawing.Size(300, 28);
            this.CodeFullPageSet.TabIndex = 2;
            this.CodeFullPageSet.Text = "默认代码块左右拉到最大";
            this.CodeFullPageSet.UseVisualStyleBackColor = true;
            this.CodeFullPageSet.CheckedChanged += new System.EventHandler(this.CodeFullPageSet_CheckedChanged);
            // 
            // checkBoxBorderLine
            // 
            this.checkBoxBorderLine.AutoSize = true;
            this.checkBoxBorderLine.Location = new System.Drawing.Point(10, 60);
            this.checkBoxBorderLine.Name = "checkBoxBorderLine";
            this.checkBoxBorderLine.Size = new System.Drawing.Size(228, 28);
            this.checkBoxBorderLine.TabIndex = 1;
            this.checkBoxBorderLine.Text = "默认加载代码竖线";
            this.checkBoxBorderLine.UseVisualStyleBackColor = true;
            this.checkBoxBorderLine.CheckedChanged += new System.EventHandler(this.checkBoxBorderLine_CheckedChanged);
            // 
            // CodeListYN
            // 
            this.CodeListYN.AutoSize = true;
            this.CodeListYN.Location = new System.Drawing.Point(10, 30);
            this.CodeListYN.Name = "CodeListYN";
            this.CodeListYN.Size = new System.Drawing.Size(228, 28);
            this.CodeListYN.TabIndex = 0;
            this.CodeListYN.Text = "默认加载代码行号";
            this.CodeListYN.UseVisualStyleBackColor = true;
            this.CodeListYN.CheckedChanged += new System.EventHandler(this.CodeListYN_CheckedChanged);
            // 
            // DefaultPreset
            // 
            this.DefaultPreset.Controls.Add(this.ExportPreset);
            this.DefaultPreset.Controls.Add(this.InportPreset);
            this.DefaultPreset.Controls.Add(this.DeletePreset);
            this.DefaultPreset.Controls.Add(this.DefaultPresetOK);
            this.DefaultPreset.Controls.Add(this.CodePresetList);
            this.DefaultPreset.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DefaultPreset.Location = new System.Drawing.Point(12, 155);
            this.DefaultPreset.Name = "DefaultPreset";
            this.DefaultPreset.Size = new System.Drawing.Size(321, 144);
            this.DefaultPreset.TabIndex = 1;
            this.DefaultPreset.TabStop = false;
            this.DefaultPreset.Text = "默认预设";
            // 
            // ExportPreset
            // 
            this.ExportPreset.BackColor = System.Drawing.Color.LemonChiffon;
            this.ExportPreset.Location = new System.Drawing.Point(235, 86);
            this.ExportPreset.Name = "ExportPreset";
            this.ExportPreset.Size = new System.Drawing.Size(75, 48);
            this.ExportPreset.TabIndex = 4;
            this.ExportPreset.Text = "导出";
            this.ExportPreset.UseVisualStyleBackColor = false;
            this.ExportPreset.Click += new System.EventHandler(this.ExportPreset_Click);
            // 
            // InportPreset
            // 
            this.InportPreset.BackColor = System.Drawing.Color.LemonChiffon;
            this.InportPreset.Location = new System.Drawing.Point(235, 34);
            this.InportPreset.Name = "InportPreset";
            this.InportPreset.Size = new System.Drawing.Size(75, 48);
            this.InportPreset.TabIndex = 3;
            this.InportPreset.Text = "导入";
            this.InportPreset.UseVisualStyleBackColor = false;
            this.InportPreset.Click += new System.EventHandler(this.InportPreset_Click);
            // 
            // DeletePreset
            // 
            this.DeletePreset.BackColor = System.Drawing.Color.OldLace;
            this.DeletePreset.Location = new System.Drawing.Point(156, 86);
            this.DeletePreset.Name = "DeletePreset";
            this.DeletePreset.Size = new System.Drawing.Size(75, 48);
            this.DeletePreset.TabIndex = 2;
            this.DeletePreset.Text = "删除";
            this.DeletePreset.UseVisualStyleBackColor = false;
            this.DeletePreset.Click += new System.EventHandler(this.DeletePreset_Click);
            // 
            // DefaultPresetOK
            // 
            this.DefaultPresetOK.BackColor = System.Drawing.Color.OldLace;
            this.DefaultPresetOK.Location = new System.Drawing.Point(156, 34);
            this.DefaultPresetOK.Name = "DefaultPresetOK";
            this.DefaultPresetOK.Size = new System.Drawing.Size(75, 48);
            this.DefaultPresetOK.TabIndex = 1;
            this.DefaultPresetOK.Text = "确定";
            this.DefaultPresetOK.UseVisualStyleBackColor = false;
            this.DefaultPresetOK.Click += new System.EventHandler(this.DefaultPresetOK_Click);
            // 
            // CodePresetList
            // 
            this.CodePresetList.FormattingEnabled = true;
            this.CodePresetList.ItemHeight = 24;
            this.CodePresetList.Location = new System.Drawing.Point(10, 34);
            this.CodePresetList.Name = "CodePresetList";
            this.CodePresetList.ScrollAlwaysVisible = true;
            this.CodePresetList.Size = new System.Drawing.Size(140, 100);
            this.CodePresetList.TabIndex = 0;
            // 
            // groupBoxFont
            // 
            this.groupBoxFont.Controls.Add(this.DefaultFontOK);
            this.groupBoxFont.Controls.Add(this.DefaultFont);
            this.groupBoxFont.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBoxFont.Location = new System.Drawing.Point(12, 434);
            this.groupBoxFont.Name = "groupBoxFont";
            this.groupBoxFont.Size = new System.Drawing.Size(321, 100);
            this.groupBoxFont.TabIndex = 2;
            this.groupBoxFont.TabStop = false;
            this.groupBoxFont.Text = "默认字体";
            // 
            // DefaultFontOK
            // 
            this.DefaultFontOK.BackColor = System.Drawing.Color.OldLace;
            this.DefaultFontOK.Location = new System.Drawing.Point(235, 27);
            this.DefaultFontOK.Name = "DefaultFontOK";
            this.DefaultFontOK.Size = new System.Drawing.Size(75, 59);
            this.DefaultFontOK.TabIndex = 1;
            this.DefaultFontOK.Text = "更改";
            this.DefaultFontOK.UseVisualStyleBackColor = false;
            this.DefaultFontOK.Click += new System.EventHandler(this.DefaultFontOK_Click);
            // 
            // DefaultFont
            // 
            this.DefaultFont.AutoSize = true;
            this.DefaultFont.Location = new System.Drawing.Point(10, 44);
            this.DefaultFont.Name = "DefaultFont";
            this.DefaultFont.Size = new System.Drawing.Size(82, 24);
            this.DefaultFont.TabIndex = 0;
            this.DefaultFont.Text = "label1";
            // 
            // openFileDialogINPreset
            // 
            this.openFileDialogINPreset.FileName = "FDscend PresetCode";
            // 
            // InportPresetGroup
            // 
            this.InportPresetGroup.Controls.Add(this.InportOK);
            this.InportPresetGroup.Controls.Add(this.InportPresetList);
            this.InportPresetGroup.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.InportPresetGroup.Location = new System.Drawing.Point(12, 305);
            this.InportPresetGroup.Name = "InportPresetGroup";
            this.InportPresetGroup.Size = new System.Drawing.Size(321, 123);
            this.InportPresetGroup.TabIndex = 3;
            this.InportPresetGroup.TabStop = false;
            this.InportPresetGroup.Text = "导入预设";
            this.InportPresetGroup.Visible = false;
            // 
            // InportPresetList
            // 
            this.InportPresetList.FormattingEnabled = true;
            this.InportPresetList.ItemHeight = 24;
            this.InportPresetList.Location = new System.Drawing.Point(10, 35);
            this.InportPresetList.Name = "InportPresetList";
            this.InportPresetList.ScrollAlwaysVisible = true;
            this.InportPresetList.Size = new System.Drawing.Size(221, 76);
            this.InportPresetList.TabIndex = 0;
            // 
            // InportOK
            // 
            this.InportOK.BackColor = System.Drawing.Color.OldLace;
            this.InportOK.Location = new System.Drawing.Point(235, 35);
            this.InportOK.Name = "InportOK";
            this.InportOK.Size = new System.Drawing.Size(75, 76);
            this.InportOK.TabIndex = 1;
            this.InportOK.Text = "确认";
            this.InportOK.UseVisualStyleBackColor = false;
            this.InportOK.Click += new System.EventHandler(this.InportOK_Click);
            // 
            // saveFileDialogExPreset
            // 
            this.saveFileDialogExPreset.FileName = "FDscend PresetCode";
            this.saveFileDialogExPreset.Title = "导出预设为";
            // 
            // CodeControlForm
            // 
            this.BackColor = System.Drawing.Color.OldLace;
            this.ClientSize = new System.Drawing.Size(345, 557);
            this.Controls.Add(this.groupBoxFont);
            this.Controls.Add(this.InportPresetGroup);
            this.Controls.Add(this.DefaultPreset);
            this.Controls.Add(this.DefaultBehave);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CodeControlForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.DefaultBehave.ResumeLayout(false);
            this.DefaultBehave.PerformLayout();
            this.DefaultPreset.ResumeLayout(false);
            this.groupBoxFont.ResumeLayout(false);
            this.groupBoxFont.PerformLayout();
            this.InportPresetGroup.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox DefaultBehave;
        private System.Windows.Forms.CheckBox CodeListYN;
        private System.Windows.Forms.CheckBox checkBoxBorderLine;
        private System.Windows.Forms.CheckBox CodeFullPageSet;
        private System.Windows.Forms.GroupBox DefaultPreset;
        private System.Windows.Forms.ListBox CodePresetList;
        private System.Windows.Forms.Button DefaultPresetOK;
        private System.Windows.Forms.Button DeletePreset;
        private System.Windows.Forms.GroupBox groupBoxFont;
        private System.Windows.Forms.Label DefaultFont;
        private System.Windows.Forms.Button DefaultFontOK;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.Button ExportPreset;
        private System.Windows.Forms.Button InportPreset;
        private System.Windows.Forms.OpenFileDialog openFileDialogINPreset;
        private System.Windows.Forms.GroupBox InportPresetGroup;
        private System.Windows.Forms.Button InportOK;
        private System.Windows.Forms.ListBox InportPresetList;
        private System.Windows.Forms.SaveFileDialog saveFileDialogExPreset;
    }
}