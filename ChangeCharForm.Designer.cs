
namespace WordAddIn1
{
    partial class ChangeCharForm
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
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.addListChar = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.labelTo = new System.Windows.Forms.Label();
            this.addTo = new System.Windows.Forms.TextBox();
            this.labelFrom = new System.Windows.Forms.Label();
            this.addFrom = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.BackColor = System.Drawing.Color.OldLace;
            this.checkedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(-1, 52);
            this.checkedListBox1.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(200, 504);
            this.checkedListBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(3, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(143, 33);
            this.label1.TabIndex = 1;
            this.label1.Text = "设置字符";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Wheat;
            this.button1.Font = new System.Drawing.Font("宋体", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(3, 565);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(200, 72);
            this.button1.TabIndex = 2;
            this.button1.Text = "替换";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // addListChar
            // 
            this.addListChar.BackColor = System.Drawing.Color.Wheat;
            this.addListChar.Font = new System.Drawing.Font("宋体", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.addListChar.Location = new System.Drawing.Point(6, 129);
            this.addListChar.Name = "addListChar";
            this.addListChar.Size = new System.Drawing.Size(188, 55);
            this.addListChar.TabIndex = 3;
            this.addListChar.Text = "+";
            this.addListChar.UseVisualStyleBackColor = false;
            this.addListChar.Click += new System.EventHandler(this.addListChar_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.labelTo);
            this.groupBox1.Controls.Add(this.addTo);
            this.groupBox1.Controls.Add(this.labelFrom);
            this.groupBox1.Controls.Add(this.addFrom);
            this.groupBox1.Controls.Add(this.addListChar);
            this.groupBox1.Location = new System.Drawing.Point(3, 663);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 193);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "增加替换";
            // 
            // labelTo
            // 
            this.labelTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelTo.Location = new System.Drawing.Point(6, 84);
            this.labelTo.Name = "labelTo";
            this.labelTo.Size = new System.Drawing.Size(80, 39);
            this.labelTo.TabIndex = 7;
            this.labelTo.Text = "To";
            this.labelTo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // addTo
            // 
            this.addTo.Location = new System.Drawing.Point(92, 84);
            this.addTo.Name = "addTo";
            this.addTo.Size = new System.Drawing.Size(102, 39);
            this.addTo.TabIndex = 6;
            // 
            // labelFrom
            // 
            this.labelFrom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelFrom.Location = new System.Drawing.Point(6, 38);
            this.labelFrom.Name = "labelFrom";
            this.labelFrom.Size = new System.Drawing.Size(80, 39);
            this.labelFrom.TabIndex = 5;
            this.labelFrom.Text = "From";
            this.labelFrom.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // addFrom
            // 
            this.addFrom.Location = new System.Drawing.Point(92, 38);
            this.addFrom.Name = "addFrom";
            this.addFrom.Size = new System.Drawing.Size(102, 39);
            this.addFrom.TabIndex = 4;
            // 
            // ChangeCharForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 28F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.OldLace;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkedListBox1);
            this.Font = new System.Drawing.Font("宋体", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.Name = "ChangeCharForm";
            this.Size = new System.Drawing.Size(208, 1020);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button addListChar;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label labelTo;
        private System.Windows.Forms.TextBox addTo;
        private System.Windows.Forms.Label labelFrom;
        private System.Windows.Forms.TextBox addFrom;
    }
}
