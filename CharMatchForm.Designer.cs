
namespace WordAddIn1
{
    partial class CharMatchForm
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
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtL = new System.Windows.Forms.Label();
            this.addLeft = new System.Windows.Forms.TextBox();
            this.addRight = new System.Windows.Forms.TextBox();
            this.txtR = new System.Windows.Forms.Label();
            this.addName = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.Label();
            this.addListChar = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.BackColor = System.Drawing.Color.OldLace;
            this.checkedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(0, 50);
            this.checkedListBox1.Margin = new System.Windows.Forms.Padding(5);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(654, 432);
            this.checkedListBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 6);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 28);
            this.label1.TabIndex = 1;
            this.label1.Text = "符号匹配";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Moccasin;
            this.button1.Font = new System.Drawing.Font("宋体", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(0, 490);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(317, 60);
            this.button1.TabIndex = 2;
            this.button1.Text = "查询";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(4, 553);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(313, 92);
            this.label2.TabIndex = 3;
            this.label2.Text = "1.勾选符号\n2.选中文字\n3.点击查询";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.addListChar);
            this.groupBox1.Controls.Add(this.addName);
            this.groupBox1.Controls.Add(this.txtName);
            this.groupBox1.Controls.Add(this.addRight);
            this.groupBox1.Controls.Add(this.txtR);
            this.groupBox1.Controls.Add(this.addLeft);
            this.groupBox1.Controls.Add(this.txtL);
            this.groupBox1.Location = new System.Drawing.Point(3, 676);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(314, 239);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "增加匹配";
            // 
            // txtL
            // 
            this.txtL.Location = new System.Drawing.Point(3, 35);
            this.txtL.Name = "txtL";
            this.txtL.Size = new System.Drawing.Size(79, 39);
            this.txtL.TabIndex = 0;
            this.txtL.Text = "Left";
            // 
            // addLeft
            // 
            this.addLeft.Location = new System.Drawing.Point(88, 35);
            this.addLeft.Name = "addLeft";
            this.addLeft.Size = new System.Drawing.Size(220, 39);
            this.addLeft.TabIndex = 1;
            // 
            // addRight
            // 
            this.addRight.Location = new System.Drawing.Point(88, 80);
            this.addRight.Name = "addRight";
            this.addRight.Size = new System.Drawing.Size(220, 39);
            this.addRight.TabIndex = 3;
            // 
            // txtR
            // 
            this.txtR.Location = new System.Drawing.Point(3, 80);
            this.txtR.Name = "txtR";
            this.txtR.Size = new System.Drawing.Size(94, 39);
            this.txtR.TabIndex = 2;
            this.txtR.Text = "Right";
            // 
            // addName
            // 
            this.addName.Location = new System.Drawing.Point(88, 122);
            this.addName.Name = "addName";
            this.addName.Size = new System.Drawing.Size(220, 39);
            this.addName.TabIndex = 5;
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(3, 122);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(79, 39);
            this.txtName.TabIndex = 4;
            this.txtName.Text = "Name";
            // 
            // addListChar
            // 
            this.addListChar.BackColor = System.Drawing.Color.Moccasin;
            this.addListChar.Font = new System.Drawing.Font("宋体", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.addListChar.Location = new System.Drawing.Point(8, 167);
            this.addListChar.Name = "addListChar";
            this.addListChar.Size = new System.Drawing.Size(300, 60);
            this.addListChar.TabIndex = 5;
            this.addListChar.Text = "+";
            this.addListChar.UseVisualStyleBackColor = false;
            this.addListChar.Click += new System.EventHandler(this.addListChar_Click);
            // 
            // CharMatchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 28F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.OldLace;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkedListBox1);
            this.Font = new System.Drawing.Font("宋体", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.Name = "CharMatchForm";
            this.Size = new System.Drawing.Size(659, 1073);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button addListChar;
        private System.Windows.Forms.TextBox addName;
        private System.Windows.Forms.Label txtName;
        private System.Windows.Forms.TextBox addRight;
        private System.Windows.Forms.Label txtR;
        private System.Windows.Forms.TextBox addLeft;
        private System.Windows.Forms.Label txtL;
    }
}
