
namespace WordAddIn1
{
    partial class AboutForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutForm));
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.help_doc = new System.Windows.Forms.Button();
            this.git_web = new System.Windows.Forms.Button();
            this.check_ver = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.download_latest = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Enabled = false;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(448, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(298, 48);
            this.label1.TabIndex = 0;
            this.label1.Text = "本插件由“分点作答”开发\r\n仅供个人使用";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::WordAddIn1.Properties.Resources._2;
            this.pictureBox1.Location = new System.Drawing.Point(-3, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(445, 453);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // help_doc
            // 
            this.help_doc.BackColor = System.Drawing.Color.OldLace;
            this.help_doc.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.help_doc.Location = new System.Drawing.Point(452, 150);
            this.help_doc.Name = "help_doc";
            this.help_doc.Size = new System.Drawing.Size(145, 71);
            this.help_doc.TabIndex = 2;
            this.help_doc.Text = "使用文档";
            this.help_doc.UseVisualStyleBackColor = false;
            this.help_doc.Click += new System.EventHandler(this.help_doc_Click);
            // 
            // git_web
            // 
            this.git_web.BackColor = System.Drawing.Color.OldLace;
            this.git_web.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.git_web.Location = new System.Drawing.Point(603, 150);
            this.git_web.Name = "git_web";
            this.git_web.Size = new System.Drawing.Size(145, 71);
            this.git_web.TabIndex = 3;
            this.git_web.Text = "项目网站";
            this.git_web.UseVisualStyleBackColor = false;
            this.git_web.Click += new System.EventHandler(this.git_web_Click);
            // 
            // check_ver
            // 
            this.check_ver.BackColor = System.Drawing.Color.OldLace;
            this.check_ver.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.check_ver.Location = new System.Drawing.Point(452, 227);
            this.check_ver.Name = "check_ver";
            this.check_ver.Size = new System.Drawing.Size(145, 71);
            this.check_ver.TabIndex = 4;
            this.check_ver.Text = "检查更新";
            this.check_ver.UseVisualStyleBackColor = false;
            this.check_ver.Click += new System.EventHandler(this.check_ver_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 12F);
            this.label2.Location = new System.Drawing.Point(448, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 24);
            this.label2.TabIndex = 5;
            this.label2.Text = "label2";
            // 
            // download_latest
            // 
            this.download_latest.BackColor = System.Drawing.Color.OldLace;
            this.download_latest.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.download_latest.Location = new System.Drawing.Point(603, 228);
            this.download_latest.Name = "download_latest";
            this.download_latest.Size = new System.Drawing.Size(143, 70);
            this.download_latest.TabIndex = 6;
            this.download_latest.Text = "下载最新";
            this.download_latest.UseVisualStyleBackColor = false;
            this.download_latest.Click += new System.EventHandler(this.download_latest_Click);
            // 
            // AboutForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.OldLace;
            this.ClientSize = new System.Drawing.Size(753, 450);
            this.Controls.Add(this.download_latest);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.check_ver);
            this.Controls.Add(this.git_web);
            this.Controls.Add(this.help_doc);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AboutForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "关于";
            this.Load += new System.EventHandler(this.AboutForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button help_doc;
        private System.Windows.Forms.Button git_web;
        private System.Windows.Forms.Button check_ver;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button download_latest;
    }
}