
namespace WordAddIn1
{
    partial class SimpleBrowser
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
            this.components = new System.ComponentModel.Container();
            this.switchMod = new System.Windows.Forms.Button();
            this.addressBar = new System.Windows.Forms.TextBox();
            this.gourl = new System.Windows.Forms.Button();
            this.webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.openFile = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).BeginInit();
            this.SuspendLayout();
            // 
            // switchMod
            // 
            this.switchMod.BackColor = System.Drawing.Color.Wheat;
            this.switchMod.Location = new System.Drawing.Point(4, 22);
            this.switchMod.Margin = new System.Windows.Forms.Padding(4);
            this.switchMod.Name = "switchMod";
            this.switchMod.Size = new System.Drawing.Size(90, 35);
            this.switchMod.TabIndex = 0;
            this.switchMod.Text = "模式";
            this.switchMod.UseVisualStyleBackColor = false;
            this.switchMod.Click += new System.EventHandler(this.switchMod_Click);
            // 
            // addressBar
            // 
            this.addressBar.Location = new System.Drawing.Point(104, 22);
            this.addressBar.Name = "addressBar";
            this.addressBar.Size = new System.Drawing.Size(298, 35);
            this.addressBar.TabIndex = 1;
            // 
            // gourl
            // 
            this.gourl.BackColor = System.Drawing.Color.Moccasin;
            this.gourl.Location = new System.Drawing.Point(412, 22);
            this.gourl.Name = "gourl";
            this.gourl.Size = new System.Drawing.Size(90, 35);
            this.gourl.TabIndex = 2;
            this.gourl.Text = "Go";
            this.gourl.UseVisualStyleBackColor = false;
            this.gourl.Click += new System.EventHandler(this.gourl_Click);
            // 
            // webView21
            // 
            this.webView21.AllowExternalDrop = true;
            this.webView21.CreationProperties = null;
            this.webView21.DefaultBackgroundColor = System.Drawing.Color.White;
            this.webView21.Location = new System.Drawing.Point(4, 74);
            this.webView21.Margin = new System.Windows.Forms.Padding(2);
            this.webView21.Name = "webView21";
            this.webView21.Size = new System.Drawing.Size(592, 599);
            this.webView21.TabIndex = 3;
            this.webView21.ZoomFactor = 1D;
            // 
            // openFile
            // 
            this.openFile.BackColor = System.Drawing.Color.OldLace;
            this.openFile.Location = new System.Drawing.Point(508, 22);
            this.openFile.Name = "openFile";
            this.openFile.Size = new System.Drawing.Size(90, 35);
            this.openFile.TabIndex = 4;
            this.openFile.Text = "文件";
            this.openFile.UseVisualStyleBackColor = false;
            this.openFile.Click += new System.EventHandler(this.openFile_Click);
            // 
            // SimpleBrowser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.openFile);
            this.Controls.Add(this.webView21);
            this.Controls.Add(this.gourl);
            this.Controls.Add(this.addressBar);
            this.Controls.Add(this.switchMod);
            this.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SimpleBrowser";
            this.Size = new System.Drawing.Size(613, 1057);
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button switchMod;
        private System.Windows.Forms.TextBox addressBar;
        private System.Windows.Forms.Button gourl;
        private Microsoft.Web.WebView2.WinForms.WebView2 webView21;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button openFile;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
