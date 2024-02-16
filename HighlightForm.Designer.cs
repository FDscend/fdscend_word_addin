
namespace WordAddIn1
{
    partial class HighlightForm
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
            this.webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.insertCode = new System.Windows.Forms.Button();
            this.onlineMod = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).BeginInit();
            this.SuspendLayout();
            // 
            // webView21
            // 
            this.webView21.AllowExternalDrop = true;
            this.webView21.CreationProperties = null;
            this.webView21.DefaultBackgroundColor = System.Drawing.Color.White;
            this.webView21.Location = new System.Drawing.Point(0, 80);
            this.webView21.Name = "webView21";
            this.webView21.Size = new System.Drawing.Size(319, 706);
            this.webView21.TabIndex = 0;
            this.webView21.ZoomFactor = 1.2D;
            // 
            // insertCode
            // 
            this.insertCode.BackColor = System.Drawing.Color.Wheat;
            this.insertCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.insertCode.Location = new System.Drawing.Point(3, 3);
            this.insertCode.Name = "insertCode";
            this.insertCode.Size = new System.Drawing.Size(122, 71);
            this.insertCode.TabIndex = 1;
            this.insertCode.Text = "插入代码";
            this.insertCode.UseVisualStyleBackColor = false;
            this.insertCode.Click += new System.EventHandler(this.insertCode_Click);
            // 
            // onlineMod
            // 
            this.onlineMod.BackColor = System.Drawing.Color.Wheat;
            this.onlineMod.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.onlineMod.Location = new System.Drawing.Point(131, 3);
            this.onlineMod.Name = "onlineMod";
            this.onlineMod.Size = new System.Drawing.Size(122, 71);
            this.onlineMod.TabIndex = 2;
            this.onlineMod.Text = "切换在线模式";
            this.onlineMod.UseVisualStyleBackColor = false;
            this.onlineMod.Click += new System.EventHandler(this.onlineMod_Click);
            // 
            // HighlightForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.OldLace;
            this.Controls.Add(this.webView21);
            this.Controls.Add(this.onlineMod);
            this.Controls.Add(this.insertCode);
            this.Name = "HighlightForm";
            this.Size = new System.Drawing.Size(322, 786);
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Web.WebView2.WinForms.WebView2 webView21;
        private System.Windows.Forms.Button insertCode;
        private System.Windows.Forms.Button onlineMod;
    }
}
