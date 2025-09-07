
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
            this.components = new System.ComponentModel.Container();
            this.webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.insertCode = new System.Windows.Forms.Button();
            this.codeFormat = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
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
            this.webView21.Size = new System.Drawing.Size(386, 706);
            this.webView21.TabIndex = 0;
            this.webView21.ZoomFactor = 1.2D;
            // 
            // insertCode
            // 
            this.insertCode.BackColor = System.Drawing.Color.Wheat;
            this.insertCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.insertCode.Location = new System.Drawing.Point(3, 3);
            this.insertCode.Name = "insertCode";
            this.insertCode.Size = new System.Drawing.Size(187, 71);
            this.insertCode.TabIndex = 1;
            this.insertCode.Text = "1. 从 WORD 复制代码";
            this.insertCode.UseVisualStyleBackColor = false;
            this.insertCode.Click += new System.EventHandler(this.insertCode_Click);
            // 
            // codeFormat
            // 
            this.codeFormat.BackColor = System.Drawing.Color.Wheat;
            this.codeFormat.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.codeFormat.Location = new System.Drawing.Point(196, 3);
            this.codeFormat.Name = "codeFormat";
            this.codeFormat.Size = new System.Drawing.Size(187, 71);
            this.codeFormat.TabIndex = 2;
            this.codeFormat.Text = "2. 粘贴高亮代码后点这个";
            this.codeFormat.UseVisualStyleBackColor = false;
            this.codeFormat.Click += new System.EventHandler(this.codeFormat_Click);
            // 
            // HighlightForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.OldLace;
            this.Controls.Add(this.codeFormat);
            this.Controls.Add(this.webView21);
            this.Controls.Add(this.insertCode);
            this.Name = "HighlightForm";
            this.Size = new System.Drawing.Size(389, 786);
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Web.WebView2.WinForms.WebView2 webView21;
        private System.Windows.Forms.Button insertCode;
        private System.Windows.Forms.Button codeFormat;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}
