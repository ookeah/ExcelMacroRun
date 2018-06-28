namespace ExcelMacroRun
{
    partial class frmMain
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtTartDir = new System.Windows.Forms.TextBox();
            this.btnSelectDir = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnExce = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.lboxFileList = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(130, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "targetDir(include *.xlsm)";
            // 
            // txtTartDir
            // 
            this.txtTartDir.Location = new System.Drawing.Point(25, 33);
            this.txtTartDir.Name = "txtTartDir";
            this.txtTartDir.Size = new System.Drawing.Size(575, 19);
            this.txtTartDir.TabIndex = 1;
            this.txtTartDir.Text = "-80char-";
            // 
            // btnSelectDir
            // 
            this.btnSelectDir.Location = new System.Drawing.Point(606, 31);
            this.btnSelectDir.Name = "btnSelectDir";
            this.btnSelectDir.Size = new System.Drawing.Size(37, 23);
            this.btnSelectDir.TabIndex = 2;
            this.btnSelectDir.Text = "...";
            this.btnSelectDir.UseVisualStyleBackColor = true;
            this.btnSelectDir.Click += new System.EventHandler(this.btnSelectDir_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "*.xlsm - List";
            // 
            // btnExce
            // 
            this.btnExce.Location = new System.Drawing.Point(300, 393);
            this.btnExce.Name = "btnExce";
            this.btnExce.Size = new System.Drawing.Size(75, 23);
            this.btnExce.TabIndex = 5;
            this.btnExce.Text = "RUN";
            this.btnExce.UseVisualStyleBackColor = true;
            this.btnExce.Click += new System.EventHandler(this.button1_Click);
            // 
            // lboxFileList
            // 
            this.lboxFileList.FormattingEnabled = true;
            this.lboxFileList.ItemHeight = 12;
            this.lboxFileList.Location = new System.Drawing.Point(99, 72);
            this.lboxFileList.Name = "lboxFileList";
            this.lboxFileList.Size = new System.Drawing.Size(501, 304);
            this.lboxFileList.TabIndex = 6;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(670, 437);
            this.Controls.Add(this.lboxFileList);
            this.Controls.Add(this.btnExce);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSelectDir);
            this.Controls.Add(this.txtTartDir);
            this.Controls.Add(this.label1);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExcelMacroRun";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtTartDir;
        private System.Windows.Forms.Button btnSelectDir;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnExce;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ListBox lboxFileList;
    }
}

