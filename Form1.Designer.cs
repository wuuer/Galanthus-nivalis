namespace WindowsFormsApplication1
{
    partial class Form1
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.Btn_SelDeliveryExcel = new System.Windows.Forms.Button();
            this.LBDeliveryFilePath = new System.Windows.Forms.Label();
            this.LBDeliveryStoreFilePath = new System.Windows.Forms.Label();
            this.lstStatus = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.LBFormatFilePath = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.LBMaterialFilePath = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.LBProjectFilePath = new System.Windows.Forms.Label();
            this.button6 = new System.Windows.Forms.Button();
            this.LBContrastFilePath = new System.Windows.Forms.Label();
            this.LBCalcFilePath = new System.Windows.Forms.Label();
            this.LBFormat2FilePath = new System.Windows.Forms.Label();
            this.button9 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.LBReplaceFilePath = new System.Windows.Forms.Label();
            this.LBToReplaceFilePath = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Btn_SelDeliveryExcel
            // 
            this.Btn_SelDeliveryExcel.AllowDrop = true;
            this.Btn_SelDeliveryExcel.Location = new System.Drawing.Point(33, 373);
            this.Btn_SelDeliveryExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Btn_SelDeliveryExcel.Name = "Btn_SelDeliveryExcel";
            this.Btn_SelDeliveryExcel.Size = new System.Drawing.Size(278, 81);
            this.Btn_SelDeliveryExcel.TabIndex = 0;
            this.Btn_SelDeliveryExcel.Text = "选择汇总表Excel";
            this.Btn_SelDeliveryExcel.UseVisualStyleBackColor = true;
            this.Btn_SelDeliveryExcel.Click += new System.EventHandler(this.Btn_SelExcel_Click);
            this.Btn_SelDeliveryExcel.DragDrop += new System.Windows.Forms.DragEventHandler(this.Btn_SelDeliveryExcel_DragDrop);
            this.Btn_SelDeliveryExcel.DragEnter += new System.Windows.Forms.DragEventHandler(this.button2_DragEnter);
            // 
            // LBDeliveryFilePath
            // 
            this.LBDeliveryFilePath.AutoSize = true;
            this.LBDeliveryFilePath.Location = new System.Drawing.Point(31, 468);
            this.LBDeliveryFilePath.Name = "LBDeliveryFilePath";
            this.LBDeliveryFilePath.Size = new System.Drawing.Size(21, 20);
            this.LBDeliveryFilePath.TabIndex = 1;
            this.LBDeliveryFilePath.Text = "   ";
            // 
            // LBDeliveryStoreFilePath
            // 
            this.LBDeliveryStoreFilePath.AutoSize = true;
            this.LBDeliveryStoreFilePath.Location = new System.Drawing.Point(31, 319);
            this.LBDeliveryStoreFilePath.Name = "LBDeliveryStoreFilePath";
            this.LBDeliveryStoreFilePath.Size = new System.Drawing.Size(21, 20);
            this.LBDeliveryStoreFilePath.TabIndex = 1;
            this.LBDeliveryStoreFilePath.Text = "   ";
            // 
            // lstStatus
            // 
            this.lstStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.lstStatus.FormattingEnabled = true;
            this.lstStatus.ItemHeight = 20;
            this.lstStatus.Location = new System.Drawing.Point(16, 568);
            this.lstStatus.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.lstStatus.Name = "lstStatus";
            this.lstStatus.Size = new System.Drawing.Size(1117, 304);
            this.lstStatus.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.AllowDrop = true;
            this.button1.Location = new System.Drawing.Point(33, 216);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(135, 81);
            this.button1.TabIndex = 0;
            this.button1.Text = "选择明细表Excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.BtnDetailExcel_Click);
            this.button1.DragDrop += new System.Windows.Forms.DragEventHandler(this.button1_DragDrop);
            this.button1.DragEnter += new System.Windows.Forms.DragEventHandler(this.button2_DragEnter);
            // 
            // button2
            // 
            this.button2.AllowDrop = true;
            this.button2.Location = new System.Drawing.Point(33, 67);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(278, 81);
            this.button2.TabIndex = 0;
            this.button2.Text = "选择格式表Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.button2.DragDrop += new System.Windows.Forms.DragEventHandler(this.button2_DragDrop);
            this.button2.DragEnter += new System.Windows.Forms.DragEventHandler(this.button2_DragEnter);
            // 
            // LBFormatFilePath
            // 
            this.LBFormatFilePath.AutoSize = true;
            this.LBFormatFilePath.Location = new System.Drawing.Point(31, 172);
            this.LBFormatFilePath.Name = "LBFormatFilePath";
            this.LBFormatFilePath.Size = new System.Drawing.Size(21, 20);
            this.LBFormatFilePath.TabIndex = 1;
            this.LBFormatFilePath.Text = "   ";
            // 
            // button3
            // 
            this.button3.AllowDrop = true;
            this.button3.Location = new System.Drawing.Point(176, 216);
            this.button3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(135, 81);
            this.button3.TabIndex = 0;
            this.button3.Text = "合并明细表Excel";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.button3.DragDrop += new System.Windows.Forms.DragEventHandler(this.button3_DragDrop);
            this.button3.DragEnter += new System.Windows.Forms.DragEventHandler(this.button2_DragEnter);
            // 
            // button4
            // 
            this.button4.AllowDrop = true;
            this.button4.Location = new System.Drawing.Point(405, 67);
            this.button4.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(278, 81);
            this.button4.TabIndex = 0;
            this.button4.Text = "选择平料表Excel";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button4.DragDrop += new System.Windows.Forms.DragEventHandler(this.button2_DragDrop);
            this.button4.DragEnter += new System.Windows.Forms.DragEventHandler(this.button2_DragEnter);
            // 
            // LBMaterialFilePath
            // 
            this.LBMaterialFilePath.AutoSize = true;
            this.LBMaterialFilePath.Location = new System.Drawing.Point(405, 171);
            this.LBMaterialFilePath.Name = "LBMaterialFilePath";
            this.LBMaterialFilePath.Size = new System.Drawing.Size(21, 20);
            this.LBMaterialFilePath.TabIndex = 4;
            this.LBMaterialFilePath.Text = "   ";
            // 
            // button5
            // 
            this.button5.AllowDrop = true;
            this.button5.Location = new System.Drawing.Point(405, 216);
            this.button5.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(278, 81);
            this.button5.TabIndex = 0;
            this.button5.Text = "选择工程表Excel(可多选)";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button1_Click);
            this.button5.DragDrop += new System.Windows.Forms.DragEventHandler(this.button5_DragDrop);
            this.button5.DragEnter += new System.Windows.Forms.DragEventHandler(this.button2_DragEnter);
            // 
            // LBProjectFilePath
            // 
            this.LBProjectFilePath.AutoSize = true;
            this.LBProjectFilePath.Location = new System.Drawing.Point(408, 317);
            this.LBProjectFilePath.Name = "LBProjectFilePath";
            this.LBProjectFilePath.Size = new System.Drawing.Size(21, 20);
            this.LBProjectFilePath.TabIndex = 5;
            this.LBProjectFilePath.Text = "   ";
            // 
            // button6
            // 
            this.button6.AllowDrop = true;
            this.button6.Location = new System.Drawing.Point(405, 373);
            this.button6.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(278, 81);
            this.button6.TabIndex = 0;
            this.button6.Text = "列出差异数Excel";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            this.button6.DragDrop += new System.Windows.Forms.DragEventHandler(this.button6_DragDrop);
            this.button6.DragEnter += new System.Windows.Forms.DragEventHandler(this.button2_DragEnter);
            // 
            // LBContrastFilePath
            // 
            this.LBContrastFilePath.AutoSize = true;
            this.LBContrastFilePath.Location = new System.Drawing.Point(402, 468);
            this.LBContrastFilePath.Name = "LBContrastFilePath";
            this.LBContrastFilePath.Size = new System.Drawing.Size(21, 20);
            this.LBContrastFilePath.TabIndex = 1;
            this.LBContrastFilePath.Text = "   ";
            // 
            // LBCalcFilePath
            // 
            this.LBCalcFilePath.AutoSize = true;
            this.LBCalcFilePath.Location = new System.Drawing.Point(927, 318);
            this.LBCalcFilePath.Name = "LBCalcFilePath";
            this.LBCalcFilePath.Size = new System.Drawing.Size(0, 20);
            this.LBCalcFilePath.TabIndex = 1;
            // 
            // LBFormat2FilePath
            // 
            this.LBFormat2FilePath.AutoSize = true;
            this.LBFormat2FilePath.Location = new System.Drawing.Point(927, 171);
            this.LBFormat2FilePath.Name = "LBFormat2FilePath";
            this.LBFormat2FilePath.Size = new System.Drawing.Size(21, 20);
            this.LBFormat2FilePath.TabIndex = 1;
            this.LBFormat2FilePath.Text = "   ";
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(744, 373);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(129, 81);
            this.button9.TabIndex = 7;
            this.button9.Text = "选择序号清单替换Excel";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(879, 373);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(129, 81);
            this.button10.TabIndex = 7;
            this.button10.Text = "选择序号清单被替换Excel";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // LBReplaceFilePath
            // 
            this.LBReplaceFilePath.AutoSize = true;
            this.LBReplaceFilePath.Location = new System.Drawing.Point(740, 468);
            this.LBReplaceFilePath.Name = "LBReplaceFilePath";
            this.LBReplaceFilePath.Size = new System.Drawing.Size(21, 20);
            this.LBReplaceFilePath.TabIndex = 1;
            this.LBReplaceFilePath.Text = "   ";
            // 
            // LBToReplaceFilePath
            // 
            this.LBToReplaceFilePath.AutoSize = true;
            this.LBToReplaceFilePath.Location = new System.Drawing.Point(740, 503);
            this.LBToReplaceFilePath.Name = "LBToReplaceFilePath";
            this.LBToReplaceFilePath.Size = new System.Drawing.Size(25, 20);
            this.LBToReplaceFilePath.TabIndex = 1;
            this.LBToReplaceFilePath.Text = "    ";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1148, 873);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.LBProjectFilePath);
            this.Controls.Add(this.LBMaterialFilePath);
            this.Controls.Add(this.lstStatus);
            this.Controls.Add(this.LBContrastFilePath);
            this.Controls.Add(this.LBFormat2FilePath);
            this.Controls.Add(this.LBToReplaceFilePath);
            this.Controls.Add(this.LBReplaceFilePath);
            this.Controls.Add(this.LBCalcFilePath);
            this.Controls.Add(this.LBFormatFilePath);
            this.Controls.Add(this.LBDeliveryStoreFilePath);
            this.Controls.Add(this.LBDeliveryFilePath);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.Btn_SelDeliveryExcel);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("微软雅黑", 8.5F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(1492, 2254);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Galanthus nivalis";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Btn_SelDeliveryExcel;
        private System.Windows.Forms.Label LBDeliveryFilePath;
        private System.Windows.Forms.Label LBDeliveryStoreFilePath;
        private System.Windows.Forms.ListBox lstStatus;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label LBFormatFilePath;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label LBMaterialFilePath;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label LBProjectFilePath;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Label LBContrastFilePath;
        private System.Windows.Forms.Label LBCalcFilePath;
        private System.Windows.Forms.Label LBFormat2FilePath;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Label LBReplaceFilePath;
        private System.Windows.Forms.Label LBToReplaceFilePath;
    }
}

