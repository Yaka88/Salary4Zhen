namespace Salary4Zhen
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.listMonth = new System.Windows.Forms.ListBox();
            this.btnCalculate = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnReport = new System.Windows.Forms.Button();
            this.btnAnalysis = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.btnImport = new System.Windows.Forms.Button();
            this.radioHang = new System.Windows.Forms.RadioButton();
            this.radioHuzhou = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // listMonth
            // 
            this.listMonth.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listMonth.FormattingEnabled = true;
            this.listMonth.ItemHeight = 29;
            this.listMonth.Items.AddRange(new object[] {
            "1月",
            "2月",
            "3月",
            "4月",
            "5月",
            "6月",
            "7月",
            "8月",
            "9月",
            "10月",
            "11月",
            "12月",
            "1月"});
            this.listMonth.Location = new System.Drawing.Point(34, 12);
            this.listMonth.Name = "listMonth";
            this.listMonth.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.listMonth.Size = new System.Drawing.Size(76, 439);
            this.listMonth.TabIndex = 0;
            this.listMonth.SelectedValueChanged += new System.EventHandler(this.listMonth_SelectedValueChanged);
            // 
            // btnCalculate
            // 
            this.btnCalculate.Location = new System.Drawing.Point(333, 128);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(116, 34);
            this.btnCalculate.TabIndex = 1;
            this.btnCalculate.Text = "Calculate";
            this.btnCalculate.UseVisualStyleBackColor = true;
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(333, 326);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(116, 34);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnReport
            // 
            this.btnReport.Location = new System.Drawing.Point(333, 189);
            this.btnReport.Name = "btnReport";
            this.btnReport.Size = new System.Drawing.Size(116, 34);
            this.btnReport.TabIndex = 3;
            this.btnReport.Text = "Report";
            this.btnReport.UseVisualStyleBackColor = true;
            this.btnReport.Click += new System.EventHandler(this.btnReport_Click);
            // 
            // btnAnalysis
            // 
            this.btnAnalysis.Location = new System.Drawing.Point(333, 258);
            this.btnAnalysis.Name = "btnAnalysis";
            this.btnAnalysis.Size = new System.Drawing.Size(116, 34);
            this.btnAnalysis.TabIndex = 4;
            this.btnAnalysis.Text = "Analysis";
            this.btnAnalysis.UseVisualStyleBackColor = true;
            this.btnAnalysis.Click += new System.EventHandler(this.btnAnalysis_Click);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 60000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(333, 70);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(116, 34);
            this.btnImport.TabIndex = 5;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // radioHang
            // 
            this.radioHang.AutoSize = true;
            this.radioHang.Checked = true;
            this.radioHang.Location = new System.Drawing.Point(150, 104);
            this.radioHang.Name = "radioHang";
            this.radioHang.Size = new System.Drawing.Size(108, 24);
            this.radioHang.TabIndex = 6;
            this.radioHang.TabStop = true;
            this.radioHang.Text = "Hangzhou";
            this.radioHang.UseVisualStyleBackColor = true;
            // 
            // radioHuzhou
            // 
            this.radioHuzhou.AutoSize = true;
            this.radioHuzhou.Location = new System.Drawing.Point(150, 158);
            this.radioHuzhou.Name = "radioHuzhou";
            this.radioHuzhou.Size = new System.Drawing.Size(90, 24);
            this.radioHuzhou.TabIndex = 7;
            this.radioHuzhou.TabStop = true;
            this.radioHuzhou.Text = "Huzhou";
            this.radioHuzhou.UseVisualStyleBackColor = true;
            this.radioHuzhou.CheckedChanged += new System.EventHandler(this.listMonth_SelectedValueChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 483);
            this.ControlBox = false;
            this.Controls.Add(this.radioHuzhou);
            this.Controls.Add(this.radioHang);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnAnalysis);
            this.Controls.Add(this.btnReport);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.listMonth);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "Salary For Zhen Only";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listMonth;
        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnReport;
        private System.Windows.Forms.Button btnAnalysis;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.RadioButton radioHang;
        private System.Windows.Forms.RadioButton radioHuzhou;
    }
}

