namespace DailyReport
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnCreateWord = new System.Windows.Forms.Button();
            this.calReportDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.ofdReadFile = new System.Windows.Forms.OpenFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.ddlFileType = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnCreateExcel = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.gvData = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.gvData)).BeginInit();
            this.SuspendLayout();
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(372, 381);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(70, 23);
            this.btnUpload.TabIndex = 0;
            this.btnUpload.Text = "读取文件";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // btnCreateWord
            // 
            this.btnCreateWord.Location = new System.Drawing.Point(455, 381);
            this.btnCreateWord.Name = "btnCreateWord";
            this.btnCreateWord.Size = new System.Drawing.Size(70, 23);
            this.btnCreateWord.TabIndex = 1;
            this.btnCreateWord.Text = "输出报告";
            this.btnCreateWord.UseVisualStyleBackColor = true;
            this.btnCreateWord.Click += new System.EventHandler(this.btnCreateWord_Click);
            // 
            // calReportDate
            // 
            this.calReportDate.Location = new System.Drawing.Point(80, 33);
            this.calReportDate.Name = "calReportDate";
            this.calReportDate.Size = new System.Drawing.Size(185, 21);
            this.calReportDate.TabIndex = 2;
            this.calReportDate.Value = new System.DateTime(2018, 8, 28, 0, 0, 0, 0);
            this.calReportDate.ValueChanged += new System.EventHandler(this.calReportDate_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "报告日期：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(306, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "文件类型：";
            // 
            // ddlFileType
            // 
            this.ddlFileType.FormattingEnabled = true;
            this.ddlFileType.Items.AddRange(new object[] {
            "每分钟节目",
            "第1动画乐园",
            "三大剧场",
            "上星频道市场份额排名",
            "34城市"});
            this.ddlFileType.Location = new System.Drawing.Point(375, 34);
            this.ddlFileType.Name = "ddlFileType";
            this.ddlFileType.Size = new System.Drawing.Size(185, 20);
            this.ddlFileType.TabIndex = 5;
            this.ddlFileType.Text = "请选择";
            this.ddlFileType.SelectedIndexChanged += new System.EventHandler(this.ddlFileType_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(269, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(11, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "*";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(564, 39);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(11, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "*";
            // 
            // btnCreateExcel
            // 
            this.btnCreateExcel.Location = new System.Drawing.Point(537, 381);
            this.btnCreateExcel.Name = "btnCreateExcel";
            this.btnCreateExcel.Size = new System.Drawing.Size(70, 23);
            this.btnCreateExcel.TabIndex = 8;
            this.btnCreateExcel.Text = "输出曲线";
            this.btnCreateExcel.UseVisualStyleBackColor = true;
            this.btnCreateExcel.Click += new System.EventHandler(this.btnCreateExcel_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(619, 381);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(70, 23);
            this.btnClear.TabIndex = 9;
            this.btnClear.Text = "清除";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // gvData
            // 
            this.gvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvData.Location = new System.Drawing.Point(13, 62);
            this.gvData.Name = "gvData";
            this.gvData.RowTemplate.Height = 23;
            this.gvData.Size = new System.Drawing.Size(675, 306);
            this.gvData.TabIndex = 10;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(702, 415);
            this.Controls.Add(this.gvData);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnCreateExcel);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ddlFileType);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.calReportDate);
            this.Controls.Add(this.btnCreateWord);
            this.Controls.Add(this.btnUpload);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "一套日报生成工具";
            ((System.ComponentModel.ISupportInitialize)(this.gvData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnCreateWord;
        private System.Windows.Forms.DateTimePicker calReportDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog ofdReadFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox ddlFileType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnCreateExcel;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.DataGridView gvData;
    }
}

