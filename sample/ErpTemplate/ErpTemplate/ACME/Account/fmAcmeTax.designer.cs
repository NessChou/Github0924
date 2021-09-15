namespace ACME
{
    partial class fmAcmeTax
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fmAcmeTax));
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.MsgDocEntry = new System.Windows.Forms.RichTextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.MsgLine = new System.Windows.Forms.RichTextBox();
            this.button64 = new System.Windows.Forms.Button();
            this.dataGridView8 = new System.Windows.Forms.DataGridView();
            this.button63 = new System.Windows.Forms.Button();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.label36 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.groupBox6.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView8)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.tabControl1);
            this.groupBox6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox6.Location = new System.Drawing.Point(0, 0);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(809, 181);
            this.groupBox6.TabIndex = 43;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "發票異常訊息";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(3, 18);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(803, 160);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.MsgDocEntry);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(795, 134);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "發票號碼";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // MsgDocEntry
            // 
            this.MsgDocEntry.AcceptsTab = true;
            this.MsgDocEntry.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MsgDocEntry.Location = new System.Drawing.Point(3, 3);
            this.MsgDocEntry.Name = "MsgDocEntry";
            this.MsgDocEntry.Size = new System.Drawing.Size(789, 128);
            this.MsgDocEntry.TabIndex = 38;
            this.MsgDocEntry.Text = "";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.MsgLine);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(795, 134);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "明細";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // MsgLine
            // 
            this.MsgLine.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MsgLine.Location = new System.Drawing.Point(3, 3);
            this.MsgLine.Name = "MsgLine";
            this.MsgLine.Size = new System.Drawing.Size(789, 128);
            this.MsgLine.TabIndex = 37;
            this.MsgLine.Text = "";
            // 
            // button64
            // 
            this.button64.Image = ((System.Drawing.Image)(resources.GetObject("button64.Image")));
            this.button64.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button64.Location = new System.Drawing.Point(400, 7);
            this.button64.Name = "button64";
            this.button64.Size = new System.Drawing.Size(100, 35);
            this.button64.TabIndex = 42;
            this.button64.Text = "2.匯出";
            this.button64.UseVisualStyleBackColor = true;
            this.button64.Click += new System.EventHandler(this.button64_Click);
            // 
            // dataGridView8
            // 
            this.dataGridView8.AllowUserToAddRows = false;
            this.dataGridView8.AllowUserToDeleteRows = false;
            this.dataGridView8.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView8.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView8.Location = new System.Drawing.Point(0, 0);
            this.dataGridView8.Name = "dataGridView8";
            this.dataGridView8.ReadOnly = true;
            this.dataGridView8.RowTemplate.Height = 24;
            this.dataGridView8.Size = new System.Drawing.Size(809, 258);
            this.dataGridView8.TabIndex = 41;
            // 
            // button63
            // 
            this.button63.Image = ((System.Drawing.Image)(resources.GetObject("button63.Image")));
            this.button63.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button63.Location = new System.Drawing.Point(178, 9);
            this.button63.Name = "button63";
            this.button63.Size = new System.Drawing.Size(97, 35);
            this.button63.TabIndex = 40;
            this.button63.Text = "1.取得資料";
            this.button63.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button63.UseVisualStyleBackColor = true;
            this.button63.Click += new System.EventHandler(this.button63_Click);
            // 
            // textBox18
            // 
            this.textBox18.Location = new System.Drawing.Point(84, 15);
            this.textBox18.MaxLength = 6;
            this.textBox18.Name = "textBox18";
            this.textBox18.Size = new System.Drawing.Size(62, 22);
            this.textBox18.TabIndex = 39;
            this.textBox18.Text = "200801";
            // 
            // label36
            // 
            this.label36.AutoSize = true;
            this.label36.Location = new System.Drawing.Point(13, 18);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(65, 12);
            this.label36.TabIndex = 38;
            this.label36.Text = "所屬年月份";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(506, 8);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 35);
            this.button1.TabIndex = 45;
            this.button1.Text = "Excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(683, 20);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(77, 12);
            this.linkLabel1.TabIndex = 46;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "開啟說明文件";
            this.linkLabel1.Click += new System.EventHandler(this.linkLabel1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(304, 12);
            this.label2.TabIndex = 48;
            this.label2.Text = "例如:所屬年月份為200907 則SAP的申報日期為 2009.08.15";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Red;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(347, 52);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(230, 12);
            this.label3.TabIndex = 49;
            this.label3.Text = "外銷方式=1 & 通關方式=1 不轉證明文件號碼";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "",
            "聿豐",
            "忠孝",
            "宇豐",
            "韋峰"});
            this.comboBox1.Location = new System.Drawing.Point(302, 22);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(66, 20);
            this.comboBox1.TabIndex = 50;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox6);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 333);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(809, 181);
            this.panel1.TabIndex = 51;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label36);
            this.panel2.Controls.Add(this.textBox18);
            this.panel2.Controls.Add(this.button63);
            this.panel2.Controls.Add(this.comboBox1);
            this.panel2.Controls.Add(this.button64);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.linkLabel1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(809, 75);
            this.panel2.TabIndex = 52;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dataGridView8);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 75);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(809, 258);
            this.panel3.TabIndex = 53;
            // 
            // fmAcmeTax
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(809, 514);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "fmAcmeTax";
            this.Text = "營業稅媒體申報";
            this.Load += new System.EventHandler(this.fmAcmeTax_Load);
            this.groupBox6.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView8)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Button button64;
        private System.Windows.Forms.DataGridView dataGridView8;
        private System.Windows.Forms.Button button63;
        private System.Windows.Forms.TextBox textBox18;
        private System.Windows.Forms.Label label36;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.RichTextBox MsgDocEntry;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.RichTextBox MsgLine;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
    }
}

