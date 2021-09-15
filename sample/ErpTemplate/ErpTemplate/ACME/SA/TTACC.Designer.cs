namespace ACME
{
    partial class TTACC
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.artextBox12 = new System.Windows.Forms.TextBox();
            this.label24 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.群組 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.客戶代碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.客戶名稱 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.傳票日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.傳票號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.銷售單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AR單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.摘要 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NTD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.匯率 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.USD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.業管 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.業務 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.離倉日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.備註 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.訂單交期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.STYPE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LORDER = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.comboBox2);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.artextBox12);
            this.panel1.Controls.Add(this.label24);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1515, 70);
            this.panel1.TabIndex = 0;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(1325, 30);
            this.button4.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(129, 29);
            this.button4.TabIndex = 247;
            this.button4.Text = "實收資本匯入";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Visible = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(301, 5);
            this.button3.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(100, 29);
            this.button3.TabIndex = 246;
            this.button3.Text = "匯出Excel";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 11);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 15);
            this.label2.TabIndex = 243;
            this.label2.Text = "日期迄";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(95, 5);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBox2.MaxLength = 8;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(89, 25);
            this.textBox2.TabIndex = 244;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(72, 11);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(15, 15);
            this.label4.TabIndex = 245;
            this.label4.Text = "~";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(592, 44);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(37, 15);
            this.label10.TabIndex = 242;
            this.label10.Text = "業助";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(387, 44);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(37, 15);
            this.label9.TabIndex = 241;
            this.label9.Text = "業務";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(632, 38);
            this.comboBox2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(153, 23);
            this.comboBox2.TabIndex = 240;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(433, 38);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(153, 23);
            this.comboBox1.TabIndex = 239;
            // 
            // artextBox12
            // 
            this.artextBox12.Location = new System.Drawing.Point(72, 40);
            this.artextBox12.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.artextBox12.Name = "artextBox12";
            this.artextBox12.Size = new System.Drawing.Size(305, 25);
            this.artextBox12.TabIndex = 238;
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(25, 44);
            this.label24.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(37, 15);
            this.label24.TabIndex = 237;
            this.label24.Text = "客戶";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(793, 11);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 15);
            this.label1.TabIndex = 226;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(608, 11);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 15);
            this.label3.TabIndex = 225;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(193, 5);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 29);
            this.button1.TabIndex = 0;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 70);
            this.panel2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1515, 700);
            this.panel2.TabIndex = 1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.群組,
            this.客戶代碼,
            this.客戶名稱,
            this.傳票日期,
            this.傳票號碼,
            this.銷售單號,
            this.AR單號,
            this.摘要,
            this.NTD,
            this.匯率,
            this.USD,
            this.業管,
            this.業務,
            this.離倉日期,
            this.備註,
            this.訂單交期,
            this.STYPE,
            this.LORDER});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1515, 700);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel|*.xls|Excelx|*.xlsx";
            // 
            // 群組
            // 
            this.群組.DataPropertyName = "群組";
            this.群組.HeaderText = "群組";
            this.群組.Name = "群組";
            this.群組.ReadOnly = true;
            this.群組.Width = 60;
            // 
            // 客戶代碼
            // 
            this.客戶代碼.DataPropertyName = "客戶代碼";
            this.客戶代碼.HeaderText = "客戶代碼";
            this.客戶代碼.Name = "客戶代碼";
            this.客戶代碼.ReadOnly = true;
            this.客戶代碼.Width = 80;
            // 
            // 客戶名稱
            // 
            this.客戶名稱.DataPropertyName = "客戶名稱";
            this.客戶名稱.HeaderText = "客戶名稱";
            this.客戶名稱.Name = "客戶名稱";
            this.客戶名稱.ReadOnly = true;
            this.客戶名稱.Width = 120;
            // 
            // 傳票日期
            // 
            this.傳票日期.DataPropertyName = "傳票日期";
            this.傳票日期.HeaderText = "傳票日期";
            this.傳票日期.Name = "傳票日期";
            this.傳票日期.ReadOnly = true;
            this.傳票日期.Width = 80;
            // 
            // 傳票號碼
            // 
            this.傳票號碼.DataPropertyName = "傳票號碼";
            this.傳票號碼.HeaderText = "傳票號碼";
            this.傳票號碼.Name = "傳票號碼";
            this.傳票號碼.ReadOnly = true;
            this.傳票號碼.Width = 80;
            // 
            // 銷售單號
            // 
            this.銷售單號.DataPropertyName = "銷售單號";
            this.銷售單號.HeaderText = "銷售單號";
            this.銷售單號.Name = "銷售單號";
            this.銷售單號.ReadOnly = true;
            this.銷售單號.Width = 80;
            // 
            // AR單號
            // 
            this.AR單號.DataPropertyName = "AR單號";
            this.AR單號.HeaderText = "AR單號";
            this.AR單號.Name = "AR單號";
            this.AR單號.ReadOnly = true;
            this.AR單號.Width = 80;
            // 
            // 摘要
            // 
            this.摘要.DataPropertyName = "摘要";
            this.摘要.HeaderText = "摘要";
            this.摘要.Name = "摘要";
            this.摘要.ReadOnly = true;
            this.摘要.Width = 200;
            // 
            // NTD
            // 
            this.NTD.DataPropertyName = "NTD";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            this.NTD.DefaultCellStyle = dataGridViewCellStyle1;
            this.NTD.HeaderText = "NTD";
            this.NTD.Name = "NTD";
            this.NTD.ReadOnly = true;
            this.NTD.Width = 65;
            // 
            // 匯率
            // 
            this.匯率.DataPropertyName = "匯率";
            this.匯率.HeaderText = "匯率";
            this.匯率.Name = "匯率";
            this.匯率.ReadOnly = true;
            this.匯率.Width = 60;
            // 
            // USD
            // 
            this.USD.DataPropertyName = "USD";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            this.USD.DefaultCellStyle = dataGridViewCellStyle2;
            this.USD.HeaderText = "USD";
            this.USD.Name = "USD";
            this.USD.ReadOnly = true;
            this.USD.Width = 60;
            // 
            // 業管
            // 
            this.業管.DataPropertyName = "業管";
            this.業管.HeaderText = "業管";
            this.業管.Name = "業管";
            this.業管.ReadOnly = true;
            // 
            // 業務
            // 
            this.業務.DataPropertyName = "業務";
            this.業務.HeaderText = "業務";
            this.業務.Name = "業務";
            this.業務.ReadOnly = true;
            // 
            // 離倉日期
            // 
            this.離倉日期.DataPropertyName = "離倉日期";
            this.離倉日期.HeaderText = "離倉日期";
            this.離倉日期.Name = "離倉日期";
            this.離倉日期.ReadOnly = true;
            // 
            // 備註
            // 
            this.備註.DataPropertyName = "備註";
            this.備註.HeaderText = "備註";
            this.備註.Name = "備註";
            this.備註.ReadOnly = true;
            // 
            // 訂單交期
            // 
            this.訂單交期.DataPropertyName = "訂單交期";
            this.訂單交期.HeaderText = "訂單交期";
            this.訂單交期.Name = "訂單交期";
            this.訂單交期.ReadOnly = true;
            // 
            // STYPE
            // 
            this.STYPE.DataPropertyName = "STYPE";
            this.STYPE.HeaderText = "TYPE";
            this.STYPE.Name = "STYPE";
            this.STYPE.ReadOnly = true;
            // 
            // LORDER
            // 
            this.LORDER.DataPropertyName = "LORDER";
            this.LORDER.HeaderText = "LORDER";
            this.LORDER.Name = "LORDER";
            this.LORDER.ReadOnly = true;
            this.LORDER.Visible = false;
            // 
            // TTACC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1515, 770);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "TTACC";
            this.Text = "預收貨款";
            this.Load += new System.EventHandler(this.TTACC_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox artextBox12;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 群組;
        private System.Windows.Forms.DataGridViewTextBoxColumn 客戶代碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 客戶名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn 傳票日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn 傳票號碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 銷售單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn AR單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 摘要;
        private System.Windows.Forms.DataGridViewTextBoxColumn NTD;
        private System.Windows.Forms.DataGridViewTextBoxColumn 匯率;
        private System.Windows.Forms.DataGridViewTextBoxColumn USD;
        private System.Windows.Forms.DataGridViewTextBoxColumn 業管;
        private System.Windows.Forms.DataGridViewTextBoxColumn 業務;
        private System.Windows.Forms.DataGridViewTextBoxColumn 離倉日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn 備註;
        private System.Windows.Forms.DataGridViewTextBoxColumn 訂單交期;
        private System.Windows.Forms.DataGridViewTextBoxColumn STYPE;
        private System.Windows.Forms.DataGridViewTextBoxColumn LORDER;
    }
}