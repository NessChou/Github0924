namespace ACME
{
    partial class ReportForm
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
            System.Windows.Forms.Label label5;
            System.Windows.Forms.Label label1;
            System.Windows.Forms.Label label2;
            System.Windows.Forms.Label label3;
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.button2 = new System.Windows.Forms.Button();
            this.INVOICE1 = new System.Windows.Forms.TextBox();
            this.INVOICE2 = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.JOBNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.客戶名稱 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.單據總類 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.收貨地 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.目的地 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.INVOICENO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.幣別 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.金額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.貿易形式 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.報單號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LINK = new System.Windows.Forms.DataGridViewLinkColumn();
            this.LINKL = new System.Windows.Forms.DataGridViewLinkColumn();
            label5 = new System.Windows.Forms.Label();
            label1 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new System.Drawing.Point(164, 9);
            label5.Name = "label5";
            label5.Size = new System.Drawing.Size(11, 12);
            label5.TabIndex = 101;
            label5.Text = "~";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(21, 9);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(53, 12);
            label1.TabIndex = 102;
            label1.Text = "起迄日期";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(21, 37);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(25, 12);
            label2.TabIndex = 110;
            label2.Text = "SO#";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(164, 37);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(11, 12);
            label3.TabIndex = 109;
            label3.Text = "~";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(279, 32);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(54, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(181, 6);
            this.textBox3.MaxLength = 8;
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(75, 22);
            this.textBox3.TabIndex = 100;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(83, 6);
            this.textBox1.MaxLength = 8;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(75, 22);
            this.textBox1.TabIndex = 99;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.JOBNO,
            this.客戶名稱,
            this.單據總類,
            this.收貨地,
            this.目的地,
            this.INVOICENO,
            this.幣別,
            this.Column3,
            this.金額,
            this.貿易形式,
            this.報單號碼,
            this.LINK,
            this.LINKL});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(868, 551);
            this.dataGridView1.TabIndex = 103;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(339, 32);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(63, 23);
            this.button2.TabIndex = 106;
            this.button2.Text = "Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // INVOICE1
            // 
            this.INVOICE1.Location = new System.Drawing.Point(83, 32);
            this.INVOICE1.MaxLength = 100;
            this.INVOICE1.Name = "INVOICE1";
            this.INVOICE1.Size = new System.Drawing.Size(75, 22);
            this.INVOICE1.TabIndex = 108;
            // 
            // INVOICE2
            // 
            this.INVOICE2.Location = new System.Drawing.Point(181, 34);
            this.INVOICE2.MaxLength = 100;
            this.INVOICE2.Name = "INVOICE2";
            this.INVOICE2.Size = new System.Drawing.Size(75, 22);
            this.INVOICE2.TabIndex = 107;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(label1);
            this.panel1.Controls.Add(label2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(label3);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.INVOICE1);
            this.panel1.Controls.Add(this.textBox3);
            this.panel1.Controls.Add(this.INVOICE2);
            this.panel1.Controls.Add(label5);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(868, 65);
            this.panel1.TabIndex = 111;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 65);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(868, 551);
            this.panel2.TabIndex = 112;
            // 
            // JOBNO
            // 
            this.JOBNO.DataPropertyName = "JOBNO";
            this.JOBNO.HeaderText = "JOBNO";
            this.JOBNO.Name = "JOBNO";
            this.JOBNO.ReadOnly = true;
            this.JOBNO.Width = 66;
            // 
            // 客戶名稱
            // 
            this.客戶名稱.DataPropertyName = "客戶名稱";
            this.客戶名稱.HeaderText = "客戶名稱";
            this.客戶名稱.Name = "客戶名稱";
            this.客戶名稱.ReadOnly = true;
            this.客戶名稱.Width = 78;
            // 
            // 單據總類
            // 
            this.單據總類.DataPropertyName = "單據總類";
            this.單據總類.HeaderText = "單據總類";
            this.單據總類.Name = "單據總類";
            this.單據總類.ReadOnly = true;
            this.單據總類.Width = 78;
            // 
            // 收貨地
            // 
            this.收貨地.DataPropertyName = "收貨地";
            this.收貨地.HeaderText = "收貨地";
            this.收貨地.Name = "收貨地";
            this.收貨地.ReadOnly = true;
            this.收貨地.Width = 66;
            // 
            // 目的地
            // 
            this.目的地.DataPropertyName = "目的地";
            this.目的地.HeaderText = "目的地";
            this.目的地.Name = "目的地";
            this.目的地.ReadOnly = true;
            this.目的地.Width = 66;
            // 
            // INVOICENO
            // 
            this.INVOICENO.DataPropertyName = "INVOICENO";
            this.INVOICENO.HeaderText = "INVOICENO";
            this.INVOICENO.Name = "INVOICENO";
            this.INVOICENO.ReadOnly = true;
            this.INVOICENO.Width = 93;
            // 
            // 幣別
            // 
            this.幣別.DataPropertyName = "幣別";
            this.幣別.HeaderText = "幣別";
            this.幣別.Name = "幣別";
            this.幣別.ReadOnly = true;
            this.幣別.Width = 54;
            // 
            // Column3
            // 
            this.Column3.DataPropertyName = "匯率";
            this.Column3.HeaderText = "匯率";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            this.Column3.Width = 54;
            // 
            // 金額
            // 
            this.金額.DataPropertyName = "金額";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            this.金額.DefaultCellStyle = dataGridViewCellStyle1;
            this.金額.HeaderText = "金額";
            this.金額.Name = "金額";
            this.金額.ReadOnly = true;
            this.金額.Width = 54;
            // 
            // 貿易形式
            // 
            this.貿易形式.DataPropertyName = "貿易形式";
            this.貿易形式.HeaderText = "貿易形式";
            this.貿易形式.Name = "貿易形式";
            this.貿易形式.ReadOnly = true;
            this.貿易形式.Width = 78;
            // 
            // 報單號碼
            // 
            this.報單號碼.DataPropertyName = "報單號碼";
            this.報單號碼.HeaderText = "報單號碼";
            this.報單號碼.Name = "報單號碼";
            this.報單號碼.ReadOnly = true;
            this.報單號碼.Width = 78;
            // 
            // LINK
            // 
            this.LINK.DataPropertyName = "LINK";
            this.LINK.HeaderText = "Invoice&Packing";
            this.LINK.Name = "LINK";
            this.LINK.ReadOnly = true;
            this.LINK.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.LINK.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.LINK.Text = "";
            this.LINK.Width = 111;
            // 
            // LINKL
            // 
            this.LINKL.DataPropertyName = "LINKL";
            this.LINKL.HeaderText = "報關單";
            this.LINKL.Name = "LINKL";
            this.LINKL.ReadOnly = true;
            this.LINKL.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.LINKL.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.LINKL.Width = 66;
            // 
            // ReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(868, 616);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "ReportForm";
            this.Text = "進出口Shipping文件查詢";
            this.Load += new System.EventHandler(this.ReportForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox INVOICE1;
        private System.Windows.Forms.TextBox INVOICE2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn JOBNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn 客戶名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn 單據總類;
        private System.Windows.Forms.DataGridViewTextBoxColumn 收貨地;
        private System.Windows.Forms.DataGridViewTextBoxColumn 目的地;
        private System.Windows.Forms.DataGridViewTextBoxColumn INVOICENO;
        private System.Windows.Forms.DataGridViewTextBoxColumn 幣別;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn 金額;
        private System.Windows.Forms.DataGridViewTextBoxColumn 貿易形式;
        private System.Windows.Forms.DataGridViewTextBoxColumn 報單號碼;
        private System.Windows.Forms.DataGridViewLinkColumn LINK;
        private System.Windows.Forms.DataGridViewLinkColumn LINKL;
    }
}