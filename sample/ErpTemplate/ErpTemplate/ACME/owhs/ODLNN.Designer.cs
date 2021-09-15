namespace ACME
{
    partial class ODLNN
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn23 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dataGridViewTextBoxColumn24 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.BU = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.簽核日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.簽核時間 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.產品編號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.數量 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CARDNAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.WHNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.金額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SALES2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.倉庫 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.createName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.客戶編號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.已匯出 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.LINENUM = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(21, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(83, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "匯出Excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(459, 4);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(70, 22);
            this.textBox1.TabIndex = 3;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(552, 4);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(73, 22);
            this.textBox2.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(535, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(11, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "~";
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "ID";
            this.dataGridViewTextBoxColumn1.HeaderText = "ID";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "SALES";
            this.dataGridViewTextBoxColumn5.HeaderText = "SALES";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // dataGridViewTextBoxColumn23
            // 
            this.dataGridViewTextBoxColumn23.DataPropertyName = "CHECKED";
            this.dataGridViewTextBoxColumn23.FalseValue = "False";
            this.dataGridViewTextBoxColumn23.HeaderText = "CHECKED";
            this.dataGridViewTextBoxColumn23.IndeterminateValue = "False";
            this.dataGridViewTextBoxColumn23.Name = "dataGridViewTextBoxColumn23";
            this.dataGridViewTextBoxColumn23.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn23.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewTextBoxColumn23.TrueValue = "True";
            // 
            // dataGridViewTextBoxColumn24
            // 
            this.dataGridViewTextBoxColumn24.DataPropertyName = "CHECKEDDATE";
            this.dataGridViewTextBoxColumn24.HeaderText = "CHECKEDDATE";
            this.dataGridViewTextBoxColumn24.Name = "dataGridViewTextBoxColumn24";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.BU,
            this.ID,
            this.簽核日期,
            this.簽核時間,
            this.產品編號,
            this.數量,
            this.CARDNAME,
            this.WHNO,
            this.金額,
            this.SALES2,
            this.倉庫,
            this.createName,
            this.客戶編號,
            this.已匯出,
            this.LINENUM});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1116, 702);
            this.dataGridView1.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(376, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "異常放貨單號";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(789, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "查詢";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(253, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(117, 23);
            this.button3.TabIndex = 10;
            this.button3.Text = "匯出Excel國外";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(870, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(101, 23);
            this.button4.TabIndex = 11;
            this.button4.Text = "異常筆數統計";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.textBox3);
            this.panel1.Controls.Add(this.checkBox1);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1116, 35);
            this.panel1.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(631, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 12);
            this.label3.TabIndex = 14;
            this.label3.Text = "JOB NO";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(681, 4);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(105, 22);
            this.textBox3.TabIndex = 13;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(110, 8);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(72, 16);
            this.checkBox1.TabIndex = 12;
            this.checkBox1.Text = "多單合併";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 35);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1116, 702);
            this.panel2.TabIndex = 13;
            // 
            // BU
            // 
            this.BU.DataPropertyName = "BU";
            this.BU.HeaderText = "BU";
            this.BU.Name = "BU";
            this.BU.ReadOnly = true;
            this.BU.Width = 46;
            // 
            // ID
            // 
            this.ID.DataPropertyName = "ID";
            this.ID.HeaderText = "異常放貨單號";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Width = 72;
            // 
            // 簽核日期
            // 
            this.簽核日期.DataPropertyName = "簽核日期";
            this.簽核日期.HeaderText = "簽核日期";
            this.簽核日期.Name = "簽核日期";
            this.簽核日期.ReadOnly = true;
            this.簽核日期.Width = 61;
            // 
            // 簽核時間
            // 
            this.簽核時間.DataPropertyName = "簽核時間";
            this.簽核時間.HeaderText = "簽核時間";
            this.簽核時間.Name = "簽核時間";
            this.簽核時間.ReadOnly = true;
            this.簽核時間.Width = 61;
            // 
            // 產品編號
            // 
            this.產品編號.DataPropertyName = "產品編號";
            this.產品編號.HeaderText = "產品編號";
            this.產品編號.Name = "產品編號";
            this.產品編號.ReadOnly = true;
            this.產品編號.Width = 61;
            // 
            // 數量
            // 
            this.數量.DataPropertyName = "數量";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            this.數量.DefaultCellStyle = dataGridViewCellStyle1;
            this.數量.HeaderText = "數量";
            this.數量.Name = "數量";
            this.數量.ReadOnly = true;
            this.數量.Width = 51;
            // 
            // CARDNAME
            // 
            this.CARDNAME.DataPropertyName = "CARDNAME";
            this.CARDNAME.HeaderText = "客戶名稱";
            this.CARDNAME.Name = "CARDNAME";
            this.CARDNAME.ReadOnly = true;
            this.CARDNAME.Width = 61;
            // 
            // WHNO
            // 
            this.WHNO.DataPropertyName = "WHNO";
            this.WHNO.HeaderText = "JOB NO";
            this.WHNO.Name = "WHNO";
            this.WHNO.ReadOnly = true;
            this.WHNO.Width = 50;
            // 
            // 金額
            // 
            this.金額.DataPropertyName = "金額";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.金額.DefaultCellStyle = dataGridViewCellStyle2;
            this.金額.HeaderText = "金額(USD)";
            this.金額.Name = "金額";
            this.金額.ReadOnly = true;
            this.金額.Width = 78;
            // 
            // SALES2
            // 
            this.SALES2.DataPropertyName = "SALES2";
            this.SALES2.HeaderText = "業務";
            this.SALES2.Name = "SALES2";
            this.SALES2.ReadOnly = true;
            this.SALES2.Visible = false;
            this.SALES2.Width = 51;
            // 
            // 倉庫
            // 
            this.倉庫.DataPropertyName = "倉庫";
            this.倉庫.HeaderText = "倉庫";
            this.倉庫.Name = "倉庫";
            this.倉庫.ReadOnly = true;
            this.倉庫.Width = 51;
            // 
            // createName
            // 
            this.createName.DataPropertyName = "createName";
            this.createName.HeaderText = "製單人";
            this.createName.Name = "createName";
            this.createName.ReadOnly = true;
            this.createName.Width = 61;
            // 
            // 客戶編號
            // 
            this.客戶編號.DataPropertyName = "客戶編號";
            this.客戶編號.HeaderText = "Column2";
            this.客戶編號.Name = "客戶編號";
            this.客戶編號.ReadOnly = true;
            this.客戶編號.Visible = false;
            this.客戶編號.Width = 74;
            // 
            // 已匯出
            // 
            this.已匯出.DataPropertyName = "已匯出";
            this.已匯出.HeaderText = "已匯出";
            this.已匯出.Name = "已匯出";
            this.已匯出.ReadOnly = true;
            this.已匯出.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.已匯出.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.已匯出.Width = 61;
            // 
            // LINENUM
            // 
            this.LINENUM.DataPropertyName = "LINENUM";
            this.LINENUM.HeaderText = "Column1";
            this.LINENUM.Name = "LINENUM";
            this.LINENUM.ReadOnly = true;
            this.LINENUM.Visible = false;
            this.LINENUM.Width = 74;
            // 
            // ODLNN
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1116, 737);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "ODLNN";
            this.Text = "異常出貨流程";
            this.Load += new System.EventHandler(this.ODLNN_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;

        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dataGridViewTextBoxColumn23;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn24;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.DataGridViewTextBoxColumn BU;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn 簽核日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn 簽核時間;
        private System.Windows.Forms.DataGridViewTextBoxColumn 產品編號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 數量;
        private System.Windows.Forms.DataGridViewTextBoxColumn CARDNAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn WHNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn 金額;
        private System.Windows.Forms.DataGridViewTextBoxColumn SALES2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 倉庫;
        private System.Windows.Forms.DataGridViewTextBoxColumn createName;
        private System.Windows.Forms.DataGridViewTextBoxColumn 客戶編號;
        private System.Windows.Forms.DataGridViewCheckBoxColumn 已匯出;
        private System.Windows.Forms.DataGridViewTextBoxColumn LINENUM;
    }
}