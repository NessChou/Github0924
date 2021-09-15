namespace ACME
{
    partial class OPCH
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.客戶代碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.傳票NO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AP發票 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.採購單 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.收貨採購單 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.總數量 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.美金單價2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.傳票備註 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.美金金額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.台幣金額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.INVOICENO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.匯率 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LC = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.label9 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(359, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(451, 11);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "匯出Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(684, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 58;
            this.label3.Text = "台幣合計";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(532, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 64;
            this.label6.Text = "美金合計";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.客戶代碼,
            this.傳票NO,
            this.AP發票,
            this.採購單,
            this.收貨採購單,
            this.總數量,
            this.美金單價2,
            this.傳票備註,
            this.美金金額,
            this.台幣金額,
            this.INVOICENO,
            this.日期,
            this.匯率,
            this.LC});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1057, 644);
            this.dataGridView1.TabIndex = 69;
            // 
            // 客戶代碼
            // 
            this.客戶代碼.DataPropertyName = "客戶代碼";
            this.客戶代碼.HeaderText = "廠商";
            this.客戶代碼.Name = "客戶代碼";
            this.客戶代碼.ReadOnly = true;
            this.客戶代碼.Width = 65;
            // 
            // 傳票NO
            // 
            this.傳票NO.DataPropertyName = "傳票NO";
            this.傳票NO.HeaderText = "傳票NO";
            this.傳票NO.Name = "傳票NO";
            this.傳票NO.ReadOnly = true;
            this.傳票NO.Width = 80;
            // 
            // AP發票
            // 
            this.AP發票.DataPropertyName = "AP發票";
            this.AP發票.HeaderText = "AP發票";
            this.AP發票.Name = "AP發票";
            this.AP發票.ReadOnly = true;
            this.AP發票.Width = 70;
            // 
            // 採購單
            // 
            this.採購單.DataPropertyName = "採購單";
            this.採購單.HeaderText = "採購單";
            this.採購單.Name = "採購單";
            this.採購單.ReadOnly = true;
            this.採購單.Width = 70;
            // 
            // 收貨採購單
            // 
            this.收貨採購單.DataPropertyName = "收貨採購單";
            this.收貨採購單.HeaderText = "收貨採購單";
            this.收貨採購單.Name = "收貨採購單";
            this.收貨採購單.ReadOnly = true;
            this.收貨採購單.Width = 90;
            // 
            // 總數量
            // 
            this.總數量.DataPropertyName = "總數量";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.總數量.DefaultCellStyle = dataGridViewCellStyle7;
            this.總數量.HeaderText = "Qty";
            this.總數量.Name = "總數量";
            this.總數量.ReadOnly = true;
            this.總數量.Width = 50;
            // 
            // 美金單價2
            // 
            this.美金單價2.DataPropertyName = "美金單價2";
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.美金單價2.DefaultCellStyle = dataGridViewCellStyle8;
            this.美金單價2.HeaderText = "U/P";
            this.美金單價2.Name = "美金單價2";
            this.美金單價2.ReadOnly = true;
            this.美金單價2.Width = 50;
            // 
            // 傳票備註
            // 
            this.傳票備註.DataPropertyName = "傳票備註";
            this.傳票備註.HeaderText = "Invoice date/NO";
            this.傳票備註.Name = "傳票備註";
            this.傳票備註.ReadOnly = true;
            this.傳票備註.Width = 340;
            // 
            // 美金金額
            // 
            this.美金金額.DataPropertyName = "美金金額";
            this.美金金額.HeaderText = "美金金額";
            this.美金金額.Name = "美金金額";
            this.美金金額.ReadOnly = true;
            this.美金金額.Width = 80;
            // 
            // 台幣金額
            // 
            this.台幣金額.DataPropertyName = "台幣金額";
            this.台幣金額.HeaderText = "台幣金額";
            this.台幣金額.Name = "台幣金額";
            this.台幣金額.ReadOnly = true;
            this.台幣金額.Width = 80;
            // 
            // INVOICENO
            // 
            this.INVOICENO.DataPropertyName = "INVOICENO";
            this.INVOICENO.HeaderText = "Invoice No";
            this.INVOICENO.Name = "INVOICENO";
            this.INVOICENO.ReadOnly = true;
            this.INVOICENO.Width = 90;
            // 
            // 日期
            // 
            this.日期.DataPropertyName = "日期";
            this.日期.HeaderText = "日期";
            this.日期.Name = "日期";
            this.日期.ReadOnly = true;
            this.日期.Width = 55;
            // 
            // 匯率
            // 
            this.匯率.DataPropertyName = "匯率";
            this.匯率.HeaderText = "Remark";
            this.匯率.Name = "匯率";
            this.匯率.ReadOnly = true;
            this.匯率.Width = 45;
            // 
            // LC
            // 
            this.LC.DataPropertyName = "LC";
            this.LC.HeaderText = "LC";
            this.LC.Name = "LC";
            this.LC.ReadOnly = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(29, 16);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(21, 12);
            this.label9.TabIndex = 71;
            this.label9.Text = "BU";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(56, 13);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(77, 20);
            this.comboBox1.TabIndex = 70;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(166, 13);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(67, 20);
            this.comboBox2.TabIndex = 72;
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Location = new System.Drawing.Point(268, 13);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(67, 20);
            this.comboBox3.TabIndex = 73;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(139, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 74;
            this.label1.Text = "年";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(245, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(17, 12);
            this.label2.TabIndex = 75;
            this.label2.Text = "月";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.comboBox3);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.comboBox2);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1057, 43);
            this.panel1.TabIndex = 76;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 43);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1057, 644);
            this.panel2.TabIndex = 77;
            // 
            // OPCH
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1057, 687);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "OPCH";
            this.Text = "BU付款金額";
            this.Load += new System.EventHandler(this.CheckPaid_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 客戶代碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 傳票NO;
        private System.Windows.Forms.DataGridViewTextBoxColumn AP發票;
        private System.Windows.Forms.DataGridViewTextBoxColumn 採購單;
        private System.Windows.Forms.DataGridViewTextBoxColumn 收貨採購單;
        private System.Windows.Forms.DataGridViewTextBoxColumn 總數量;
        private System.Windows.Forms.DataGridViewTextBoxColumn 美金單價2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 傳票備註;
        private System.Windows.Forms.DataGridViewTextBoxColumn 美金金額;
        private System.Windows.Forms.DataGridViewTextBoxColumn 台幣金額;
        private System.Windows.Forms.DataGridViewTextBoxColumn INVOICENO;
        private System.Windows.Forms.DataGridViewTextBoxColumn 日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn 匯率;
        private System.Windows.Forms.DataGridViewTextBoxColumn LC;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;

    }
}