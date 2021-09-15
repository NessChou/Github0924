namespace ACME
{
    partial class POTATOAR
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnPrintTest = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.付款人 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.全雞 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.半雞 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.應稅金額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.稅額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.免稅金額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.運費 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.總計 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.取貨日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.交易方式 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.統編 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.發票號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AFEE = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.button2 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button5 = new System.Windows.Forms.Button();
            this.comboBox4 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnPrintTest
            // 
            this.btnPrintTest.Location = new System.Drawing.Point(604, 3);
            this.btnPrintTest.Name = "btnPrintTest";
            this.btnPrintTest.Size = new System.Drawing.Size(61, 23);
            this.btnPrintTest.TabIndex = 2;
            this.btnPrintTest.Text = "列印發票";
            this.btnPrintTest.UseVisualStyleBackColor = true;
            this.btnPrintTest.Click += new System.EventHandler(this.btnPrintTest_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(403, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(61, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(62, 6);
            this.textBox1.MaxLength = 8;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(80, 22);
            this.textBox1.TabIndex = 5;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(170, 6);
            this.textBox2.MaxLength = 8;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(80, 22);
            this.textBox2.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(150, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(11, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "~";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.付款人,
            this.全雞,
            this.半雞,
            this.應稅金額,
            this.稅額,
            this.免稅金額,
            this.運費,
            this.總計,
            this.取貨日期,
            this.交易方式,
            this.統編,
            this.發票號碼,
            this.AFEE});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1282, 492);
            this.dataGridView1.TabIndex = 3;
            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            this.dataGridView1.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridView1_RowPrePaint);
            // 
            // ID
            // 
            this.ID.DataPropertyName = "ID";
            this.ID.HeaderText = "PO";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Width = 35;
            // 
            // 付款人
            // 
            this.付款人.DataPropertyName = "付款人";
            this.付款人.HeaderText = "付款人";
            this.付款人.Name = "付款人";
            this.付款人.ReadOnly = true;
            this.付款人.Width = 70;
            // 
            // 全雞
            // 
            this.全雞.DataPropertyName = "全雞";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            this.全雞.DefaultCellStyle = dataGridViewCellStyle1;
            this.全雞.HeaderText = "全雞數量";
            this.全雞.Name = "全雞";
            this.全雞.ReadOnly = true;
            this.全雞.Width = 80;
            // 
            // 半雞
            // 
            this.半雞.DataPropertyName = "半雞";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.半雞.DefaultCellStyle = dataGridViewCellStyle2;
            this.半雞.HeaderText = "半雞數量";
            this.半雞.Name = "半雞";
            this.半雞.ReadOnly = true;
            this.半雞.Width = 80;
            // 
            // 應稅金額
            // 
            this.應稅金額.DataPropertyName = "應稅金額";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle3.Format = "N0";
            dataGridViewCellStyle3.NullValue = null;
            this.應稅金額.DefaultCellStyle = dataGridViewCellStyle3;
            this.應稅金額.HeaderText = "應稅金額";
            this.應稅金額.Name = "應稅金額";
            this.應稅金額.Width = 80;
            // 
            // 稅額
            // 
            this.稅額.DataPropertyName = "稅額";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle4.Format = "N0";
            dataGridViewCellStyle4.NullValue = null;
            this.稅額.DefaultCellStyle = dataGridViewCellStyle4;
            this.稅額.HeaderText = "稅額";
            this.稅額.Name = "稅額";
            this.稅額.Width = 55;
            // 
            // 免稅金額
            // 
            this.免稅金額.DataPropertyName = "免稅金額";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle5.Format = "N0";
            dataGridViewCellStyle5.NullValue = null;
            this.免稅金額.DefaultCellStyle = dataGridViewCellStyle5;
            this.免稅金額.HeaderText = "免稅金額";
            this.免稅金額.Name = "免稅金額";
            this.免稅金額.Width = 80;
            // 
            // 運費
            // 
            this.運費.DataPropertyName = "運費";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.運費.DefaultCellStyle = dataGridViewCellStyle6;
            this.運費.HeaderText = "運費";
            this.運費.Name = "運費";
            this.運費.ReadOnly = true;
            this.運費.Width = 55;
            // 
            // 總計
            // 
            this.總計.DataPropertyName = "總計";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle7.Format = "N0";
            dataGridViewCellStyle7.NullValue = null;
            this.總計.DefaultCellStyle = dataGridViewCellStyle7;
            this.總計.HeaderText = "總計";
            this.總計.Name = "總計";
            this.總計.ReadOnly = true;
            this.總計.Width = 55;
            // 
            // 取貨日期
            // 
            this.取貨日期.DataPropertyName = "取貨日期";
            this.取貨日期.HeaderText = "取貨日期";
            this.取貨日期.Name = "取貨日期";
            this.取貨日期.ReadOnly = true;
            this.取貨日期.Width = 80;
            // 
            // 交易方式
            // 
            this.交易方式.DataPropertyName = "交易方式";
            this.交易方式.HeaderText = "交易方式";
            this.交易方式.Name = "交易方式";
            this.交易方式.ReadOnly = true;
            this.交易方式.Width = 80;
            // 
            // 統編
            // 
            this.統編.DataPropertyName = "統編";
            this.統編.HeaderText = "統一編號";
            this.統編.Name = "統編";
            this.統編.ReadOnly = true;
            this.統編.Width = 80;
            // 
            // 發票號碼
            // 
            this.發票號碼.DataPropertyName = "發票號碼";
            this.發票號碼.HeaderText = "發票號碼";
            this.發票號碼.Name = "發票號碼";
            this.發票號碼.ReadOnly = true;
            this.發票號碼.Width = 80;
            // 
            // AFEE
            // 
            this.AFEE.DataPropertyName = "AFEE";
            this.AFEE.HeaderText = "發票開立完畢";
            this.AFEE.Name = "AFEE";
            this.AFEE.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.AFEE.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.AFEE.TrueValue = "True";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(470, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(61, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "發票取號";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "取貨日期";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button5);
            this.panel1.Controls.Add(this.comboBox4);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.comboBox3);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.comboBox2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btnPrintTest);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1282, 55);
            this.panel1.TabIndex = 10;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(671, 3);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(52, 23);
            this.button5.TabIndex = 19;
            this.button5.Text = "預覽";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // comboBox4
            // 
            this.comboBox4.FormattingEnabled = true;
            this.comboBox4.Items.AddRange(new object[] {
            "升序",
            "降序"});
            this.comboBox4.Location = new System.Drawing.Point(946, 7);
            this.comboBox4.Name = "comboBox4";
            this.comboBox4.Size = new System.Drawing.Size(84, 20);
            this.comboBox4.TabIndex = 18;
            this.comboBox4.SelectedIndexChanged += new System.EventHandler(this.comboBox4_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(809, 10);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 12);
            this.label5.TabIndex = 17;
            this.label5.Text = "排序";
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Items.AddRange(new object[] {
            "PO",
            "付款人",
            "全雞數量",
            "全雞單價",
            "半雞數量",
            "半雞單價",
            "應稅金額",
            "稅額",
            "免稅金額",
            "運費",
            "總計",
            "取貨日期",
            "交易方式",
            "統一編號",
            "發票號碼",
            "發票開立完畢"});
            this.comboBox3.Location = new System.Drawing.Point(848, 6);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(92, 20);
            this.comboBox3.TabIndex = 16;
            this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(242, 35);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 15;
            this.label4.Text = "交易方式";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "已開立發票",
            "未開立發票",
            "全部"});
            this.comboBox2.Location = new System.Drawing.Point(301, 32);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(101, 20);
            this.comboBox2.TabIndex = 14;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 35);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 13;
            this.label3.Text = "發票開立狀況";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "已開立發票",
            "未開立發票",
            "全部"});
            this.comboBox1.Location = new System.Drawing.Point(97, 32);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(101, 20);
            this.comboBox1.TabIndex = 12;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(729, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(61, 23);
            this.button4.TabIndex = 11;
            this.button4.Text = "EXCEL";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(537, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(61, 23);
            this.button3.TabIndex = 10;
            this.button3.Text = "修改";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 55);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1282, 492);
            this.panel2.TabIndex = 11;
            // 
            // POTATOAR
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1282, 547);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "POTATOAR";
            this.Text = "列印發票";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.POTATOAR_FormClosed);
            this.Load += new System.EventHandler(this.POTATOAR_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnPrintTest;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.ComboBox comboBox4;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn 付款人;
        private System.Windows.Forms.DataGridViewTextBoxColumn 全雞;
        private System.Windows.Forms.DataGridViewTextBoxColumn 半雞;
        private System.Windows.Forms.DataGridViewTextBoxColumn 應稅金額;
        private System.Windows.Forms.DataGridViewTextBoxColumn 稅額;
        private System.Windows.Forms.DataGridViewTextBoxColumn 免稅金額;
        private System.Windows.Forms.DataGridViewTextBoxColumn 運費;
        private System.Windows.Forms.DataGridViewTextBoxColumn 總計;
        private System.Windows.Forms.DataGridViewTextBoxColumn 取貨日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn 交易方式;
        private System.Windows.Forms.DataGridViewTextBoxColumn 統編;
        private System.Windows.Forms.DataGridViewTextBoxColumn 發票號碼;
        private System.Windows.Forms.DataGridViewCheckBoxColumn AFEE;
        private System.Windows.Forms.Button button5;
    }
}