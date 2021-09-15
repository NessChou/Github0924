namespace ACME
{
    partial class ACCEEP
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.申請日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.公司名稱 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.部門代號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.部門名稱 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.申請人名稱 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EEP單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.送簽文件路徑 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.狀態 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EEP = new System.Windows.Forms.DataGridViewLinkColumn();
            this.LISTID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.comboBox3);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.comboBox2);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.textBox5);
            this.panel1.Controls.Add(this.textBox6);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.textBox3);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBox4);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(900, 123);
            this.panel1.TabIndex = 0;
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Items.AddRange(new object[] {
            "結案",
            "進行中",
            "取回",
            "作廢"});
            this.comboBox3.Location = new System.Drawing.Point(361, 42);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(96, 20);
            this.comboBox3.TabIndex = 244;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(314, 45);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(41, 12);
            this.label9.TabIndex = 243;
            this.label9.Text = "申請人";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "",
            "結案",
            "進行中",
            "取回",
            "作廢"});
            this.comboBox2.Location = new System.Drawing.Point(361, 11);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(96, 20);
            this.comboBox2.TabIndex = 242;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(314, 14);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 12);
            this.label8.TabIndex = 238;
            this.label8.Text = "狀態";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(28, 97);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 234;
            this.label6.Text = "送簽日期";
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(87, 94);
            this.textBox5.MaxLength = 8;
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(87, 22);
            this.textBox5.TabIndex = 237;
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(197, 94);
            this.textBox6.MaxLength = 8;
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(87, 22);
            this.textBox6.TabIndex = 235;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(180, 97);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(11, 12);
            this.label7.TabIndex = 236;
            this.label7.Text = "~";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(32, 40);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 230;
            this.label4.Text = "部門代碼";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(87, 37);
            this.textBox1.MaxLength = 8;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(87, 22);
            this.textBox1.TabIndex = 233;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(87, 66);
            this.textBox3.MaxLength = 20;
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(87, 22);
            this.textBox3.TabIndex = 231;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(180, 44);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(11, 12);
            this.label5.TabIndex = 232;
            this.label5.Text = "~";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(88, 6);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(196, 20);
            this.comboBox1.TabIndex = 229;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 12);
            this.label3.TabIndex = 228;
            this.label3.Text = "COMPANY";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(566, 11);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(73, 28);
            this.button2.TabIndex = 227;
            this.button2.Text = "EXCEL";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 223;
            this.label1.Text = "單號查詢";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(197, 69);
            this.textBox4.MaxLength = 20;
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(87, 22);
            this.textBox4.TabIndex = 226;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(197, 37);
            this.textBox2.MaxLength = 8;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(87, 22);
            this.textBox2.TabIndex = 224;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(180, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(11, 12);
            this.label2.TabIndex = 225;
            this.label2.Text = "~";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(487, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(73, 28);
            this.button1.TabIndex = 0;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 123);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(900, 438);
            this.panel2.TabIndex = 1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.申請日期,
            this.公司名稱,
            this.部門代號,
            this.部門名稱,
            this.申請人名稱,
            this.EEP單號,
            this.送簽文件路徑,
            this.狀態,
            this.EEP,
            this.LISTID});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(900, 438);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            // 
            // 申請日期
            // 
            this.申請日期.DataPropertyName = "申請日期";
            this.申請日期.HeaderText = "申請日期";
            this.申請日期.Name = "申請日期";
            this.申請日期.Width = 80;
            // 
            // 公司名稱
            // 
            this.公司名稱.DataPropertyName = "公司名稱";
            this.公司名稱.HeaderText = "公司名稱";
            this.公司名稱.Name = "公司名稱";
            this.公司名稱.Width = 140;
            // 
            // 部門代號
            // 
            this.部門代號.DataPropertyName = "部門代號";
            this.部門代號.HeaderText = "部門代號";
            this.部門代號.Name = "部門代號";
            this.部門代號.Width = 80;
            // 
            // 部門名稱
            // 
            this.部門名稱.DataPropertyName = "部門名稱";
            this.部門名稱.HeaderText = "部門名稱";
            this.部門名稱.Name = "部門名稱";
            // 
            // 申請人名稱
            // 
            this.申請人名稱.DataPropertyName = "申請人名稱";
            this.申請人名稱.HeaderText = "申請人名稱";
            this.申請人名稱.Name = "申請人名稱";
            // 
            // EEP單號
            // 
            this.EEP單號.DataPropertyName = "EEP單號";
            this.EEP單號.HeaderText = "EEP單號";
            this.EEP單號.Name = "EEP單號";
            // 
            // 送簽文件路徑
            // 
            this.送簽文件路徑.HeaderText = "送簽文件路徑";
            this.送簽文件路徑.Name = "送簽文件路徑";
            this.送簽文件路徑.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.送簽文件路徑.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.送簽文件路徑.Text = "送簽文件";
            this.送簽文件路徑.UseColumnTextForLinkValue = true;
            // 
            // 狀態
            // 
            this.狀態.DataPropertyName = "狀態";
            this.狀態.HeaderText = "狀態";
            this.狀態.Name = "狀態";
            this.狀態.Width = 60;
            // 
            // EEP
            // 
            this.EEP.HeaderText = "";
            this.EEP.Name = "EEP";
            this.EEP.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.EEP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.EEP.Text = "簽核記錄";
            this.EEP.UseColumnTextForLinkValue = true;
            this.EEP.Width = 80;
            // 
            // LISTID
            // 
            this.LISTID.DataPropertyName = "LISTID";
            this.LISTID.HeaderText = "LISTID";
            this.LISTID.Name = "LISTID";
            this.LISTID.Visible = false;
            // 
            // ACCEEP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(900, 561);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "ACCEEP";
            this.Text = "費用電子簽呈查詢表";
            this.Load += new System.EventHandler(this.ACCAR_Load);
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
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.DataGridViewTextBoxColumn 申請日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn 公司名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn 部門代號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 部門名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn 申請人名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn EEP單號;
        private System.Windows.Forms.DataGridViewLinkColumn 送簽文件路徑;
        private System.Windows.Forms.DataGridViewTextBoxColumn 狀態;
        private System.Windows.Forms.DataGridViewLinkColumn EEP;
        private System.Windows.Forms.DataGridViewTextBoxColumn LISTID;
    }
}