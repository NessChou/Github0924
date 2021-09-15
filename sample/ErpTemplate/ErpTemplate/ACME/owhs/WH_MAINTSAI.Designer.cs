namespace ACME
{
    partial class WH_MAINTSAI
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
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.JOBNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CARDNAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.INVNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.INVDATE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DOCENTRY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.RENT = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.RUSH = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.CARDINFO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.WHNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.WH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.板數 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DDATE = new GenericDataGridView.GenericDataGridView.CalendarColumn();
            this.RDATE = new GenericDataGridView.GenericDataGridView.CalendarColumn();
            this.MEMO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.button3 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(919, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.JOBNO,
            this.CARDNAME,
            this.INVNO,
            this.INVDATE,
            this.DOCENTRY,
            this.RENT,
            this.RUSH,
            this.CARDINFO,
            this.WHNO,
            this.WH,
            this.板數,
            this.DDATE,
            this.RDATE,
            this.MEMO});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1329, 551);
            this.dataGridView1.TabIndex = 3;
            this.dataGridView1.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView1_DataError);
            // 
            // JOBNO
            // 
            this.JOBNO.DataPropertyName = "JOBNO";
            this.JOBNO.HeaderText = "SI工單號碼";
            this.JOBNO.Name = "JOBNO";
            // 
            // CARDNAME
            // 
            this.CARDNAME.DataPropertyName = "CARDNAME";
            this.CARDNAME.HeaderText = "供應商";
            this.CARDNAME.Name = "CARDNAME";
            // 
            // INVNO
            // 
            this.INVNO.DataPropertyName = "INVNO";
            this.INVNO.HeaderText = "原廠Invoice No";
            this.INVNO.Name = "INVNO";
            this.INVNO.Width = 110;
            // 
            // INVDATE
            // 
            this.INVDATE.DataPropertyName = "INVDATE";
            this.INVDATE.HeaderText = "Invoice 日期";
            this.INVDATE.Name = "INVDATE";
            // 
            // DOCENTRY
            // 
            this.DOCENTRY.DataPropertyName = "DOCENTRY";
            this.DOCENTRY.HeaderText = "收貨採購單號";
            this.DOCENTRY.Name = "DOCENTRY";
            // 
            // RENT
            // 
            this.RENT.DataPropertyName = "RENT";
            this.RENT.HeaderText = "免倉租";
            this.RENT.Items.AddRange(new object[] {
            "",
            "V"});
            this.RENT.Name = "RENT";
            this.RENT.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.RENT.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.RENT.Width = 70;
            // 
            // RUSH
            // 
            this.RUSH.DataPropertyName = "RUSH";
            this.RUSH.HeaderText = "急貨";
            this.RUSH.Items.AddRange(new object[] {
            "",
            "V"});
            this.RUSH.Name = "RUSH";
            this.RUSH.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.RUSH.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.RUSH.Width = 70;
            // 
            // CARDINFO
            // 
            this.CARDINFO.DataPropertyName = "CARDINFO";
            this.CARDINFO.HeaderText = "客戶資料";
            this.CARDINFO.Name = "CARDINFO";
            // 
            // WHNO
            // 
            this.WHNO.DataPropertyName = "WHNO";
            this.WHNO.HeaderText = "WH工單號碼";
            this.WHNO.Name = "WHNO";
            // 
            // WH
            // 
            this.WH.DataPropertyName = "WH";
            this.WH.HeaderText = "倉庫";
            this.WH.Name = "WH";
            // 
            // 板數
            // 
            this.板數.DataPropertyName = "PACK";
            this.板數.HeaderText = "板數";
            this.板數.Name = "板數";
            this.板數.Width = 60;
            // 
            // DDATE
            // 
            this.DDATE.DataPropertyName = "DDATE";
            this.DDATE.HeaderText = "預計進倉日期";
            this.DDATE.Name = "DDATE";
            this.DDATE.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.DDATE.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // RDATE
            // 
            this.RDATE.DataPropertyName = "RDATE";
            this.RDATE.HeaderText = "預計出貨日期";
            this.RDATE.Name = "RDATE";
            this.RDATE.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.RDATE.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // MEMO
            // 
            this.MEMO.DataPropertyName = "MEMO";
            this.MEMO.HeaderText = "備註";
            this.MEMO.Name = "MEMO";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(90, 13);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(114, 22);
            this.textBox1.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "SI工單號碼";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(297, 13);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 22);
            this.textBox2.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(210, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "原廠Invoice No";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(403, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "收貨採購單號";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(490, 13);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(100, 22);
            this.textBox3.TabIndex = 8;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(1000, 11);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "更新";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.comboBox2);
            this.panel1.Controls.Add(this.checkBox2);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.checkBox1);
            this.panel1.Controls.Add(this.textBox7);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.textBox6);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.textBox4);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.textBox3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1329, 80);
            this.panel1.TabIndex = 12;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(815, 13);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(98, 20);
            this.comboBox2.TabIndex = 148;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(316, 54);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(48, 16);
            this.checkBox2.TabIndex = 89;
            this.checkBox2.Text = "急貨";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(1081, 11);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 88;
            this.button3.Text = "匯出Excel";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(231, 54);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(60, 16);
            this.checkBox1.TabIndex = 20;
            this.checkBox1.Text = "免倉租";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(145, 49);
            this.textBox7.MaxLength = 8;
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(59, 22);
            this.textBox7.TabIndex = 19;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(123, 54);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(11, 12);
            this.label7.TabIndex = 18;
            this.label7.Text = "~";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(57, 49);
            this.textBox6.MaxLength = 8;
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(59, 22);
            this.textBox6.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 54);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(29, 12);
            this.label6.TabIndex = 16;
            this.label6.Text = "日期";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(780, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 12);
            this.label5.TabIndex = 14;
            this.label5.Text = "倉庫";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(674, 13);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(100, 22);
            this.textBox4.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(596, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "WH工單號碼";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 80);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1329, 551);
            this.panel2.TabIndex = 13;
            // 
            // WH_MAINTSAI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1329, 631);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "WH_MAINTSAI";
            this.Text = "進貨免倉明細查詢";
            this.Load += new System.EventHandler(this.WH_MAINTSAI_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.DataGridViewTextBoxColumn JOBNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn CARDNAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn INVNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn INVDATE;
        private System.Windows.Forms.DataGridViewTextBoxColumn DOCENTRY;
        private System.Windows.Forms.DataGridViewComboBoxColumn RENT;
        private System.Windows.Forms.DataGridViewComboBoxColumn RUSH;
        private System.Windows.Forms.DataGridViewTextBoxColumn CARDINFO;
        private System.Windows.Forms.DataGridViewTextBoxColumn WHNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn WH;
        private System.Windows.Forms.DataGridViewTextBoxColumn 板數;
        private GenericDataGridView.GenericDataGridView.CalendarColumn DDATE;
        private GenericDataGridView.GenericDataGridView.CalendarColumn RDATE;
        private System.Windows.Forms.DataGridViewTextBoxColumn MEMO;
    }
}