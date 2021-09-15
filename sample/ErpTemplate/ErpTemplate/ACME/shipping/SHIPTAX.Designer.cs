namespace ACME
{
    partial class SHIPTAX
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tAXCHECKCheckBox = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnTaxCount = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnEmail = new System.Windows.Forms.Button();
            this.工單號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.報單號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.預計抵達日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.INV金額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.關稅百分比 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.關稅 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.推貿費 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.營業稅 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.總額 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(297, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnEmail);
            this.panel1.Controls.Add(this.tAXCHECKCheckBox);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.btnTaxCount);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(968, 57);
            this.panel1.TabIndex = 1;
            // 
            // tAXCHECKCheckBox
            // 
            this.tAXCHECKCheckBox.Location = new System.Drawing.Point(508, 11);
            this.tAXCHECKCheckBox.Name = "tAXCHECKCheckBox";
            this.tAXCHECKCheckBox.Size = new System.Drawing.Size(80, 24);
            this.tAXCHECKCheckBox.TabIndex = 250;
            this.tAXCHECKCheckBox.Text = "先放後稅";
            this.tAXCHECKCheckBox.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "結關日";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(165, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(11, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "~";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(182, 13);
            this.textBox2.MaxLength = 8;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 22);
            this.textBox2.TabIndex = 2;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(59, 13);
            this.textBox1.MaxLength = 8;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 22);
            this.textBox1.TabIndex = 1;
            // 
            // btnTaxCount
            // 
            this.btnTaxCount.Location = new System.Drawing.Point(378, 12);
            this.btnTaxCount.Name = "btnTaxCount";
            this.btnTaxCount.Size = new System.Drawing.Size(75, 23);
            this.btnTaxCount.TabIndex = 0;
            this.btnTaxCount.Text = "稅金計算";
            this.btnTaxCount.UseVisualStyleBackColor = true;
            this.btnTaxCount.Click += new System.EventHandler(this.btnTaxCount_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 57);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(968, 574);
            this.panel2.TabIndex = 2;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(968, 574);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dataGridView2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(960, 548);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "未請款資料";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataGridView2
            // 
            dataGridViewCellStyle1.NullValue = null;
            this.dataGridView2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.工單號碼,
            this.報單號碼,
            this.預計抵達日期,
            this.INV金額,
            this.關稅百分比,
            this.關稅,
            this.推貿費,
            this.營業稅,
            this.總額});
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.Location = new System.Drawing.Point(3, 3);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.Size = new System.Drawing.Size(954, 542);
            this.dataGridView2.TabIndex = 1;
            this.dataGridView2.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dataGridView2_MouseDoubleClick);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(960, 548);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "已請款資料";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(954, 542);
            this.dataGridView1.TabIndex = 0;
            // 
            // btnEmail
            // 
            this.btnEmail.Location = new System.Drawing.Point(630, 13);
            this.btnEmail.Name = "btnEmail";
            this.btnEmail.Size = new System.Drawing.Size(75, 23);
            this.btnEmail.TabIndex = 251;
            this.btnEmail.Text = "寄信";
            this.btnEmail.UseVisualStyleBackColor = true;
            this.btnEmail.Click += new System.EventHandler(this.btnEmail_Click);
            // 
            // 工單號碼
            // 
            this.工單號碼.DataPropertyName = "工單號碼";
            this.工單號碼.HeaderText = "工單號碼";
            this.工單號碼.Name = "工單號碼";
            this.工單號碼.Width = 78;
            // 
            // 報單號碼
            // 
            this.報單號碼.DataPropertyName = "報單號碼";
            this.報單號碼.HeaderText = "報單號碼";
            this.報單號碼.Name = "報單號碼";
            this.報單號碼.Width = 78;
            // 
            // 預計抵達日期
            // 
            this.預計抵達日期.DataPropertyName = "預計抵達日期";
            this.預計抵達日期.HeaderText = "預計抵達日期";
            this.預計抵達日期.Name = "預計抵達日期";
            this.預計抵達日期.Width = 102;
            // 
            // INV金額
            // 
            this.INV金額.DataPropertyName = "INV金額";
            dataGridViewCellStyle2.Format = "N2";
            dataGridViewCellStyle2.NullValue = null;
            this.INV金額.DefaultCellStyle = dataGridViewCellStyle2;
            this.INV金額.HeaderText = "INV金額(美金)";
            this.INV金額.Name = "INV金額";
            this.INV金額.Width = 106;
            // 
            // 關稅百分比
            // 
            this.關稅百分比.DataPropertyName = "關稅百分比";
            this.關稅百分比.HeaderText = "關稅(%)";
            this.關稅百分比.Name = "關稅百分比";
            this.關稅百分比.Width = 71;
            // 
            // 關稅
            // 
            this.關稅.DataPropertyName = "關稅";
            dataGridViewCellStyle3.Format = "N2";
            dataGridViewCellStyle3.NullValue = null;
            this.關稅.DefaultCellStyle = dataGridViewCellStyle3;
            this.關稅.HeaderText = "關稅(美金)";
            this.關稅.Name = "關稅";
            this.關稅.Width = 86;
            // 
            // 推貿費
            // 
            this.推貿費.DataPropertyName = "推貿費";
            dataGridViewCellStyle4.Format = "N2";
            dataGridViewCellStyle4.NullValue = null;
            this.推貿費.DefaultCellStyle = dataGridViewCellStyle4;
            this.推貿費.HeaderText = "推貿費(美金)";
            this.推貿費.Name = "推貿費";
            this.推貿費.Width = 98;
            // 
            // 營業稅
            // 
            this.營業稅.DataPropertyName = "營業稅";
            dataGridViewCellStyle5.Format = "N2";
            dataGridViewCellStyle5.NullValue = null;
            this.營業稅.DefaultCellStyle = dataGridViewCellStyle5;
            this.營業稅.HeaderText = "營業稅(美金)";
            this.營業稅.Name = "營業稅";
            this.營業稅.Width = 98;
            // 
            // 總額
            // 
            this.總額.DataPropertyName = "總額";
            dataGridViewCellStyle6.Format = "N0";
            dataGridViewCellStyle6.NullValue = null;
            this.總額.DefaultCellStyle = dataGridViewCellStyle6;
            this.總額.HeaderText = "總額(台幣)";
            this.總額.Name = "總額";
            this.總額.Width = 86;
            // 
            // SHIPTAX
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(968, 631);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "SHIPTAX";
            this.Text = "先放後稅";
            this.Load += new System.EventHandler(this.SHIPTAX_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.CheckBox tAXCHECKCheckBox;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button btnTaxCount;
        private System.Windows.Forms.Button btnEmail;
        private System.Windows.Forms.DataGridViewTextBoxColumn 工單號碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 報單號碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 預計抵達日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn INV金額;
        private System.Windows.Forms.DataGridViewTextBoxColumn 關稅百分比;
        private System.Windows.Forms.DataGridViewTextBoxColumn 關稅;
        private System.Windows.Forms.DataGridViewTextBoxColumn 推貿費;
        private System.Windows.Forms.DataGridViewTextBoxColumn 營業稅;
        private System.Windows.Forms.DataGridViewTextBoxColumn 總額;
    }
}