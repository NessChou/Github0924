namespace ACME
{
    partial class SHIPBOMDAN
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.型號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.數量 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.JOBNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.型號2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.收貨採購單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.進項發票號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.路徑 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.檔案名稱 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.報單號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.path2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PRINT = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.PDATE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.型號,
            this.數量});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(404, 267);
            this.dataGridView1.TabIndex = 0;
            // 
            // 型號
            // 
            this.型號.HeaderText = "型號";
            this.型號.Name = "型號";
            // 
            // 數量
            // 
            this.數量.HeaderText = "數量";
            this.數量.Name = "數量";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(38, 71);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(118, 49);
            this.button1.TabIndex = 1;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel6);
            this.panel1.Controls.Add(this.panel5);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1194, 267);
            this.panel1.TabIndex = 2;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.label2);
            this.panel6.Controls.Add(this.label1);
            this.panel6.Controls.Add(this.textBox2);
            this.panel6.Controls.Add(this.textBox1);
            this.panel6.Controls.Add(this.button1);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel6.Location = new System.Drawing.Point(404, 0);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(790, 267);
            this.panel6.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(97, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(11, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "~";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "年";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(114, 12);
            this.textBox2.MaxLength = 4;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(53, 22);
            this.textBox2.TabIndex = 3;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(38, 12);
            this.textBox1.MaxLength = 4;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(53, 22);
            this.textBox1.TabIndex = 2;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.dataGridView1);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel5.Location = new System.Drawing.Point(0, 0);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(404, 267);
            this.panel5.TabIndex = 2;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView3);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 267);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1194, 298);
            this.panel2.TabIndex = 3;
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToAddRows = false;
            this.dataGridView3.AllowUserToDeleteRows = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.JOBNO,
            this.型號2,
            this.收貨採購單號,
            this.進項發票號碼,
            this.路徑,
            this.檔案名稱,
            this.Column1,
            this.報單號碼,
            this.Column3,
            this.path2,
            this.PRINT,
            this.PDATE});
            this.dataGridView3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView3.Location = new System.Drawing.Point(0, 0);
            this.dataGridView3.MultiSelect = false;
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.RowTemplate.Height = 24;
            this.dataGridView3.Size = new System.Drawing.Size(1194, 298);
            this.dataGridView3.TabIndex = 3;
            this.dataGridView3.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellContentClick);
            this.dataGridView3.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellDoubleClick);
            this.dataGridView3.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellValueChanged);
            this.dataGridView3.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dataGridView3_MouseDoubleClick);
            // 
            // JOBNO
            // 
            this.JOBNO.DataPropertyName = "JOBNO";
            this.JOBNO.HeaderText = "工單號碼";
            this.JOBNO.Name = "JOBNO";
            // 
            // 型號2
            // 
            this.型號2.DataPropertyName = "型號";
            this.型號2.HeaderText = "型號";
            this.型號2.Name = "型號2";
            // 
            // 收貨採購單號
            // 
            this.收貨採購單號.DataPropertyName = "收貨採購單號";
            this.收貨採購單號.HeaderText = "收貨採購單號";
            this.收貨採購單號.Name = "收貨採購單號";
            // 
            // 進項發票號碼
            // 
            this.進項發票號碼.DataPropertyName = "進項發票號碼";
            this.進項發票號碼.HeaderText = "進項發票號碼";
            this.進項發票號碼.Name = "進項發票號碼";
            // 
            // 路徑
            // 
            this.路徑.DataPropertyName = "路徑";
            this.路徑.HeaderText = "路徑";
            this.路徑.Name = "路徑";
            this.路徑.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.路徑.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.路徑.Visible = false;
            // 
            // 檔案名稱
            // 
            this.檔案名稱.DataPropertyName = "檔案名稱";
            this.檔案名稱.HeaderText = "檔案名稱";
            this.檔案名稱.Name = "檔案名稱";
            this.檔案名稱.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.檔案名稱.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.檔案名稱.Width = 130;
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "path";
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.Visible = false;
            // 
            // 報單號碼
            // 
            this.報單號碼.DataPropertyName = "報單號碼";
            this.報單號碼.HeaderText = "報單號碼";
            this.報單號碼.Name = "報單號碼";
            // 
            // Column3
            // 
            this.Column3.DataPropertyName = "報單下載";
            this.Column3.HeaderText = "報單下載";
            this.Column3.Name = "Column3";
            this.Column3.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Column3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.Column3.Width = 200;
            // 
            // path2
            // 
            this.path2.DataPropertyName = "path2";
            this.path2.HeaderText = "Column4";
            this.path2.Name = "path2";
            this.path2.Visible = false;
            // 
            // PRINT
            // 
            this.PRINT.DataPropertyName = "PRINT";
            this.PRINT.FalseValue = "N";
            this.PRINT.HeaderText = "列印確認";
            this.PRINT.Name = "PRINT";
            this.PRINT.TrueValue = "V";
            this.PRINT.Width = 80;
            // 
            // PDATE
            // 
            this.PDATE.DataPropertyName = "PDATE";
            this.PDATE.HeaderText = "列印時間";
            this.PDATE.Name = "PDATE";
            this.PDATE.ReadOnly = true;
            this.PDATE.Width = 80;
            // 
            // SHIPBOMDAN
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1194, 565);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "SHIPBOMDAN";
            this.Text = "報單號碼搜尋";
            this.Load += new System.EventHandler(this.SHIPBOMDAN_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 型號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 數量;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridViewTextBoxColumn JOBNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn 型號2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 收貨採購單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 進項發票號碼;
        private System.Windows.Forms.DataGridViewLinkColumn 路徑;
        private System.Windows.Forms.DataGridViewLinkColumn 檔案名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 報單號碼;
        private System.Windows.Forms.DataGridViewLinkColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn path2;
        private System.Windows.Forms.DataGridViewCheckBoxColumn PRINT;
        private System.Windows.Forms.DataGridViewTextBoxColumn PDATE;
    }
}