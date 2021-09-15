namespace ACME
{
    partial class WHOPCH
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
            this.label4 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.收貨採購單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SHIPPING工單號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.進項發票號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.發票日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.檔案名稱 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.報關號碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.路徑 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.報關檔案 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.path = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PATH2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.textBox4);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1208, 67);
            this.panel1.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(325, 15);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(103, 12);
            this.label4.TabIndex = 4;
            this.label4.Text = "SHIPPING工單號碼";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(434, 4);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox4.Size = new System.Drawing.Size(189, 57);
            this.textBox4.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "AU發票號碼";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(120, 4);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(199, 60);
            this.textBox1.TabIndex = 1;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(656, 9);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 0;
            this.button3.Text = "查詢";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 67);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1208, 563);
            this.panel2.TabIndex = 1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.收貨採購單號,
            this.SHIPPING工單號碼,
            this.進項發票號碼,
            this.發票日期,
            this.檔案名稱,
            this.報關號碼,
            this.路徑,
            this.報關檔案,
            this.path,
            this.PATH2});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1208, 563);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // 收貨採購單號
            // 
            this.收貨採購單號.DataPropertyName = "收貨採購單號";
            this.收貨採購單號.HeaderText = "收貨採購單號";
            this.收貨採購單號.Name = "收貨採購單號";
            this.收貨採購單號.Width = 72;
            // 
            // SHIPPING工單號碼
            // 
            this.SHIPPING工單號碼.DataPropertyName = "SHIPPING工單號碼";
            this.SHIPPING工單號碼.HeaderText = "SHIPPING工單號碼";
            this.SHIPPING工單號碼.Name = "SHIPPING工單號碼";
            this.SHIPPING工單號碼.Width = 85;
            // 
            // 進項發票號碼
            // 
            this.進項發票號碼.DataPropertyName = "進項發票號碼";
            this.進項發票號碼.HeaderText = "AU發票號碼";
            this.進項發票號碼.Name = "進項發票號碼";
            this.進項發票號碼.Width = 76;
            // 
            // 發票日期
            // 
            this.發票日期.DataPropertyName = "發票日期";
            this.發票日期.HeaderText = "發票日期";
            this.發票日期.Name = "發票日期";
            this.發票日期.Width = 61;
            // 
            // 檔案名稱
            // 
            this.檔案名稱.DataPropertyName = "檔案名稱";
            this.檔案名稱.HeaderText = "收採檔案";
            this.檔案名稱.Name = "檔案名稱";
            this.檔案名稱.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.檔案名稱.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.檔案名稱.Width = 61;
            // 
            // 報關號碼
            // 
            this.報關號碼.DataPropertyName = "報關號碼";
            this.報關號碼.HeaderText = "報關號碼";
            this.報關號碼.Name = "報關號碼";
            this.報關號碼.Width = 61;
            // 
            // 路徑
            // 
            this.路徑.DataPropertyName = "路徑";
            this.路徑.HeaderText = "路徑";
            this.路徑.Name = "路徑";
            this.路徑.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.路徑.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.路徑.Visible = false;
            this.路徑.Width = 51;
            // 
            // 報關檔案
            // 
            this.報關檔案.DataPropertyName = "檔案名稱2";
            this.報關檔案.HeaderText = "報關檔案";
            this.報關檔案.Name = "報關檔案";
            this.報關檔案.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.報關檔案.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.報關檔案.Width = 61;
            // 
            // path
            // 
            this.path.DataPropertyName = "path";
            this.path.HeaderText = "Column1";
            this.path.Name = "path";
            this.path.Visible = false;
            this.path.Width = 74;
            // 
            // PATH2
            // 
            this.PATH2.DataPropertyName = "PATH2";
            this.PATH2.HeaderText = "PATH2";
            this.PATH2.Name = "PATH2";
            this.PATH2.Visible = false;
            this.PATH2.Width = 65;
            // 
            // WHOPCH
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1208, 630);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "WHOPCH";
            this.Text = "發票號碼對應收貨採購單文件";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridViewTextBoxColumn 收貨採購單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn SHIPPING工單號碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 進項發票號碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 發票日期;
        private System.Windows.Forms.DataGridViewLinkColumn 檔案名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn 報關號碼;
        private System.Windows.Forms.DataGridViewLinkColumn 路徑;
        private System.Windows.Forms.DataGridViewLinkColumn 報關檔案;
        private System.Windows.Forms.DataGridViewTextBoxColumn path;
        private System.Windows.Forms.DataGridViewTextBoxColumn PATH2;
    }
}