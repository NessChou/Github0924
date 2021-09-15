namespace ACME
{
    partial class GBBOM2
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle19 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle20 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle21 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.組合料號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.品名規格 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.發票品名 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.數量 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.成本 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.組合品項 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.建議售價 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.毛利 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.子料號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.子發票品名 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.子數量 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.子成本 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.子售價 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.組合料號,
            this.品名規格,
            this.發票品名,
            this.數量,
            this.成本,
            this.組合品項,
            this.建議售價,
            this.毛利,
            this.子料號,
            this.子發票品名,
            this.子數量,
            this.子成本,
            this.子售價});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(954, 558);
            this.dataGridView1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(954, 42);
            this.panel1.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 42);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(954, 558);
            this.panel2.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(54, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(109, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "匯出EXCEL";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // 組合料號
            // 
            this.組合料號.DataPropertyName = "組合料號";
            this.組合料號.HeaderText = "組合料號";
            this.組合料號.Name = "組合料號";
            this.組合料號.ReadOnly = true;
            this.組合料號.Width = 78;
            // 
            // 品名規格
            // 
            this.品名規格.DataPropertyName = "品名規格";
            this.品名規格.HeaderText = "品名規格";
            this.品名規格.Name = "品名規格";
            this.品名規格.ReadOnly = true;
            this.品名規格.Width = 78;
            // 
            // 發票品名
            // 
            this.發票品名.DataPropertyName = "發票品名";
            this.發票品名.HeaderText = "發票品名";
            this.發票品名.Name = "發票品名";
            this.發票品名.ReadOnly = true;
            this.發票品名.Width = 78;
            // 
            // 數量
            // 
            this.數量.DataPropertyName = "數量";
            dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle15.Format = "N0";
            dataGridViewCellStyle15.NullValue = null;
            this.數量.DefaultCellStyle = dataGridViewCellStyle15;
            this.數量.HeaderText = "數量";
            this.數量.Name = "數量";
            this.數量.ReadOnly = true;
            this.數量.Width = 54;
            // 
            // 成本
            // 
            this.成本.DataPropertyName = "成本";
            dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle16.Format = "N4";
            this.成本.DefaultCellStyle = dataGridViewCellStyle16;
            this.成本.HeaderText = "成本";
            this.成本.Name = "成本";
            this.成本.ReadOnly = true;
            this.成本.Width = 54;
            // 
            // 組合品項
            // 
            this.組合品項.DataPropertyName = "組合品項";
            this.組合品項.HeaderText = "組合品項";
            this.組合品項.Name = "組合品項";
            this.組合品項.ReadOnly = true;
            this.組合品項.Width = 78;
            // 
            // 建議售價
            // 
            this.建議售價.DataPropertyName = "建議售價";
            dataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle17.Format = "N4";
            this.建議售價.DefaultCellStyle = dataGridViewCellStyle17;
            this.建議售價.HeaderText = "建議售價";
            this.建議售價.Name = "建議售價";
            this.建議售價.ReadOnly = true;
            this.建議售價.Width = 78;
            // 
            // 毛利
            // 
            this.毛利.DataPropertyName = "毛利";
            dataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle18.Format = "N4";
            this.毛利.DefaultCellStyle = dataGridViewCellStyle18;
            this.毛利.HeaderText = "毛利";
            this.毛利.Name = "毛利";
            this.毛利.ReadOnly = true;
            this.毛利.Width = 54;
            // 
            // 子料號
            // 
            this.子料號.DataPropertyName = "子料號";
            this.子料號.HeaderText = "料號";
            this.子料號.Name = "子料號";
            this.子料號.ReadOnly = true;
            this.子料號.Width = 54;
            // 
            // 子發票品名
            // 
            this.子發票品名.DataPropertyName = "子發票品名";
            this.子發票品名.HeaderText = "發票品名";
            this.子發票品名.Name = "子發票品名";
            this.子發票品名.ReadOnly = true;
            this.子發票品名.Width = 78;
            // 
            // 子數量
            // 
            this.子數量.DataPropertyName = "子數量";
            dataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle19.Format = "N0";
            dataGridViewCellStyle19.NullValue = "0";
            this.子數量.DefaultCellStyle = dataGridViewCellStyle19;
            this.子數量.HeaderText = "數量";
            this.子數量.Name = "子數量";
            this.子數量.ReadOnly = true;
            this.子數量.Width = 54;
            // 
            // 子成本
            // 
            this.子成本.DataPropertyName = "子成本";
            dataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle20.Format = "N4";
            dataGridViewCellStyle20.NullValue = null;
            this.子成本.DefaultCellStyle = dataGridViewCellStyle20;
            this.子成本.HeaderText = "成本";
            this.子成本.Name = "子成本";
            this.子成本.ReadOnly = true;
            this.子成本.Width = 54;
            // 
            // 子售價
            // 
            this.子售價.DataPropertyName = "子售價";
            dataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle21.Format = "N4";
            dataGridViewCellStyle21.NullValue = null;
            this.子售價.DefaultCellStyle = dataGridViewCellStyle21;
            this.子售價.HeaderText = "售價";
            this.子售價.Name = "子售價";
            this.子售價.ReadOnly = true;
            this.子售價.Width = 54;
            // 
            // GBBOM2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(954, 600);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "GBBOM2";
            this.Text = "組合商品報表";
            this.Load += new System.EventHandler(this.GBBOM2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 組合料號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 品名規格;
        private System.Windows.Forms.DataGridViewTextBoxColumn 發票品名;
        private System.Windows.Forms.DataGridViewTextBoxColumn 數量;
        private System.Windows.Forms.DataGridViewTextBoxColumn 成本;
        private System.Windows.Forms.DataGridViewTextBoxColumn 組合品項;
        private System.Windows.Forms.DataGridViewTextBoxColumn 建議售價;
        private System.Windows.Forms.DataGridViewTextBoxColumn 毛利;
        private System.Windows.Forms.DataGridViewTextBoxColumn 子料號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 子發票品名;
        private System.Windows.Forms.DataGridViewTextBoxColumn 子數量;
        private System.Windows.Forms.DataGridViewTextBoxColumn 子成本;
        private System.Windows.Forms.DataGridViewTextBoxColumn 子售價;
    }
}