namespace ACME
{
    partial class PROVER
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.專案代碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.專案名稱 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.類型 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.母件編號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.子件編號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.產品名稱 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.採購單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.數量 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.單價 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.採購成本 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.已付採購成本 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.未付採購成本 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.預估成本 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.專案代碼,
            this.專案名稱,
            this.類型,
            this.母件編號,
            this.子件編號,
            this.產品名稱,
            this.採購單號,
            this.數量,
            this.單價,
            this.採購成本,
            this.已付採購成本,
            this.未付採購成本,
            this.預估成本});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1235, 571);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridView1_RowPrePaint);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(72, 14);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(86, 20);
            this.comboBox1.TabIndex = 1;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "版本";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1235, 46);
            this.panel1.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 46);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1235, 571);
            this.panel2.TabIndex = 4;
            // 
            // 專案代碼
            // 
            this.專案代碼.DataPropertyName = "專案代碼";
            this.專案代碼.HeaderText = "專案代碼";
            this.專案代碼.Name = "專案代碼";
            this.專案代碼.ReadOnly = true;
            this.專案代碼.Width = 78;
            // 
            // 專案名稱
            // 
            this.專案名稱.DataPropertyName = "專案名稱";
            this.專案名稱.HeaderText = "專案名稱";
            this.專案名稱.Name = "專案名稱";
            this.專案名稱.ReadOnly = true;
            this.專案名稱.Width = 120;
            // 
            // 類型
            // 
            this.類型.DataPropertyName = "類型";
            this.類型.HeaderText = "類型";
            this.類型.Name = "類型";
            this.類型.ReadOnly = true;
            this.類型.Width = 54;
            // 
            // 母件編號
            // 
            this.母件編號.DataPropertyName = "母件編號";
            this.母件編號.HeaderText = "母件編號";
            this.母件編號.Name = "母件編號";
            this.母件編號.ReadOnly = true;
            // 
            // 子件編號
            // 
            this.子件編號.DataPropertyName = "子件編號";
            this.子件編號.HeaderText = "子件編號";
            this.子件編號.Name = "子件編號";
            this.子件編號.ReadOnly = true;
            this.子件編號.Width = 120;
            // 
            // 產品名稱
            // 
            this.產品名稱.DataPropertyName = "產品名稱";
            this.產品名稱.HeaderText = "產品名稱";
            this.產品名稱.Name = "產品名稱";
            this.產品名稱.ReadOnly = true;
            this.產品名稱.Width = 120;
            // 
            // 採購單號
            // 
            this.採購單號.DataPropertyName = "採購單號";
            this.採購單號.HeaderText = "採購單號";
            this.採購單號.Name = "採購單號";
            this.採購單號.ReadOnly = true;
            this.採購單號.Width = 78;
            // 
            // 數量
            // 
            this.數量.DataPropertyName = "數量";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N2";
            this.數量.DefaultCellStyle = dataGridViewCellStyle1;
            this.數量.HeaderText = "數量";
            this.數量.Name = "數量";
            this.數量.ReadOnly = true;
            this.數量.Width = 60;
            // 
            // 單價
            // 
            this.單價.DataPropertyName = "單價";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            this.單價.DefaultCellStyle = dataGridViewCellStyle2;
            this.單價.HeaderText = "單價";
            this.單價.Name = "單價";
            this.單價.ReadOnly = true;
            this.單價.Width = 70;
            // 
            // 採購成本
            // 
            this.採購成本.DataPropertyName = "採購成本";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle3.Format = "N0";
            this.採購成本.DefaultCellStyle = dataGridViewCellStyle3;
            this.採購成本.HeaderText = "採購成本";
            this.採購成本.Name = "採購成本";
            this.採購成本.ReadOnly = true;
            this.採購成本.Width = 78;
            // 
            // 已付採購成本
            // 
            this.已付採購成本.DataPropertyName = "已付採購成本";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle4.Format = "N0";
            this.已付採購成本.DefaultCellStyle = dataGridViewCellStyle4;
            this.已付採購成本.HeaderText = "已付採購成本";
            this.已付採購成本.Name = "已付採購成本";
            this.已付採購成本.ReadOnly = true;
            this.已付採購成本.Width = 110;
            // 
            // 未付採購成本
            // 
            this.未付採購成本.DataPropertyName = "未付採購成本";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle5.Format = "N0";
            this.未付採購成本.DefaultCellStyle = dataGridViewCellStyle5;
            this.未付採購成本.HeaderText = "未付採購成本";
            this.未付採購成本.Name = "未付採購成本";
            this.未付採購成本.ReadOnly = true;
            this.未付採購成本.Width = 110;
            // 
            // 預估成本
            // 
            this.預估成本.DataPropertyName = "預估成本";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle6.Format = "N0";
            this.預估成本.DefaultCellStyle = dataGridViewCellStyle6;
            this.預估成本.HeaderText = "預估成本";
            this.預估成本.Name = "預估成本";
            this.預估成本.ReadOnly = true;
            this.預估成本.Width = 78;
            // 
            // PROVER
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1235, 617);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "PROVER";
            this.Text = "版本";
            this.Load += new System.EventHandler(this.PROVER_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 專案代碼;
        private System.Windows.Forms.DataGridViewTextBoxColumn 專案名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn 類型;
        private System.Windows.Forms.DataGridViewTextBoxColumn 母件編號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 子件編號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 產品名稱;
        private System.Windows.Forms.DataGridViewTextBoxColumn 採購單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 數量;
        private System.Windows.Forms.DataGridViewTextBoxColumn 單價;
        private System.Windows.Forms.DataGridViewTextBoxColumn 採購成本;
        private System.Windows.Forms.DataGridViewTextBoxColumn 已付採購成本;
        private System.Windows.Forms.DataGridViewTextBoxColumn 未付採購成本;
        private System.Windows.Forms.DataGridViewTextBoxColumn 預估成本;
    }
}