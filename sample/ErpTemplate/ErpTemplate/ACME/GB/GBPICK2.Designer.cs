namespace ACME
{
    partial class GBPICK2
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
            this.dataGridView6 = new System.Windows.Forms.DataGridView();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.批發 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.撿貨單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.撿貨日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.備貨單 = new System.Windows.Forms.DataGridViewImageColumn();
            this.colEdit2 = new System.Windows.Forms.DataGridViewImageColumn();
            this.逢泰EzCat = new System.Windows.Forms.DataGridViewImageColumn();
            this.逢泰出貨主檔 = new System.Windows.Forms.DataGridViewImageColumn();
            this.簡訊 = new System.Windows.Forms.DataGridViewImageColumn();
            this.結案 = new System.Windows.Forms.DataGridViewImageColumn();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView6
            // 
            this.dataGridView6.AllowUserToAddRows = false;
            this.dataGridView6.AllowUserToDeleteRows = false;
            this.dataGridView6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView6.Location = new System.Drawing.Point(1418, 416);
            this.dataGridView6.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView6.Name = "dataGridView6";
            this.dataGridView6.ReadOnly = true;
            this.dataGridView6.RowTemplate.Height = 24;
            this.dataGridView6.Size = new System.Drawing.Size(54, 39);
            this.dataGridView6.TabIndex = 19;
            this.dataGridView6.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.批發,
            this.撿貨單號,
            this.撿貨日期,
            this.備貨單,
            this.colEdit2,
            this.逢泰EzCat,
            this.逢泰出貨主檔,
            this.簡訊,
            this.結案});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1344, 708);
            this.dataGridView1.TabIndex = 20;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // 批發
            // 
            this.批發.DataPropertyName = "批發";
            this.批發.HeaderText = "";
            this.批發.Name = "批發";
            this.批發.ReadOnly = true;
            this.批發.Width = 50;
            // 
            // 撿貨單號
            // 
            this.撿貨單號.DataPropertyName = "撿貨單號";
            this.撿貨單號.HeaderText = "撿貨單號";
            this.撿貨單號.Name = "撿貨單號";
            this.撿貨單號.ReadOnly = true;
            this.撿貨單號.Width = 150;
            // 
            // 撿貨日期
            // 
            this.撿貨日期.DataPropertyName = "撿貨日期";
            this.撿貨日期.HeaderText = "撿貨日期";
            this.撿貨日期.Name = "撿貨日期";
            this.撿貨日期.ReadOnly = true;
            // 
            // 備貨單
            // 
            this.備貨單.HeaderText = "備貨單Excel";
            this.備貨單.Image = global::ACME.Properties.Resources.app1;
            this.備貨單.Name = "備貨單";
            this.備貨單.ReadOnly = true;
            this.備貨單.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.備貨單.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.備貨單.Width = 130;
            // 
            // colEdit2
            // 
            this.colEdit2.HeaderText = "匯出EZCAT";
            this.colEdit2.Image = global::ACME.Properties.Resources.addfile2;
            this.colEdit2.Name = "colEdit2";
            this.colEdit2.ReadOnly = true;
            this.colEdit2.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colEdit2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colEdit2.Width = 120;
            // 
            // 逢泰EzCat
            // 
            this.逢泰EzCat.HeaderText = "逢泰EZCAT";
            this.逢泰EzCat.Image = global::ACME.Properties.Resources.bnUpload;
            this.逢泰EzCat.Name = "逢泰EzCat";
            this.逢泰EzCat.ReadOnly = true;
            // 
            // 逢泰出貨主檔
            // 
            this.逢泰出貨主檔.HeaderText = "逢泰出貨主檔";
            this.逢泰出貨主檔.Image = global::ACME.Properties.Resources.Yes;
            this.逢泰出貨主檔.Name = "逢泰出貨主檔";
            this.逢泰出貨主檔.ReadOnly = true;
            // 
            // 簡訊
            // 
            this.簡訊.HeaderText = "簡訊";
            this.簡訊.Image = global::ACME.Properties.Resources.bnPrint_Image;
            this.簡訊.Name = "簡訊";
            this.簡訊.ReadOnly = true;
            // 
            // 結案
            // 
            this.結案.HeaderText = "結案";
            this.結案.Image = global::ACME.Properties.Resources.app1;
            this.結案.Name = "結案";
            this.結案.ReadOnly = true;
            this.結案.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.結案.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.結案.Width = 70;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.checkBox1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1344, 41);
            this.panel2.TabIndex = 22;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 41);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1344, 708);
            this.panel1.TabIndex = 23;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(28, 12);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(75, 20);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "子料號";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // GBPICK2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1344, 749);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.dataGridView6);
            this.Font = new System.Drawing.Font("新細明體", 12F);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "GBPICK2";
            this.Text = "理貨單";
            this.Load += new System.EventHandler(this.GBPICK2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView6;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 批發;
        private System.Windows.Forms.DataGridViewTextBoxColumn 撿貨單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 撿貨日期;
        private System.Windows.Forms.DataGridViewImageColumn 備貨單;
        private System.Windows.Forms.DataGridViewImageColumn colEdit2;
        private System.Windows.Forms.DataGridViewImageColumn 逢泰EzCat;
        private System.Windows.Forms.DataGridViewImageColumn 逢泰出貨主檔;
        private System.Windows.Forms.DataGridViewImageColumn 簡訊;
        private System.Windows.Forms.DataGridViewImageColumn 結案;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Panel panel1;
    }
}