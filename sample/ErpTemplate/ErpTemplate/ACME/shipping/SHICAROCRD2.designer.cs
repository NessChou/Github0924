namespace ACME
{
    partial class SHICAROCRD2
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.panel2 = new System.Windows.Forms.Panel();
            this.車型 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.長 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.寬 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.高 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colEdit2 = new System.Windows.Forms.DataGridViewImageColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.車型,
            this.長,
            this.寬,
            this.高,
            this.colEdit2});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(429, 489);
            this.dataGridView1.TabIndex = 59;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(429, 489);
            this.panel2.TabIndex = 66;
            // 
            // 車型
            // 
            this.車型.DataPropertyName = "車型";
            this.車型.HeaderText = "車型";
            this.車型.Name = "車型";
            this.車型.Width = 54;
            // 
            // 長
            // 
            this.長.DataPropertyName = "長";
            this.長.HeaderText = "長";
            this.長.Name = "長";
            this.長.Width = 42;
            // 
            // 寬
            // 
            this.寬.DataPropertyName = "寬";
            this.寬.HeaderText = "寬";
            this.寬.Name = "寬";
            this.寬.Width = 42;
            // 
            // 高
            // 
            this.高.DataPropertyName = "高";
            this.高.HeaderText = "高";
            this.高.Name = "高";
            this.高.Width = 42;
            // 
            // colEdit2
            // 
            this.colEdit2.HeaderText = "";
            this.colEdit2.Image = global::ACME.Properties.Resources.Yes;
            this.colEdit2.Name = "colEdit2";
            this.colEdit2.Width = 21;
            // 
            // SHICAROCRD2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(429, 489);
            this.Controls.Add(this.panel2);
            this.Name = "SHICAROCRD2";
            this.Text = "車型";
            this.Load += new System.EventHandler(this.APS1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 車型;
        private System.Windows.Forms.DataGridViewTextBoxColumn 長;
        private System.Windows.Forms.DataGridViewTextBoxColumn 寬;
        private System.Windows.Forms.DataGridViewTextBoxColumn 高;
        private System.Windows.Forms.DataGridViewImageColumn colEdit2;
    }
}