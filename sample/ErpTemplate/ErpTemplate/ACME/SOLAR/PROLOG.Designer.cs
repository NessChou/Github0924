namespace ACME
{
    partial class PROLOG
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
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.母件編號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.子件編號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DOCTYPE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.舊值 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.新值 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.修改者 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.修改時間 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.母件編號,
            this.子件編號,
            this.DOCTYPE,
            this.舊值,
            this.新值,
            this.修改者,
            this.修改時間});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(892, 545);
            this.dataGridView1.TabIndex = 0;
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "版本";
            this.Column1.HeaderText = "版本";
            this.Column1.Name = "Column1";
            this.Column1.Width = 54;
            // 
            // 母件編號
            // 
            this.母件編號.DataPropertyName = "母件編號";
            this.母件編號.HeaderText = "母件編號";
            this.母件編號.Name = "母件編號";
            this.母件編號.Width = 78;
            // 
            // 子件編號
            // 
            this.子件編號.DataPropertyName = "子件編號";
            this.子件編號.HeaderText = "子件編號";
            this.子件編號.Name = "子件編號";
            this.子件編號.Width = 78;
            // 
            // DOCTYPE
            // 
            this.DOCTYPE.DataPropertyName = "DOCTYPE";
            this.DOCTYPE.HeaderText = "類型";
            this.DOCTYPE.Name = "DOCTYPE";
            this.DOCTYPE.ReadOnly = true;
            this.DOCTYPE.Width = 54;
            // 
            // 舊值
            // 
            this.舊值.DataPropertyName = "舊值";
            this.舊值.HeaderText = "舊值";
            this.舊值.Name = "舊值";
            this.舊值.ReadOnly = true;
            this.舊值.Width = 54;
            // 
            // 新值
            // 
            this.新值.DataPropertyName = "新值";
            this.新值.HeaderText = "新值";
            this.新值.Name = "新值";
            this.新值.ReadOnly = true;
            this.新值.Width = 54;
            // 
            // 修改者
            // 
            this.修改者.DataPropertyName = "修改者";
            this.修改者.HeaderText = "修改者";
            this.修改者.Name = "修改者";
            this.修改者.ReadOnly = true;
            this.修改者.Width = 66;
            // 
            // 修改時間
            // 
            this.修改時間.DataPropertyName = "修改時間";
            this.修改時間.HeaderText = "修改時間";
            this.修改時間.Name = "修改時間";
            this.修改時間.ReadOnly = true;
            this.修改時間.Width = 78;
            // 
            // PROLOG
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(892, 545);
            this.Controls.Add(this.dataGridView1);
            this.Name = "PROLOG";
            this.Text = "LOG";
            this.Load += new System.EventHandler(this.PROLOG_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 母件編號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 子件編號;
        private System.Windows.Forms.DataGridViewTextBoxColumn DOCTYPE;
        private System.Windows.Forms.DataGridViewTextBoxColumn 舊值;
        private System.Windows.Forms.DataGridViewTextBoxColumn 新值;
        private System.Windows.Forms.DataGridViewTextBoxColumn 修改者;
        private System.Windows.Forms.DataGridViewTextBoxColumn 修改時間;
    }
}