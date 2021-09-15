namespace ACME
{
    partial class ESCOEEP
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
            this.流程 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.單號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.流程階段 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MAILHEAD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MAILTEMP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MAILTO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.流程,
            this.單號,
            this.流程階段,
            this.MAILHEAD,
            this.MAILTEMP,
            this.MAILTO});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1010, 626);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseDoubleClick);
            // 
            // 流程
            // 
            this.流程.DataPropertyName = "流程";
            this.流程.HeaderText = "流程";
            this.流程.Name = "流程";
            this.流程.ReadOnly = true;
            this.流程.Width = 54;
            // 
            // 單號
            // 
            this.單號.DataPropertyName = "單號";
            this.單號.HeaderText = "單號";
            this.單號.Name = "單號";
            this.單號.ReadOnly = true;
            this.單號.Width = 54;
            // 
            // 流程階段
            // 
            this.流程階段.DataPropertyName = "流程階段";
            this.流程階段.HeaderText = "流程階段";
            this.流程階段.Name = "流程階段";
            this.流程階段.ReadOnly = true;
            this.流程階段.Width = 78;
            // 
            // MAILHEAD
            // 
            this.MAILHEAD.DataPropertyName = "MAILHEAD";
            this.MAILHEAD.HeaderText = "MAILHEAD";
            this.MAILHEAD.Name = "MAILHEAD";
            this.MAILHEAD.ReadOnly = true;
            this.MAILHEAD.Visible = false;
            this.MAILHEAD.Width = 90;
            // 
            // MAILTEMP
            // 
            this.MAILTEMP.DataPropertyName = "MAILTEMP";
            this.MAILTEMP.HeaderText = "MAILTEMP";
            this.MAILTEMP.Name = "MAILTEMP";
            this.MAILTEMP.ReadOnly = true;
            this.MAILTEMP.Visible = false;
            this.MAILTEMP.Width = 89;
            // 
            // MAILTO
            // 
            this.MAILTO.DataPropertyName = "MAILTO";
            this.MAILTO.HeaderText = "MAILTO";
            this.MAILTO.Name = "MAILTO";
            this.MAILTO.ReadOnly = true;
            this.MAILTO.Visible = false;
            this.MAILTO.Width = 74;
            // 
            // SAEEP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1010, 626);
            this.Controls.Add(this.dataGridView1);
            this.Name = "SAEEP";
            this.Text = "流程未簽核提醒";
            this.Load += new System.EventHandler(this.SAEEP_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 流程;
        private System.Windows.Forms.DataGridViewTextBoxColumn 單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 流程階段;
        private System.Windows.Forms.DataGridViewTextBoxColumn MAILHEAD;
        private System.Windows.Forms.DataGridViewTextBoxColumn MAILTEMP;
        private System.Windows.Forms.DataGridViewTextBoxColumn MAILTO;
    }
}