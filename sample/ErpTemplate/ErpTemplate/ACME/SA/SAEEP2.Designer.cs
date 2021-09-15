namespace ACME
{
    partial class SAEEP2
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
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
            this.流程,
            this.單號,
            this.流程階段,
            this.MAILHEAD,
            this.MAILTEMP,
            this.MAILTO});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1347, 736);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseDoubleClick);
            // 
            // 流程
            // 
            this.流程.DataPropertyName = "流程";
            this.流程.HeaderText = "流程";
            this.流程.Name = "流程";
            this.流程.ReadOnly = true;
            this.流程.Width = 62;
            // 
            // 單號
            // 
            this.單號.DataPropertyName = "單號";
            this.單號.HeaderText = "單號";
            this.單號.Name = "單號";
            this.單號.ReadOnly = true;
            this.單號.Width = 62;
            // 
            // 流程階段
            // 
            this.流程階段.DataPropertyName = "流程階段";
            this.流程階段.HeaderText = "流程階段";
            this.流程階段.Name = "流程階段";
            this.流程階段.ReadOnly = true;
            this.流程階段.Width = 92;
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
            this.MAILTO.Width = 88;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1347, 46);
            this.panel1.TabIndex = 1;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "單號",
            "LISTID"});
            this.comboBox1.Location = new System.Drawing.Point(176, 8);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(116, 23);
            this.comboBox1.TabIndex = 3;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(341, 5);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(132, 25);
            this.textBox1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(31, 4);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 29);
            this.button1.TabIndex = 0;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 46);
            this.panel2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1347, 736);
            this.panel2.TabIndex = 2;
            // 
            // SAEEP2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1347, 782);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "SAEEP2";
            this.Text = "流程未簽核提醒";
            this.Load += new System.EventHandler(this.SAEEP_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 流程;
        private System.Windows.Forms.DataGridViewTextBoxColumn 單號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 流程階段;
        private System.Windows.Forms.DataGridViewTextBoxColumn MAILHEAD;
        private System.Windows.Forms.DataGridViewTextBoxColumn MAILTEMP;
        private System.Windows.Forms.DataGridViewTextBoxColumn MAILTO;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}