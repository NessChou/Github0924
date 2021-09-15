namespace ACME
{
    partial class RMAODLN
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
            this.components = new System.ComponentModel.Container();
            this.rm = new ACME.ACMEDataSet.rm();
            this.pARAMSBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.pARAMSTableAdapter = new ACME.ACMEDataSet.rmTableAdapters.PARAMSTableAdapter();
            this.pARAMSDataGridView = new System.Windows.Forms.DataGridView();
            this.PARAM_KIND = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.rm)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pARAMSBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pARAMSDataGridView)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // rm
            // 
            this.rm.DataSetName = "rm";
            this.rm.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // pARAMSBindingSource
            // 
            this.pARAMSBindingSource.DataMember = "PARAMS";
            this.pARAMSBindingSource.DataSource = this.rm;
            // 
            // pARAMSTableAdapter
            // 
            this.pARAMSTableAdapter.ClearBeforeFill = true;
            // 
            // pARAMSDataGridView
            // 
            this.pARAMSDataGridView.AutoGenerateColumns = false;
            this.pARAMSDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.pARAMSDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PARAM_KIND,
            this.dataGridViewTextBoxColumn3});
            this.pARAMSDataGridView.DataSource = this.pARAMSBindingSource;
            this.pARAMSDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pARAMSDataGridView.Location = new System.Drawing.Point(0, 0);
            this.pARAMSDataGridView.Name = "pARAMSDataGridView";
            this.pARAMSDataGridView.RowTemplate.Height = 24;
            this.pARAMSDataGridView.Size = new System.Drawing.Size(618, 576);
            this.pARAMSDataGridView.TabIndex = 1;
            this.pARAMSDataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.pARAMSDataGridView_DefaultValuesNeeded);
            // 
            // PARAM_KIND
            // 
            this.PARAM_KIND.DataPropertyName = "PARAM_KIND";
            this.PARAM_KIND.HeaderText = "PARAM_KIND";
            this.PARAM_KIND.Name = "PARAM_KIND";
            this.PARAM_KIND.Visible = false;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "PARAM_NO";
            this.dataGridViewTextBoxColumn3.HeaderText = "倉庫";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(18, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "存檔";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel2);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(624, 637);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "倉庫維護";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.pARAMSDataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 58);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(618, 576);
            this.panel2.TabIndex = 4;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(618, 40);
            this.panel1.TabIndex = 3;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(47, 21);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 20);
            this.comboBox1.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "倉庫";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "RMANO";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(64, 48);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(163, 127);
            this.textBox1.TabIndex = 7;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(249, 46);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "EXCEL";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.label1);
            this.panel3.Controls.Add(this.comboBox1);
            this.panel3.Controls.Add(this.textBox1);
            this.panel3.Controls.Add(this.button2);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(363, 637);
            this.panel3.TabIndex = 9;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.groupBox1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(363, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(624, 637);
            this.panel4.TabIndex = 10;
            // 
            // RMAODLN
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(987, 637);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Name = "RMAODLN";
            this.Text = "收貨工單";
            this.Load += new System.EventHandler(this.RMAODLN_Load);
            ((System.ComponentModel.ISupportInitialize)(this.rm)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pARAMSBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pARAMSDataGridView)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private ACMEDataSet.rm rm;
        private System.Windows.Forms.BindingSource pARAMSBindingSource;
        private ACMEDataSet.rmTableAdapters.PARAMSTableAdapter pARAMSTableAdapter;
        private System.Windows.Forms.DataGridView pARAMSDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn PARAM_KIND;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
    }
}