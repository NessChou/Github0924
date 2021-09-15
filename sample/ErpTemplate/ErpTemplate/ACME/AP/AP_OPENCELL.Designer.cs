namespace ACME
{
    partial class AP_OPENCELL
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
            this.aP_OPENCELLBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.lC = new ACME.ACMEDataSet.LC();
            this.aP_OPENCELLDataGridView = new System.Windows.Forms.DataGridView();
            this.OPENCELL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.KIT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PARTNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.aP_OPENCELLTableAdapter = new ACME.ACMEDataSet.LCTableAdapters.AP_OPENCELLTableAdapter();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.btnOcTconExcelExport = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnOcExcelExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.aP_OPENCELLBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lC)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aP_OPENCELLDataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // aP_OPENCELLBindingSource
            // 
            this.aP_OPENCELLBindingSource.DataMember = "AP_OPENCELL";
            this.aP_OPENCELLBindingSource.DataSource = this.lC;
            // 
            // lC
            // 
            this.lC.DataSetName = "LC";
            this.lC.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // aP_OPENCELLDataGridView
            // 
            this.aP_OPENCELLDataGridView.AutoGenerateColumns = false;
            this.aP_OPENCELLDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.aP_OPENCELLDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.aP_OPENCELLDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.OPENCELL,
            this.KIT,
            this.PARTNO});
            this.aP_OPENCELLDataGridView.DataSource = this.aP_OPENCELLBindingSource;
            this.aP_OPENCELLDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.aP_OPENCELLDataGridView.Location = new System.Drawing.Point(0, 0);
            this.aP_OPENCELLDataGridView.Name = "aP_OPENCELLDataGridView";
            this.aP_OPENCELLDataGridView.RowTemplate.Height = 24;
            this.aP_OPENCELLDataGridView.Size = new System.Drawing.Size(823, 481);
            this.aP_OPENCELLDataGridView.TabIndex = 1;
            this.aP_OPENCELLDataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.aP_OPENCELLDataGridView_CellValueChanged);
            // 
            // OPENCELL
            // 
            this.OPENCELL.DataPropertyName = "OPENCELL";
            this.OPENCELL.HeaderText = "項目編號";
            this.OPENCELL.Name = "OPENCELL";
            this.OPENCELL.Width = 78;
            // 
            // KIT
            // 
            this.KIT.DataPropertyName = "KIT";
            this.KIT.HeaderText = "對應T-CON";
            this.KIT.Name = "KIT";
            this.KIT.Width = 89;
            // 
            // PARTNO
            // 
            this.PARTNO.DataPropertyName = "PARTNO";
            this.PARTNO.HeaderText = "PARTNO";
            this.PARTNO.Name = "PARTNO";
            this.PARTNO.Width = 75;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "項目編號";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(83, 6);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(134, 22);
            this.textBox1.TabIndex = 3;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // aP_OPENCELLTableAdapter
            // 
            this.aP_OPENCELLTableAdapter.ClearBeforeFill = true;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.btnOcExcelExport);
            this.panel1.Controls.Add(this.btnOcTconExcelExport);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(823, 115);
            this.panel1.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Location = new System.Drawing.Point(337, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(474, 100);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "匯入";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "China",
            "Taiwan"});
            this.comboBox1.Location = new System.Drawing.Point(197, 22);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(87, 20);
            this.comboBox1.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(162, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 14;
            this.label4.Text = "產地";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(300, 21);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 56);
            this.button4.TabIndex = 13;
            this.button4.Text = "匯入";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 12);
            this.label2.TabIndex = 10;
            this.label2.Text = "PARTNO";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(63, 22);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(93, 22);
            this.textBox2.TabIndex = 10;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 49);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(89, 23);
            this.button3.TabIndex = 8;
            this.button3.Text = "查詢";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btnOcTconExcelExport
            // 
            this.btnOcTconExcelExport.Location = new System.Drawing.Point(14, 78);
            this.btnOcTconExcelExport.Name = "btnOcTconExcelExport";
            this.btnOcTconExcelExport.Size = new System.Drawing.Size(87, 23);
            this.btnOcTconExcelExport.TabIndex = 7;
            this.btnOcTconExcelExport.Text = "OC對應TCON";
            this.btnOcTconExcelExport.UseVisualStyleBackColor = true;
            this.btnOcTconExcelExport.Click += new System.EventHandler(this.btnExcelExport_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(198, 49);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "匯出EXCEL";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(107, 49);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(85, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "存檔";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.aP_OPENCELLDataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 115);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(823, 481);
            this.panel2.TabIndex = 5;
            // 
            // btnOcExcelExport
            // 
            this.btnOcExcelExport.Location = new System.Drawing.Point(107, 78);
            this.btnOcExcelExport.Name = "btnOcExcelExport";
            this.btnOcExcelExport.Size = new System.Drawing.Size(87, 23);
            this.btnOcExcelExport.TabIndex = 7;
            this.btnOcExcelExport.Text = "沒有TCON";
            this.btnOcExcelExport.UseVisualStyleBackColor = true;
            this.btnOcExcelExport.Click += new System.EventHandler(this.btnExcelExport_Click);
            // 
            // AP_OPENCELL
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(823, 596);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "AP_OPENCELL";
            this.Text = "OpenCell對應T-CON";
            this.Load += new System.EventHandler(this.AP_OPENCELL_Load);
            ((System.ComponentModel.ISupportInitialize)(this.aP_OPENCELLBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lC)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aP_OPENCELLDataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private ACMEDataSet.LC lC;
        private System.Windows.Forms.BindingSource aP_OPENCELLBindingSource;
        private ACMEDataSet.LCTableAdapters.AP_OPENCELLTableAdapter aP_OPENCELLTableAdapter;
        private System.Windows.Forms.DataGridView aP_OPENCELLDataGridView;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridViewTextBoxColumn OPENCELL;
        private System.Windows.Forms.DataGridViewTextBoxColumn KIT;
        private System.Windows.Forms.DataGridViewTextBoxColumn PARTNO;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button btnOcTconExcelExport;
        private System.Windows.Forms.Button btnOcExcelExport;
    }
}