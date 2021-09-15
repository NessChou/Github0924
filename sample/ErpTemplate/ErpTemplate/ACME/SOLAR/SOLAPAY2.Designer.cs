namespace ACME
{
    partial class SOLAPAY2
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
            this.sOLAR = new ACME.ACMEDataSet.SOLAR();
            this.sOLAR_PAY1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.sOLAR_PAY1TableAdapter = new ACME.ACMEDataSet.SOLARTableAdapters.SOLAR_PAY1TableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.SOLARTableAdapters.TableAdapterManager();
            this.sOLAR_PAY1DataGridView = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.JOBNO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn14 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PAYCHECK = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_PAY1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_PAY1DataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // sOLAR
            // 
            this.sOLAR.DataSetName = "SOLAR";
            this.sOLAR.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // sOLAR_PAY1BindingSource
            // 
            this.sOLAR_PAY1BindingSource.DataMember = "SOLAR_PAY1";
            this.sOLAR_PAY1BindingSource.DataSource = this.sOLAR;
            // 
            // sOLAR_PAY1TableAdapter
            // 
            this.sOLAR_PAY1TableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.SOLAR_PAY1TableAdapter = this.sOLAR_PAY1TableAdapter;
            this.tableAdapterManager.SOLAR_PAYDownloadTableAdapter = null;
            this.tableAdapterManager.SOLAR_PAYTableAdapter = null;
            this.tableAdapterManager.SOLAR_PROBOM2TableAdapter = null;
            this.tableAdapterManager.SOLAR_PROBOMDownloadTableAdapter = null;
            this.tableAdapterManager.SOLAR_PROBOMTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.SOLARTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // sOLAR_PAY1DataGridView
            // 
            this.sOLAR_PAY1DataGridView.AllowUserToAddRows = false;
            this.sOLAR_PAY1DataGridView.AllowUserToDeleteRows = false;
            this.sOLAR_PAY1DataGridView.AutoGenerateColumns = false;
            this.sOLAR_PAY1DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.sOLAR_PAY1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sOLAR_PAY1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.JOBNO,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn11,
            this.dataGridViewTextBoxColumn14,
            this.PAYCHECK});
            this.sOLAR_PAY1DataGridView.DataSource = this.sOLAR_PAY1BindingSource;
            this.sOLAR_PAY1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sOLAR_PAY1DataGridView.Location = new System.Drawing.Point(0, 0);
            this.sOLAR_PAY1DataGridView.Name = "sOLAR_PAY1DataGridView";
            this.sOLAR_PAY1DataGridView.RowTemplate.Height = 24;
            this.sOLAR_PAY1DataGridView.Size = new System.Drawing.Size(1216, 576);
            this.sOLAR_PAY1DataGridView.TabIndex = 2;
            this.sOLAR_PAY1DataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.sOLAR_PAY1DataGridView_CellValueChanged);
            this.sOLAR_PAY1DataGridView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.sOLAR_PAY1DataGridView_MouseDoubleClick);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1216, 35);
            this.panel1.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(403, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "預付請款",
            "採購請款"});
            this.comboBox1.Location = new System.Drawing.Point(284, 6);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(102, 20);
            this.comboBox1.TabIndex = 6;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(178, 4);
            this.textBox2.MaxLength = 8;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(88, 22);
            this.textBox2.TabIndex = 5;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(55, 4);
            this.textBox1.MaxLength = 8;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(88, 22);
            this.textBox1.TabIndex = 4;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.sOLAR_PAY1DataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 35);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1216, 576);
            this.panel2.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "填表日";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(158, 7);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(11, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "~";
            // 
            // JOBNO
            // 
            this.JOBNO.DataPropertyName = "ShippingCode";
            this.JOBNO.HeaderText = "JOBNO";
            this.JOBNO.Name = "JOBNO";
            this.JOBNO.ReadOnly = true;
            this.JOBNO.Width = 66;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "DOCDATE";
            this.dataGridViewTextBoxColumn2.HeaderText = "填表日";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.Width = 66;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "CARDNAME";
            this.dataGridViewTextBoxColumn3.HeaderText = "廠商名稱";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            this.dataGridViewTextBoxColumn3.Width = 78;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "PRJID";
            this.dataGridViewTextBoxColumn4.HeaderText = "專案代碼";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            this.dataGridViewTextBoxColumn4.Width = 78;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "PRJNAME";
            this.dataGridViewTextBoxColumn5.HeaderText = "專案名稱";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            this.dataGridViewTextBoxColumn5.Width = 78;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "OPQTY";
            this.dataGridViewTextBoxColumn9.HeaderText = "採購數量";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.ReadOnly = true;
            this.dataGridViewTextBoxColumn9.Width = 78;
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "OPPRICE";
            this.dataGridViewTextBoxColumn10.HeaderText = "採購單價";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            this.dataGridViewTextBoxColumn10.ReadOnly = true;
            this.dataGridViewTextBoxColumn10.Width = 78;
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.DataPropertyName = "OPAMT";
            this.dataGridViewTextBoxColumn11.HeaderText = "採購金額";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            this.dataGridViewTextBoxColumn11.ReadOnly = true;
            this.dataGridViewTextBoxColumn11.Width = 78;
            // 
            // dataGridViewTextBoxColumn14
            // 
            this.dataGridViewTextBoxColumn14.DataPropertyName = "AMT";
            this.dataGridViewTextBoxColumn14.HeaderText = "請款金額";
            this.dataGridViewTextBoxColumn14.Name = "dataGridViewTextBoxColumn14";
            this.dataGridViewTextBoxColumn14.ReadOnly = true;
            this.dataGridViewTextBoxColumn14.Width = 78;
            // 
            // PAYCHECK
            // 
            this.PAYCHECK.DataPropertyName = "PAYCHECK";
            this.PAYCHECK.HeaderText = "已付款";
            this.PAYCHECK.Name = "PAYCHECK";
            this.PAYCHECK.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.PAYCHECK.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.PAYCHECK.Width = 66;
            // 
            // SOLAPAY2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1216, 611);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "SOLAPAY2";
            this.Text = "支付通知單列表";
            this.Load += new System.EventHandler(this.SOLAPAY2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_PAY1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_PAY1DataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private ACMEDataSet.SOLAR sOLAR;
        private System.Windows.Forms.BindingSource sOLAR_PAY1BindingSource;
        private ACMEDataSet.SOLARTableAdapters.SOLAR_PAY1TableAdapter sOLAR_PAY1TableAdapter;
        private ACMEDataSet.SOLARTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.DataGridView sOLAR_PAY1DataGridView;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn JOBNO;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn14;
        private System.Windows.Forms.DataGridViewCheckBoxColumn PAYCHECK;

    }
}