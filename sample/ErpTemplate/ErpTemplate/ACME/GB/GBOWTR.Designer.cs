namespace ACME
{
    partial class GBOWTR
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.Label modAdjNOLabel;
            System.Windows.Forms.Label modAdjNameLabel;
            System.Windows.Forms.Label adjustTypeLabel;
            System.Windows.Forms.Label remarkLabel;
            System.Windows.Forms.Label label1;
            this.cHOICE = new ACME.ACMEDataSet.CHOICE();
            this.stkModAdjMainBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.stkModAdjMainTableAdapter = new ACME.ACMEDataSet.CHOICETableAdapters.stkModAdjMainTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.CHOICETableAdapters.TableAdapterManager();
            this.modAdjNOTextBox = new System.Windows.Forms.TextBox();
            this.stkModAdjSubBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.stkModAdjSubTableAdapter = new ACME.ACMEDataSet.CHOICETableAdapters.stkModAdjSubTableAdapter();
            this.stkModAdjSubDataGridView = new System.Windows.Forms.DataGridView();
            this.modAdjNameTextBox = new System.Windows.Forms.TextBox();
            this.adjustTypeTextBox = new System.Windows.Forms.TextBox();
            this.remarkTextBox = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            modAdjNOLabel = new System.Windows.Forms.Label();
            modAdjNameLabel = new System.Windows.Forms.Label();
            adjustTypeLabel = new System.Windows.Forms.Label();
            remarkLabel = new System.Windows.Forms.Label();
            label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.cHOICE)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.stkModAdjMainBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.stkModAdjSubBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.stkModAdjSubDataGridView)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Size = new System.Drawing.Size(834, 476);
            this.panel1.Controls.SetChildIndex(this.panel2, 0);
            this.panel1.Controls.SetChildIndex(this.panel3, 0);
            // 
            // modAdjNOLabel
            // 
            modAdjNOLabel.AutoSize = true;
            modAdjNOLabel.Location = new System.Drawing.Point(16, 15);
            modAdjNOLabel.Name = "modAdjNOLabel";
            modAdjNOLabel.Size = new System.Drawing.Size(29, 12);
            modAdjNOLabel.TabIndex = 1;
            modAdjNOLabel.Text = "編號";
            // 
            // modAdjNameLabel
            // 
            modAdjNameLabel.AutoSize = true;
            modAdjNameLabel.Location = new System.Drawing.Point(220, 15);
            modAdjNameLabel.Name = "modAdjNameLabel";
            modAdjNameLabel.Size = new System.Drawing.Size(29, 12);
            modAdjNameLabel.TabIndex = 4;
            modAdjNameLabel.Text = "名稱";
            // 
            // adjustTypeLabel
            // 
            adjustTypeLabel.AutoSize = true;
            adjustTypeLabel.Location = new System.Drawing.Point(20, 41);
            adjustTypeLabel.Name = "adjustTypeLabel";
            adjustTypeLabel.Size = new System.Drawing.Size(53, 12);
            adjustTypeLabel.TabIndex = 6;
            adjustTypeLabel.Text = "調整類別";
            // 
            // remarkLabel
            // 
            remarkLabel.AutoSize = true;
            remarkLabel.Location = new System.Drawing.Point(20, 71);
            remarkLabel.Name = "remarkLabel";
            remarkLabel.Size = new System.Drawing.Size(29, 12);
            remarkLabel.TabIndex = 10;
            remarkLabel.Text = "備註";
            // 
            // cHOICE
            // 
            this.cHOICE.DataSetName = "CHOICE";
            this.cHOICE.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // stkModAdjMainBindingSource
            // 
            this.stkModAdjMainBindingSource.DataMember = "stkModAdjMain";
            this.stkModAdjMainBindingSource.DataSource = this.cHOICE;
            // 
            // stkModAdjMainTableAdapter
            // 
            this.stkModAdjMainTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.stkModAdjMainTableAdapter = this.stkModAdjMainTableAdapter;
            this.tableAdapterManager.stkModAdjSubTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.CHOICETableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // modAdjNOTextBox
            // 
            this.modAdjNOTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.stkModAdjMainBindingSource, "ModAdjNO", true));
            this.modAdjNOTextBox.Location = new System.Drawing.Point(91, 12);
            this.modAdjNOTextBox.Name = "modAdjNOTextBox";
            this.modAdjNOTextBox.Size = new System.Drawing.Size(100, 22);
            this.modAdjNOTextBox.TabIndex = 2;
            // 
            // stkModAdjSubBindingSource
            // 
            this.stkModAdjSubBindingSource.DataMember = "stkModAdjMain_stkModAdjSub";
            this.stkModAdjSubBindingSource.DataSource = this.stkModAdjMainBindingSource;
            // 
            // stkModAdjSubTableAdapter
            // 
            this.stkModAdjSubTableAdapter.ClearBeforeFill = true;
            // 
            // stkModAdjSubDataGridView
            // 
            this.stkModAdjSubDataGridView.AllowUserToAddRows = false;
            this.stkModAdjSubDataGridView.AllowUserToDeleteRows = false;
            this.stkModAdjSubDataGridView.AutoGenerateColumns = false;
            this.stkModAdjSubDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.stkModAdjSubDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.stkModAdjSubDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn9});
            this.stkModAdjSubDataGridView.DataSource = this.stkModAdjSubBindingSource;
            this.stkModAdjSubDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.stkModAdjSubDataGridView.Location = new System.Drawing.Point(0, 0);
            this.stkModAdjSubDataGridView.Name = "stkModAdjSubDataGridView";
            this.stkModAdjSubDataGridView.ReadOnly = true;
            this.stkModAdjSubDataGridView.RowTemplate.Height = 24;
            this.stkModAdjSubDataGridView.Size = new System.Drawing.Size(834, 354);
            this.stkModAdjSubDataGridView.TabIndex = 3;
            // 
            // modAdjNameTextBox
            // 
            this.modAdjNameTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.stkModAdjMainBindingSource, "ModAdjName", true));
            this.modAdjNameTextBox.Location = new System.Drawing.Point(266, 12);
            this.modAdjNameTextBox.Name = "modAdjNameTextBox";
            this.modAdjNameTextBox.Size = new System.Drawing.Size(207, 22);
            this.modAdjNameTextBox.TabIndex = 5;
            // 
            // adjustTypeTextBox
            // 
            this.adjustTypeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.stkModAdjMainBindingSource, "AdjustType", true));
            this.adjustTypeTextBox.Location = new System.Drawing.Point(91, 38);
            this.adjustTypeTextBox.Name = "adjustTypeTextBox";
            this.adjustTypeTextBox.Size = new System.Drawing.Size(100, 22);
            this.adjustTypeTextBox.TabIndex = 7;
            // 
            // remarkTextBox
            // 
            this.remarkTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.stkModAdjMainBindingSource, "Remark", true));
            this.remarkTextBox.Location = new System.Drawing.Point(91, 68);
            this.remarkTextBox.Name = "remarkTextBox";
            this.remarkTextBox.Size = new System.Drawing.Size(257, 22);
            this.remarkTextBox.TabIndex = 11;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(label1);
            this.panel2.Controls.Add(this.textBox1);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(modAdjNOLabel);
            this.panel2.Controls.Add(this.modAdjNOTextBox);
            this.panel2.Controls.Add(this.modAdjNameTextBox);
            this.panel2.Controls.Add(remarkLabel);
            this.panel2.Controls.Add(modAdjNameLabel);
            this.panel2.Controls.Add(this.remarkTextBox);
            this.panel2.Controls.Add(this.adjustTypeTextBox);
            this.panel2.Controls.Add(adjustTypeLabel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(834, 100);
            this.panel2.TabIndex = 14;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.stkModAdjSubDataGridView);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 100);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(834, 354);
            this.panel3.TabIndex = 15;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(671, 10);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 14;
            this.button1.Text = "修改";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(565, 9);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 22);
            this.textBox1.TabIndex = 15;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(506, 15);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(53, 12);
            label1.TabIndex = 16;
            label1.Text = "計劃數量";
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "SerNo";
            this.dataGridViewTextBoxColumn2.HeaderText = "欄號";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.Width = 54;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "ProdID";
            this.dataGridViewTextBoxColumn3.HeaderText = "產品編號";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            this.dataGridViewTextBoxColumn3.Width = 78;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "ProdName";
            this.dataGridViewTextBoxColumn4.HeaderText = "品名規格";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            this.dataGridViewTextBoxColumn4.Width = 78;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "WareHouseID";
            this.dataGridViewTextBoxColumn5.HeaderText = "倉庫";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            this.dataGridViewTextBoxColumn5.Width = 54;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "Quantity";
            this.dataGridViewTextBoxColumn6.HeaderText = "數量";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.ReadOnly = true;
            this.dataGridViewTextBoxColumn6.Width = 54;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "ItemRemark";
            this.dataGridViewTextBoxColumn9.HeaderText = "分錄備註";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.ReadOnly = true;
            this.dataGridViewTextBoxColumn9.Width = 78;
            // 
            // GBOWTR
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(834, 515);
            this.Name = "GBOWTR";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.cHOICE)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.stkModAdjMainBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.stkModAdjSubBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.stkModAdjSubDataGridView)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.CHOICE cHOICE;
        private System.Windows.Forms.BindingSource stkModAdjMainBindingSource;
        private ACMEDataSet.CHOICETableAdapters.stkModAdjMainTableAdapter stkModAdjMainTableAdapter;
        private ACMEDataSet.CHOICETableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.TextBox modAdjNOTextBox;
        private System.Windows.Forms.BindingSource stkModAdjSubBindingSource;
        private ACMEDataSet.CHOICETableAdapters.stkModAdjSubTableAdapter stkModAdjSubTableAdapter;
        private System.Windows.Forms.DataGridView stkModAdjSubDataGridView;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox modAdjNameTextBox;
        private System.Windows.Forms.TextBox remarkTextBox;
        private System.Windows.Forms.TextBox adjustTypeTextBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
    }
}
