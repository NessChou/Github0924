namespace ACME
{
    partial class KITPACKING
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
            this.packingListDKITBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.packingListDKITBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ship = new ACME.ACMEDataSet.ship();
            this.packingListDKITTableAdapter = new ACME.ACMEDataSet.shipTableAdapters.PackingListDKITTableAdapter();
            this.button1 = new System.Windows.Forms.Button();
            this.packingListDKITDataGridView = new System.Windows.Forms.DataGridView();
            this.SeqNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.QTY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Net1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Gross1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.packingListDKITBindingNavigator)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.packingListDKITBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.packingListDKITDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // packingListDKITBindingNavigator
            // 
            this.packingListDKITBindingNavigator.AddNewItem = null;
            this.packingListDKITBindingNavigator.BindingSource = this.packingListDKITBindingSource;
            this.packingListDKITBindingNavigator.CountItem = null;
            this.packingListDKITBindingNavigator.DeleteItem = null;
            this.packingListDKITBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.packingListDKITBindingNavigator.MoveFirstItem = null;
            this.packingListDKITBindingNavigator.MoveLastItem = null;
            this.packingListDKITBindingNavigator.MoveNextItem = null;
            this.packingListDKITBindingNavigator.MovePreviousItem = null;
            this.packingListDKITBindingNavigator.Name = "packingListDKITBindingNavigator";
            this.packingListDKITBindingNavigator.PositionItem = null;
            this.packingListDKITBindingNavigator.Size = new System.Drawing.Size(875, 25);
            this.packingListDKITBindingNavigator.TabIndex = 0;
            this.packingListDKITBindingNavigator.Text = "bindingNavigator1";
            // 
            // packingListDKITBindingSource
            // 
            this.packingListDKITBindingSource.DataMember = "PackingListDKIT";
            this.packingListDKITBindingSource.DataSource = this.ship;
            // 
            // ship
            // 
            this.ship.DataSetName = "ship";
            this.ship.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // packingListDKITTableAdapter
            // 
            this.packingListDKITTableAdapter.ClearBeforeFill = true;
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(21, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "存檔";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // packingListDKITDataGridView
            // 
            this.packingListDKITDataGridView.AutoGenerateColumns = false;
            this.packingListDKITDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.packingListDKITDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SeqNo,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn5,
            this.QTY,
            this.Net1,
            this.Gross1});
            this.packingListDKITDataGridView.DataSource = this.packingListDKITBindingSource;
            this.packingListDKITDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.packingListDKITDataGridView.Location = new System.Drawing.Point(0, 25);
            this.packingListDKITDataGridView.Name = "packingListDKITDataGridView";
            this.packingListDKITDataGridView.RowTemplate.Height = 24;
            this.packingListDKITDataGridView.Size = new System.Drawing.Size(875, 534);
            this.packingListDKITDataGridView.TabIndex = 3;
            this.packingListDKITDataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.packingListDKITDataGridView_DefaultValuesNeeded);
            // 
            // SeqNo
            // 
            this.SeqNo.DataPropertyName = "SeqNo";
            this.SeqNo.HeaderText = "No";
            this.SeqNo.Name = "SeqNo";
            this.SeqNo.ReadOnly = true;
            this.SeqNo.Width = 50;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "PLT";
            this.dataGridViewTextBoxColumn9.HeaderText = "PLT";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "CNo";
            this.dataGridViewTextBoxColumn10.HeaderText = "C/No";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "KIT";
            this.dataGridViewTextBoxColumn5.HeaderText = "子料件";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // QTY
            // 
            this.QTY.DataPropertyName = "QTY";
            this.QTY.HeaderText = "QTY";
            this.QTY.Name = "QTY";
            // 
            // Net1
            // 
            this.Net1.DataPropertyName = "Net";
            this.Net1.HeaderText = "Net";
            this.Net1.Name = "Net1";
            // 
            // Gross1
            // 
            this.Gross1.DataPropertyName = "Gross";
            this.Gross1.HeaderText = "Gross";
            this.Gross1.Name = "Gross1";
            // 
            // KITPACKING
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(875, 559);
            this.Controls.Add(this.packingListDKITDataGridView);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.packingListDKITBindingNavigator);
            this.Name = "KITPACKING";
            this.Text = "KITPACKING";
            this.Load += new System.EventHandler(this.KITPACKING_Load);
            ((System.ComponentModel.ISupportInitialize)(this.packingListDKITBindingNavigator)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.packingListDKITBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.packingListDKITDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship ship;
        private System.Windows.Forms.BindingSource packingListDKITBindingSource;
        private ACMEDataSet.shipTableAdapters.PackingListDKITTableAdapter packingListDKITTableAdapter;
        private System.Windows.Forms.BindingNavigator packingListDKITBindingNavigator;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView packingListDKITDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn SeqNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn QTY;
        private System.Windows.Forms.DataGridViewTextBoxColumn Net1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Gross1;
    }
}