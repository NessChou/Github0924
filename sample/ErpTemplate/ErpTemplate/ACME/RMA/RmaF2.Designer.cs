namespace ACME
{
    partial class RmaF2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RmaF2));
            this.rm = new ACME.ACMEDataSet.rm();
            this.rMA_CTR1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.rMA_CTR1TableAdapter = new ACME.ACMEDataSet.rmTableAdapters.RMA_CTR1TableAdapter();
            this.rMA_CTR1BindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.rMA_CTR1BindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.rMA_CTR1DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.rm)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CTR1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CTR1BindingNavigator)).BeginInit();
            this.rMA_CTR1BindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CTR1DataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // rm
            // 
            this.rm.DataSetName = "rm";
            this.rm.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // rMA_CTR1BindingSource
            // 
            this.rMA_CTR1BindingSource.DataMember = "RMA_CTR1";
            this.rMA_CTR1BindingSource.DataSource = this.rm;
            // 
            // rMA_CTR1TableAdapter
            // 
            this.rMA_CTR1TableAdapter.ClearBeforeFill = true;
            // 
            // rMA_CTR1BindingNavigator
            // 
            this.rMA_CTR1BindingNavigator.AddNewItem = null;
            this.rMA_CTR1BindingNavigator.BindingSource = this.rMA_CTR1BindingSource;
            this.rMA_CTR1BindingNavigator.CountItem = null;
            this.rMA_CTR1BindingNavigator.DeleteItem = null;
            this.rMA_CTR1BindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.rMA_CTR1BindingNavigatorSaveItem});
            this.rMA_CTR1BindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.rMA_CTR1BindingNavigator.MoveFirstItem = null;
            this.rMA_CTR1BindingNavigator.MoveLastItem = null;
            this.rMA_CTR1BindingNavigator.MoveNextItem = null;
            this.rMA_CTR1BindingNavigator.MovePreviousItem = null;
            this.rMA_CTR1BindingNavigator.Name = "rMA_CTR1BindingNavigator";
            this.rMA_CTR1BindingNavigator.PositionItem = null;
            this.rMA_CTR1BindingNavigator.Size = new System.Drawing.Size(1155, 25);
            this.rMA_CTR1BindingNavigator.TabIndex = 0;
            this.rMA_CTR1BindingNavigator.Text = "bindingNavigator1";
            // 
            // rMA_CTR1BindingNavigatorSaveItem
            // 
            this.rMA_CTR1BindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.rMA_CTR1BindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("rMA_CTR1BindingNavigatorSaveItem.Image")));
            this.rMA_CTR1BindingNavigatorSaveItem.Name = "rMA_CTR1BindingNavigatorSaveItem";
            this.rMA_CTR1BindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.rMA_CTR1BindingNavigatorSaveItem.Text = "儲存資料";
            this.rMA_CTR1BindingNavigatorSaveItem.Click += new System.EventHandler(this.rMA_CTR1BindingNavigatorSaveItem_Click);
            // 
            // rMA_CTR1DataGridView
            // 
            this.rMA_CTR1DataGridView.AllowUserToAddRows = false;
            this.rMA_CTR1DataGridView.AutoGenerateColumns = false;
            this.rMA_CTR1DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.rMA_CTR1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.rMA_CTR1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7,
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn11,
            this.dataGridViewTextBoxColumn12});
            this.rMA_CTR1DataGridView.DataSource = this.rMA_CTR1BindingSource;
            this.rMA_CTR1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rMA_CTR1DataGridView.Location = new System.Drawing.Point(0, 25);
            this.rMA_CTR1DataGridView.Name = "rMA_CTR1DataGridView";
            this.rMA_CTR1DataGridView.RowTemplate.Height = 24;
            this.rMA_CTR1DataGridView.Size = new System.Drawing.Size(1155, 550);
            this.rMA_CTR1DataGridView.TabIndex = 2;
            this.rMA_CTR1DataGridView.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.rMA_CTR1DataGridView_RowPostPaint);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "RMA NO";
            this.dataGridViewTextBoxColumn1.HeaderText = "RMA NO";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.Width = 69;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "S/N";
            this.dataGridViewTextBoxColumn2.HeaderText = "S/N";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 47;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "Model";
            this.dataGridViewTextBoxColumn3.HeaderText = "Model";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 60;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "Ver";
            this.dataGridViewTextBoxColumn4.HeaderText = "Ver";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.Width = 47;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "W/C";
            this.dataGridViewTextBoxColumn5.HeaderText = "W/C";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.Width = 52;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "IQC/CLR/FR";
            this.dataGridViewTextBoxColumn6.HeaderText = "IQC/CLR/FR";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.Width = 93;
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "Customer Complain";
            this.dataGridViewTextBoxColumn7.HeaderText = "Customer Complain";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.Width = 114;
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.DataPropertyName = "ACMEPOINT Confirm";
            this.dataGridViewTextBoxColumn8.HeaderText = "ACMEPOINT Confirm";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            this.dataGridViewTextBoxColumn8.Width = 127;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "ACMEPOINT Judge";
            this.dataGridViewTextBoxColumn9.HeaderText = "ACMEPOINT Judge";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.Width = 115;
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "產地";
            this.dataGridViewTextBoxColumn10.HeaderText = "產地";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            this.dataGridViewTextBoxColumn10.Width = 51;
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.DataPropertyName = "REMARK";
            this.dataGridViewTextBoxColumn11.HeaderText = "REMARK";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            this.dataGridViewTextBoxColumn11.Width = 79;
            // 
            // dataGridViewTextBoxColumn12
            // 
            this.dataGridViewTextBoxColumn12.DataPropertyName = "ID";
            this.dataGridViewTextBoxColumn12.HeaderText = "ID";
            this.dataGridViewTextBoxColumn12.Name = "dataGridViewTextBoxColumn12";
            this.dataGridViewTextBoxColumn12.Visible = false;
            this.dataGridViewTextBoxColumn12.Width = 42;
            // 
            // RmaF2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1155, 575);
            this.Controls.Add(this.rMA_CTR1DataGridView);
            this.Controls.Add(this.rMA_CTR1BindingNavigator);
            this.Name = "RmaF2";
            this.Text = "複判明細";
            this.Load += new System.EventHandler(this.RmaF2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.rm)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CTR1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CTR1BindingNavigator)).EndInit();
            this.rMA_CTR1BindingNavigator.ResumeLayout(false);
            this.rMA_CTR1BindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CTR1DataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.rm rm;
        private System.Windows.Forms.BindingSource rMA_CTR1BindingSource;
        private ACMEDataSet.rmTableAdapters.RMA_CTR1TableAdapter rMA_CTR1TableAdapter;
        private System.Windows.Forms.BindingNavigator rMA_CTR1BindingNavigator;
        private System.Windows.Forms.ToolStripButton rMA_CTR1BindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView rMA_CTR1DataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;

    }
}