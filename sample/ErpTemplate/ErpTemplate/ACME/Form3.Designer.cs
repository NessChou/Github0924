namespace ACME
{
    partial class Form3
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
            this.ship2 = new ACME.ACMEDataSet.ship2();
            this.shipping_OQUTDownloadBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.shipping_OQUTDownloadTableAdapter = new ACME.ACMEDataSet.ship2TableAdapters.Shipping_OQUTDownloadTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.ship2TableAdapters.TableAdapterManager();
            this.fillToolStrip = new System.Windows.Forms.ToolStrip();
            this.shippingcodeToolStripLabel = new System.Windows.Forms.ToolStripLabel();
            this.shippingcodeToolStripTextBox = new System.Windows.Forms.ToolStripTextBox();
            this.fillToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.shipping_OQUTDownloadDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.shipping_OQUTDownload2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.shipping_OQUTDownload2TableAdapter = new ACME.ACMEDataSet.ship2TableAdapters.Shipping_OQUTDownload2TableAdapter();
            this.fillToolStrip1 = new System.Windows.Forms.ToolStrip();
            this.shippingCodeToolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.shippingCodeToolStripTextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.fillToolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.shipping_OQUTDownload2DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ship2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadBindingSource)).BeginInit();
            this.fillToolStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownload2BindingSource)).BeginInit();
            this.fillToolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownload2DataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.shipping_OQUTDownload2DataGridView);
            this.panel1.Controls.Add(this.shipping_OQUTDownloadDataGridView);
            this.panel1.Size = new System.Drawing.Size(993, 580);
            this.panel1.Controls.SetChildIndex(this.shipping_OQUTDownloadDataGridView, 0);
            this.panel1.Controls.SetChildIndex(this.shipping_OQUTDownload2DataGridView, 0);
            // 
            // ship2
            // 
            this.ship2.DataSetName = "ship2";
            this.ship2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // shipping_OQUTDownloadBindingSource
            // 
            this.shipping_OQUTDownloadBindingSource.DataMember = "Shipping_OQUTDownload";
            this.shipping_OQUTDownloadBindingSource.DataSource = this.ship2;
            // 
            // shipping_OQUTDownloadTableAdapter
            // 
            this.shipping_OQUTDownloadTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.Shipping_OQUT1TableAdapter = null;
            this.tableAdapterManager.Shipping_OQUTDownload2TableAdapter = this.shipping_OQUTDownload2TableAdapter;
            this.tableAdapterManager.Shipping_OQUTDownloadTableAdapter = this.shipping_OQUTDownloadTableAdapter;
            this.tableAdapterManager.Shipping_OQUTTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.ship2TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // fillToolStrip
            // 
            this.fillToolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.shippingcodeToolStripLabel,
            this.shippingcodeToolStripTextBox,
            this.fillToolStripButton});
            this.fillToolStrip.Location = new System.Drawing.Point(0, 39);
            this.fillToolStrip.Name = "fillToolStrip";
            this.fillToolStrip.Size = new System.Drawing.Size(993, 25);
            this.fillToolStrip.TabIndex = 3;
            this.fillToolStrip.Text = "fillToolStrip";
            // 
            // shippingcodeToolStripLabel
            // 
            this.shippingcodeToolStripLabel.Name = "shippingcodeToolStripLabel";
            this.shippingcodeToolStripLabel.Size = new System.Drawing.Size(89, 22);
            this.shippingcodeToolStripLabel.Text = "shippingcode:";
            // 
            // shippingcodeToolStripTextBox
            // 
            this.shippingcodeToolStripTextBox.Name = "shippingcodeToolStripTextBox";
            this.shippingcodeToolStripTextBox.Size = new System.Drawing.Size(100, 25);
            // 
            // fillToolStripButton
            // 
            this.fillToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.fillToolStripButton.Name = "fillToolStripButton";
            this.fillToolStripButton.Size = new System.Drawing.Size(27, 22);
            this.fillToolStripButton.Text = "Fill";
            this.fillToolStripButton.Click += new System.EventHandler(this.fillToolStripButton_Click);
            // 
            // shipping_OQUTDownloadDataGridView
            // 
            this.shipping_OQUTDownloadDataGridView.AutoGenerateColumns = false;
            this.shipping_OQUTDownloadDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.shipping_OQUTDownloadDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5});
            this.shipping_OQUTDownloadDataGridView.DataSource = this.shipping_OQUTDownloadBindingSource;
            this.shipping_OQUTDownloadDataGridView.Location = new System.Drawing.Point(457, 298);
            this.shipping_OQUTDownloadDataGridView.Name = "shipping_OQUTDownloadDataGridView";
            this.shipping_OQUTDownloadDataGridView.RowTemplate.Height = 24;
            this.shipping_OQUTDownloadDataGridView.Size = new System.Drawing.Size(300, 220);
            this.shipping_OQUTDownloadDataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "ID";
            this.dataGridViewTextBoxColumn1.HeaderText = "ID";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "shippingcode";
            this.dataGridViewTextBoxColumn2.HeaderText = "shippingcode";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "seq";
            this.dataGridViewTextBoxColumn3.HeaderText = "seq";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "filename";
            this.dataGridViewTextBoxColumn4.HeaderText = "filename";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "path";
            this.dataGridViewTextBoxColumn5.HeaderText = "path";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // shipping_OQUTDownload2BindingSource
            // 
            this.shipping_OQUTDownload2BindingSource.DataMember = "Shipping_OQUTDownload2";
            this.shipping_OQUTDownload2BindingSource.DataSource = this.ship2;
            // 
            // shipping_OQUTDownload2TableAdapter
            // 
            this.shipping_OQUTDownload2TableAdapter.ClearBeforeFill = true;
            // 
            // fillToolStrip1
            // 
            this.fillToolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.shippingCodeToolStripLabel1,
            this.shippingCodeToolStripTextBox1,
            this.fillToolStripButton1});
            this.fillToolStrip1.Location = new System.Drawing.Point(0, 64);
            this.fillToolStrip1.Name = "fillToolStrip1";
            this.fillToolStrip1.Size = new System.Drawing.Size(993, 25);
            this.fillToolStrip1.TabIndex = 4;
            this.fillToolStrip1.Text = "fillToolStrip1";
            // 
            // shippingCodeToolStripLabel1
            // 
            this.shippingCodeToolStripLabel1.Name = "shippingCodeToolStripLabel1";
            this.shippingCodeToolStripLabel1.Size = new System.Drawing.Size(93, 16);
            this.shippingCodeToolStripLabel1.Text = "ShippingCode:";
            // 
            // shippingCodeToolStripTextBox1
            // 
            this.shippingCodeToolStripTextBox1.Name = "shippingCodeToolStripTextBox1";
            this.shippingCodeToolStripTextBox1.Size = new System.Drawing.Size(100, 23);
            // 
            // fillToolStripButton1
            // 
            this.fillToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.fillToolStripButton1.Name = "fillToolStripButton1";
            this.fillToolStripButton1.Size = new System.Drawing.Size(27, 20);
            this.fillToolStripButton1.Text = "Fill";
            this.fillToolStripButton1.Click += new System.EventHandler(this.fillToolStripButton1_Click);
            // 
            // shipping_OQUTDownload2DataGridView
            // 
            this.shipping_OQUTDownload2DataGridView.AutoGenerateColumns = false;
            this.shipping_OQUTDownload2DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.shipping_OQUTDownload2DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7,
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn11,
            this.dataGridViewTextBoxColumn12});
            this.shipping_OQUTDownload2DataGridView.DataSource = this.shipping_OQUTDownload2BindingSource;
            this.shipping_OQUTDownload2DataGridView.Location = new System.Drawing.Point(662, 189);
            this.shipping_OQUTDownload2DataGridView.Name = "shipping_OQUTDownload2DataGridView";
            this.shipping_OQUTDownload2DataGridView.RowTemplate.Height = 24;
            this.shipping_OQUTDownload2DataGridView.Size = new System.Drawing.Size(300, 220);
            this.shipping_OQUTDownload2DataGridView.TabIndex = 2;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "ID";
            this.dataGridViewTextBoxColumn6.HeaderText = "ID";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "shippingcode";
            this.dataGridViewTextBoxColumn7.HeaderText = "shippingcode";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.DataPropertyName = "seq";
            this.dataGridViewTextBoxColumn8.HeaderText = "seq";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "filename";
            this.dataGridViewTextBoxColumn9.HeaderText = "filename";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "path";
            this.dataGridViewTextBoxColumn10.HeaderText = "path";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.DataPropertyName = "CARDCODE";
            this.dataGridViewTextBoxColumn11.HeaderText = "CARDCODE";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            // 
            // dataGridViewTextBoxColumn12
            // 
            this.dataGridViewTextBoxColumn12.DataPropertyName = "CARDNAME";
            this.dataGridViewTextBoxColumn12.HeaderText = "CARDNAME";
            this.dataGridViewTextBoxColumn12.Name = "dataGridViewTextBoxColumn12";
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(993, 619);
            this.Controls.Add(this.fillToolStrip1);
            this.Controls.Add(this.fillToolStrip);
            this.Name = "Form3";
            this.Controls.SetChildIndex(this.panel1, 0);
            this.Controls.SetChildIndex(this.fillToolStrip, 0);
            this.Controls.SetChildIndex(this.fillToolStrip1, 0);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ship2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadBindingSource)).EndInit();
            this.fillToolStrip.ResumeLayout(false);
            this.fillToolStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownload2BindingSource)).EndInit();
            this.fillToolStrip1.ResumeLayout(false);
            this.fillToolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownload2DataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship2 ship2;
        private System.Windows.Forms.BindingSource shipping_OQUTDownloadBindingSource;
        private ACMEDataSet.ship2TableAdapters.Shipping_OQUTDownloadTableAdapter shipping_OQUTDownloadTableAdapter;
        private ACMEDataSet.ship2TableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.ToolStrip fillToolStrip;
        private System.Windows.Forms.ToolStripLabel shippingcodeToolStripLabel;
        private System.Windows.Forms.ToolStripTextBox shippingcodeToolStripTextBox;
        private System.Windows.Forms.ToolStripButton fillToolStripButton;
        private System.Windows.Forms.DataGridView shipping_OQUTDownloadDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private ACMEDataSet.ship2TableAdapters.Shipping_OQUTDownload2TableAdapter shipping_OQUTDownload2TableAdapter;
        private System.Windows.Forms.BindingSource shipping_OQUTDownload2BindingSource;
        private System.Windows.Forms.ToolStrip fillToolStrip1;
        private System.Windows.Forms.ToolStripLabel shippingCodeToolStripLabel1;
        private System.Windows.Forms.ToolStripTextBox shippingCodeToolStripTextBox1;
        private System.Windows.Forms.ToolStripButton fillToolStripButton1;
        private System.Windows.Forms.DataGridView shipping_OQUTDownload2DataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;
    }
}
