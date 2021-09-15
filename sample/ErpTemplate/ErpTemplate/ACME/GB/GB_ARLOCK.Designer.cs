namespace ACME
{
    partial class GB_ARLOCK
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
            System.Windows.Forms.Label sTARTDATELabel;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GB_ARLOCK));
            this.pOTATO = new ACME.ACMEDataSet.POTATO();
            this.gB_DATELOCKBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gB_DATELOCKTableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_DATELOCKTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager();
            this.gB_DATELOCKBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.gB_DATELOCKBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.eNDDATETextBox = new System.Windows.Forms.TextBox();
            this.dOCDATETextBox = new System.Windows.Forms.TextBox();
            this.lOGUSERTextBox = new System.Windows.Forms.TextBox();
            sTARTDATELabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_DATELOCKBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_DATELOCKBindingNavigator)).BeginInit();
            this.gB_DATELOCKBindingNavigator.SuspendLayout();
            this.SuspendLayout();
            // 
            // sTARTDATELabel
            // 
            sTARTDATELabel.AutoSize = true;
            sTARTDATELabel.Location = new System.Drawing.Point(31, 35);
            sTARTDATELabel.Name = "sTARTDATELabel";
            sTARTDATELabel.Size = new System.Drawing.Size(53, 12);
            sTARTDATELabel.TabIndex = 1;
            sTARTDATELabel.Text = "鎖單日期";
            // 
            // pOTATO
            // 
            this.pOTATO.DataSetName = "POTATO";
            this.pOTATO.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // gB_DATELOCKBindingSource
            // 
            this.gB_DATELOCKBindingSource.DataMember = "GB_DATELOCK";
            this.gB_DATELOCKBindingSource.DataSource = this.pOTATO;
            // 
            // gB_DATELOCKTableAdapter
            // 
            this.gB_DATELOCKTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.GB_DATELOCKTableAdapter = this.gB_DATELOCKTableAdapter;
            this.tableAdapterManager.GB_FRIEND1TableAdapter = null;
            this.tableAdapterManager.GB_FRIENDTableAdapter = null;
            this.tableAdapterManager.GB_INVTRACKTableAdapter = null;
            this.tableAdapterManager.GB_OCRDTableAdapter = null;
            this.tableAdapterManager.GB_POTATO1TableAdapter = null;
            this.tableAdapterManager.GB_POTATO21TableAdapter = null;
            this.tableAdapterManager.GB_POTATO2TableAdapter = null;
            this.tableAdapterManager.GB_POTATOTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // gB_DATELOCKBindingNavigator
            // 
            this.gB_DATELOCKBindingNavigator.AddNewItem = null;
            this.gB_DATELOCKBindingNavigator.BindingSource = this.gB_DATELOCKBindingSource;
            this.gB_DATELOCKBindingNavigator.CountItem = null;
            this.gB_DATELOCKBindingNavigator.DeleteItem = null;
            this.gB_DATELOCKBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.gB_DATELOCKBindingNavigatorSaveItem});
            this.gB_DATELOCKBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.gB_DATELOCKBindingNavigator.MoveFirstItem = null;
            this.gB_DATELOCKBindingNavigator.MoveLastItem = null;
            this.gB_DATELOCKBindingNavigator.MoveNextItem = null;
            this.gB_DATELOCKBindingNavigator.MovePreviousItem = null;
            this.gB_DATELOCKBindingNavigator.Name = "gB_DATELOCKBindingNavigator";
            this.gB_DATELOCKBindingNavigator.PositionItem = null;
            this.gB_DATELOCKBindingNavigator.Size = new System.Drawing.Size(364, 25);
            this.gB_DATELOCKBindingNavigator.TabIndex = 0;
            this.gB_DATELOCKBindingNavigator.Text = "bindingNavigator1";
            // 
            // gB_DATELOCKBindingNavigatorSaveItem
            // 
            this.gB_DATELOCKBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gB_DATELOCKBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("gB_DATELOCKBindingNavigatorSaveItem.Image")));
            this.gB_DATELOCKBindingNavigatorSaveItem.Name = "gB_DATELOCKBindingNavigatorSaveItem";
            this.gB_DATELOCKBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.gB_DATELOCKBindingNavigatorSaveItem.Text = "儲存資料";
            this.gB_DATELOCKBindingNavigatorSaveItem.Click += new System.EventHandler(this.gB_DATELOCKBindingNavigatorSaveItem_Click);
            // 
            // eNDDATETextBox
            // 
            this.eNDDATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gB_DATELOCKBindingSource, "ENDDATE", true));
            this.eNDDATETextBox.Location = new System.Drawing.Point(99, 32);
            this.eNDDATETextBox.MaxLength = 8;
            this.eNDDATETextBox.Name = "eNDDATETextBox";
            this.eNDDATETextBox.Size = new System.Drawing.Size(100, 22);
            this.eNDDATETextBox.TabIndex = 4;
            // 
            // dOCDATETextBox
            // 
            this.dOCDATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gB_DATELOCKBindingSource, "DOCDATE", true));
            this.dOCDATETextBox.Location = new System.Drawing.Point(99, 71);
            this.dOCDATETextBox.Name = "dOCDATETextBox";
            this.dOCDATETextBox.Size = new System.Drawing.Size(0, 22);
            this.dOCDATETextBox.TabIndex = 6;
            // 
            // lOGUSERTextBox
            // 
            this.lOGUSERTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gB_DATELOCKBindingSource, "LOGUSER", true));
            this.lOGUSERTextBox.Location = new System.Drawing.Point(252, 71);
            this.lOGUSERTextBox.Name = "lOGUSERTextBox";
            this.lOGUSERTextBox.Size = new System.Drawing.Size(0, 22);
            this.lOGUSERTextBox.TabIndex = 8;
            // 
            // GB_ARLOCK
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(364, 75);
            this.Controls.Add(this.lOGUSERTextBox);
            this.Controls.Add(this.dOCDATETextBox);
            this.Controls.Add(this.eNDDATETextBox);
            this.Controls.Add(sTARTDATELabel);
            this.Controls.Add(this.gB_DATELOCKBindingNavigator);
            this.Name = "GB_ARLOCK";
            this.Text = "鎖單日期";
            this.Load += new System.EventHandler(this.GB_ARLOCK_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_DATELOCKBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_DATELOCKBindingNavigator)).EndInit();
            this.gB_DATELOCKBindingNavigator.ResumeLayout(false);
            this.gB_DATELOCKBindingNavigator.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.POTATO pOTATO;
        private System.Windows.Forms.BindingSource gB_DATELOCKBindingSource;
        private ACMEDataSet.POTATOTableAdapters.GB_DATELOCKTableAdapter gB_DATELOCKTableAdapter;
        private ACMEDataSet.POTATOTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.BindingNavigator gB_DATELOCKBindingNavigator;
        private System.Windows.Forms.ToolStripButton gB_DATELOCKBindingNavigatorSaveItem;
        private System.Windows.Forms.TextBox eNDDATETextBox;
        private System.Windows.Forms.TextBox dOCDATETextBox;
        private System.Windows.Forms.TextBox lOGUSERTextBox;
    }
}