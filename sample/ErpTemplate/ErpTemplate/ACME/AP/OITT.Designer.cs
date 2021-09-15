namespace ACME
{
    partial class OITT
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
            System.Windows.Forms.Label codeLabel;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OITT));
            this.sAP = new ACME.ACMEDataSet.SAP();
            this.oITTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.oITTTableAdapter = new ACME.ACMEDataSet.SAPTableAdapters.OITTTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.SAPTableAdapters.TableAdapterManager();
            this.iTT1TableAdapter = new ACME.ACMEDataSet.SAPTableAdapters.ITT1TableAdapter();
            this.oITTBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.oITTBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.codeTextBox = new System.Windows.Forms.TextBox();
            this.iTT1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.iTT1DataGridView = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ChildNum = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.產品編號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Warehouse = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Price = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Currency = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PriceList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OrigPrice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OrigCurr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IssueMthd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Object = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PrncpInput = new System.Windows.Forms.DataGridViewTextBoxColumn();
            codeLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.sAP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.oITTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.oITTBindingNavigator)).BeginInit();
            this.oITTBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iTT1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iTT1DataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // codeLabel
            // 
            codeLabel.AutoSize = true;
            codeLabel.Location = new System.Drawing.Point(12, 24);
            codeLabel.Name = "codeLabel";
            codeLabel.Size = new System.Drawing.Size(53, 12);
            codeLabel.TabIndex = 1;
            codeLabel.Text = "產品號碼";
            // 
            // sAP
            // 
            this.sAP.DataSetName = "SAP";
            this.sAP.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // oITTBindingSource
            // 
            this.oITTBindingSource.DataMember = "OITT";
            this.oITTBindingSource.DataSource = this.sAP;
            // 
            // oITTTableAdapter
            // 
            this.oITTTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.ITT1TableAdapter = this.iTT1TableAdapter;
            this.tableAdapterManager.OITTTableAdapter = this.oITTTableAdapter;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.SAPTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // iTT1TableAdapter
            // 
            this.iTT1TableAdapter.ClearBeforeFill = true;
            // 
            // oITTBindingNavigator
            // 
            this.oITTBindingNavigator.AddNewItem = null;
            this.oITTBindingNavigator.BindingSource = this.oITTBindingSource;
            this.oITTBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.oITTBindingNavigator.DeleteItem = null;
            this.oITTBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bindingNavigatorMoveFirstItem,
            this.bindingNavigatorMovePreviousItem,
            this.bindingNavigatorSeparator,
            this.bindingNavigatorPositionItem,
            this.bindingNavigatorCountItem,
            this.bindingNavigatorSeparator1,
            this.bindingNavigatorMoveNextItem,
            this.bindingNavigatorMoveLastItem,
            this.bindingNavigatorSeparator2,
            this.oITTBindingNavigatorSaveItem});
            this.oITTBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.oITTBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.oITTBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.oITTBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.oITTBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.oITTBindingNavigator.Name = "oITTBindingNavigator";
            this.oITTBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.oITTBindingNavigator.Size = new System.Drawing.Size(824, 25);
            this.oITTBindingNavigator.TabIndex = 0;
            this.oITTBindingNavigator.Text = "bindingNavigator1";
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(27, 22);
            this.bindingNavigatorCountItem.Text = "/{0}";
            this.bindingNavigatorCountItem.ToolTipText = "項目總數";
            // 
            // bindingNavigatorMoveFirstItem
            // 
            this.bindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveFirstItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveFirstItem.Image")));
            this.bindingNavigatorMoveFirstItem.Name = "bindingNavigatorMoveFirstItem";
            this.bindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveFirstItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveFirstItem.Text = "移到最前面";
            // 
            // bindingNavigatorMovePreviousItem
            // 
            this.bindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMovePreviousItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMovePreviousItem.Image")));
            this.bindingNavigatorMovePreviousItem.Name = "bindingNavigatorMovePreviousItem";
            this.bindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMovePreviousItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMovePreviousItem.Text = "移到上一個";
            // 
            // bindingNavigatorSeparator
            // 
            this.bindingNavigatorSeparator.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorPositionItem
            // 
            this.bindingNavigatorPositionItem.AccessibleName = "位置";
            this.bindingNavigatorPositionItem.AutoSize = false;
            this.bindingNavigatorPositionItem.Name = "bindingNavigatorPositionItem";
            this.bindingNavigatorPositionItem.Size = new System.Drawing.Size(50, 23);
            this.bindingNavigatorPositionItem.Text = "0";
            this.bindingNavigatorPositionItem.ToolTipText = "目前的位置";
            // 
            // bindingNavigatorSeparator1
            // 
            this.bindingNavigatorSeparator1.Name = "bindingNavigatorSeparator1";
            this.bindingNavigatorSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorMoveNextItem
            // 
            this.bindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveNextItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveNextItem.Image")));
            this.bindingNavigatorMoveNextItem.Name = "bindingNavigatorMoveNextItem";
            this.bindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveNextItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveNextItem.Text = "移到下一個";
            // 
            // bindingNavigatorMoveLastItem
            // 
            this.bindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveLastItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveLastItem.Image")));
            this.bindingNavigatorMoveLastItem.Name = "bindingNavigatorMoveLastItem";
            this.bindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveLastItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveLastItem.Text = "移到最後面";
            // 
            // bindingNavigatorSeparator2
            // 
            this.bindingNavigatorSeparator2.Name = "bindingNavigatorSeparator2";
            this.bindingNavigatorSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // oITTBindingNavigatorSaveItem
            // 
            this.oITTBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.oITTBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("oITTBindingNavigatorSaveItem.Image")));
            this.oITTBindingNavigatorSaveItem.Name = "oITTBindingNavigatorSaveItem";
            this.oITTBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.oITTBindingNavigatorSaveItem.Text = "儲存資料";
            this.oITTBindingNavigatorSaveItem.Click += new System.EventHandler(this.oITTBindingNavigatorSaveItem_Click);
            // 
            // codeTextBox
            // 
            this.codeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.oITTBindingSource, "Code", true));
            this.codeTextBox.Location = new System.Drawing.Point(80, 21);
            this.codeTextBox.Name = "codeTextBox";
            this.codeTextBox.ReadOnly = true;
            this.codeTextBox.Size = new System.Drawing.Size(444, 22);
            this.codeTextBox.TabIndex = 2;
            // 
            // iTT1BindingSource
            // 
            this.iTT1BindingSource.DataMember = "OITT_ITT1";
            this.iTT1BindingSource.DataSource = this.oITTBindingSource;
            // 
            // iTT1DataGridView
            // 
            this.iTT1DataGridView.AutoGenerateColumns = false;
            this.iTT1DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.iTT1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.iTT1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.ChildNum,
            this.產品編號,
            this.Quantity,
            this.Warehouse,
            this.Price,
            this.Currency,
            this.PriceList,
            this.OrigPrice,
            this.OrigCurr,
            this.IssueMthd,
            this.Object,
            this.PrncpInput});
            this.iTT1DataGridView.DataSource = this.iTT1BindingSource;
            this.iTT1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.iTT1DataGridView.Location = new System.Drawing.Point(0, 0);
            this.iTT1DataGridView.Name = "iTT1DataGridView";
            this.iTT1DataGridView.RowTemplate.Height = 24;
            this.iTT1DataGridView.Size = new System.Drawing.Size(824, 514);
            this.iTT1DataGridView.TabIndex = 3;
            this.iTT1DataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.iTT1DataGridView_DefaultValuesNeeded);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(codeLabel);
            this.panel1.Controls.Add(this.codeTextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(824, 54);
            this.panel1.TabIndex = 4;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.iTT1DataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 79);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(824, 514);
            this.panel2.TabIndex = 5;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Father";
            this.dataGridViewTextBoxColumn1.HeaderText = "Father";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.Visible = false;
            this.dataGridViewTextBoxColumn1.Width = 59;
            // 
            // ChildNum
            // 
            this.ChildNum.DataPropertyName = "ChildNum";
            this.ChildNum.HeaderText = "Line";
            this.ChildNum.Name = "ChildNum";
            this.ChildNum.ReadOnly = true;
            this.ChildNum.Width = 51;
            // 
            // 產品編號
            // 
            this.產品編號.DataPropertyName = "Code";
            this.產品編號.HeaderText = "產品編號";
            this.產品編號.Name = "產品編號";
            this.產品編號.Width = 78;
            // 
            // Quantity
            // 
            this.Quantity.DataPropertyName = "Quantity";
            this.Quantity.HeaderText = "數量";
            this.Quantity.Name = "Quantity";
            this.Quantity.Width = 54;
            // 
            // Warehouse
            // 
            this.Warehouse.DataPropertyName = "Warehouse";
            this.Warehouse.HeaderText = "Warehouse";
            this.Warehouse.Name = "Warehouse";
            this.Warehouse.Visible = false;
            this.Warehouse.Width = 82;
            // 
            // Price
            // 
            this.Price.DataPropertyName = "Price";
            this.Price.HeaderText = "單價";
            this.Price.Name = "Price";
            this.Price.Width = 54;
            // 
            // Currency
            // 
            this.Currency.DataPropertyName = "Currency";
            this.Currency.HeaderText = "Currency";
            this.Currency.Name = "Currency";
            this.Currency.Visible = false;
            this.Currency.Width = 74;
            // 
            // PriceList
            // 
            this.PriceList.DataPropertyName = "PriceList";
            this.PriceList.HeaderText = "PriceList";
            this.PriceList.Name = "PriceList";
            this.PriceList.Visible = false;
            this.PriceList.Width = 70;
            // 
            // OrigPrice
            // 
            this.OrigPrice.DataPropertyName = "OrigPrice";
            this.OrigPrice.HeaderText = "OrigPrice";
            this.OrigPrice.Name = "OrigPrice";
            this.OrigPrice.Visible = false;
            this.OrigPrice.Width = 74;
            // 
            // OrigCurr
            // 
            this.OrigCurr.DataPropertyName = "OrigCurr";
            this.OrigCurr.HeaderText = "OrigCurr";
            this.OrigCurr.Name = "OrigCurr";
            this.OrigCurr.Visible = false;
            this.OrigCurr.Width = 73;
            // 
            // IssueMthd
            // 
            this.IssueMthd.DataPropertyName = "IssueMthd";
            this.IssueMthd.HeaderText = "IssueMthd";
            this.IssueMthd.Name = "IssueMthd";
            this.IssueMthd.Visible = false;
            this.IssueMthd.Width = 78;
            // 
            // Object
            // 
            this.Object.DataPropertyName = "Object";
            this.Object.HeaderText = "Object";
            this.Object.Name = "Object";
            this.Object.Visible = false;
            this.Object.Width = 60;
            // 
            // PrncpInput
            // 
            this.PrncpInput.DataPropertyName = "PrncpInput";
            this.PrncpInput.HeaderText = "PrncpInput";
            this.PrncpInput.Name = "PrncpInput";
            this.PrncpInput.Visible = false;
            this.PrncpInput.Width = 82;
            // 
            // OITT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(824, 593);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.oITTBindingNavigator);
            this.Name = "OITT";
            this.Text = "SAP物料表";
            this.Load += new System.EventHandler(this.OITT_Load);
            ((System.ComponentModel.ISupportInitialize)(this.sAP)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.oITTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.oITTBindingNavigator)).EndInit();
            this.oITTBindingNavigator.ResumeLayout(false);
            this.oITTBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iTT1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iTT1DataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.SAP sAP;
        private System.Windows.Forms.BindingSource oITTBindingSource;
        private ACMEDataSet.SAPTableAdapters.OITTTableAdapter oITTTableAdapter;
        private ACMEDataSet.SAPTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.BindingNavigator oITTBindingNavigator;
        private System.Windows.Forms.ToolStripLabel bindingNavigatorCountItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveFirstItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMovePreviousItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator;
        private System.Windows.Forms.ToolStripTextBox bindingNavigatorPositionItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator1;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveNextItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveLastItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator2;
        private System.Windows.Forms.ToolStripButton oITTBindingNavigatorSaveItem;
        private ACMEDataSet.SAPTableAdapters.ITT1TableAdapter iTT1TableAdapter;
        private System.Windows.Forms.TextBox codeTextBox;
        private System.Windows.Forms.BindingSource iTT1BindingSource;
        private System.Windows.Forms.DataGridView iTT1DataGridView;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ChildNum;
        private System.Windows.Forms.DataGridViewTextBoxColumn 產品編號;
        private System.Windows.Forms.DataGridViewTextBoxColumn Quantity;
        private System.Windows.Forms.DataGridViewTextBoxColumn Warehouse;
        private System.Windows.Forms.DataGridViewTextBoxColumn Price;
        private System.Windows.Forms.DataGridViewTextBoxColumn Currency;
        private System.Windows.Forms.DataGridViewTextBoxColumn PriceList;
        private System.Windows.Forms.DataGridViewTextBoxColumn OrigPrice;
        private System.Windows.Forms.DataGridViewTextBoxColumn OrigCurr;
        private System.Windows.Forms.DataGridViewTextBoxColumn IssueMthd;
        private System.Windows.Forms.DataGridViewTextBoxColumn Object;
        private System.Windows.Forms.DataGridViewTextBoxColumn PrncpInput;
    }
}