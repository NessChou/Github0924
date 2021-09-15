namespace ACME
{
    partial class GB_OCRD
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GB_OCRD));
            this.POTATO = new ACME.ACMEDataSet.POTATO();
            this.gB_OCRDBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gB_OCRDTableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_OCRDTableAdapter();
            this.gB_OCRDBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorDeleteItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.gB_OCRDBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.gB_OCRDDataGridView = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.gB_OCRD2DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn13 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn14 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn15 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn16 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.QTY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PRICE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AMOUNT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gB_OCRD2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.panel3 = new System.Windows.Forms.Panel();
            this.button6 = new System.Windows.Forms.Button();
            this.gB_OCRD2TableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_OCRD2TableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager();
            this.panel5 = new System.Windows.Forms.Panel();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SCOM = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MEMO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.POTATO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRDBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRDBindingNavigator)).BeginInit();
            this.gB_OCRDBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRDDataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRD2DataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRD2BindingSource)).BeginInit();
            this.panel3.SuspendLayout();
            this.panel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // POTATO
            // 
            this.POTATO.DataSetName = "POTATO";
            this.POTATO.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // gB_OCRDBindingSource
            // 
            this.gB_OCRDBindingSource.DataMember = "GB_OCRD";
            this.gB_OCRDBindingSource.DataSource = this.POTATO;
            // 
            // gB_OCRDTableAdapter
            // 
            this.gB_OCRDTableAdapter.ClearBeforeFill = true;
            // 
            // gB_OCRDBindingNavigator
            // 
            this.gB_OCRDBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.gB_OCRDBindingNavigator.BindingSource = this.gB_OCRDBindingSource;
            this.gB_OCRDBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.gB_OCRDBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.gB_OCRDBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bindingNavigatorMoveFirstItem,
            this.bindingNavigatorMovePreviousItem,
            this.bindingNavigatorSeparator,
            this.bindingNavigatorPositionItem,
            this.bindingNavigatorCountItem,
            this.bindingNavigatorSeparator1,
            this.bindingNavigatorMoveNextItem,
            this.bindingNavigatorMoveLastItem,
            this.bindingNavigatorSeparator2,
            this.bindingNavigatorAddNewItem,
            this.bindingNavigatorDeleteItem,
            this.gB_OCRDBindingNavigatorSaveItem});
            this.gB_OCRDBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.gB_OCRDBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.gB_OCRDBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.gB_OCRDBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.gB_OCRDBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.gB_OCRDBindingNavigator.Name = "gB_OCRDBindingNavigator";
            this.gB_OCRDBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.gB_OCRDBindingNavigator.Size = new System.Drawing.Size(1228, 25);
            this.gB_OCRDBindingNavigator.TabIndex = 0;
            this.gB_OCRDBindingNavigator.Text = "bindingNavigator1";
            // 
            // bindingNavigatorAddNewItem
            // 
            this.bindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorAddNewItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorAddNewItem.Image")));
            this.bindingNavigatorAddNewItem.Name = "bindingNavigatorAddNewItem";
            this.bindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorAddNewItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorAddNewItem.Text = "加入新的";
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(28, 22);
            this.bindingNavigatorCountItem.Text = "/{0}";
            this.bindingNavigatorCountItem.ToolTipText = "項目總數";
            // 
            // bindingNavigatorDeleteItem
            // 
            this.bindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorDeleteItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorDeleteItem.Image")));
            this.bindingNavigatorDeleteItem.Name = "bindingNavigatorDeleteItem";
            this.bindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorDeleteItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorDeleteItem.Text = "刪除";
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
            // gB_OCRDBindingNavigatorSaveItem
            // 
            this.gB_OCRDBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gB_OCRDBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("gB_OCRDBindingNavigatorSaveItem.Image")));
            this.gB_OCRDBindingNavigatorSaveItem.Name = "gB_OCRDBindingNavigatorSaveItem";
            this.gB_OCRDBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.gB_OCRDBindingNavigatorSaveItem.Text = "儲存資料";
            this.gB_OCRDBindingNavigatorSaveItem.Click += new System.EventHandler(this.gB_OCRDBindingNavigatorSaveItem_Click);
            // 
            // gB_OCRDDataGridView
            // 
            this.gB_OCRDDataGridView.AutoGenerateColumns = false;
            this.gB_OCRDDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gB_OCRDDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7,
            this.dataGridViewTextBoxColumn8,
            this.SCOM,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.MEMO});
            this.gB_OCRDDataGridView.DataSource = this.gB_OCRDBindingSource;
            this.gB_OCRDDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gB_OCRDDataGridView.Location = new System.Drawing.Point(0, 0);
            this.gB_OCRDDataGridView.Name = "gB_OCRDDataGridView";
            this.gB_OCRDDataGridView.RowTemplate.Height = 24;
            this.gB_OCRDDataGridView.Size = new System.Drawing.Size(1228, 472);
            this.gB_OCRDDataGridView.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 497);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1228, 267);
            this.panel1.TabIndex = 2;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1228, 267);
            this.panel2.TabIndex = 3;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.gB_OCRD2DataGridView);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 31);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1228, 236);
            this.panel4.TabIndex = 1;
            // 
            // gB_OCRD2DataGridView
            // 
            this.gB_OCRD2DataGridView.AutoGenerateColumns = false;
            this.gB_OCRD2DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gB_OCRD2DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn13,
            this.dataGridViewTextBoxColumn14,
            this.dataGridViewTextBoxColumn15,
            this.dataGridViewTextBoxColumn16,
            this.QTY,
            this.PRICE,
            this.AMOUNT});
            this.gB_OCRD2DataGridView.DataSource = this.gB_OCRD2BindingSource;
            this.gB_OCRD2DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gB_OCRD2DataGridView.Location = new System.Drawing.Point(0, 0);
            this.gB_OCRD2DataGridView.Name = "gB_OCRD2DataGridView";
            this.gB_OCRD2DataGridView.RowTemplate.Height = 24;
            this.gB_OCRD2DataGridView.Size = new System.Drawing.Size(1228, 236);
            this.gB_OCRD2DataGridView.TabIndex = 0;
            this.gB_OCRD2DataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gB_OCRD2DataGridView_CellValueChanged);
            // 
            // dataGridViewTextBoxColumn13
            // 
            this.dataGridViewTextBoxColumn13.DataPropertyName = "ITEMCODE";
            this.dataGridViewTextBoxColumn13.HeaderText = "產品編號";
            this.dataGridViewTextBoxColumn13.Name = "dataGridViewTextBoxColumn13";
            this.dataGridViewTextBoxColumn13.Width = 80;
            // 
            // dataGridViewTextBoxColumn14
            // 
            this.dataGridViewTextBoxColumn14.DataPropertyName = "ITEMNAME";
            this.dataGridViewTextBoxColumn14.HeaderText = "產品名稱";
            this.dataGridViewTextBoxColumn14.Name = "dataGridViewTextBoxColumn14";
            this.dataGridViewTextBoxColumn14.Width = 140;
            // 
            // dataGridViewTextBoxColumn15
            // 
            this.dataGridViewTextBoxColumn15.DataPropertyName = "STARTDATE";
            this.dataGridViewTextBoxColumn15.HeaderText = "開始日期";
            this.dataGridViewTextBoxColumn15.MaxInputLength = 8;
            this.dataGridViewTextBoxColumn15.Name = "dataGridViewTextBoxColumn15";
            this.dataGridViewTextBoxColumn15.Width = 80;
            // 
            // dataGridViewTextBoxColumn16
            // 
            this.dataGridViewTextBoxColumn16.DataPropertyName = "ENDDATE";
            this.dataGridViewTextBoxColumn16.HeaderText = "截止日期";
            this.dataGridViewTextBoxColumn16.MaxInputLength = 8;
            this.dataGridViewTextBoxColumn16.Name = "dataGridViewTextBoxColumn16";
            this.dataGridViewTextBoxColumn16.Width = 80;
            // 
            // QTY
            // 
            this.QTY.DataPropertyName = "QTY";
            this.QTY.HeaderText = "數量";
            this.QTY.Name = "QTY";
            this.QTY.Width = 60;
            // 
            // PRICE
            // 
            this.PRICE.DataPropertyName = "PRICE";
            this.PRICE.HeaderText = "單價";
            this.PRICE.Name = "PRICE";
            this.PRICE.Width = 60;
            // 
            // AMOUNT
            // 
            this.AMOUNT.DataPropertyName = "AMOUNT";
            this.AMOUNT.HeaderText = "金額";
            this.AMOUNT.Name = "AMOUNT";
            this.AMOUNT.Width = 60;
            // 
            // gB_OCRD2BindingSource
            // 
            this.gB_OCRD2BindingSource.DataMember = "GB_OCRD_GB_OCRD2";
            this.gB_OCRD2BindingSource.DataSource = this.gB_OCRDBindingSource;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.button6);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1228, 31);
            this.panel3.TabIndex = 0;
            // 
            // button6
            // 
            this.button6.BackgroundImage = global::ACME.Properties.Resources.tw12_sp1b;
            this.button6.ForeColor = System.Drawing.Color.White;
            this.button6.Location = new System.Drawing.Point(40, 5);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(150, 23);
            this.button6.TabIndex = 16;
            this.button6.Text = "選擇項目大宗(箱為單位)";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // gB_OCRD2TableAdapter
            // 
            this.gB_OCRD2TableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.GB_DATELOCKTableAdapter = null;
            this.tableAdapterManager.GB_FRIEND1TableAdapter = null;
            this.tableAdapterManager.GB_FRIENDTableAdapter = null;
            this.tableAdapterManager.GB_INVTRACKTableAdapter = null;
            this.tableAdapterManager.GB_OCRD2TableAdapter = this.gB_OCRD2TableAdapter;
            this.tableAdapterManager.GB_OCRDTableAdapter = this.gB_OCRDTableAdapter;
            this.tableAdapterManager.GB_POTATO1TableAdapter = null;
            this.tableAdapterManager.GB_POTATO21TableAdapter = null;
            this.tableAdapterManager.GB_POTATO2TableAdapter = null;
            this.tableAdapterManager.GB_POTATOTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.gB_OCRDDataGridView);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(0, 25);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(1228, 472);
            this.panel5.TabIndex = 3;
            // 
            // ID
            // 
            this.ID.DataPropertyName = "ID";
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Width = 40;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "ORDNAME";
            this.dataGridViewTextBoxColumn2.HeaderText = "訂購人";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 80;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "ORDCOM";
            this.dataGridViewTextBoxColumn3.HeaderText = "訂購人公司";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "ORDTEL";
            this.dataGridViewTextBoxColumn4.HeaderText = "訂購人電話";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.Width = 90;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "ORDEMAIL";
            this.dataGridViewTextBoxColumn5.HeaderText = "訂購人EMAIL";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.Width = 120;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "PAYMAN";
            this.dataGridViewTextBoxColumn6.HeaderText = "付款人";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.Width = 80;
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "UNIT";
            this.dataGridViewTextBoxColumn7.HeaderText = "統一編號";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.Width = 80;
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.DataPropertyName = "SPERSON";
            this.dataGridViewTextBoxColumn8.HeaderText = "收貨人";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            this.dataGridViewTextBoxColumn8.Width = 80;
            // 
            // SCOM
            // 
            this.SCOM.DataPropertyName = "SCOM";
            this.SCOM.HeaderText = "收貨人公司";
            this.SCOM.Name = "SCOM";
            this.SCOM.Width = 90;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "SADDRESS";
            this.dataGridViewTextBoxColumn9.HeaderText = "收貨人地址";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.Width = 200;
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "STEL";
            this.dataGridViewTextBoxColumn10.HeaderText = "收貨人電話";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            this.dataGridViewTextBoxColumn10.Width = 90;
            // 
            // MEMO
            // 
            this.MEMO.DataPropertyName = "MEMO";
            this.MEMO.HeaderText = "備註";
            this.MEMO.Name = "MEMO";
            this.MEMO.Width = 200;
            // 
            // GB_OCRD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1228, 764);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.gB_OCRDBindingNavigator);
            this.Name = "GB_OCRD";
            this.Text = "客戶檔維護";
            this.Load += new System.EventHandler(this.GB_OCRD_Load);
            ((System.ComponentModel.ISupportInitialize)(this.POTATO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRDBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRDBindingNavigator)).EndInit();
            this.gB_OCRDBindingNavigator.ResumeLayout(false);
            this.gB_OCRDBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRDDataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRD2DataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_OCRD2BindingSource)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.POTATO POTATO;
        private System.Windows.Forms.BindingSource gB_OCRDBindingSource;
        private ACMEDataSet.POTATOTableAdapters.GB_OCRDTableAdapter gB_OCRDTableAdapter;
        private System.Windows.Forms.BindingNavigator gB_OCRDBindingNavigator;
        private System.Windows.Forms.ToolStripButton bindingNavigatorAddNewItem;
        private System.Windows.Forms.ToolStripLabel bindingNavigatorCountItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorDeleteItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveFirstItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMovePreviousItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator;
        private System.Windows.Forms.ToolStripTextBox bindingNavigatorPositionItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator1;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveNextItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveLastItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator2;
        private System.Windows.Forms.ToolStripButton gB_OCRDBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView gB_OCRDDataGridView;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.BindingSource gB_OCRD2BindingSource;
        private ACMEDataSet.POTATOTableAdapters.GB_OCRD2TableAdapter gB_OCRD2TableAdapter;
        private ACMEDataSet.POTATOTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.DataGridView gB_OCRD2DataGridView;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn13;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn14;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn15;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn16;
        private System.Windows.Forms.DataGridViewTextBoxColumn QTY;
        private System.Windows.Forms.DataGridViewTextBoxColumn PRICE;
        private System.Windows.Forms.DataGridViewTextBoxColumn AMOUNT;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn SCOM;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn MEMO;
    }
}