namespace ACME
{
    partial class SQUTOITM
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SQUTOITM));
            this.ship = new ACME.ACMEDataSet.ship();
            this.shipping_OQUT5BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.shipping_OQUT5TableAdapter = new ACME.ACMEDataSet.shipTableAdapters.Shipping_OQUT5TableAdapter();
            this.shipping_OQUT5BindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.shipping_OQUT5BindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.shipping_OQUT5DataGridView = new System.Windows.Forms.DataGridView();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textITEMNAME = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.ITEMCODE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT5BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT5BindingNavigator)).BeginInit();
            this.shipping_OQUT5BindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT5DataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ship
            // 
            this.ship.DataSetName = "ship";
            this.ship.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // shipping_OQUT5BindingSource
            // 
            this.shipping_OQUT5BindingSource.DataMember = "Shipping_OQUT5";
            this.shipping_OQUT5BindingSource.DataSource = this.ship;
            // 
            // shipping_OQUT5TableAdapter
            // 
            this.shipping_OQUT5TableAdapter.ClearBeforeFill = true;
            // 
            // shipping_OQUT5BindingNavigator
            // 
            this.shipping_OQUT5BindingNavigator.AddNewItem = null;
            this.shipping_OQUT5BindingNavigator.BindingSource = this.shipping_OQUT5BindingSource;
            this.shipping_OQUT5BindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.shipping_OQUT5BindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.shipping_OQUT5BindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bindingNavigatorMoveFirstItem,
            this.bindingNavigatorMovePreviousItem,
            this.bindingNavigatorSeparator,
            this.bindingNavigatorPositionItem,
            this.bindingNavigatorCountItem,
            this.bindingNavigatorSeparator1,
            this.bindingNavigatorMoveNextItem,
            this.bindingNavigatorMoveLastItem,
            this.bindingNavigatorSeparator2,
            this.bindingNavigatorDeleteItem,
            this.shipping_OQUT5BindingNavigatorSaveItem});
            this.shipping_OQUT5BindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.shipping_OQUT5BindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.shipping_OQUT5BindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.shipping_OQUT5BindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.shipping_OQUT5BindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.shipping_OQUT5BindingNavigator.Name = "shipping_OQUT5BindingNavigator";
            this.shipping_OQUT5BindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.shipping_OQUT5BindingNavigator.Size = new System.Drawing.Size(709, 25);
            this.shipping_OQUT5BindingNavigator.TabIndex = 0;
            this.shipping_OQUT5BindingNavigator.Text = "bindingNavigator1";
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
            // shipping_OQUT5BindingNavigatorSaveItem
            // 
            this.shipping_OQUT5BindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.shipping_OQUT5BindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("shipping_OQUT5BindingNavigatorSaveItem.Image")));
            this.shipping_OQUT5BindingNavigatorSaveItem.Name = "shipping_OQUT5BindingNavigatorSaveItem";
            this.shipping_OQUT5BindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.shipping_OQUT5BindingNavigatorSaveItem.Text = "儲存資料";
            this.shipping_OQUT5BindingNavigatorSaveItem.Click += new System.EventHandler(this.shipping_OQUT5BindingNavigatorSaveItem_Click);
            // 
            // shipping_OQUT5DataGridView
            // 
            this.shipping_OQUT5DataGridView.AllowUserToAddRows = false;
            this.shipping_OQUT5DataGridView.AutoGenerateColumns = false;
            this.shipping_OQUT5DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.shipping_OQUT5DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ITEMCODE,
            this.dataGridViewTextBoxColumn3});
            this.shipping_OQUT5DataGridView.DataSource = this.shipping_OQUT5BindingSource;
            this.shipping_OQUT5DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.shipping_OQUT5DataGridView.Location = new System.Drawing.Point(0, 0);
            this.shipping_OQUT5DataGridView.Name = "shipping_OQUT5DataGridView";
            this.shipping_OQUT5DataGridView.RowTemplate.Height = 24;
            this.shipping_OQUT5DataGridView.Size = new System.Drawing.Size(709, 567);
            this.shipping_OQUT5DataGridView.TabIndex = 1;
            this.shipping_OQUT5DataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.shipping_OQUT5DataGridView_DefaultValuesNeeded);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(298, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 13;
            this.button2.Text = "新增";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(379, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "返回";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 12;
            this.label2.Text = "新項目名稱";
            // 
            // textITEMNAME
            // 
            this.textITEMNAME.Location = new System.Drawing.Point(83, 6);
            this.textITEMNAME.Name = "textITEMNAME";
            this.textITEMNAME.Size = new System.Drawing.Size(209, 22);
            this.textITEMNAME.TabIndex = 10;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.textITEMNAME);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(709, 33);
            this.panel1.TabIndex = 14;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.shipping_OQUT5DataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 58);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(709, 567);
            this.panel2.TabIndex = 15;
            // 
            // ITEMCODE
            // 
            this.ITEMCODE.DataPropertyName = "ITEMCODE";
            this.ITEMCODE.HeaderText = "項目編號";
            this.ITEMCODE.Name = "ITEMCODE";
            this.ITEMCODE.ReadOnly = true;
            this.ITEMCODE.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "ITEMNAME";
            this.dataGridViewTextBoxColumn3.HeaderText = "項目名稱";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.dataGridViewTextBoxColumn3.Width = 300;
            // 
            // SQUTOITM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(709, 625);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.shipping_OQUT5BindingNavigator);
            this.Name = "SQUTOITM";
            this.Text = "新項目編號";
            this.Load += new System.EventHandler(this.SQUTOITM_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ship)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT5BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT5BindingNavigator)).EndInit();
            this.shipping_OQUT5BindingNavigator.ResumeLayout(false);
            this.shipping_OQUT5BindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT5DataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship ship;
        private System.Windows.Forms.BindingSource shipping_OQUT5BindingSource;
        private ACMEDataSet.shipTableAdapters.Shipping_OQUT5TableAdapter shipping_OQUT5TableAdapter;
        private System.Windows.Forms.BindingNavigator shipping_OQUT5BindingNavigator;
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
        private System.Windows.Forms.ToolStripButton shipping_OQUT5BindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView shipping_OQUT5DataGridView;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textITEMNAME;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ITEMCODE;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
    }
}