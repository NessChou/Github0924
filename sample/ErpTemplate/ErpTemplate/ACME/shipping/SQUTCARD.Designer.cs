namespace ACME
{
    partial class SQUTCARD
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SQUTCARD));
            this.ship = new ACME.ACMEDataSet.ship();
            this.shipping_OQUT4BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.shipping_OQUT4TableAdapter = new ACME.ACMEDataSet.shipTableAdapters.Shipping_OQUT4TableAdapter();
            this.shipping_OQUT4BindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.shipping_OQUT4BindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.shipping_OQUT4DataGridView = new System.Windows.Forms.DataGridView();
            this.CARDCODE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.textCARDNAME = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT4BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT4BindingNavigator)).BeginInit();
            this.shipping_OQUT4BindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT4DataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ship
            // 
            this.ship.DataSetName = "ship";
            this.ship.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // shipping_OQUT4BindingSource
            // 
            this.shipping_OQUT4BindingSource.DataMember = "Shipping_OQUT4";
            this.shipping_OQUT4BindingSource.DataSource = this.ship;
            // 
            // shipping_OQUT4TableAdapter
            // 
            this.shipping_OQUT4TableAdapter.ClearBeforeFill = true;
            // 
            // shipping_OQUT4BindingNavigator
            // 
            this.shipping_OQUT4BindingNavigator.AddNewItem = null;
            this.shipping_OQUT4BindingNavigator.BindingSource = this.shipping_OQUT4BindingSource;
            this.shipping_OQUT4BindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.shipping_OQUT4BindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.shipping_OQUT4BindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.shipping_OQUT4BindingNavigatorSaveItem});
            this.shipping_OQUT4BindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.shipping_OQUT4BindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.shipping_OQUT4BindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.shipping_OQUT4BindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.shipping_OQUT4BindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.shipping_OQUT4BindingNavigator.Name = "shipping_OQUT4BindingNavigator";
            this.shipping_OQUT4BindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.shipping_OQUT4BindingNavigator.Size = new System.Drawing.Size(719, 25);
            this.shipping_OQUT4BindingNavigator.TabIndex = 0;
            this.shipping_OQUT4BindingNavigator.Text = "bindingNavigator1";
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
            // shipping_OQUT4BindingNavigatorSaveItem
            // 
            this.shipping_OQUT4BindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.shipping_OQUT4BindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("shipping_OQUT4BindingNavigatorSaveItem.Image")));
            this.shipping_OQUT4BindingNavigatorSaveItem.Name = "shipping_OQUT4BindingNavigatorSaveItem";
            this.shipping_OQUT4BindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.shipping_OQUT4BindingNavigatorSaveItem.Text = "儲存資料";
            this.shipping_OQUT4BindingNavigatorSaveItem.Click += new System.EventHandler(this.shipping_OQUT4BindingNavigatorSaveItem_Click);
            // 
            // shipping_OQUT4DataGridView
            // 
            this.shipping_OQUT4DataGridView.AllowUserToAddRows = false;
            this.shipping_OQUT4DataGridView.AutoGenerateColumns = false;
            this.shipping_OQUT4DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.shipping_OQUT4DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CARDCODE,
            this.dataGridViewTextBoxColumn3});
            this.shipping_OQUT4DataGridView.DataSource = this.shipping_OQUT4BindingSource;
            this.shipping_OQUT4DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.shipping_OQUT4DataGridView.Location = new System.Drawing.Point(0, 0);
            this.shipping_OQUT4DataGridView.Name = "shipping_OQUT4DataGridView";
            this.shipping_OQUT4DataGridView.RowTemplate.Height = 24;
            this.shipping_OQUT4DataGridView.Size = new System.Drawing.Size(719, 586);
            this.shipping_OQUT4DataGridView.TabIndex = 1;
            this.shipping_OQUT4DataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.shipping_OQUT4DataGridView_DefaultValuesNeeded);
            // 
            // CARDCODE
            // 
            this.CARDCODE.DataPropertyName = "CARDCODE";
            this.CARDCODE.HeaderText = "新供應商編號";
            this.CARDCODE.Name = "CARDCODE";
            this.CARDCODE.ReadOnly = true;
            this.CARDCODE.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "CARDNAME";
            this.dataGridViewTextBoxColumn3.HeaderText = "新供應商名稱";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.dataGridViewTextBoxColumn3.Width = 300;
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(391, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "返回";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // textCARDNAME
            // 
            this.textCARDNAME.Location = new System.Drawing.Point(95, 7);
            this.textCARDNAME.Name = "textCARDNAME";
            this.textCARDNAME.Size = new System.Drawing.Size(209, 22);
            this.textCARDNAME.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "新供應商名稱";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(310, 5);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "新增";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.textCARDNAME);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(719, 39);
            this.panel1.TabIndex = 8;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.shipping_OQUT4DataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 64);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(719, 586);
            this.panel2.TabIndex = 9;
            // 
            // SQUTCARD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(719, 650);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.shipping_OQUT4BindingNavigator);
            this.Name = "SQUTCARD";
            this.Text = "供應商";
            this.Load += new System.EventHandler(this.SQUTCARD_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ship)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT4BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT4BindingNavigator)).EndInit();
            this.shipping_OQUT4BindingNavigator.ResumeLayout(false);
            this.shipping_OQUT4BindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT4DataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship ship;
        private System.Windows.Forms.BindingSource shipping_OQUT4BindingSource;
        private ACMEDataSet.shipTableAdapters.Shipping_OQUT4TableAdapter shipping_OQUT4TableAdapter;
        private System.Windows.Forms.BindingNavigator shipping_OQUT4BindingNavigator;
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
        private System.Windows.Forms.ToolStripButton shipping_OQUT4BindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView shipping_OQUT4DataGridView;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textCARDNAME;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn CARDCODE;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
    }
}