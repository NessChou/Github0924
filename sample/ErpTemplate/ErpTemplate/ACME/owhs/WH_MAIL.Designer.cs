namespace ACME
{
    partial class WH_MAIL
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WH_MAIL));
            System.Windows.Forms.Label cEMAILLabel;
            System.Windows.Forms.Label sEMAILLabel;
            this.wh = new ACME.ACMEDataSet.wh();
            this.wH_MAILBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.wH_MAILTableAdapter = new ACME.ACMEDataSet.whTableAdapters.WH_MAILTableAdapter();
            this.wH_MAILBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.wH_MAILBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.wH_MAILDataGridView = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cEMAILTextBox = new System.Windows.Forms.TextBox();
            this.sEMAILTextBox = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            cEMAILLabel = new System.Windows.Forms.Label();
            sEMAILLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.wh)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_MAILBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_MAILBindingNavigator)).BeginInit();
            this.wH_MAILBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.wH_MAILDataGridView)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // wh
            // 
            this.wh.DataSetName = "wh";
            this.wh.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // wH_MAILBindingSource
            // 
            this.wH_MAILBindingSource.DataMember = "WH_MAIL";
            this.wH_MAILBindingSource.DataSource = this.wh;
            // 
            // wH_MAILTableAdapter
            // 
            this.wH_MAILTableAdapter.ClearBeforeFill = true;
            // 
            // wH_MAILBindingNavigator
            // 
            this.wH_MAILBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.wH_MAILBindingNavigator.BindingSource = this.wH_MAILBindingSource;
            this.wH_MAILBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.wH_MAILBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.wH_MAILBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.wH_MAILBindingNavigatorSaveItem});
            this.wH_MAILBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.wH_MAILBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.wH_MAILBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.wH_MAILBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.wH_MAILBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.wH_MAILBindingNavigator.Name = "wH_MAILBindingNavigator";
            this.wH_MAILBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.wH_MAILBindingNavigator.Size = new System.Drawing.Size(880, 25);
            this.wH_MAILBindingNavigator.TabIndex = 0;
            this.wH_MAILBindingNavigator.Text = "bindingNavigator1";
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
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(27, 22);
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
            // wH_MAILBindingNavigatorSaveItem
            // 
            this.wH_MAILBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.wH_MAILBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("wH_MAILBindingNavigatorSaveItem.Image")));
            this.wH_MAILBindingNavigatorSaveItem.Name = "wH_MAILBindingNavigatorSaveItem";
            this.wH_MAILBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.wH_MAILBindingNavigatorSaveItem.Text = "儲存資料";
            this.wH_MAILBindingNavigatorSaveItem.Click += new System.EventHandler(this.wH_MAILBindingNavigatorSaveItem_Click);
            // 
            // wH_MAILDataGridView
            // 
            this.wH_MAILDataGridView.AutoGenerateColumns = false;
            this.wH_MAILDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.wH_MAILDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.wH_MAILDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3});
            this.wH_MAILDataGridView.DataSource = this.wH_MAILBindingSource;
            this.wH_MAILDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wH_MAILDataGridView.Location = new System.Drawing.Point(0, 0);
            this.wH_MAILDataGridView.Name = "wH_MAILDataGridView";
            this.wH_MAILDataGridView.RowTemplate.Height = 24;
            this.wH_MAILDataGridView.Size = new System.Drawing.Size(880, 222);
            this.wH_MAILDataGridView.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.wH_MAILDataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 25);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(880, 222);
            this.panel2.TabIndex = 7;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "WHCODE";
            this.dataGridViewTextBoxColumn2.HeaderText = "倉庫代碼";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 78;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "WHNAME";
            this.dataGridViewTextBoxColumn3.HeaderText = "倉庫名稱";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 78;
            // 
            // cEMAILTextBox
            // 
            this.cEMAILTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.wH_MAILBindingSource, "CEMAIL", true));
            this.cEMAILTextBox.Location = new System.Drawing.Point(456, 9);
            this.cEMAILTextBox.Multiline = true;
            this.cEMAILTextBox.Name = "cEMAILTextBox";
            this.cEMAILTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.cEMAILTextBox.Size = new System.Drawing.Size(257, 188);
            this.cEMAILTextBox.TabIndex = 5;
            // 
            // cEMAILLabel
            // 
            cEMAILLabel.AutoSize = true;
            cEMAILLabel.Location = new System.Drawing.Point(405, 9);
            cEMAILLabel.Name = "cEMAILLabel";
            cEMAILLabel.Size = new System.Drawing.Size(45, 12);
            cEMAILLabel.TabIndex = 4;
            cEMAILLabel.Text = "CC人員";
            // 
            // sEMAILTextBox
            // 
            this.sEMAILTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.wH_MAILBindingSource, "SEMAIL", true));
            this.sEMAILTextBox.Location = new System.Drawing.Point(99, 6);
            this.sEMAILTextBox.Multiline = true;
            this.sEMAILTextBox.Name = "sEMAILTextBox";
            this.sEMAILTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.sEMAILTextBox.Size = new System.Drawing.Size(257, 191);
            this.sEMAILTextBox.TabIndex = 3;
            // 
            // sEMAILLabel
            // 
            sEMAILLabel.AutoSize = true;
            sEMAILLabel.Location = new System.Drawing.Point(24, 9);
            sEMAILLabel.Name = "sEMAILLabel";
            sEMAILLabel.Size = new System.Drawing.Size(56, 12);
            sEMAILLabel.TabIndex = 2;
            sEMAILLabel.Text = "主要窗口 ";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(sEMAILLabel);
            this.panel1.Controls.Add(this.sEMAILTextBox);
            this.panel1.Controls.Add(cEMAILLabel);
            this.panel1.Controls.Add(this.cEMAILTextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 247);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(880, 250);
            this.panel1.TabIndex = 6;
            // 
            // WH_MAIL
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(880, 497);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.wH_MAILBindingNavigator);
            this.Name = "WH_MAIL";
            this.Text = "收貨工單窗口";
            this.Load += new System.EventHandler(this.WH_MAIL_Load);
            ((System.ComponentModel.ISupportInitialize)(this.wh)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_MAILBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_MAILBindingNavigator)).EndInit();
            this.wH_MAILBindingNavigator.ResumeLayout(false);
            this.wH_MAILBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.wH_MAILDataGridView)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.wh wh;
        private System.Windows.Forms.BindingSource wH_MAILBindingSource;
        private ACMEDataSet.whTableAdapters.WH_MAILTableAdapter wH_MAILTableAdapter;
        private System.Windows.Forms.BindingNavigator wH_MAILBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton wH_MAILBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView wH_MAILDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox cEMAILTextBox;
        private System.Windows.Forms.TextBox sEMAILTextBox;
        private System.Windows.Forms.Panel panel1;
    }
}