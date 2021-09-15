namespace ACME
{
    partial class HRDIN
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HRDIN));
            this.hR = new ACME.ACMEDataSet.HR();
            this.dINBENDON_USERBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dINBENDON_USERTableAdapter = new ACME.ACMEDataSet.HRTableAdapters.DINBENDON_USERTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.HRTableAdapters.TableAdapterManager();
            this.dINBENDON_USERBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.dINBENDON_USERBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripTextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.dINBENDON_USERDataGridView = new System.Windows.Forms.DataGridView();
            this.日期 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DINNER = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.hR)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_USERBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_USERBindingNavigator)).BeginInit();
            this.dINBENDON_USERBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_USERDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // hR
            // 
            this.hR.DataSetName = "HR";
            this.hR.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dINBENDON_USERBindingSource
            // 
            this.dINBENDON_USERBindingSource.DataMember = "DINBENDON_USER";
            this.dINBENDON_USERBindingSource.DataSource = this.hR;
            // 
            // dINBENDON_USERTableAdapter
            // 
            this.dINBENDON_USERTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.DINBENDON_USERTableAdapter = this.dINBENDON_USERTableAdapter;
            this.tableAdapterManager.HR_BUTableAdapter = null;
            this.tableAdapterManager.HR_Main104TableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.HRTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // dINBENDON_USERBindingNavigator
            // 
            this.dINBENDON_USERBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.dINBENDON_USERBindingNavigator.BindingSource = this.dINBENDON_USERBindingSource;
            this.dINBENDON_USERBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.dINBENDON_USERBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.dINBENDON_USERBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.dINBENDON_USERBindingNavigatorSaveItem,
            this.toolStripLabel1,
            this.toolStripTextBox1});
            this.dINBENDON_USERBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.dINBENDON_USERBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.dINBENDON_USERBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.dINBENDON_USERBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.dINBENDON_USERBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.dINBENDON_USERBindingNavigator.Name = "dINBENDON_USERBindingNavigator";
            this.dINBENDON_USERBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.dINBENDON_USERBindingNavigator.Size = new System.Drawing.Size(856, 25);
            this.dINBENDON_USERBindingNavigator.TabIndex = 0;
            this.dINBENDON_USERBindingNavigator.Text = "bindingNavigator1";
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
            // dINBENDON_USERBindingNavigatorSaveItem
            // 
            this.dINBENDON_USERBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.dINBENDON_USERBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("dINBENDON_USERBindingNavigatorSaveItem.Image")));
            this.dINBENDON_USERBindingNavigatorSaveItem.Name = "dINBENDON_USERBindingNavigatorSaveItem";
            this.dINBENDON_USERBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.dINBENDON_USERBindingNavigatorSaveItem.Text = "儲存資料";
            this.dINBENDON_USERBindingNavigatorSaveItem.Click += new System.EventHandler(this.dINBENDON_USERBindingNavigatorSaveItem_Click);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(31, 22);
            this.toolStripLabel1.Text = "日期";
            // 
            // toolStripTextBox1
            // 
            this.toolStripTextBox1.Name = "toolStripTextBox1";
            this.toolStripTextBox1.Size = new System.Drawing.Size(100, 25);
            this.toolStripTextBox1.TextChanged += new System.EventHandler(this.toolStripTextBox1_TextChanged);
            // 
            // dINBENDON_USERDataGridView
            // 
            this.dINBENDON_USERDataGridView.AutoGenerateColumns = false;
            this.dINBENDON_USERDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dINBENDON_USERDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.日期,
            this.dataGridViewTextBoxColumn7,
            this.DINNER});
            this.dINBENDON_USERDataGridView.DataSource = this.dINBENDON_USERBindingSource;
            this.dINBENDON_USERDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dINBENDON_USERDataGridView.Location = new System.Drawing.Point(0, 25);
            this.dINBENDON_USERDataGridView.Name = "dINBENDON_USERDataGridView";
            this.dINBENDON_USERDataGridView.RowTemplate.Height = 24;
            this.dINBENDON_USERDataGridView.Size = new System.Drawing.Size(856, 565);
            this.dINBENDON_USERDataGridView.TabIndex = 1;
            this.dINBENDON_USERDataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.dINBENDON_USERDataGridView_DefaultValuesNeeded);
            // 
            // 日期
            // 
            this.日期.DataPropertyName = "datetime";
            this.日期.HeaderText = "日期";
            this.日期.Name = "日期";
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "userid";
            this.dataGridViewTextBoxColumn7.HeaderText = "使用者";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            // 
            // DINNER
            // 
            this.DINNER.DataPropertyName = "name";
            this.DINNER.HeaderText = "晚餐";
            this.DINNER.Name = "DINNER";
            // 
            // HRDIN
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(856, 590);
            this.Controls.Add(this.dINBENDON_USERDataGridView);
            this.Controls.Add(this.dINBENDON_USERBindingNavigator);
            this.Name = "HRDIN";
            this.Text = "HRDIN";
            this.Load += new System.EventHandler(this.HRDIN_Load);
            ((System.ComponentModel.ISupportInitialize)(this.hR)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_USERBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_USERBindingNavigator)).EndInit();
            this.dINBENDON_USERBindingNavigator.ResumeLayout(false);
            this.dINBENDON_USERBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_USERDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.HR hR;
        private System.Windows.Forms.BindingSource dINBENDON_USERBindingSource;
        private ACMEDataSet.HRTableAdapters.DINBENDON_USERTableAdapter dINBENDON_USERTableAdapter;
        private ACMEDataSet.HRTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.BindingNavigator dINBENDON_USERBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton dINBENDON_USERBindingNavigatorSaveItem;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripTextBox toolStripTextBox1;
        private System.Windows.Forms.DataGridView dINBENDON_USERDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 日期;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn DINNER;
    }
}