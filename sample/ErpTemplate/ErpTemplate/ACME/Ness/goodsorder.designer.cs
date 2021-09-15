
namespace ACME
{
    partial class goodsorder
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.labend = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.labstart = new System.Windows.Forms.Label();
            this.labelweight = new System.Windows.Forms.Label();
            this.dgvGoods = new System.Windows.Forms.DataGridView();
            this.DeleteCheck = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.stackCheck = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.stacktag = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tag = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvSort = new System.Windows.Forms.DataGridView();
            this.btnOrder3 = new System.Windows.Forms.Button();
            this.btnOrder2 = new System.Windows.Forms.Button();
            this.btnsort = new System.Windows.Forms.Button();
            this.btnOrder1 = new System.Windows.Forms.Button();
            this.cobend = new System.Windows.Forms.ComboBox();
            this.cobstart = new System.Windows.Forms.ComboBox();
            this.cobTruckRegion = new System.Windows.Forms.ComboBox();
            this.cobTruckType = new System.Windows.Forms.ComboBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnStack = new System.Windows.Forms.Button();
            this.btnImportExcel = new System.Windows.Forms.Button();
            this.btnInsert = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGoods)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSort)).BeginInit();
            this.SuspendLayout();
            // 
            // labend
            // 
            this.labend.AutoSize = true;
            this.labend.Location = new System.Drawing.Point(390, 73);
            this.labend.Name = "labend";
            this.labend.Size = new System.Drawing.Size(29, 12);
            this.labend.TabIndex = 27;
            this.labend.Text = "終點";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(220, 139);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 26;
            this.label1.Text = "車型";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(220, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 25;
            this.label2.Text = "廠商";
            // 
            // labstart
            // 
            this.labstart.AutoSize = true;
            this.labstart.Location = new System.Drawing.Point(220, 73);
            this.labstart.Name = "labstart";
            this.labstart.Size = new System.Drawing.Size(29, 12);
            this.labstart.TabIndex = 24;
            this.labstart.Text = "起點";
            // 
            // labelweight
            // 
            this.labelweight.AutoSize = true;
            this.labelweight.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelweight.Location = new System.Drawing.Point(45, 139);
            this.labelweight.Name = "labelweight";
            this.labelweight.Size = new System.Drawing.Size(44, 16);
            this.labelweight.TabIndex = 23;
            this.labelweight.Text = "總重:";
            // 
            // dgvGoods
            // 
            this.dgvGoods.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvGoods.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DeleteCheck,
            this.stackCheck,
            this.type,
            this.num,
            this.stacktag,
            this.tag});
            this.dgvGoods.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvGoods.Location = new System.Drawing.Point(64, 173);
            this.dgvGoods.Name = "dgvGoods";
            this.dgvGoods.RowTemplate.Height = 24;
            this.dgvGoods.Size = new System.Drawing.Size(576, 570);
            this.dgvGoods.TabIndex = 22;
            // 
            // DeleteCheck
            // 
            this.DeleteCheck.DataPropertyName = "DeleteCheck";
            this.DeleteCheck.HeaderText = "刪";
            this.DeleteCheck.Name = "DeleteCheck";
            this.DeleteCheck.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.DeleteCheck.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.DeleteCheck.Width = 25;
            // 
            // stackCheck
            // 
            this.stackCheck.DataPropertyName = "stackCheck";
            this.stackCheck.HeaderText = "堆疊";
            this.stackCheck.Name = "stackCheck";
            this.stackCheck.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.stackCheck.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.stackCheck.Width = 50;
            // 
            // type
            // 
            this.type.DataPropertyName = "type";
            this.type.HeaderText = "樣式";
            this.type.Name = "type";
            this.type.Width = 75;
            // 
            // num
            // 
            this.num.DataPropertyName = "num";
            this.num.HeaderText = "板號";
            this.num.Name = "num";
            this.num.Width = 65;
            // 
            // stacktag
            // 
            this.stacktag.DataPropertyName = "stacktag";
            this.stacktag.HeaderText = "已堆疊";
            this.stacktag.Name = "stacktag";
            this.stacktag.ReadOnly = true;
            this.stacktag.Visible = false;
            // 
            // tag
            // 
            this.tag.DataPropertyName = "tag";
            this.tag.HeaderText = "已排序";
            this.tag.Name = "tag";
            this.tag.Visible = false;
            // 
            // dgvSort
            // 
            this.dgvSort.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSort.Location = new System.Drawing.Point(693, 173);
            this.dgvSort.Name = "dgvSort";
            this.dgvSort.RowTemplate.Height = 24;
            this.dgvSort.Size = new System.Drawing.Size(482, 570);
            this.dgvSort.TabIndex = 21;
            // 
            // btnOrder3
            // 
            this.btnOrder3.Location = new System.Drawing.Point(843, 98);
            this.btnOrder3.Name = "btnOrder3";
            this.btnOrder3.Size = new System.Drawing.Size(80, 49);
            this.btnOrder3.TabIndex = 19;
            this.btnOrder3.Text = "排序法三";
            this.btnOrder3.UseVisualStyleBackColor = true;
            this.btnOrder3.Click += new System.EventHandler(this.btnOrder_Click);
            // 
            // btnOrder2
            // 
            this.btnOrder2.Location = new System.Drawing.Point(740, 98);
            this.btnOrder2.Name = "btnOrder2";
            this.btnOrder2.Size = new System.Drawing.Size(80, 49);
            this.btnOrder2.TabIndex = 20;
            this.btnOrder2.Text = "排序法二";
            this.btnOrder2.UseVisualStyleBackColor = true;
            this.btnOrder2.Click += new System.EventHandler(this.btnOrder_Click);
            // 
            // btnsort
            // 
            this.btnsort.Location = new System.Drawing.Point(637, 36);
            this.btnsort.Name = "btnsort";
            this.btnsort.Size = new System.Drawing.Size(80, 49);
            this.btnsort.TabIndex = 18;
            this.btnsort.Text = "排序";
            this.btnsort.UseVisualStyleBackColor = true;
            this.btnsort.Click += new System.EventHandler(this.btnOrder_Click);
            // 
            // btnOrder1
            // 
            this.btnOrder1.Location = new System.Drawing.Point(637, 98);
            this.btnOrder1.Name = "btnOrder1";
            this.btnOrder1.Size = new System.Drawing.Size(80, 49);
            this.btnOrder1.TabIndex = 17;
            this.btnOrder1.Text = "排序法一";
            this.btnOrder1.UseVisualStyleBackColor = true;
            this.btnOrder1.Click += new System.EventHandler(this.btnOrder_Click);
            // 
            // cobend
            // 
            this.cobend.FormattingEnabled = true;
            this.cobend.Location = new System.Drawing.Point(429, 65);
            this.cobend.Name = "cobend";
            this.cobend.Size = new System.Drawing.Size(125, 20);
            this.cobend.TabIndex = 16;
            this.cobend.SelectedIndexChanged += new System.EventHandler(this.cobend_SelectedIndexChanged);
            // 
            // cobstart
            // 
            this.cobstart.FormattingEnabled = true;
            this.cobstart.Location = new System.Drawing.Point(259, 65);
            this.cobstart.Name = "cobstart";
            this.cobstart.Size = new System.Drawing.Size(125, 20);
            this.cobstart.TabIndex = 15;
            this.cobstart.SelectedIndexChanged += new System.EventHandler(this.cobstart_SelectedIndexChanged);
            // 
            // cobTruckRegion
            // 
            this.cobTruckRegion.FormattingEnabled = true;
            this.cobTruckRegion.Location = new System.Drawing.Point(259, 26);
            this.cobTruckRegion.Name = "cobTruckRegion";
            this.cobTruckRegion.Size = new System.Drawing.Size(125, 20);
            this.cobTruckRegion.TabIndex = 14;
            this.cobTruckRegion.SelectedIndexChanged += new System.EventHandler(this.cobTruckRegion_SelectedIndexChanged);
            // 
            // cobTruckType
            // 
            this.cobTruckType.FormattingEnabled = true;
            this.cobTruckType.Location = new System.Drawing.Point(259, 131);
            this.cobTruckType.Name = "cobTruckType";
            this.cobTruckType.Size = new System.Drawing.Size(295, 20);
            this.cobTruckType.TabIndex = 13;
            this.cobTruckType.SelectedIndexChanged += new System.EventHandler(this.cobTruckType_SelectedIndexChanged);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(141, 26);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(73, 38);
            this.btnDelete.TabIndex = 11;
            this.btnDelete.Text = "刪除貨物";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnStack
            // 
            this.btnStack.Location = new System.Drawing.Point(141, 79);
            this.btnStack.Name = "btnStack";
            this.btnStack.Size = new System.Drawing.Size(73, 38);
            this.btnStack.TabIndex = 10;
            this.btnStack.Text = "堆疊全選";
            this.btnStack.UseVisualStyleBackColor = true;
            this.btnStack.Click += new System.EventHandler(this.btnStack_Click);
            // 
            // btnImportExcel
            // 
            this.btnImportExcel.Location = new System.Drawing.Point(48, 79);
            this.btnImportExcel.Name = "btnImportExcel";
            this.btnImportExcel.Size = new System.Drawing.Size(73, 38);
            this.btnImportExcel.TabIndex = 12;
            this.btnImportExcel.Text = "匯入excel";
            this.btnImportExcel.UseVisualStyleBackColor = true;
            this.btnImportExcel.Click += new System.EventHandler(this.btnImportExcel_Click);
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(48, 26);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(73, 38);
            this.btnInsert.TabIndex = 9;
            this.btnInsert.Text = "新增貨物";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // goodsorder
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1221, 769);
            this.Controls.Add(this.labend);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.labstart);
            this.Controls.Add(this.labelweight);
            this.Controls.Add(this.dgvGoods);
            this.Controls.Add(this.dgvSort);
            this.Controls.Add(this.btnOrder3);
            this.Controls.Add(this.btnOrder2);
            this.Controls.Add(this.btnsort);
            this.Controls.Add(this.btnOrder1);
            this.Controls.Add(this.cobend);
            this.Controls.Add(this.cobstart);
            this.Controls.Add(this.cobTruckRegion);
            this.Controls.Add(this.cobTruckType);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnStack);
            this.Controls.Add(this.btnImportExcel);
            this.Controls.Add(this.btnInsert);
            this.Name = "goodsorder";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgvGoods)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSort)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labend;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label labstart;
        private System.Windows.Forms.Label labelweight;
        private System.Windows.Forms.DataGridView dgvGoods;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DeleteCheck;
        private System.Windows.Forms.DataGridViewCheckBoxColumn stackCheck;
        private System.Windows.Forms.DataGridViewTextBoxColumn type;
        private System.Windows.Forms.DataGridViewTextBoxColumn num;
        private System.Windows.Forms.DataGridViewTextBoxColumn stacktag;
        private System.Windows.Forms.DataGridViewTextBoxColumn tag;
        private System.Windows.Forms.DataGridView dgvSort;
        private System.Windows.Forms.Button btnOrder3;
        private System.Windows.Forms.Button btnOrder2;
        private System.Windows.Forms.Button btnsort;
        private System.Windows.Forms.Button btnOrder1;
        private System.Windows.Forms.ComboBox cobend;
        private System.Windows.Forms.ComboBox cobstart;
        private System.Windows.Forms.ComboBox cobTruckRegion;
        private System.Windows.Forms.ComboBox cobTruckType;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnStack;
        private System.Windows.Forms.Button btnImportExcel;
        private System.Windows.Forms.Button btnInsert;
    }
}

