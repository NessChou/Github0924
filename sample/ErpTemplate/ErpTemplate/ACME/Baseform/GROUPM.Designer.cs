namespace ACME
{
    partial class GROUPM
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
            System.Windows.Forms.Label shippingCodeLabel;
            System.Windows.Forms.Label qUARTERLabel;
            System.Windows.Forms.Label gNAMELabel;
            System.Windows.Forms.Label gDATELabel;
            System.Windows.Forms.Label aMTLabel;
            System.Windows.Forms.Label mEMBERLabel;
            this.shippingCodeTextBox = new System.Windows.Forms.TextBox();
            this.gROUPMBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.uSERS = new ACME.ACMEDataSet.USERS();
            this.qUARTERTextBox = new System.Windows.Forms.TextBox();
            this.gNAMETextBox = new System.Windows.Forms.TextBox();
            this.gDATETextBox = new System.Windows.Forms.TextBox();
            this.aMTTextBox = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.mEMBERTextBox = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.gROUPD1DataGridView = new System.Windows.Forms.DataGridView();
            this.gROUPD1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.gROUPDDataGridView = new System.Windows.Forms.DataGridView();
            this.LINENUM = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gROUPDBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gROUPMTableAdapter = new ACME.ACMEDataSet.USERSTableAdapters.GROUPMTableAdapter();
            this.gROUPDTableAdapter = new ACME.ACMEDataSet.USERSTableAdapters.GROUPDTableAdapter();
            this.gROUPD1TableAdapter = new ACME.ACMEDataSet.USERSTableAdapters.GROUPD1TableAdapter();
            this.LINENUM1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.AMT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GRADE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AMT2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TIME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            shippingCodeLabel = new System.Windows.Forms.Label();
            qUARTERLabel = new System.Windows.Forms.Label();
            gNAMELabel = new System.Windows.Forms.Label();
            gDATELabel = new System.Windows.Forms.Label();
            aMTLabel = new System.Windows.Forms.Label();
            mEMBERLabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPMBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uSERS)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPD1DataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPD1BindingSource)).BeginInit();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPDDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPDBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Size = new System.Drawing.Size(1082, 465);
            this.panel1.Controls.SetChildIndex(this.panel2, 0);
            this.panel1.Controls.SetChildIndex(this.panel3, 0);
            // 
            // shippingCodeLabel
            // 
            shippingCodeLabel.AutoSize = true;
            shippingCodeLabel.Location = new System.Drawing.Point(12, 10);
            shippingCodeLabel.Name = "shippingCodeLabel";
            shippingCodeLabel.Size = new System.Drawing.Size(44, 12);
            shippingCodeLabel.TabIndex = 1;
            shippingCodeLabel.Text = "JOB NO";
            // 
            // qUARTERLabel
            // 
            qUARTERLabel.AutoSize = true;
            qUARTERLabel.Location = new System.Drawing.Point(198, 40);
            qUARTERLabel.Name = "qUARTERLabel";
            qUARTERLabel.Size = new System.Drawing.Size(17, 12);
            qUARTERLabel.TabIndex = 3;
            qUARTERLabel.Text = "季";
            // 
            // gNAMELabel
            // 
            gNAMELabel.AutoSize = true;
            gNAMELabel.Location = new System.Drawing.Point(220, 10);
            gNAMELabel.Name = "gNAMELabel";
            gNAMELabel.Size = new System.Drawing.Size(49, 12);
            gNAMELabel.TabIndex = 5;
            gNAMELabel.Text = "GNAME:";
            // 
            // gDATELabel
            // 
            gDATELabel.AutoSize = true;
            gDATELabel.Location = new System.Drawing.Point(12, 40);
            gDATELabel.Name = "gDATELabel";
            gDATELabel.Size = new System.Drawing.Size(29, 12);
            gDATELabel.TabIndex = 7;
            gDATELabel.Text = "日期";
            // 
            // aMTLabel
            // 
            aMTLabel.AutoSize = true;
            aMTLabel.Location = new System.Drawing.Point(354, 40);
            aMTLabel.Name = "aMTLabel";
            aMTLabel.Size = new System.Drawing.Size(29, 12);
            aMTLabel.TabIndex = 9;
            aMTLabel.Text = "金額";
            // 
            // mEMBERLabel
            // 
            mEMBERLabel.AutoSize = true;
            mEMBERLabel.Location = new System.Drawing.Point(564, 10);
            mEMBERLabel.Name = "mEMBERLabel";
            mEMBERLabel.Size = new System.Drawing.Size(58, 12);
            mEMBERLabel.TabIndex = 10;
            mEMBERLabel.Text = "MEMBER:";
            // 
            // shippingCodeTextBox
            // 
            this.shippingCodeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gROUPMBindingSource, "ShippingCode", true));
            this.shippingCodeTextBox.Location = new System.Drawing.Point(62, 7);
            this.shippingCodeTextBox.Name = "shippingCodeTextBox";
            this.shippingCodeTextBox.Size = new System.Drawing.Size(135, 22);
            this.shippingCodeTextBox.TabIndex = 2;
            // 
            // gROUPMBindingSource
            // 
            this.gROUPMBindingSource.DataMember = "GROUPM";
            this.gROUPMBindingSource.DataSource = this.uSERS;
            // 
            // uSERS
            // 
            this.uSERS.DataSetName = "USERS";
            this.uSERS.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // qUARTERTextBox
            // 
            this.qUARTERTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gROUPMBindingSource, "QUARTER", true));
            this.qUARTERTextBox.Location = new System.Drawing.Point(231, 35);
            this.qUARTERTextBox.Name = "qUARTERTextBox";
            this.qUARTERTextBox.Size = new System.Drawing.Size(100, 22);
            this.qUARTERTextBox.TabIndex = 4;
            // 
            // gNAMETextBox
            // 
            this.gNAMETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gROUPMBindingSource, "GNAME", true));
            this.gNAMETextBox.Location = new System.Drawing.Point(275, 7);
            this.gNAMETextBox.Name = "gNAMETextBox";
            this.gNAMETextBox.Size = new System.Drawing.Size(283, 22);
            this.gNAMETextBox.TabIndex = 6;
            // 
            // gDATETextBox
            // 
            this.gDATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gROUPMBindingSource, "GDATE", true));
            this.gDATETextBox.Location = new System.Drawing.Point(64, 37);
            this.gDATETextBox.Name = "gDATETextBox";
            this.gDATETextBox.Size = new System.Drawing.Size(100, 22);
            this.gDATETextBox.TabIndex = 8;
            // 
            // aMTTextBox
            // 
            this.aMTTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gROUPMBindingSource, "AMT", true));
            this.aMTTextBox.Location = new System.Drawing.Point(393, 35);
            this.aMTTextBox.Name = "aMTTextBox";
            this.aMTTextBox.Size = new System.Drawing.Size(100, 22);
            this.aMTTextBox.TabIndex = 10;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.button2);
            this.panel2.Controls.Add(mEMBERLabel);
            this.panel2.Controls.Add(this.mEMBERTextBox);
            this.panel2.Controls.Add(shippingCodeLabel);
            this.panel2.Controls.Add(this.shippingCodeTextBox);
            this.panel2.Controls.Add(aMTLabel);
            this.panel2.Controls.Add(this.qUARTERTextBox);
            this.panel2.Controls.Add(this.aMTTextBox);
            this.panel2.Controls.Add(qUARTERLabel);
            this.panel2.Controls.Add(gDATELabel);
            this.panel2.Controls.Add(this.gNAMETextBox);
            this.panel2.Controls.Add(this.gDATETextBox);
            this.panel2.Controls.Add(gNAMELabel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1082, 79);
            this.panel2.TabIndex = 11;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(514, 40);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 12;
            this.button2.Text = "EXCEL";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // mEMBERTextBox
            // 
            this.mEMBERTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.gROUPMBindingSource, "MEMBER", true));
            this.mEMBERTextBox.Location = new System.Drawing.Point(638, 7);
            this.mEMBERTextBox.Multiline = true;
            this.mEMBERTextBox.Name = "mEMBERTextBox";
            this.mEMBERTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.mEMBERTextBox.Size = new System.Drawing.Size(383, 66);
            this.mEMBERTextBox.TabIndex = 11;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.tabControl1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 79);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1082, 364);
            this.panel3.TabIndex = 12;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1082, 364);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.gROUPD1DataGridView);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1074, 338);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // gROUPD1DataGridView
            // 
            this.gROUPD1DataGridView.AutoGenerateColumns = false;
            this.gROUPD1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gROUPD1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.LINENUM1,
            this.dataGridViewTextBoxColumn7,
            this.AMT,
            this.GRADE,
            this.AMT2,
            this.TIME});
            this.gROUPD1DataGridView.DataSource = this.gROUPD1BindingSource;
            this.gROUPD1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gROUPD1DataGridView.Location = new System.Drawing.Point(3, 3);
            this.gROUPD1DataGridView.Name = "gROUPD1DataGridView";
            this.gROUPD1DataGridView.RowTemplate.Height = 24;
            this.gROUPD1DataGridView.Size = new System.Drawing.Size(1068, 332);
            this.gROUPD1DataGridView.TabIndex = 0;
            this.gROUPD1DataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.gROUPD1DataGridView_DataError);
            this.gROUPD1DataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.gROUPD1DataGridView_DefaultValuesNeeded);
            // 
            // gROUPD1BindingSource
            // 
            this.gROUPD1BindingSource.DataMember = "GROUPM_GROUPD1";
            this.gROUPD1BindingSource.DataSource = this.gROUPMBindingSource;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.gROUPDDataGridView);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(826, 338);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // gROUPDDataGridView
            // 
            this.gROUPDDataGridView.AutoGenerateColumns = false;
            this.gROUPDDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.gROUPDDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gROUPDDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.LINENUM,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4});
            this.gROUPDDataGridView.DataSource = this.gROUPDBindingSource;
            this.gROUPDDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gROUPDDataGridView.Location = new System.Drawing.Point(3, 3);
            this.gROUPDDataGridView.Name = "gROUPDDataGridView";
            this.gROUPDDataGridView.RowTemplate.Height = 24;
            this.gROUPDDataGridView.Size = new System.Drawing.Size(820, 332);
            this.gROUPDDataGridView.TabIndex = 0;
            this.gROUPDDataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.gROUPDDataGridView_DataError);
            this.gROUPDDataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.gROUPDDataGridView_DefaultValuesNeeded);
            // 
            // LINENUM
            // 
            this.LINENUM.DataPropertyName = "LINENUM";
            this.LINENUM.HeaderText = "LINENUM";
            this.LINENUM.Name = "LINENUM";
            this.LINENUM.Width = 82;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "ITEMNAME";
            this.dataGridViewTextBoxColumn3.HeaderText = "ITEMNAME";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 91;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "AMT";
            this.dataGridViewTextBoxColumn4.HeaderText = "AMT";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.Width = 55;
            // 
            // gROUPDBindingSource
            // 
            this.gROUPDBindingSource.DataMember = "GROUPM_GROUPD";
            this.gROUPDBindingSource.DataSource = this.gROUPMBindingSource;
            // 
            // gROUPMTableAdapter
            // 
            this.gROUPMTableAdapter.ClearBeforeFill = true;
            // 
            // gROUPDTableAdapter
            // 
            this.gROUPDTableAdapter.ClearBeforeFill = true;
            // 
            // gROUPD1TableAdapter
            // 
            this.gROUPD1TableAdapter.ClearBeforeFill = true;
            // 
            // LINENUM1
            // 
            this.LINENUM1.DataPropertyName = "LINENUM";
            this.LINENUM1.HeaderText = "LINENUM";
            this.LINENUM1.Name = "LINENUM1";
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "USERS";
            this.dataGridViewTextBoxColumn7.HeaderText = "USERS";
            this.dataGridViewTextBoxColumn7.Items.AddRange(new object[] {
            "anniechung",
            "applechen",
            "danielliu",
            "bettytseng",
            "davidhuang",
            "eileenchiu",
            "estheryeh",
            "evahsu",
            "fionalai",
            "gabrielyang",
            "jackywang",
            "jameswu",
            "johnnytsai",
            "kerrychan",
            "lilylee",
            "lleytonchen",
            "maggieweng",
            "michelleko",
            "monicalin",
            "pattyliu",
            "seanchen",
            "sharonhuang",
            "shirleyjuan",
            "sunnywang",
            "tonywu",
            "ulatsai",
            "viviweng"});
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // AMT
            // 
            this.AMT.DataPropertyName = "AMT";
            this.AMT.HeaderText = "公里數";
            this.AMT.Name = "AMT";
            // 
            // GRADE
            // 
            this.GRADE.DataPropertyName = "GRADE";
            this.GRADE.HeaderText = "GRADE";
            this.GRADE.Name = "GRADE";
            // 
            // AMT2
            // 
            this.AMT2.DataPropertyName = "AMT2";
            this.AMT2.HeaderText = "AMT2";
            this.AMT2.Name = "AMT2";
            // 
            // TIME
            // 
            this.TIME.DataPropertyName = "TIME";
            this.TIME.HeaderText = "TIME";
            this.TIME.Name = "TIME";
            // 
            // GROUPM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(1082, 503);
            this.Name = "GROUPM";
            this.Load += new System.EventHandler(this.GROUPM_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPMBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uSERS)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gROUPD1DataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPD1BindingSource)).EndInit();
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gROUPDDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gROUPDBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.USERS uSERS;
        private System.Windows.Forms.BindingSource gROUPMBindingSource;
        private ACMEDataSet.USERSTableAdapters.GROUPMTableAdapter gROUPMTableAdapter;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox shippingCodeTextBox;
        private System.Windows.Forms.TextBox qUARTERTextBox;
        private System.Windows.Forms.TextBox aMTTextBox;
        private System.Windows.Forms.TextBox gNAMETextBox;
        private System.Windows.Forms.TextBox gDATETextBox;
        private System.Windows.Forms.BindingSource gROUPDBindingSource;
        private ACMEDataSet.USERSTableAdapters.GROUPDTableAdapter gROUPDTableAdapter;
        private System.Windows.Forms.DataGridView gROUPDDataGridView;
        private System.Windows.Forms.BindingSource gROUPD1BindingSource;
        private ACMEDataSet.USERSTableAdapters.GROUPD1TableAdapter gROUPD1TableAdapter;
        private System.Windows.Forms.DataGridView gROUPD1DataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn LINENUM;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.TextBox mEMBERTextBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridViewTextBoxColumn LINENUM1;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn AMT;
        private System.Windows.Forms.DataGridViewTextBoxColumn GRADE;
        private System.Windows.Forms.DataGridViewTextBoxColumn AMT2;
        private System.Windows.Forms.DataGridViewTextBoxColumn TIME;

    }
}
