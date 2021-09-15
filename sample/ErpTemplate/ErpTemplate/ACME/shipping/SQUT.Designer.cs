namespace ACME
{
    partial class SQUT
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
            System.Windows.Forms.Label cARDCODELabel;
            System.Windows.Forms.Label iTEMCODELabel;
            System.Windows.Forms.Label dOCDATELabel;
            System.Windows.Forms.Label eNDDATELabel;
            System.Windows.Forms.Label sHIPWAYLabel;
            System.Windows.Forms.Label createNameLabel;
            System.Windows.Forms.Label tRADELabel;
            System.Windows.Forms.Label tERMLabel;
            System.Windows.Forms.Label label11;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SQUT));
            this.ship = new ACME.ACMEDataSet.ship();
            this.shipping_OQUTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.shipping_OQUTTableAdapter = new ACME.ACMEDataSet.shipTableAdapters.Shipping_OQUTTableAdapter();
            this.shippingCodeTextBox = new System.Windows.Forms.TextBox();
            this.cARDNAMETextBox = new System.Windows.Forms.TextBox();
            this.iTEMCODETextBox = new System.Windows.Forms.TextBox();
            this.iTEMNAMETextBox = new System.Windows.Forms.TextBox();
            this.dOCDATETextBox = new System.Windows.Forms.TextBox();
            this.eNDDATETextBox = new System.Windows.Forms.TextBox();
            this.shipping_OQUT1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.shipping_OQUT1TableAdapter = new ACME.ACMEDataSet.shipTableAdapters.Shipping_OQUT1TableAdapter();
            this.shipping_OQUT1DataGridView = new System.Windows.Forms.DataGridView();
            this.QTYPE = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.createNameTextBox = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.cARDCODETextBox = new System.Windows.Forms.TextBox();
            this.tRADETextBox = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button25 = new System.Windows.Forms.Button();
            this.shipping_OQUTDownloadDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LINK = new System.Windows.Forms.DataGridViewLinkColumn();
            this.shipping_OQUTDownloadBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.mEMOTextBox = new System.Windows.Forms.TextBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.mAILTextBox = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tERMTextBox = new System.Windows.Forms.TextBox();
            this.comboBox5 = new System.Windows.Forms.ComboBox();
            this.sHIPWAYTextBox = new System.Windows.Forms.TextBox();
            this.comboBox6 = new System.Windows.Forms.ComboBox();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button7 = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.shipping_OQUTDownloadTableAdapter = new ACME.ACMEDataSet.shipTableAdapters.Shipping_OQUTDownloadTableAdapter();
            shippingCodeLabel = new System.Windows.Forms.Label();
            cARDCODELabel = new System.Windows.Forms.Label();
            iTEMCODELabel = new System.Windows.Forms.Label();
            dOCDATELabel = new System.Windows.Forms.Label();
            eNDDATELabel = new System.Windows.Forms.Label();
            sHIPWAYLabel = new System.Windows.Forms.Label();
            createNameLabel = new System.Windows.Forms.Label();
            tRADELabel = new System.Windows.Forms.Label();
            tERMLabel = new System.Windows.Forms.Label();
            label11 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT1DataGridView)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadBindingSource)).BeginInit();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Size = new System.Drawing.Size(963, 525);
            this.panel1.Controls.SetChildIndex(this.panel2, 0);
            this.panel1.Controls.SetChildIndex(this.panel3, 0);
            // 
            // shippingCodeLabel
            // 
            shippingCodeLabel.AutoSize = true;
            shippingCodeLabel.Location = new System.Drawing.Point(14, 9);
            shippingCodeLabel.Name = "shippingCodeLabel";
            shippingCodeLabel.Size = new System.Drawing.Size(44, 12);
            shippingCodeLabel.TabIndex = 1;
            shippingCodeLabel.Text = "JOB NO";
            // 
            // cARDCODELabel
            // 
            cARDCODELabel.AutoSize = true;
            cARDCODELabel.Location = new System.Drawing.Point(5, 56);
            cARDCODELabel.Name = "cARDCODELabel";
            cARDCODELabel.Size = new System.Drawing.Size(65, 12);
            cARDCODELabel.TabIndex = 3;
            cARDCODELabel.Text = "供應商編號";
            // 
            // iTEMCODELabel
            // 
            iTEMCODELabel.AutoSize = true;
            iTEMCODELabel.Location = new System.Drawing.Point(14, 147);
            iTEMCODELabel.Name = "iTEMCODELabel";
            iTEMCODELabel.Size = new System.Drawing.Size(53, 12);
            iTEMCODELabel.TabIndex = 7;
            iTEMCODELabel.Text = "運輸路線";
            // 
            // dOCDATELabel
            // 
            dOCDATELabel.AutoSize = true;
            dOCDATELabel.Location = new System.Drawing.Point(197, 9);
            dOCDATELabel.Name = "dOCDATELabel";
            dOCDATELabel.Size = new System.Drawing.Size(53, 12);
            dOCDATELabel.TabIndex = 11;
            dOCDATELabel.Text = "報價日期";
            // 
            // eNDDATELabel
            // 
            eNDDATELabel.AutoSize = true;
            eNDDATELabel.Location = new System.Drawing.Point(398, 9);
            eNDDATELabel.Name = "eNDDATELabel";
            eNDDATELabel.Size = new System.Drawing.Size(53, 12);
            eNDDATELabel.TabIndex = 13;
            eNDDATELabel.Text = "有效日期";
            // 
            // sHIPWAYLabel
            // 
            sHIPWAYLabel.AutoSize = true;
            sHIPWAYLabel.Location = new System.Drawing.Point(14, 100);
            sHIPWAYLabel.Name = "sHIPWAYLabel";
            sHIPWAYLabel.Size = new System.Drawing.Size(53, 12);
            sHIPWAYLabel.TabIndex = 15;
            sHIPWAYLabel.Text = "運輸方式";
            // 
            // createNameLabel
            // 
            createNameLabel.AutoSize = true;
            createNameLabel.Location = new System.Drawing.Point(590, 9);
            createNameLabel.Name = "createNameLabel";
            createNameLabel.Size = new System.Drawing.Size(41, 12);
            createNameLabel.TabIndex = 16;
            createNameLabel.Text = "製單者";
            // 
            // tRADELabel
            // 
            tRADELabel.AutoSize = true;
            tRADELabel.Location = new System.Drawing.Point(169, 100);
            tRADELabel.Name = "tRADELabel";
            tRADELabel.Size = new System.Drawing.Size(53, 12);
            tRADELabel.TabIndex = 71;
            tRADELabel.Text = "貿易條件";
            // 
            // tERMLabel
            // 
            tERMLabel.AutoSize = true;
            tERMLabel.Location = new System.Drawing.Point(453, 102);
            tERMLabel.Name = "tERMLabel";
            tERMLabel.Size = new System.Drawing.Size(53, 12);
            tERMLabel.TabIndex = 72;
            tERMLabel.Text = "貿易形式";
            // 
            // label11
            // 
            label11.AutoSize = true;
            label11.Location = new System.Drawing.Point(14, 177);
            label11.Name = "label11";
            label11.Size = new System.Drawing.Size(34, 12);
            label11.TabIndex = 86;
            label11.Text = "MAIL";
            // 
            // ship
            // 
            this.ship.DataSetName = "ship";
            this.ship.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // shipping_OQUTBindingSource
            // 
            this.shipping_OQUTBindingSource.DataMember = "Shipping_OQUT";
            this.shipping_OQUTBindingSource.DataSource = this.ship;
            // 
            // shipping_OQUTTableAdapter
            // 
            this.shipping_OQUTTableAdapter.ClearBeforeFill = true;
            // 
            // shippingCodeTextBox
            // 
            this.shippingCodeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "ShippingCode", true));
            this.shippingCodeTextBox.Location = new System.Drawing.Point(76, 6);
            this.shippingCodeTextBox.Name = "shippingCodeTextBox";
            this.shippingCodeTextBox.Size = new System.Drawing.Size(100, 22);
            this.shippingCodeTextBox.TabIndex = 2;
            // 
            // cARDNAMETextBox
            // 
            this.cARDNAMETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "CARDNAME", true));
            this.cARDNAMETextBox.Location = new System.Drawing.Point(161, 53);
            this.cARDNAMETextBox.Name = "cARDNAMETextBox";
            this.cARDNAMETextBox.Size = new System.Drawing.Size(267, 22);
            this.cARDNAMETextBox.TabIndex = 6;
            // 
            // iTEMCODETextBox
            // 
            this.iTEMCODETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "ITEMCODE", true));
            this.iTEMCODETextBox.Location = new System.Drawing.Point(76, 144);
            this.iTEMCODETextBox.Name = "iTEMCODETextBox";
            this.iTEMCODETextBox.Size = new System.Drawing.Size(86, 22);
            this.iTEMCODETextBox.TabIndex = 8;
            // 
            // iTEMNAMETextBox
            // 
            this.iTEMNAMETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "ITEMNAME", true));
            this.iTEMNAMETextBox.Location = new System.Drawing.Point(161, 144);
            this.iTEMNAMETextBox.Name = "iTEMNAMETextBox";
            this.iTEMNAMETextBox.Size = new System.Drawing.Size(267, 22);
            this.iTEMNAMETextBox.TabIndex = 10;
            // 
            // dOCDATETextBox
            // 
            this.dOCDATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "DOCDATE", true));
            this.dOCDATETextBox.Location = new System.Drawing.Point(265, 6);
            this.dOCDATETextBox.Name = "dOCDATETextBox";
            this.dOCDATETextBox.Size = new System.Drawing.Size(100, 22);
            this.dOCDATETextBox.TabIndex = 12;
            // 
            // eNDDATETextBox
            // 
            this.eNDDATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "ENDDATE", true));
            this.eNDDATETextBox.Location = new System.Drawing.Point(462, 6);
            this.eNDDATETextBox.Name = "eNDDATETextBox";
            this.eNDDATETextBox.Size = new System.Drawing.Size(100, 22);
            this.eNDDATETextBox.TabIndex = 14;
            // 
            // shipping_OQUT1BindingSource
            // 
            this.shipping_OQUT1BindingSource.DataMember = "Shipping_OQUT_Shipping_OQUT1";
            this.shipping_OQUT1BindingSource.DataSource = this.shipping_OQUTBindingSource;
            // 
            // shipping_OQUT1TableAdapter
            // 
            this.shipping_OQUT1TableAdapter.ClearBeforeFill = true;
            // 
            // shipping_OQUT1DataGridView
            // 
            this.shipping_OQUT1DataGridView.AutoGenerateColumns = false;
            this.shipping_OQUT1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.shipping_OQUT1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.QTYPE});
            this.shipping_OQUT1DataGridView.DataSource = this.shipping_OQUT1BindingSource;
            this.shipping_OQUT1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.shipping_OQUT1DataGridView.Location = new System.Drawing.Point(3, 3);
            this.shipping_OQUT1DataGridView.Name = "shipping_OQUT1DataGridView";
            this.shipping_OQUT1DataGridView.RowTemplate.Height = 24;
            this.shipping_OQUT1DataGridView.Size = new System.Drawing.Size(949, 252);
            this.shipping_OQUT1DataGridView.TabIndex = 16;
            this.shipping_OQUT1DataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.shipping_OQUT1DataGridView_CellValueChanged);
            // 
            // QTYPE
            // 
            this.QTYPE.DataPropertyName = "QTYPE";
            this.QTYPE.HeaderText = "成本結構";
            this.QTYPE.Items.AddRange(new object[] {
            "出口地Local Charge",
            "目的地Local Charge",
            "卡車派送"});
            this.QTYPE.Name = "QTYPE";
            this.QTYPE.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.QTYPE.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.QTYPE.Width = 200;
            // 
            // createNameTextBox
            // 
            this.createNameTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "CreateName", true));
            this.createNameTextBox.Location = new System.Drawing.Point(637, 6);
            this.createNameTextBox.Name = "createNameTextBox";
            this.createNameTextBox.Size = new System.Drawing.Size(100, 22);
            this.createNameTextBox.TabIndex = 17;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(592, 131);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(100, 52);
            this.button3.TabIndex = 59;
            this.button3.Text = "新項目編號維護";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(592, 37);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 51);
            this.button1.TabIndex = 64;
            this.button1.Text = "新供應商維護";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // cARDCODETextBox
            // 
            this.cARDCODETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "CARDCODE", true));
            this.cARDCODETextBox.Location = new System.Drawing.Point(76, 53);
            this.cARDCODETextBox.Name = "cARDCODETextBox";
            this.cARDCODETextBox.Size = new System.Drawing.Size(86, 22);
            this.cARDCODETextBox.TabIndex = 67;
            // 
            // tRADETextBox
            // 
            this.tRADETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "TRADE", true));
            this.tRADETextBox.Location = new System.Drawing.Point(228, 97);
            this.tRADETextBox.Name = "tRADETextBox";
            this.tRADETextBox.Size = new System.Drawing.Size(185, 22);
            this.tRADETextBox.TabIndex = 72;
            this.tRADETextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tRADETextBox_KeyDown);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(963, 284);
            this.tabControl1.TabIndex = 73;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.shipping_OQUT1DataGridView);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(955, 258);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "成本結構";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.AutoScroll = true;
            this.tabPage2.Controls.Add(this.button25);
            this.tabPage2.Controls.Add(this.shipping_OQUTDownloadDataGridView);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(955, 258);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "附件上傳";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button25
            // 
            this.button25.Image = ((System.Drawing.Image)(resources.GetObject("button25.Image")));
            this.button25.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button25.Location = new System.Drawing.Point(473, 6);
            this.button25.Name = "button25";
            this.button25.Size = new System.Drawing.Size(99, 44);
            this.button25.TabIndex = 4;
            this.button25.Text = "上傳檔案";
            this.button25.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button25.UseVisualStyleBackColor = true;
            this.button25.Click += new System.EventHandler(this.button25_Click);
            // 
            // shipping_OQUTDownloadDataGridView
            // 
            this.shipping_OQUTDownloadDataGridView.AllowUserToAddRows = false;
            this.shipping_OQUTDownloadDataGridView.AutoGenerateColumns = false;
            this.shipping_OQUTDownloadDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.shipping_OQUTDownloadDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.shipping_OQUTDownloadDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.LINK});
            this.shipping_OQUTDownloadDataGridView.DataSource = this.shipping_OQUTDownloadBindingSource;
            this.shipping_OQUTDownloadDataGridView.Location = new System.Drawing.Point(6, 6);
            this.shipping_OQUTDownloadDataGridView.Name = "shipping_OQUTDownloadDataGridView";
            this.shipping_OQUTDownloadDataGridView.RowTemplate.Height = 24;
            this.shipping_OQUTDownloadDataGridView.Size = new System.Drawing.Size(461, 270);
            this.shipping_OQUTDownloadDataGridView.TabIndex = 0;
            this.shipping_OQUTDownloadDataGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.shipping_OQUTDownloadDataGridView_CellContentClick);
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "seq";
            this.dataGridViewTextBoxColumn4.HeaderText = "";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            this.dataGridViewTextBoxColumn4.Width = 19;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "filename";
            this.dataGridViewTextBoxColumn5.HeaderText = "檔名";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            this.dataGridViewTextBoxColumn5.Width = 54;
            // 
            // LINK
            // 
            this.LINK.HeaderText = "LINK";
            this.LINK.Name = "LINK";
            this.LINK.ReadOnly = true;
            this.LINK.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.LINK.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.LINK.Text = "讀取檔案";
            this.LINK.UseColumnTextForLinkValue = true;
            this.LINK.Width = 57;
            // 
            // shipping_OQUTDownloadBindingSource
            // 
            this.shipping_OQUTDownloadBindingSource.DataMember = "Shipping_OQUT_Shipping_OQUTDownload";
            this.shipping_OQUTDownloadBindingSource.DataSource = this.shipping_OQUTBindingSource;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.mEMOTextBox);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(955, 258);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "備註";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // mEMOTextBox
            // 
            this.mEMOTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "MEMO", true));
            this.mEMOTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mEMOTextBox.Location = new System.Drawing.Point(3, 3);
            this.mEMOTextBox.Multiline = true;
            this.mEMOTextBox.Name = "mEMOTextBox";
            this.mEMOTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.mEMOTextBox.Size = new System.Drawing.Size(949, 252);
            this.mEMOTextBox.TabIndex = 1;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.mAILTextBox);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(955, 258);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "信件內容";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // mAILTextBox
            // 
            this.mAILTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "MAIL", true));
            this.mAILTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mAILTextBox.Location = new System.Drawing.Point(3, 3);
            this.mAILTextBox.Multiline = true;
            this.mAILTextBox.Name = "mAILTextBox";
            this.mAILTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.mAILTextBox.Size = new System.Drawing.Size(949, 252);
            this.mAILTextBox.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tERMTextBox);
            this.panel2.Controls.Add(this.comboBox5);
            this.panel2.Controls.Add(this.sHIPWAYTextBox);
            this.panel2.Controls.Add(this.comboBox6);
            this.panel2.Controls.Add(this.button6);
            this.panel2.Controls.Add(this.button5);
            this.panel2.Controls.Add(this.button4);
            this.panel2.Controls.Add(this.button2);
            this.panel2.Controls.Add(label11);
            this.panel2.Controls.Add(this.textBox2);
            this.panel2.Controls.Add(this.button7);
            this.panel2.Controls.Add(tERMLabel);
            this.panel2.Controls.Add(shippingCodeLabel);
            this.panel2.Controls.Add(this.shippingCodeTextBox);
            this.panel2.Controls.Add(cARDCODELabel);
            this.panel2.Controls.Add(tRADELabel);
            this.panel2.Controls.Add(this.cARDNAMETextBox);
            this.panel2.Controls.Add(this.tRADETextBox);
            this.panel2.Controls.Add(this.iTEMCODETextBox);
            this.panel2.Controls.Add(iTEMCODELabel);
            this.panel2.Controls.Add(this.iTEMNAMETextBox);
            this.panel2.Controls.Add(this.cARDCODETextBox);
            this.panel2.Controls.Add(this.dOCDATETextBox);
            this.panel2.Controls.Add(dOCDATELabel);
            this.panel2.Controls.Add(this.eNDDATETextBox);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(eNDDATELabel);
            this.panel2.Controls.Add(sHIPWAYLabel);
            this.panel2.Controls.Add(this.button3);
            this.panel2.Controls.Add(this.createNameTextBox);
            this.panel2.Controls.Add(createNameLabel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(963, 219);
            this.panel2.TabIndex = 74;
            // 
            // tERMTextBox
            // 
            this.tERMTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "TERM", true));
            this.tERMTextBox.Location = new System.Drawing.Point(512, 97);
            this.tERMTextBox.Name = "tERMTextBox";
            this.tERMTextBox.Size = new System.Drawing.Size(100, 22);
            this.tERMTextBox.TabIndex = 96;
            // 
            // comboBox5
            // 
            this.comboBox5.FormattingEnabled = true;
            this.comboBox5.Location = new System.Drawing.Point(512, 97);
            this.comboBox5.Name = "comboBox5";
            this.comboBox5.Size = new System.Drawing.Size(121, 20);
            this.comboBox5.TabIndex = 95;
            this.comboBox5.SelectedIndexChanged += new System.EventHandler(this.comboBox5_SelectedIndexChanged);
            this.comboBox5.MouseClick += new System.Windows.Forms.MouseEventHandler(this.comboBox5_MouseClick);
            // 
            // sHIPWAYTextBox
            // 
            this.sHIPWAYTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.shipping_OQUTBindingSource, "SHIPWAY", true));
            this.sHIPWAYTextBox.Location = new System.Drawing.Point(73, 95);
            this.sHIPWAYTextBox.Name = "sHIPWAYTextBox";
            this.sHIPWAYTextBox.Size = new System.Drawing.Size(70, 22);
            this.sHIPWAYTextBox.TabIndex = 94;
            // 
            // comboBox6
            // 
            this.comboBox6.FormattingEnabled = true;
            this.comboBox6.Location = new System.Drawing.Point(72, 95);
            this.comboBox6.Name = "comboBox6";
            this.comboBox6.Size = new System.Drawing.Size(90, 20);
            this.comboBox6.TabIndex = 93;
            this.comboBox6.SelectedIndexChanged += new System.EventHandler(this.comboBox6_SelectedIndexChanged);
            this.comboBox6.MouseClick += new System.Windows.Forms.MouseEventHandler(this.comboBox6_MouseClick);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(434, 160);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(100, 23);
            this.button6.TabIndex = 92;
            this.button6.Text = "SAP項目編號";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(434, 131);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(100, 23);
            this.button5.TabIndex = 91;
            this.button5.Text = "新項目編號";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(434, 38);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(100, 23);
            this.button4.TabIndex = 90;
            this.button4.Text = "新供應商";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(434, 66);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(100, 23);
            this.button2.TabIndex = 89;
            this.button2.Text = "SAP供應商";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(76, 177);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(203, 22);
            this.textBox2.TabIndex = 85;
            // 
            // button7
            // 
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button7.Location = new System.Drawing.Point(285, 175);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(55, 22);
            this.button7.TabIndex = 84;
            this.button7.Text = "寄信";
            this.button7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.tabControl1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 219);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(963, 284);
            this.panel3.TabIndex = 75;
            // 
            // shipping_OQUTDownloadTableAdapter
            // 
            this.shipping_OQUTDownloadTableAdapter.ClearBeforeFill = true;
            // 
            // SQUT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(963, 564);
            this.Name = "SQUT";
            this.Text = "供應商報價系統";
            this.Load += new System.EventHandler(this.SQUT_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUT1DataGridView)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shipping_OQUTDownloadBindingSource)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship ship;
        private System.Windows.Forms.BindingSource shipping_OQUTBindingSource;
        private ACMEDataSet.shipTableAdapters.Shipping_OQUTTableAdapter shipping_OQUTTableAdapter;
        private System.Windows.Forms.TextBox shippingCodeTextBox;
        private System.Windows.Forms.TextBox eNDDATETextBox;
        private System.Windows.Forms.TextBox dOCDATETextBox;
        private System.Windows.Forms.TextBox iTEMNAMETextBox;
        private System.Windows.Forms.TextBox iTEMCODETextBox;
        private System.Windows.Forms.TextBox cARDNAMETextBox;
        private System.Windows.Forms.BindingSource shipping_OQUT1BindingSource;
        private ACMEDataSet.shipTableAdapters.Shipping_OQUT1TableAdapter shipping_OQUT1TableAdapter;
        private System.Windows.Forms.DataGridView shipping_OQUT1DataGridView;
        private System.Windows.Forms.TextBox createNameTextBox;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox cARDCODETextBox;
        private System.Windows.Forms.TextBox tRADETextBox;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.BindingSource shipping_OQUTDownloadBindingSource;
        private ACMEDataSet.shipTableAdapters.Shipping_OQUTDownloadTableAdapter shipping_OQUTDownloadTableAdapter;
        private System.Windows.Forms.DataGridView shipping_OQUTDownloadDataGridView;
        private System.Windows.Forms.Button button25;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewLinkColumn LINK;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TextBox mEMOTextBox;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.TextBox mAILTextBox;
        private System.Windows.Forms.DataGridViewComboBoxColumn QTYPE;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.ComboBox comboBox5;
        private System.Windows.Forms.TextBox sHIPWAYTextBox;
        private System.Windows.Forms.ComboBox comboBox6;
        private System.Windows.Forms.TextBox tERMTextBox;
    }
}
