namespace ACME
{
    partial class SOLAROPCH
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
            System.Windows.Forms.Label createNameLabel;
            System.Windows.Forms.Label dOCDATELabel;
            System.Windows.Forms.Label label1;
            System.Windows.Forms.Label label2;
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.sOLAR = new ACME.ACMEDataSet.SOLAR();
            this.sOLAR_OPCHBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.sOLAR_OPCHTableAdapter = new ACME.ACMEDataSet.SOLARTableAdapters.SOLAR_OPCHTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.SOLARTableAdapters.TableAdapterManager();
            this.shippingCodeTextBox = new System.Windows.Forms.TextBox();
            this.createNameTextBox = new System.Windows.Forms.TextBox();
            this.dOCDATETextBox = new System.Windows.Forms.TextBox();
            this.sOLAR_OPCH1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.sOLAR_OPCH1TableAdapter = new ACME.ACMEDataSet.SOLARTableAdapters.SOLAR_OPCH1TableAdapter();
            this.sOLAR_OPCH1DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LINENUM = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel6 = new System.Windows.Forms.Panel();
            this.sOLAR_OPCH2DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DOCENTRY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PATH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FILENAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sOLAR_OPCH2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.panel5 = new System.Windows.Forms.Panel();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button2 = new System.Windows.Forms.Button();
            this.sOLAR_OPCH2TableAdapter = new ACME.ACMEDataSet.SOLARTableAdapters.SOLAR_OPCH2TableAdapter();
            this.contextMenuStrip3 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            shippingCodeLabel = new System.Windows.Forms.Label();
            createNameLabel = new System.Windows.Forms.Label();
            dOCDATELabel = new System.Windows.Forms.Label();
            label1 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCHBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH1DataGridView)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel4.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH2DataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH2BindingSource)).BeginInit();
            this.panel5.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.contextMenuStrip3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Size = new System.Drawing.Size(999, 471);
            this.panel1.Controls.SetChildIndex(this.panel2, 0);
            // 
            // shippingCodeLabel
            // 
            shippingCodeLabel.AutoSize = true;
            shippingCodeLabel.Location = new System.Drawing.Point(12, 24);
            shippingCodeLabel.Name = "shippingCodeLabel";
            shippingCodeLabel.Size = new System.Drawing.Size(17, 12);
            shippingCodeLabel.TabIndex = 1;
            shippingCodeLabel.Text = "ID";
            // 
            // createNameLabel
            // 
            createNameLabel.AutoSize = true;
            createNameLabel.Location = new System.Drawing.Point(240, 24);
            createNameLabel.Name = "createNameLabel";
            createNameLabel.Size = new System.Drawing.Size(41, 12);
            createNameLabel.TabIndex = 3;
            createNameLabel.Text = "製單人";
            // 
            // dOCDATELabel
            // 
            dOCDATELabel.AutoSize = true;
            dOCDATELabel.Location = new System.Drawing.Point(435, 24);
            dOCDATELabel.Name = "dOCDATELabel";
            dOCDATELabel.Size = new System.Drawing.Size(53, 12);
            dOCDATELabel.TabIndex = 5;
            dOCDATELabel.Text = "製單日期";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(12, 68);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(53, 12);
            label1.TabIndex = 10;
            label1.Text = "採購單號";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(17, 7);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(53, 12);
            label2.TabIndex = 14;
            label2.Text = "附件SIZE";
            // 
            // sOLAR
            // 
            this.sOLAR.DataSetName = "SOLAR";
            this.sOLAR.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // sOLAR_OPCHBindingSource
            // 
            this.sOLAR_OPCHBindingSource.DataMember = "SOLAR_OPCH";
            this.sOLAR_OPCHBindingSource.DataSource = this.sOLAR;
            // 
            // sOLAR_OPCHTableAdapter
            // 
            this.sOLAR_OPCHTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.SOLAR_OPCH1TableAdapter = null;
            this.tableAdapterManager.SOLAR_OPCH2TableAdapter = null;
            this.tableAdapterManager.SOLAR_OPCHTableAdapter = this.sOLAR_OPCHTableAdapter;
            this.tableAdapterManager.SOLAR_PAY1TableAdapter = null;
            this.tableAdapterManager.SOLAR_PAYDownloadTableAdapter = null;
            this.tableAdapterManager.SOLAR_PAYTableAdapter = null;
            this.tableAdapterManager.SOLAR_PROBOM2TableAdapter = null;
            this.tableAdapterManager.SOLAR_PROBOMDownloadTableAdapter = null;
            this.tableAdapterManager.SOLAR_PROBOMTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.SOLARTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // shippingCodeTextBox
            // 
            this.shippingCodeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sOLAR_OPCHBindingSource, "ShippingCode", true));
            this.shippingCodeTextBox.Location = new System.Drawing.Point(82, 21);
            this.shippingCodeTextBox.Name = "shippingCodeTextBox";
            this.shippingCodeTextBox.Size = new System.Drawing.Size(100, 22);
            this.shippingCodeTextBox.TabIndex = 2;
            // 
            // createNameTextBox
            // 
            this.createNameTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sOLAR_OPCHBindingSource, "CreateName", true));
            this.createNameTextBox.Location = new System.Drawing.Point(314, 21);
            this.createNameTextBox.Name = "createNameTextBox";
            this.createNameTextBox.Size = new System.Drawing.Size(100, 22);
            this.createNameTextBox.TabIndex = 4;
            // 
            // dOCDATETextBox
            // 
            this.dOCDATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sOLAR_OPCHBindingSource, "DOCDATE", true));
            this.dOCDATETextBox.Location = new System.Drawing.Point(503, 21);
            this.dOCDATETextBox.Name = "dOCDATETextBox";
            this.dOCDATETextBox.Size = new System.Drawing.Size(100, 22);
            this.dOCDATETextBox.TabIndex = 6;
            // 
            // sOLAR_OPCH1BindingSource
            // 
            this.sOLAR_OPCH1BindingSource.DataMember = "SOLAR_OPCH_SOLAR_OPCH1";
            this.sOLAR_OPCH1BindingSource.DataSource = this.sOLAR_OPCHBindingSource;
            // 
            // sOLAR_OPCH1TableAdapter
            // 
            this.sOLAR_OPCH1TableAdapter.ClearBeforeFill = true;
            // 
            // sOLAR_OPCH1DataGridView
            // 
            this.sOLAR_OPCH1DataGridView.AutoGenerateColumns = false;
            this.sOLAR_OPCH1DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.sOLAR_OPCH1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sOLAR_OPCH1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7,
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn11,
            this.LINENUM});
            this.sOLAR_OPCH1DataGridView.DataSource = this.sOLAR_OPCH1BindingSource;
            this.sOLAR_OPCH1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sOLAR_OPCH1DataGridView.Location = new System.Drawing.Point(3, 3);
            this.sOLAR_OPCH1DataGridView.Name = "sOLAR_OPCH1DataGridView";
            this.sOLAR_OPCH1DataGridView.RowTemplate.Height = 24;
            this.sOLAR_OPCH1DataGridView.Size = new System.Drawing.Size(985, 323);
            this.sOLAR_OPCH1DataGridView.TabIndex = 7;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "NO";
            this.dataGridViewTextBoxColumn3.HeaderText = "NO";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 46;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "CARDNAME";
            this.dataGridViewTextBoxColumn4.HeaderText = "廠商";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.Width = 54;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "ITEMCODE";
            this.dataGridViewTextBoxColumn5.HeaderText = "料號";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.Width = 54;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "QTY";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N2";
            this.dataGridViewTextBoxColumn6.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewTextBoxColumn6.HeaderText = "數量";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.Width = 54;
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "AMT";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            this.dataGridViewTextBoxColumn7.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewTextBoxColumn7.HeaderText = "總金額";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.Width = 66;
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.DataPropertyName = "PAYAMT";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle3.Format = "N0";
            this.dataGridViewTextBoxColumn8.DefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewTextBoxColumn8.HeaderText = "請款金額";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            this.dataGridViewTextBoxColumn8.Width = 78;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "PRJCODE";
            this.dataGridViewTextBoxColumn9.HeaderText = "專案";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.Width = 54;
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "MEMO";
            this.dataGridViewTextBoxColumn10.HeaderText = "備註";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            this.dataGridViewTextBoxColumn10.Width = 54;
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.DataPropertyName = "DOCENTRY";
            this.dataGridViewTextBoxColumn11.HeaderText = "採購單號";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            this.dataGridViewTextBoxColumn11.Width = 78;
            // 
            // LINENUM
            // 
            this.LINENUM.DataPropertyName = "LINENUM";
            this.LINENUM.HeaderText = "採購列號";
            this.LINENUM.Name = "LINENUM";
            this.LINENUM.Width = 78;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(188, 63);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(93, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "選取採購單號";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(82, 65);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 22);
            this.textBox1.TabIndex = 9;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(999, 449);
            this.panel2.TabIndex = 11;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.tabControl1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 94);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(999, 355);
            this.panel4.TabIndex = 12;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(999, 355);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.sOLAR_OPCH1DataGridView);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(991, 329);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "明細檔";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.AutoScroll = true;
            this.tabPage2.Controls.Add(this.panel6);
            this.tabPage2.Controls.Add(this.panel5);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(826, 306);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "附件";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.sOLAR_OPCH2DataGridView);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel6.Location = new System.Drawing.Point(3, 38);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(820, 265);
            this.panel6.TabIndex = 2;
            // 
            // sOLAR_OPCH2DataGridView
            // 
            this.sOLAR_OPCH2DataGridView.AutoGenerateColumns = false;
            this.sOLAR_OPCH2DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.sOLAR_OPCH2DataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.sOLAR_OPCH2DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sOLAR_OPCH2DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn12,
            this.DOCENTRY,
            this.PATH,
            this.FILENAME});
            this.sOLAR_OPCH2DataGridView.DataSource = this.sOLAR_OPCH2BindingSource;
            this.sOLAR_OPCH2DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sOLAR_OPCH2DataGridView.Location = new System.Drawing.Point(0, 0);
            this.sOLAR_OPCH2DataGridView.Name = "sOLAR_OPCH2DataGridView";
            this.sOLAR_OPCH2DataGridView.RowTemplate.Height = 24;
            this.sOLAR_OPCH2DataGridView.Size = new System.Drawing.Size(820, 265);
            this.sOLAR_OPCH2DataGridView.TabIndex = 0;
            // 
            // dataGridViewTextBoxColumn12
            // 
            this.dataGridViewTextBoxColumn12.DataPropertyName = "NO";
            this.dataGridViewTextBoxColumn12.HeaderText = "NO";
            this.dataGridViewTextBoxColumn12.Name = "dataGridViewTextBoxColumn12";
            this.dataGridViewTextBoxColumn12.Width = 46;
            // 
            // DOCENTRY
            // 
            this.DOCENTRY.DataPropertyName = "DOCENTRY";
            this.DOCENTRY.HeaderText = "採購單號";
            this.DOCENTRY.Name = "DOCENTRY";
            this.DOCENTRY.Width = 78;
            // 
            // PATH
            // 
            this.PATH.DataPropertyName = "PATH";
            this.PATH.HeaderText = "PATH";
            this.PATH.Name = "PATH";
            this.PATH.Visible = false;
            this.PATH.Width = 59;
            // 
            // FILENAME
            // 
            this.FILENAME.DataPropertyName = "FILENAME";
            this.FILENAME.HeaderText = "檔名";
            this.FILENAME.Name = "FILENAME";
            this.FILENAME.Width = 54;
            // 
            // sOLAR_OPCH2BindingSource
            // 
            this.sOLAR_OPCH2BindingSource.DataMember = "SOLAR_OPCH_SOLAR_OPCH2";
            this.sOLAR_OPCH2BindingSource.DataSource = this.sOLAR_OPCHBindingSource;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.textBox2);
            this.panel5.Controls.Add(label2);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel5.Location = new System.Drawing.Point(3, 3);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(820, 35);
            this.panel5.TabIndex = 1;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(76, 4);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(75, 22);
            this.textBox2.TabIndex = 15;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dataGridView1);
            this.panel3.Controls.Add(this.button2);
            this.panel3.Controls.Add(shippingCodeLabel);
            this.panel3.Controls.Add(dOCDATELabel);
            this.panel3.Controls.Add(this.shippingCodeTextBox);
            this.panel3.Controls.Add(this.button1);
            this.panel3.Controls.Add(this.createNameTextBox);
            this.panel3.Controls.Add(this.dOCDATETextBox);
            this.panel3.Controls.Add(label1);
            this.panel3.Controls.Add(this.textBox1);
            this.panel3.Controls.Add(createNameLabel);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(999, 94);
            this.panel3.TabIndex = 11;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(699, 21);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(43, 22);
            this.dataGridView1.TabIndex = 12;
            this.dataGridView1.Visible = false;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(287, 63);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(93, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "MAIL";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // sOLAR_OPCH2TableAdapter
            // 
            this.sOLAR_OPCH2TableAdapter.ClearBeforeFill = true;
            // 
            // contextMenuStrip3
            // 
            this.contextMenuStrip3.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.contextMenuStrip3.Name = "contextMenuStrip1";
            this.contextMenuStrip3.Size = new System.Drawing.Size(113, 26);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(112, 22);
            this.toolStripMenuItem1.Text = "複製列";
            // 
            // SOLAROPCH
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(999, 510);
            this.Name = "SOLAROPCH";
            this.Text = "進貨附加文件通知";
            this.Load += new System.EventHandler(this.SOLAROPCH_Load);
            this.Controls.SetChildIndex(this.panel1, 0);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCHBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH1DataGridView)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH2DataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sOLAR_OPCH2BindingSource)).EndInit();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.contextMenuStrip3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.SOLAR sOLAR;
        private System.Windows.Forms.BindingSource sOLAR_OPCHBindingSource;
        private ACMEDataSet.SOLARTableAdapters.SOLAR_OPCHTableAdapter sOLAR_OPCHTableAdapter;
        private ACMEDataSet.SOLARTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.TextBox shippingCodeTextBox;
        private System.Windows.Forms.TextBox dOCDATETextBox;
        private System.Windows.Forms.TextBox createNameTextBox;
        private System.Windows.Forms.BindingSource sOLAR_OPCH1BindingSource;
        private ACMEDataSet.SOLARTableAdapters.SOLAR_OPCH1TableAdapter sOLAR_OPCH1TableAdapter;
        private System.Windows.Forms.DataGridView sOLAR_OPCH1DataGridView;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.BindingSource sOLAR_OPCH2BindingSource;
        private ACMEDataSet.SOLARTableAdapters.SOLAR_OPCH2TableAdapter sOLAR_OPCH2TableAdapter;
        private System.Windows.Forms.DataGridView sOLAR_OPCH2DataGridView;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.TextBox textBox2;
        public System.Windows.Forms.ContextMenuStrip contextMenuStrip3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;
        private System.Windows.Forms.DataGridViewTextBoxColumn DOCENTRY;
        private System.Windows.Forms.DataGridViewTextBoxColumn PATH;
        private System.Windows.Forms.DataGridViewTextBoxColumn FILENAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
        private System.Windows.Forms.DataGridViewTextBoxColumn LINENUM;
    }
}
