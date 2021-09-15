
namespace ACME
{
    partial class ACCAPINVOICE
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ACCAPINVOICE));
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txbShipDateStart = new System.Windows.Forms.TextBox();
            this.txbShipDateEnd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCustNumber = new System.Windows.Forms.Button();
            this.btnCheckOPCH = new System.Windows.Forms.Button();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.cmbBU = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txbCardCode = new System.Windows.Forms.TextBox();
            this.txbDocDate = new System.Windows.Forms.TextBox();
            this.txbOriCurrencyAmount = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dgvAccApInvoice = new System.Windows.Forms.DataGridView();
            this.DocDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OPDNDocEntry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.por1BaseEntry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.oporDocEntry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CardCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CardName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.UnTax = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.VatSumSy = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DocTotalSy = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_acme_inv = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.shipdate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_PC_BSINV = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.u_acme_shipday = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Currency = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OriCurrencyAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_ACME_RATE1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.u_acme_lc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TaxIdNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnCancelCardCode = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAccApInvoice)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.btnCancelCardCode);
            this.panel1.Controls.Add(this.btnCustNumber);
            this.panel1.Controls.Add(this.btnCheckOPCH);
            this.panel1.Controls.Add(this.btnQuery);
            this.panel1.Controls.Add(this.btnImport);
            this.panel1.Controls.Add(this.cmbBU);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txbCardCode);
            this.panel1.Controls.Add(this.txbDocDate);
            this.panel1.Controls.Add(this.txbOriCurrencyAmount);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1171, 113);
            this.panel1.TabIndex = 0;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txbShipDateStart);
            this.groupBox1.Controls.Add(this.txbShipDateEnd);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.groupBox1.Location = new System.Drawing.Point(21, 9);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(330, 49);
            this.groupBox1.TabIndex = 58;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "收貨採購日期";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(20, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(22, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "起";
            // 
            // txbShipDateStart
            // 
            this.txbShipDateStart.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbShipDateStart.Location = new System.Drawing.Point(48, 17);
            this.txbShipDateStart.Name = "txbShipDateStart";
            this.txbShipDateStart.Size = new System.Drawing.Size(112, 23);
            this.txbShipDateStart.TabIndex = 0;
            // 
            // txbShipDateEnd
            // 
            this.txbShipDateEnd.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbShipDateEnd.Location = new System.Drawing.Point(204, 17);
            this.txbShipDateEnd.Name = "txbShipDateEnd";
            this.txbShipDateEnd.Size = new System.Drawing.Size(112, 23);
            this.txbShipDateEnd.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(176, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(22, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "訖";
            // 
            // btnCustNumber
            // 
            this.btnCustNumber.ForeColor = System.Drawing.SystemColors.ActiveBorder;
            this.btnCustNumber.Image = ((System.Drawing.Image)(resources.GetObject("btnCustNumber.Image")));
            this.btnCustNumber.Location = new System.Drawing.Point(353, 64);
            this.btnCustNumber.Name = "btnCustNumber";
            this.btnCustNumber.Size = new System.Drawing.Size(32, 23);
            this.btnCustNumber.TabIndex = 57;
            this.btnCustNumber.Text = "y";
            this.btnCustNumber.UseVisualStyleBackColor = true;
            this.btnCustNumber.Click += new System.EventHandler(this.btnCustNumber_Click);
            // 
            // btnCheckOPCH
            // 
            this.btnCheckOPCH.Location = new System.Drawing.Point(1059, 11);
            this.btnCheckOPCH.Name = "btnCheckOPCH";
            this.btnCheckOPCH.Size = new System.Drawing.Size(66, 35);
            this.btnCheckOPCH.TabIndex = 4;
            this.btnCheckOPCH.Text = "檢核";
            this.btnCheckOPCH.UseVisualStyleBackColor = true;
            this.btnCheckOPCH.Visible = false;
            // 
            // btnQuery
            // 
            this.btnQuery.Location = new System.Drawing.Point(987, 52);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(66, 35);
            this.btnQuery.TabIndex = 4;
            this.btnQuery.Text = "查詢";
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(1059, 52);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(77, 35);
            this.btnImport.TabIndex = 4;
            this.btnImport.Text = "轉入AP發票";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // cmbBU
            // 
            this.cmbBU.Font = new System.Drawing.Font("新細明體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cmbBU.FormattingEnabled = true;
            this.cmbBU.Items.AddRange(new object[] {
            "AUO全部",
            "ADP全部",
            "ADP+AUO全部"});
            this.cmbBU.Location = new System.Drawing.Point(553, 56);
            this.cmbBU.Name = "cmbBU";
            this.cmbBU.Size = new System.Drawing.Size(112, 23);
            this.cmbBU.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("新細明體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(693, 27);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 15);
            this.label4.TabIndex = 2;
            this.label4.Text = "原幣金額加總";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("新細明體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(18, 67);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(67, 15);
            this.label6.TabIndex = 2;
            this.label6.Text = "廠商編號";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("新細明體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(484, 27);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(67, 15);
            this.label5.TabIndex = 2;
            this.label5.Text = "過帳日期";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("新細明體", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(504, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 15);
            this.label3.TabIndex = 2;
            this.label3.Text = "BU";
            // 
            // txbCardCode
            // 
            this.txbCardCode.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbCardCode.Location = new System.Drawing.Point(91, 62);
            this.txbCardCode.Name = "txbCardCode";
            this.txbCardCode.Size = new System.Drawing.Size(246, 23);
            this.txbCardCode.TabIndex = 0;
            // 
            // txbDocDate
            // 
            this.txbDocDate.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbDocDate.Location = new System.Drawing.Point(553, 23);
            this.txbDocDate.Name = "txbDocDate";
            this.txbDocDate.Size = new System.Drawing.Size(112, 23);
            this.txbDocDate.TabIndex = 0;
            // 
            // txbOriCurrencyAmount
            // 
            this.txbOriCurrencyAmount.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbOriCurrencyAmount.Location = new System.Drawing.Point(796, 24);
            this.txbOriCurrencyAmount.Name = "txbOriCurrencyAmount";
            this.txbOriCurrencyAmount.Size = new System.Drawing.Size(145, 23);
            this.txbOriCurrencyAmount.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dgvAccApInvoice);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 113);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1171, 503);
            this.panel2.TabIndex = 1;
            // 
            // dgvAccApInvoice
            // 
            this.dgvAccApInvoice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAccApInvoice.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DocDate,
            this.OPDNDocEntry,
            this.por1BaseEntry,
            this.oporDocEntry,
            this.CardCode,
            this.CardName,
            this.Quantity,
            this.UnTax,
            this.VatSumSy,
            this.DocTotalSy,
            this.U_acme_inv,
            this.shipdate,
            this.U_PC_BSINV,
            this.u_acme_shipday,
            this.Currency,
            this.OriCurrencyAmount,
            this.U_ACME_RATE1,
            this.u_acme_lc,
            this.TaxIdNumber});
            this.dgvAccApInvoice.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvAccApInvoice.Location = new System.Drawing.Point(0, 0);
            this.dgvAccApInvoice.Name = "dgvAccApInvoice";
            this.dgvAccApInvoice.RowTemplate.Height = 24;
            this.dgvAccApInvoice.Size = new System.Drawing.Size(1171, 503);
            this.dgvAccApInvoice.TabIndex = 0;
            // 
            // DocDate
            // 
            this.DocDate.DataPropertyName = "DocDate";
            this.DocDate.HeaderText = "過帳日期";
            this.DocDate.Name = "DocDate";
            // 
            // OPDNDocEntry
            // 
            this.OPDNDocEntry.DataPropertyName = "OPDNDocEntry";
            this.OPDNDocEntry.HeaderText = "收貨採購單單號";
            this.OPDNDocEntry.Name = "OPDNDocEntry";
            // 
            // por1BaseEntry
            // 
            this.por1BaseEntry.DataPropertyName = "pdn1BaseEntry";
            this.por1BaseEntry.HeaderText = "採購報價";
            this.por1BaseEntry.Name = "por1BaseEntry";
            // 
            // oporDocEntry
            // 
            this.oporDocEntry.DataPropertyName = "por1Docentry";
            this.oporDocEntry.HeaderText = "採購單";
            this.oporDocEntry.Name = "oporDocEntry";
            // 
            // CardCode
            // 
            this.CardCode.DataPropertyName = "CardCode";
            this.CardCode.HeaderText = "廠商編號";
            this.CardCode.Name = "CardCode";
            // 
            // CardName
            // 
            this.CardName.DataPropertyName = "CardName";
            this.CardName.HeaderText = "廠商名稱";
            this.CardName.Name = "CardName";
            // 
            // Quantity
            // 
            this.Quantity.DataPropertyName = "Quantity";
            this.Quantity.HeaderText = "數量";
            this.Quantity.Name = "Quantity";
            // 
            // UnTax
            // 
            this.UnTax.DataPropertyName = "UnTax";
            this.UnTax.HeaderText = "未稅總計";
            this.UnTax.Name = "UnTax";
            // 
            // VatSumSy
            // 
            this.VatSumSy.DataPropertyName = "VatSumSy";
            this.VatSumSy.HeaderText = "稅額";
            this.VatSumSy.Name = "VatSumSy";
            // 
            // DocTotalSy
            // 
            this.DocTotalSy.DataPropertyName = "DocTotalSy";
            this.DocTotalSy.HeaderText = "總計";
            this.DocTotalSy.Name = "DocTotalSy";
            // 
            // U_acme_inv
            // 
            this.U_acme_inv.DataPropertyName = "U_acme_inv";
            this.U_acme_inv.HeaderText = "InvoiceNo";
            this.U_acme_inv.Name = "U_acme_inv";
            // 
            // shipdate
            // 
            this.shipdate.DataPropertyName = "shipdate";
            this.shipdate.HeaderText = "日期";
            this.shipdate.Name = "shipdate";
            // 
            // U_PC_BSINV
            // 
            this.U_PC_BSINV.DataPropertyName = "U_PC_BSINV";
            this.U_PC_BSINV.HeaderText = "收採發票號碼";
            this.U_PC_BSINV.Name = "U_PC_BSINV";
            // 
            // u_acme_shipday
            // 
            this.u_acme_shipday.DataPropertyName = "u_acme_shipday";
            this.u_acme_shipday.HeaderText = "發票日期";
            this.u_acme_shipday.Name = "u_acme_shipday";
            // 
            // Currency
            // 
            this.Currency.DataPropertyName = "Currency";
            this.Currency.HeaderText = "原始幣別";
            this.Currency.Name = "Currency";
            // 
            // OriCurrencyAmount
            // 
            this.OriCurrencyAmount.DataPropertyName = "OriCurrencyAmount";
            this.OriCurrencyAmount.HeaderText = "原幣金額";
            this.OriCurrencyAmount.Name = "OriCurrencyAmount";
            // 
            // U_ACME_RATE1
            // 
            this.U_ACME_RATE1.DataPropertyName = "U_ACME_RATE1";
            this.U_ACME_RATE1.HeaderText = "匯率";
            this.U_ACME_RATE1.Name = "U_ACME_RATE1";
            // 
            // u_acme_lc
            // 
            this.u_acme_lc.DataPropertyName = "u_acme_lc";
            this.u_acme_lc.HeaderText = "LC";
            this.u_acme_lc.Name = "u_acme_lc";
            // 
            // TaxIdNumber
            // 
            this.TaxIdNumber.DataPropertyName = "TaxIdNumber";
            this.TaxIdNumber.HeaderText = "統一編號";
            this.TaxIdNumber.Name = "TaxIdNumber";
            // 
            // btnCancelCardCode
            // 
            this.btnCancelCardCode.ForeColor = System.Drawing.SystemColors.ActiveBorder;
            this.btnCancelCardCode.Image = global::ACME.Properties.Resources.bnCancelEdit_Image;
            this.btnCancelCardCode.Location = new System.Drawing.Point(391, 64);
            this.btnCancelCardCode.Name = "btnCancelCardCode";
            this.btnCancelCardCode.Size = new System.Drawing.Size(32, 23);
            this.btnCancelCardCode.TabIndex = 57;
            this.btnCancelCardCode.Text = "y";
            this.btnCancelCardCode.UseVisualStyleBackColor = true;
            this.btnCancelCardCode.Click += new System.EventHandler(this.btnCancelCardCode_Click);
            // 
            // ACCAPINVOICE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1171, 616);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "ACCAPINVOICE";
            this.Text = "ACCAPINVOICE";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAccApInvoice)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txbDocDate;
        private System.Windows.Forms.TextBox txbShipDateStart;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dgvAccApInvoice;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.ComboBox cmbBU;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txbOriCurrencyAmount;
        private System.Windows.Forms.TextBox txbShipDateEnd;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnCustNumber;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txbCardCode;
        private System.Windows.Forms.Button btnCheckOPCH;
        private System.Windows.Forms.DataGridViewTextBoxColumn DocDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn OPDNDocEntry;
        private System.Windows.Forms.DataGridViewTextBoxColumn por1BaseEntry;
        private System.Windows.Forms.DataGridViewTextBoxColumn oporDocEntry;
        private System.Windows.Forms.DataGridViewTextBoxColumn CardCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn CardName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Quantity;
        private System.Windows.Forms.DataGridViewTextBoxColumn UnTax;
        private System.Windows.Forms.DataGridViewTextBoxColumn VatSumSy;
        private System.Windows.Forms.DataGridViewTextBoxColumn DocTotalSy;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_acme_inv;
        private System.Windows.Forms.DataGridViewTextBoxColumn shipdate;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_PC_BSINV;
        private System.Windows.Forms.DataGridViewTextBoxColumn u_acme_shipday;
        private System.Windows.Forms.DataGridViewTextBoxColumn Currency;
        private System.Windows.Forms.DataGridViewTextBoxColumn OriCurrencyAmount;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_ACME_RATE1;
        private System.Windows.Forms.DataGridViewTextBoxColumn u_acme_lc;
        private System.Windows.Forms.DataGridViewTextBoxColumn TaxIdNumber;
        private System.Windows.Forms.Button btnCancelCardCode;
    }
}