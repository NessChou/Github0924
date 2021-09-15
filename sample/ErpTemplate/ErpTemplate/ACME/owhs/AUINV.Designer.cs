namespace ACME
{
    partial class AUINV
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.Label a2Label;
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ShippingCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FrgnName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SeqNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ItemCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BoxCheck = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ShipDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.A5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.a22 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.A3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.A7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DeCust = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.A6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.download = new System.Windows.Forms.DataGridViewLinkColumn();
            this.結案 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.panel5 = new System.Windows.Forms.Panel();
            this.btnExcelImport = new System.Windows.Forms.Button();
            this.btnExportExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.tSHNO = new System.Windows.Forms.TextBox();
            this.tWHNO = new System.Windows.Forms.TextBox();
            this.tSDATE = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tEDATE = new System.Windows.Forms.TextBox();
            this.cVER = new System.Windows.Forms.ComboBox();
            this.tINV = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cMODEL = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel6 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dgvFile = new System.Windows.Forms.DataGridView();
            this.ta2 = new System.Windows.Forms.TextBox();
            a2Label = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFile)).BeginInit();
            this.SuspendLayout();
            // 
            // a2Label
            // 
            a2Label.AutoSize = true;
            a2Label.Location = new System.Drawing.Point(3, 9);
            a2Label.Name = "a2Label";
            a2Label.Size = new System.Drawing.Size(89, 12);
            a2Label.TabIndex = 4;
            a2Label.Text = "異常箱號及情況";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(227, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(40, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(173, 8);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(48, 16);
            this.checkBox1.TabIndex = 1;
            this.checkBox1.Text = "未結";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ShippingCode,
            this.FrgnName,
            this.SeqNo,
            this.ItemCode,
            this.BoxCheck,
            this.ShipDate,
            this.A5,
            this.a22,
            this.A3,
            this.A7,
            this.DeCust,
            this.A6,
            this.download,
            this.結案});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1127, 586);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentDoubleClick);
            // 
            // ShippingCode
            // 
            this.ShippingCode.DataPropertyName = "WH工單";
            this.ShippingCode.HeaderText = "WH工單";
            this.ShippingCode.Name = "ShippingCode";
            this.ShippingCode.Width = 57;
            // 
            // FrgnName
            // 
            this.FrgnName.DataPropertyName = "SI工單";
            this.FrgnName.HeaderText = "SI工單";
            this.FrgnName.Name = "FrgnName";
            this.FrgnName.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.FrgnName.Width = 60;
            // 
            // SeqNo
            // 
            this.SeqNo.DataPropertyName = "序號";
            this.SeqNo.HeaderText = "序號";
            this.SeqNo.Name = "SeqNo";
            this.SeqNo.Visible = false;
            this.SeqNo.Width = 54;
            // 
            // ItemCode
            // 
            this.ItemCode.DataPropertyName = "產品編號";
            this.ItemCode.HeaderText = "產品編號";
            this.ItemCode.Name = "ItemCode";
            this.ItemCode.Width = 61;
            // 
            // BoxCheck
            // 
            this.BoxCheck.DataPropertyName = "原廠INVOICE";
            this.BoxCheck.HeaderText = "原廠Invoice";
            this.BoxCheck.Name = "BoxCheck";
            this.BoxCheck.Width = 82;
            // 
            // ShipDate
            // 
            this.ShipDate.DataPropertyName = "異常數量";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.ShipDate.DefaultCellStyle = dataGridViewCellStyle1;
            this.ShipDate.HeaderText = "異常數量";
            this.ShipDate.Name = "ShipDate";
            this.ShipDate.Width = 61;
            // 
            // A5
            // 
            this.A5.DataPropertyName = "進貨日期";
            this.A5.HeaderText = "進貨日期";
            this.A5.Name = "A5";
            this.A5.Width = 61;
            // 
            // a22
            // 
            this.a22.DataPropertyName = "a2";
            this.a22.HeaderText = "異常情況";
            this.a22.Name = "a22";
            this.a22.Visible = false;
            this.a22.Width = 78;
            // 
            // A3
            // 
            this.A3.DataPropertyName = "後續處理情形";
            this.A3.HeaderText = "後續處理情形";
            this.A3.Name = "A3";
            this.A3.Width = 72;
            // 
            // A7
            // 
            this.A7.DataPropertyName = "SHIPPING備註";
            this.A7.HeaderText = "Shipping備註";
            this.A7.Name = "A7";
            this.A7.Width = 78;
            // 
            // DeCust
            // 
            this.DeCust.DataPropertyName = "客戶別";
            this.DeCust.HeaderText = "客戶別";
            this.DeCust.Name = "DeCust";
            this.DeCust.Visible = false;
            this.DeCust.Width = 61;
            // 
            // A6
            // 
            this.A6.DataPropertyName = "運送貨代";
            this.A6.HeaderText = "運送貨代";
            this.A6.Name = "A6";
            this.A6.Width = 61;
            // 
            // download
            // 
            this.download.DataPropertyName = "下載";
            this.download.HeaderText = "下載";
            this.download.Name = "download";
            this.download.ReadOnly = true;
            this.download.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.download.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.download.Text = "下載檔案";
            this.download.Width = 51;
            // 
            // 結案
            // 
            this.結案.DataPropertyName = "結案";
            this.結案.FalseValue = "false";
            this.結案.HeaderText = "結案";
            this.結案.IndeterminateValue = "false";
            this.結案.Name = "結案";
            this.結案.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.結案.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.結案.TrueValue = "true";
            this.結案.Width = 51;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel7);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1370, 32);
            this.panel1.TabIndex = 3;
            // 
            // panel7
            // 
            this.panel7.Controls.Add(this.panel5);
            this.panel7.Controls.Add(this.panel6);
            this.panel7.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel7.Location = new System.Drawing.Point(0, 0);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(1370, 32);
            this.panel7.TabIndex = 0;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.btnExcelImport);
            this.panel5.Controls.Add(this.btnExportExcel);
            this.panel5.Controls.Add(this.label1);
            this.panel5.Controls.Add(this.button2);
            this.panel5.Controls.Add(this.checkBox1);
            this.panel5.Controls.Add(this.tSHNO);
            this.panel5.Controls.Add(this.button1);
            this.panel5.Controls.Add(this.tWHNO);
            this.panel5.Controls.Add(this.tSDATE);
            this.panel5.Controls.Add(this.label7);
            this.panel5.Controls.Add(this.label2);
            this.panel5.Controls.Add(this.label6);
            this.panel5.Controls.Add(this.tEDATE);
            this.panel5.Controls.Add(this.cVER);
            this.panel5.Controls.Add(this.tINV);
            this.panel5.Controls.Add(this.label5);
            this.panel5.Controls.Add(this.label3);
            this.panel5.Controls.Add(this.cMODEL);
            this.panel5.Controls.Add(this.label4);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(0, 0);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(1127, 32);
            this.panel5.TabIndex = 3;
            // 
            // btnExcelImport
            // 
            this.btnExcelImport.Location = new System.Drawing.Point(320, 3);
            this.btnExcelImport.Name = "btnExcelImport";
            this.btnExcelImport.Size = new System.Drawing.Size(73, 23);
            this.btnExcelImport.TabIndex = 14;
            this.btnExcelImport.Text = "匯入Excel";
            this.btnExcelImport.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnExcelImport.UseVisualStyleBackColor = true;
            this.btnExcelImport.Click += new System.EventHandler(this.btnExcelImport_Click);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.Location = new System.Drawing.Point(399, 3);
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(73, 23);
            this.btnExportExcel.TabIndex = 14;
            this.btnExportExcel.Text = "匯出Excel";
            this.btnExportExcel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnExportExcel.UseVisualStyleBackColor = true;
            this.btnExportExcel.Click += new System.EventHandler(this.button3_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "日期";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(274, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(40, 23);
            this.button2.TabIndex = 13;
            this.button2.Text = "更新";
            this.button2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tSHNO
            // 
            this.tSHNO.Location = new System.Drawing.Point(1120, 6);
            this.tSHNO.Name = "tSHNO";
            this.tSHNO.Size = new System.Drawing.Size(100, 22);
            this.tSHNO.TabIndex = 12;
            // 
            // tWHNO
            // 
            this.tWHNO.Location = new System.Drawing.Point(965, 6);
            this.tWHNO.Name = "tWHNO";
            this.tWHNO.Size = new System.Drawing.Size(100, 22);
            this.tWHNO.TabIndex = 11;
            // 
            // tSDATE
            // 
            this.tSDATE.Location = new System.Drawing.Point(38, 3);
            this.tSDATE.MaxLength = 8;
            this.tSDATE.Name = "tSDATE";
            this.tSDATE.Size = new System.Drawing.Size(55, 22);
            this.tSDATE.TabIndex = 3;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(1071, 12);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(39, 12);
            this.label7.TabIndex = 10;
            this.label7.Text = "SI工單";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(95, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(11, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "~";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(911, 12);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 12);
            this.label6.TabIndex = 3;
            this.label6.Text = "WH工單";
            // 
            // tEDATE
            // 
            this.tEDATE.Location = new System.Drawing.Point(112, 3);
            this.tEDATE.MaxLength = 8;
            this.tEDATE.Name = "tEDATE";
            this.tEDATE.Size = new System.Drawing.Size(55, 22);
            this.tEDATE.TabIndex = 5;
            // 
            // cVER
            // 
            this.cVER.FormattingEnabled = true;
            this.cVER.Location = new System.Drawing.Point(822, 8);
            this.cVER.Name = "cVER";
            this.cVER.Size = new System.Drawing.Size(71, 20);
            this.cVER.TabIndex = 9;
            // 
            // tINV
            // 
            this.tINV.Location = new System.Drawing.Point(542, 6);
            this.tINV.Name = "tINV";
            this.tINV.Size = new System.Drawing.Size(100, 22);
            this.tINV.TabIndex = 3;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(787, 12);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 12);
            this.label5.TabIndex = 8;
            this.label5.Text = "版本";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(479, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "Invoice No";
            // 
            // cMODEL
            // 
            this.cMODEL.FormattingEnabled = true;
            this.cMODEL.Location = new System.Drawing.Point(683, 8);
            this.cMODEL.Name = "cMODEL";
            this.cMODEL.Size = new System.Drawing.Size(95, 20);
            this.cMODEL.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(648, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "型號";
            // 
            // panel6
            // 
            this.panel6.Controls.Add(a2Label);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel6.Location = new System.Drawing.Point(1127, 0);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(243, 32);
            this.panel6.TabIndex = 5;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 32);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1370, 586);
            this.panel2.TabIndex = 4;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.dataGridView1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1127, 586);
            this.panel4.TabIndex = 5;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dgvFile);
            this.panel3.Controls.Add(this.ta2);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel3.Location = new System.Drawing.Point(1127, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(243, 586);
            this.panel3.TabIndex = 4;
            // 
            // dgvFile
            // 
            this.dgvFile.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFile.Location = new System.Drawing.Point(0, 245);
            this.dgvFile.Name = "dgvFile";
            this.dgvFile.RowTemplate.Height = 24;
            this.dgvFile.Size = new System.Drawing.Size(240, 341);
            this.dgvFile.TabIndex = 4;
            this.dgvFile.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvFile_CellContentClick);
            // 
            // ta2
            // 
            this.ta2.Location = new System.Drawing.Point(0, 0);
            this.ta2.Multiline = true;
            this.ta2.Name = "ta2";
            this.ta2.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.ta2.Size = new System.Drawing.Size(240, 244);
            this.ta2.TabIndex = 3;
            // 
            // AUINV
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1370, 618);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "AUINV";
            this.Text = "收貨異常";
            this.Load += new System.EventHandler(this.AUINV_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFile)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TextBox ta2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tSDATE;
        private System.Windows.Forms.TextBox tEDATE;
        private System.Windows.Forms.TextBox tINV;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cVER;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cMODEL;
        private System.Windows.Forms.TextBox tSHNO;
        private System.Windows.Forms.TextBox tWHNO;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Button btnExportExcel;
        private System.Windows.Forms.Button btnExcelImport;
        private System.Windows.Forms.DataGridView dgvFile;
        private System.Windows.Forms.DataGridViewTextBoxColumn ShippingCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn FrgnName;
        private System.Windows.Forms.DataGridViewTextBoxColumn SeqNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn ItemCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn BoxCheck;
        private System.Windows.Forms.DataGridViewTextBoxColumn ShipDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn A5;
        private System.Windows.Forms.DataGridViewTextBoxColumn a22;
        private System.Windows.Forms.DataGridViewTextBoxColumn A3;
        private System.Windows.Forms.DataGridViewTextBoxColumn A7;
        private System.Windows.Forms.DataGridViewTextBoxColumn DeCust;
        private System.Windows.Forms.DataGridViewTextBoxColumn A6;
        private System.Windows.Forms.DataGridViewLinkColumn download;
        private System.Windows.Forms.DataGridViewCheckBoxColumn 結案;
    }
}