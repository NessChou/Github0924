
namespace ACME
{
    partial class StockStatus
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.txbItemcode = new System.Windows.Forms.TextBox();
            this.cmbSize = new System.Windows.Forms.ComboBox();
            this.cmbType = new System.Windows.Forms.ComboBox();
            this.cmbModel = new System.Windows.Forms.ComboBox();
            this.cmbGrade = new System.Windows.Forms.ComboBox();
            this.cmbVersion = new System.Windows.Forms.ComboBox();
            this.cmbBU = new System.Windows.Forms.ComboBox();
            this.ckbOnHandGreatThenZero = new System.Windows.Forms.CheckBox();
            this.ckbZeroUnshow = new System.Windows.Forms.CheckBox();
            this.ckbOCTcon = new System.Windows.Forms.CheckBox();
            this.ckbUndeliverGreaterThenZero = new System.Windows.Forms.CheckBox();
            this.btnExcel = new System.Windows.Forms.Button();
            this.btnSort = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dgvStockStatus = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvStockStatus)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txbItemcode);
            this.panel1.Controls.Add(this.cmbSize);
            this.panel1.Controls.Add(this.cmbType);
            this.panel1.Controls.Add(this.cmbModel);
            this.panel1.Controls.Add(this.cmbGrade);
            this.panel1.Controls.Add(this.cmbVersion);
            this.panel1.Controls.Add(this.cmbBU);
            this.panel1.Controls.Add(this.ckbOnHandGreatThenZero);
            this.panel1.Controls.Add(this.ckbZeroUnshow);
            this.panel1.Controls.Add(this.ckbOCTcon);
            this.panel1.Controls.Add(this.ckbUndeliverGreaterThenZero);
            this.panel1.Controls.Add(this.btnExcel);
            this.panel1.Controls.Add(this.btnSort);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1522, 140);
            this.panel1.TabIndex = 6;
            // 
            // txbItemcode
            // 
            this.txbItemcode.Location = new System.Drawing.Point(777, 31);
            this.txbItemcode.Multiline = true;
            this.txbItemcode.Name = "txbItemcode";
            this.txbItemcode.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txbItemcode.Size = new System.Drawing.Size(177, 100);
            this.txbItemcode.TabIndex = 25;
            // 
            // cmbSize
            // 
            this.cmbSize.FormattingEnabled = true;
            this.cmbSize.Location = new System.Drawing.Point(69, 31);
            this.cmbSize.Name = "cmbSize";
            this.cmbSize.Size = new System.Drawing.Size(121, 20);
            this.cmbSize.TabIndex = 24;
            this.cmbSize.SelectedIndexChanged += new System.EventHandler(this.cmbSize_SelectedIndexChanged);
            this.cmbSize.TextChanged += new System.EventHandler(this.cmbSize_TextChanged);
            // 
            // cmbType
            // 
            this.cmbType.FormattingEnabled = true;
            this.cmbType.Location = new System.Drawing.Point(270, 31);
            this.cmbType.Name = "cmbType";
            this.cmbType.Size = new System.Drawing.Size(121, 20);
            this.cmbType.TabIndex = 24;
            this.cmbType.SelectedIndexChanged += new System.EventHandler(this.cmbType_SelectedIndexChanged);
            this.cmbType.TextChanged += new System.EventHandler(this.cmbType_TextChanged);
            // 
            // cmbModel
            // 
            this.cmbModel.FormattingEnabled = true;
            this.cmbModel.Location = new System.Drawing.Point(1215, 26);
            this.cmbModel.Name = "cmbModel";
            this.cmbModel.Size = new System.Drawing.Size(121, 20);
            this.cmbModel.TabIndex = 24;
            this.cmbModel.Visible = false;
            // 
            // cmbGrade
            // 
            this.cmbGrade.FormattingEnabled = true;
            this.cmbGrade.Location = new System.Drawing.Point(1028, 80);
            this.cmbGrade.Name = "cmbGrade";
            this.cmbGrade.Size = new System.Drawing.Size(121, 20);
            this.cmbGrade.TabIndex = 24;
            this.cmbGrade.Visible = false;
            // 
            // cmbVersion
            // 
            this.cmbVersion.FormattingEnabled = true;
            this.cmbVersion.Location = new System.Drawing.Point(1245, 80);
            this.cmbVersion.Name = "cmbVersion";
            this.cmbVersion.Size = new System.Drawing.Size(121, 20);
            this.cmbVersion.TabIndex = 24;
            this.cmbVersion.Visible = false;
            // 
            // cmbBU
            // 
            this.cmbBU.FormattingEnabled = true;
            this.cmbBU.Location = new System.Drawing.Point(455, 32);
            this.cmbBU.Name = "cmbBU";
            this.cmbBU.Size = new System.Drawing.Size(121, 20);
            this.cmbBU.TabIndex = 24;
            // 
            // ckbOnHandGreatThenZero
            // 
            this.ckbOnHandGreatThenZero.AutoSize = true;
            this.ckbOnHandGreatThenZero.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ckbOnHandGreatThenZero.Location = new System.Drawing.Point(296, 83);
            this.ckbOnHandGreatThenZero.Name = "ckbOnHandGreatThenZero";
            this.ckbOnHandGreatThenZero.Size = new System.Drawing.Size(111, 20);
            this.ckbOnHandGreatThenZero.TabIndex = 23;
            this.ckbOnHandGreatThenZero.Text = "現有數量大於零";
            this.ckbOnHandGreatThenZero.UseVisualStyleBackColor = true;
            // 
            // ckbZeroUnshow
            // 
            this.ckbZeroUnshow.AutoSize = true;
            this.ckbZeroUnshow.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ckbZeroUnshow.Location = new System.Drawing.Point(643, 82);
            this.ckbZeroUnshow.Name = "ckbZeroUnshow";
            this.ckbZeroUnshow.Size = new System.Drawing.Size(87, 20);
            this.ckbZeroUnshow.TabIndex = 22;
            this.ckbZeroUnshow.Text = "零值不顯示";
            this.ckbZeroUnshow.UseVisualStyleBackColor = true;
            // 
            // ckbOCTcon
            // 
            this.ckbOCTcon.AutoSize = true;
            this.ckbOCTcon.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ckbOCTcon.Location = new System.Drawing.Point(530, 83);
            this.ckbOCTcon.Name = "ckbOCTcon";
            this.ckbOCTcon.Size = new System.Drawing.Size(107, 20);
            this.ckbOCTcon.TabIndex = 21;
            this.ckbOCTcon.Text = "O/C對應T-con";
            this.ckbOCTcon.UseVisualStyleBackColor = true;
            // 
            // ckbUndeliverGreaterThenZero
            // 
            this.ckbUndeliverGreaterThenZero.AutoSize = true;
            this.ckbUndeliverGreaterThenZero.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ckbUndeliverGreaterThenZero.Location = new System.Drawing.Point(413, 83);
            this.ckbUndeliverGreaterThenZero.Name = "ckbUndeliverGreaterThenZero";
            this.ckbUndeliverGreaterThenZero.Size = new System.Drawing.Size(111, 20);
            this.ckbUndeliverGreaterThenZero.TabIndex = 20;
            this.ckbUndeliverGreaterThenZero.Text = "訂單未交大於零";
            this.ckbUndeliverGreaterThenZero.UseVisualStyleBackColor = true;
            // 
            // btnExcel
            // 
            this.btnExcel.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnExcel.Location = new System.Drawing.Point(188, 76);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(87, 31);
            this.btnExcel.TabIndex = 19;
            this.btnExcel.Text = "匯出Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btnSort
            // 
            this.btnSort.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnSort.Location = new System.Drawing.Point(78, 76);
            this.btnSort.Name = "btnSort";
            this.btnSort.Size = new System.Drawing.Size(87, 31);
            this.btnSort.TabIndex = 18;
            this.btnSort.Text = "查詢";
            this.btnSort.UseVisualStyleBackColor = true;
            this.btnSort.Click += new System.EventHandler(this.btnSort_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(669, 28);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 25);
            this.label3.TabIndex = 16;
            this.label3.Text = "多筆料號";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Rounded MT Bold", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(410, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 23);
            this.label1.TabIndex = 16;
            this.label1.Text = "BU";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Arial Rounded MT Bold", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(1155, 76);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(85, 23);
            this.label9.TabIndex = 17;
            this.label9.Text = "Verison";
            this.label9.Visible = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial Rounded MT Bold", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(950, 76);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(72, 23);
            this.label7.TabIndex = 15;
            this.label7.Text = "Grade";
            this.label7.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial Rounded MT Bold", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(1140, 23);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(69, 23);
            this.label5.TabIndex = 14;
            this.label5.Text = "Model";
            this.label5.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial Rounded MT Bold", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(205, 26);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 23);
            this.label4.TabIndex = 13;
            this.label4.Text = "Type";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial Rounded MT Bold", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 23);
            this.label2.TabIndex = 12;
            this.label2.Text = "Size";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dgvStockStatus);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 140);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1522, 471);
            this.panel2.TabIndex = 7;
            // 
            // dgvStockStatus
            // 
            this.dgvStockStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvStockStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvStockStatus.Location = new System.Drawing.Point(0, 0);
            this.dgvStockStatus.Name = "dgvStockStatus";
            this.dgvStockStatus.RowTemplate.Height = 24;
            this.dgvStockStatus.Size = new System.Drawing.Size(1522, 471);
            this.dgvStockStatus.TabIndex = 0;
            // 
            // StockStatus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1522, 611);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "StockStatus";
            this.Text = "貨況查詢";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvStockStatus)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox cmbBU;
        private System.Windows.Forms.CheckBox ckbOnHandGreatThenZero;
        private System.Windows.Forms.CheckBox ckbZeroUnshow;
        private System.Windows.Forms.CheckBox ckbOCTcon;
        private System.Windows.Forms.CheckBox ckbUndeliverGreaterThenZero;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Button btnSort;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dgvStockStatus;
        private System.Windows.Forms.TextBox txbItemcode;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbSize;
        private System.Windows.Forms.ComboBox cmbType;
        private System.Windows.Forms.ComboBox cmbModel;
        private System.Windows.Forms.ComboBox cmbGrade;
        private System.Windows.Forms.ComboBox cmbVersion;
    }
}