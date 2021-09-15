namespace ACME
{
    partial class SHIPPACK
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.來源Invoice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.ShippingCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PLNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Doctentry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SeqNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PackageNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DescGoods = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Net = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Gross = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MeasurmentCM = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TREETYPE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.VISORDER = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SOID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PACKMARK = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SeqNo2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.來源Invoice});
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(226, 254);
            this.dataGridView1.TabIndex = 0;
            // 
            // 來源Invoice
            // 
            this.來源Invoice.HeaderText = "來源Invoice";
            this.來源Invoice.Name = "來源Invoice";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(297, 101);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(124, 59);
            this.button1.TabIndex = 1;
            this.button1.Text = "1.預覽";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(367, 31);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(205, 22);
            this.textBox1.TabIndex = 2;
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ShippingCode,
            this.PLNo,
            this.Doctentry,
            this.SeqNo,
            this.PackageNo,
            this.CNo,
            this.DescGoods,
            this.Quantity,
            this.Net,
            this.Gross,
            this.MeasurmentCM,
            this.TREETYPE,
            this.VISORDER,
            this.SOID,
            this.PACKMARK,
            this.SeqNo2});
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.Location = new System.Drawing.Point(0, 0);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.Size = new System.Drawing.Size(1050, 391);
            this.dataGridView2.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(286, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "目的Invoice";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(427, 101);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(124, 59);
            this.button2.TabIndex = 5;
            this.button2.Text = "2.匯入資料";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1050, 272);
            this.panel1.TabIndex = 6;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 272);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1050, 391);
            this.panel2.TabIndex = 7;
            // 
            // ShippingCode
            // 
            this.ShippingCode.DataPropertyName = "ShippingCode";
            this.ShippingCode.HeaderText = "ShippingCode";
            this.ShippingCode.Name = "ShippingCode";
            this.ShippingCode.ReadOnly = true;
            this.ShippingCode.Visible = false;
            this.ShippingCode.Width = 97;
            // 
            // PLNo
            // 
            this.PLNo.DataPropertyName = "PLNo";
            this.PLNo.HeaderText = "PLNo";
            this.PLNo.Name = "PLNo";
            this.PLNo.ReadOnly = true;
            // 
            // Doctentry
            // 
            this.Doctentry.DataPropertyName = "Doctentry";
            this.Doctentry.HeaderText = "Doctentry";
            this.Doctentry.Name = "Doctentry";
            this.Doctentry.ReadOnly = true;
            this.Doctentry.Visible = false;
            this.Doctentry.Width = 76;
            // 
            // SeqNo
            // 
            this.SeqNo.DataPropertyName = "SeqNo";
            this.SeqNo.HeaderText = "SeqNo";
            this.SeqNo.Name = "SeqNo";
            this.SeqNo.ReadOnly = true;
            this.SeqNo.Visible = false;
            this.SeqNo.Width = 61;
            // 
            // PackageNo
            // 
            this.PackageNo.DataPropertyName = "PackageNo";
            this.PackageNo.HeaderText = "PackageNo";
            this.PackageNo.Name = "PackageNo";
            this.PackageNo.ReadOnly = true;
            this.PackageNo.Width = 60;
            // 
            // CNo
            // 
            this.CNo.DataPropertyName = "CNo";
            this.CNo.HeaderText = "C/No";
            this.CNo.Name = "CNo";
            this.CNo.ReadOnly = true;
            this.CNo.Width = 60;
            // 
            // DescGoods
            // 
            this.DescGoods.DataPropertyName = "DescGoods";
            this.DescGoods.HeaderText = "Description Of Goods";
            this.DescGoods.Name = "DescGoods";
            this.DescGoods.ReadOnly = true;
            this.DescGoods.Width = 260;
            // 
            // Quantity
            // 
            this.Quantity.DataPropertyName = "Quantity";
            this.Quantity.HeaderText = "Quantity(pcs)";
            this.Quantity.Name = "Quantity";
            this.Quantity.ReadOnly = true;
            // 
            // Net
            // 
            this.Net.DataPropertyName = "Net";
            this.Net.HeaderText = "Net";
            this.Net.Name = "Net";
            this.Net.ReadOnly = true;
            // 
            // Gross
            // 
            this.Gross.DataPropertyName = "Gross";
            this.Gross.HeaderText = "Gross";
            this.Gross.Name = "Gross";
            this.Gross.ReadOnly = true;
            // 
            // MeasurmentCM
            // 
            this.MeasurmentCM.DataPropertyName = "MeasurmentCM";
            this.MeasurmentCM.HeaderText = "MeasurmentCM";
            this.MeasurmentCM.Name = "MeasurmentCM";
            this.MeasurmentCM.ReadOnly = true;
            // 
            // TREETYPE
            // 
            this.TREETYPE.DataPropertyName = "TREETYPE";
            this.TREETYPE.HeaderText = "TREETYPE";
            this.TREETYPE.Name = "TREETYPE";
            this.TREETYPE.ReadOnly = true;
            this.TREETYPE.Visible = false;
            this.TREETYPE.Width = 87;
            // 
            // VISORDER
            // 
            this.VISORDER.DataPropertyName = "VISORDER";
            this.VISORDER.HeaderText = "VISORDER";
            this.VISORDER.Name = "VISORDER";
            this.VISORDER.ReadOnly = true;
            this.VISORDER.Visible = false;
            this.VISORDER.Width = 87;
            // 
            // SOID
            // 
            this.SOID.DataPropertyName = "SOID";
            this.SOID.HeaderText = "SOID";
            this.SOID.Name = "SOID";
            this.SOID.ReadOnly = true;
            this.SOID.Visible = false;
            this.SOID.Width = 56;
            // 
            // PACKMARK
            // 
            this.PACKMARK.DataPropertyName = "PACKMARK";
            this.PACKMARK.HeaderText = "PACKMARK";
            this.PACKMARK.Name = "PACKMARK";
            this.PACKMARK.ReadOnly = true;
            this.PACKMARK.Visible = false;
            this.PACKMARK.Width = 94;
            // 
            // SeqNo2
            // 
            this.SeqNo2.DataPropertyName = "SeqNo2";
            this.SeqNo2.HeaderText = "SeqNo2";
            this.SeqNo2.Name = "SeqNo2";
            this.SeqNo2.ReadOnly = true;
            this.SeqNo2.Visible = false;
            this.SeqNo2.Width = 67;
            // 
            // SHIPPACK
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1050, 663);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "SHIPPACK";
            this.Text = "PACKING匯入";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 來源Invoice;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ShippingCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn PLNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn Doctentry;
        private System.Windows.Forms.DataGridViewTextBoxColumn SeqNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn PackageNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn CNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn DescGoods;
        private System.Windows.Forms.DataGridViewTextBoxColumn Quantity;
        private System.Windows.Forms.DataGridViewTextBoxColumn Net;
        private System.Windows.Forms.DataGridViewTextBoxColumn Gross;
        private System.Windows.Forms.DataGridViewTextBoxColumn MeasurmentCM;
        private System.Windows.Forms.DataGridViewTextBoxColumn TREETYPE;
        private System.Windows.Forms.DataGridViewTextBoxColumn VISORDER;
        private System.Windows.Forms.DataGridViewTextBoxColumn SOID;
        private System.Windows.Forms.DataGridViewTextBoxColumn PACKMARK;
        private System.Windows.Forms.DataGridViewTextBoxColumn SeqNo2;
    }
}