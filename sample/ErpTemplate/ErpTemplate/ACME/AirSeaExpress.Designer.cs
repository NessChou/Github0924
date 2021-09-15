
namespace ACME
{
    partial class AirSeaExpress
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AirSeaExpress));
            this.panel1 = new System.Windows.Forms.Panel();
            this.txbCardName = new System.Windows.Forms.TextBox();
            this.txbCardCode1 = new System.Windows.Forms.TextBox();
            this.btnGetCardCode1 = new System.Windows.Forms.Button();
            this.txbCloseDay = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnSearch = new System.Windows.Forms.Button();
            this.cmbReceiveDay = new System.Windows.Forms.ComboBox();
            this.cmbCardCode2 = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dgvAirSeaExpress = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAirSeaExpress)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txbCardName);
            this.panel1.Controls.Add(this.txbCardCode1);
            this.panel1.Controls.Add(this.btnGetCardCode1);
            this.panel1.Controls.Add(this.txbCloseDay);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btnSearch);
            this.panel1.Controls.Add(this.cmbReceiveDay);
            this.panel1.Controls.Add(this.cmbCardCode2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1408, 95);
            this.panel1.TabIndex = 0;
            // 
            // txbCardName
            // 
            this.txbCardName.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbCardName.Location = new System.Drawing.Point(173, 19);
            this.txbCardName.Name = "txbCardName";
            this.txbCardName.Size = new System.Drawing.Size(159, 25);
            this.txbCardName.TabIndex = 56;
            // 
            // txbCardCode1
            // 
            this.txbCardCode1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbCardCode1.Location = new System.Drawing.Point(91, 19);
            this.txbCardCode1.Name = "txbCardCode1";
            this.txbCardCode1.Size = new System.Drawing.Size(76, 25);
            this.txbCardCode1.TabIndex = 56;
            // 
            // btnGetCardCode1
            // 
            this.btnGetCardCode1.ForeColor = System.Drawing.SystemColors.ActiveBorder;
            this.btnGetCardCode1.Image = ((System.Drawing.Image)(resources.GetObject("btnGetCardCode1.Image")));
            this.btnGetCardCode1.Location = new System.Drawing.Point(338, 18);
            this.btnGetCardCode1.Name = "btnGetCardCode1";
            this.btnGetCardCode1.Size = new System.Drawing.Size(31, 27);
            this.btnGetCardCode1.TabIndex = 55;
            this.btnGetCardCode1.Text = "...";
            this.btnGetCardCode1.UseVisualStyleBackColor = true;
            this.btnGetCardCode1.Click += new System.EventHandler(this.btnGetCardCode1_Click);
            // 
            // txbCloseDay
            // 
            this.txbCloseDay.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txbCloseDay.Location = new System.Drawing.Point(462, 21);
            this.txbCloseDay.Name = "txbCloseDay";
            this.txbCloseDay.Size = new System.Drawing.Size(141, 25);
            this.txbCloseDay.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(615, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 20);
            this.label4.TabIndex = 2;
            this.label4.Text = "運送方式";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(383, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 20);
            this.label3.TabIndex = 2;
            this.label3.Text = "日期區間";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(1257, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "費用廠商";
            this.label1.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(12, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "廠商編號";
            // 
            // btnSearch
            // 
            this.btnSearch.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnSearch.Location = new System.Drawing.Point(1097, 14);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(108, 32);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "查詢";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // cmbReceiveDay
            // 
            this.cmbReceiveDay.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cmbReceiveDay.FormattingEnabled = true;
            this.cmbReceiveDay.Items.AddRange(new object[] {
            "SEA",
            "AIR"});
            this.cmbReceiveDay.Location = new System.Drawing.Point(694, 19);
            this.cmbReceiveDay.Name = "cmbReceiveDay";
            this.cmbReceiveDay.Size = new System.Drawing.Size(150, 28);
            this.cmbReceiveDay.TabIndex = 0;
            this.cmbReceiveDay.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // cmbCardCode2
            // 
            this.cmbCardCode2.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cmbCardCode2.FormattingEnabled = true;
            this.cmbCardCode2.Location = new System.Drawing.Point(1335, 39);
            this.cmbCardCode2.Name = "cmbCardCode2";
            this.cmbCardCode2.Size = new System.Drawing.Size(57, 24);
            this.cmbCardCode2.TabIndex = 0;
            this.cmbCardCode2.Visible = false;
            this.cmbCardCode2.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dgvAirSeaExpress);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 95);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1408, 520);
            this.panel2.TabIndex = 1;
            // 
            // dgvAirSeaExpress
            // 
            this.dgvAirSeaExpress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAirSeaExpress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvAirSeaExpress.Location = new System.Drawing.Point(0, 0);
            this.dgvAirSeaExpress.Name = "dgvAirSeaExpress";
            this.dgvAirSeaExpress.RowTemplate.Height = 24;
            this.dgvAirSeaExpress.Size = new System.Drawing.Size(1408, 520);
            this.dgvAirSeaExpress.TabIndex = 0;
            // 
            // AirSeaExpress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1408, 615);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "AirSeaExpress";
            this.Text = "AirExpress";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAirSeaExpress)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dgvAirSeaExpress;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbReceiveDay;
        private System.Windows.Forms.ComboBox cmbCardCode2;
        private System.Windows.Forms.TextBox txbCloseDay;
        private System.Windows.Forms.TextBox txbCardCode1;
        private System.Windows.Forms.Button btnGetCardCode1;
        private System.Windows.Forms.TextBox txbCardName;
    }
}