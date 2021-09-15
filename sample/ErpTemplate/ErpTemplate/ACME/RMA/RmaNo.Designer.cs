namespace ACME
{
    partial class RmaNo
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
            this.components = new System.ComponentModel.Container();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.COMPANY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_RMA_NO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_CusName_s = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_RMODEL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_RVER = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_RGRADE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_RQUINITY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.U_AUO_RMA_NO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Contractid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.listBox3 = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.COMPANY,
            this.U_RMA_NO,
            this.U_CusName_s,
            this.U_RMODEL,
            this.U_RVER,
            this.U_RGRADE,
            this.U_RQUINITY,
            this.U_AUO_RMA_NO,
            this.Contractid});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(835, 514);
            this.dataGridView1.TabIndex = 0;
            // 
            // COMPANY
            // 
            this.COMPANY.DataPropertyName = "COMPANY";
            this.COMPANY.HeaderText = "公司";
            this.COMPANY.Name = "COMPANY";
            this.COMPANY.Width = 51;
            // 
            // U_RMA_NO
            // 
            this.U_RMA_NO.DataPropertyName = "U_RMA_NO";
            this.U_RMA_NO.HeaderText = "RMA_No";
            this.U_RMA_NO.Name = "U_RMA_NO";
            this.U_RMA_NO.Width = 76;
            // 
            // U_CusName_s
            // 
            this.U_CusName_s.DataPropertyName = "U_CusName_s";
            this.U_CusName_s.HeaderText = "客戶簡稱";
            this.U_CusName_s.Name = "U_CusName_s";
            this.U_CusName_s.Width = 61;
            // 
            // U_RMODEL
            // 
            this.U_RMODEL.DataPropertyName = "U_RMODEL";
            this.U_RMODEL.HeaderText = "型號";
            this.U_RMODEL.Name = "U_RMODEL";
            this.U_RMODEL.Width = 51;
            // 
            // U_RVER
            // 
            this.U_RVER.DataPropertyName = "U_RVER";
            this.U_RVER.HeaderText = "Ver.";
            this.U_RVER.Name = "U_RVER";
            this.U_RVER.Width = 50;
            // 
            // U_RGRADE
            // 
            this.U_RGRADE.DataPropertyName = "U_RGRADE";
            this.U_RGRADE.HeaderText = "Grade";
            this.U_RGRADE.Name = "U_RGRADE";
            this.U_RGRADE.Width = 58;
            // 
            // U_RQUINITY
            // 
            this.U_RQUINITY.DataPropertyName = "U_RQUINITY";
            this.U_RQUINITY.HeaderText = "Q\'TY";
            this.U_RQUINITY.Name = "U_RQUINITY";
            this.U_RQUINITY.Width = 55;
            // 
            // U_AUO_RMA_NO
            // 
            this.U_AUO_RMA_NO.DataPropertyName = "U_AUO_RMA_NO";
            this.U_AUO_RMA_NO.HeaderText = "AUO RMA No";
            this.U_AUO_RMA_NO.Name = "U_AUO_RMA_NO";
            this.U_AUO_RMA_NO.Width = 79;
            // 
            // Contractid
            // 
            this.Contractid.DataPropertyName = "Contractid";
            this.Contractid.HeaderText = "契約號碼";
            this.Contractid.Name = "Contractid";
            this.Contractid.Visible = false;
            this.Contractid.Width = 61;
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(355, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "確定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 12;
            this.listBox1.Location = new System.Drawing.Point(682, 12);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(18, 16);
            this.listBox1.TabIndex = 2;
            this.listBox1.Visible = false;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(56, 11);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 22);
            this.textBox1.TabIndex = 3;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "RMANO";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.listBox3);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.listBox2);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.listBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(835, 40);
            this.panel1.TabIndex = 5;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(447, 10);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(99, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "批量匯入資料";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 12;
            this.listBox2.Location = new System.Drawing.Point(638, 13);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(18, 16);
            this.listBox2.TabIndex = 7;
            this.listBox2.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(162, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "Vender Rma No";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(249, 11);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 22);
            this.textBox2.TabIndex = 5;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 40);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(835, 514);
            this.panel2.TabIndex = 6;
            // 
            // listBox3
            // 
            this.listBox3.FormattingEnabled = true;
            this.listBox3.ItemHeight = 12;
            this.listBox3.Location = new System.Drawing.Point(717, 13);
            this.listBox3.Name = "listBox3";
            this.listBox3.Size = new System.Drawing.Size(18, 16);
            this.listBox3.TabIndex = 9;
            this.listBox3.Visible = false;
            // 
            // RmaNo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(835, 554);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "RmaNo";
            this.Text = "RMA資料";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.DataGridViewTextBoxColumn COMPANY;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_RMA_NO;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_CusName_s;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_RMODEL;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_RVER;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_RGRADE;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_RQUINITY;
        private System.Windows.Forms.DataGridViewTextBoxColumn U_AUO_RMA_NO;
        private System.Windows.Forms.DataGridViewTextBoxColumn Contractid;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ListBox listBox3;
    }
}