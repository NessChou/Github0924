namespace ACME
{
    partial class RmaCarton
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
            this.rMA_CARTONDataGridView = new System.Windows.Forms.DataGridView();
            this.DOCMONTH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DOCDATE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AUIN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AUOUT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CUSTIN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CUSTOUT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DOCYEAR = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.rMA_CARTONBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.rm = new ACME.ACMEDataSet.rm();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.button2 = new System.Windows.Forms.Button();
            this.rMA_CARTONTableAdapter = new ACME.ACMEDataSet.rmTableAdapters.RMA_CARTONTableAdapter();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CARTONDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CARTONBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rm)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // rMA_CARTONDataGridView
            // 
            this.rMA_CARTONDataGridView.AutoGenerateColumns = false;
            this.rMA_CARTONDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DOCMONTH,
            this.DOCDATE,
            this.AUIN,
            this.AUOUT,
            this.CUSTIN,
            this.CUSTOUT,
            this.DOCYEAR});
            this.rMA_CARTONDataGridView.DataSource = this.rMA_CARTONBindingSource;
            this.rMA_CARTONDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rMA_CARTONDataGridView.Location = new System.Drawing.Point(0, 31);
            this.rMA_CARTONDataGridView.Name = "rMA_CARTONDataGridView";
            this.rMA_CARTONDataGridView.RowTemplate.Height = 24;
            this.rMA_CARTONDataGridView.Size = new System.Drawing.Size(814, 621);
            this.rMA_CARTONDataGridView.TabIndex = 2;
            this.rMA_CARTONDataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.rMA_CARTONDataGridView_DefaultValuesNeeded);
            // 
            // DOCMONTH
            // 
            this.DOCMONTH.DataPropertyName = "DOCMONTH";
            this.DOCMONTH.HeaderText = "日";
            this.DOCMONTH.Name = "DOCMONTH";
            this.DOCMONTH.Visible = false;
            // 
            // DOCDATE
            // 
            this.DOCDATE.DataPropertyName = "DOCDATE";
            this.DOCDATE.HeaderText = "日";
            this.DOCDATE.Name = "DOCDATE";
            this.DOCDATE.Width = 50;
            // 
            // AUIN
            // 
            this.AUIN.DataPropertyName = "AUIN";
            this.AUIN.HeaderText = "AU還回";
            this.AUIN.Name = "AUIN";
            this.AUIN.Width = 80;
            // 
            // AUOUT
            // 
            this.AUOUT.DataPropertyName = "AUOUT";
            this.AUOUT.HeaderText = "出貨AU";
            this.AUOUT.Name = "AUOUT";
            this.AUOUT.Width = 80;
            // 
            // CUSTIN
            // 
            this.CUSTIN.DataPropertyName = "CUSTIN";
            this.CUSTIN.HeaderText = "客戶退回";
            this.CUSTIN.Name = "CUSTIN";
            this.CUSTIN.Width = 80;
            // 
            // CUSTOUT
            // 
            this.CUSTOUT.DataPropertyName = "CUSTOUT";
            this.CUSTOUT.HeaderText = "還貨客戶";
            this.CUSTOUT.Name = "CUSTOUT";
            this.CUSTOUT.Width = 80;
            // 
            // DOCYEAR
            // 
            this.DOCYEAR.DataPropertyName = "DOCYEAR";
            this.DOCYEAR.HeaderText = "DOCYEAR";
            this.DOCYEAR.Name = "DOCYEAR";
            this.DOCYEAR.Visible = false;
            // 
            // rMA_CARTONBindingSource
            // 
            this.rMA_CARTONBindingSource.DataMember = "RMA_CARTON";
            this.rMA_CARTONBindingSource.DataSource = this.rm;
            // 
            // rm
            // 
            this.rm.DataSetName = "rm";
            this.rm.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(255, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "查詢";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(135, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "月";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(158, 7);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(66, 20);
            this.comboBox2.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 10;
            this.label1.Text = "年";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(39, 7);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(79, 20);
            this.comboBox1.TabIndex = 9;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(336, 6);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 13;
            this.button2.Text = "存檔";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // rMA_CARTONTableAdapter
            // 
            this.rMA_CARTONTableAdapter.ClearBeforeFill = true;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.dataGridView2);
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.comboBox2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(814, 31);
            this.panel1.TabIndex = 14;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(713, 4);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.Size = new System.Drawing.Size(27, 22);
            this.dataGridView2.TabIndex = 18;
            this.dataGridView2.Visible = false;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(495, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 26);
            this.button4.TabIndex = 16;
            this.button4.Text = "報表-月";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(417, 6);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 15;
            this.button3.Text = "報表-年";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(576, 6);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(27, 22);
            this.dataGridView1.TabIndex = 14;
            this.dataGridView1.Visible = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.rMA_CARTONDataGridView);
            this.panel2.Controls.Add(this.panel1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(814, 652);
            this.panel2.TabIndex = 15;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(828, 684);
            this.tabControl1.TabIndex = 16;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.panel2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(820, 658);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "箱數管理";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.button5);
            this.tabPage2.Controls.Add(this.textBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(820, 658);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "每日收貨發信";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "日期";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(137, 12);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 1;
            this.button5.Text = "寄信";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(57, 13);
            this.textBox1.MaxLength = 8;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(74, 22);
            this.textBox1.TabIndex = 0;
            // 
            // RmaCarton
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(828, 684);
            this.Controls.Add(this.tabControl1);
            this.Name = "RmaCarton";
            this.Text = "箱數管理";
            this.Load += new System.EventHandler(this.RmaCarton_Load);
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CARTONDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rMA_CARTONBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rm)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private ACME.ACMEDataSet.rm rm;
        private System.Windows.Forms.BindingSource rMA_CARTONBindingSource;
        private ACME.ACMEDataSet.rmTableAdapters.RMA_CARTONTableAdapter rMA_CARTONTableAdapter;
        private System.Windows.Forms.DataGridView rMA_CARTONDataGridView;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.DataGridViewTextBoxColumn DOCMONTH;
        private System.Windows.Forms.DataGridViewTextBoxColumn DOCDATE;
        private System.Windows.Forms.DataGridViewTextBoxColumn AUIN;
        private System.Windows.Forms.DataGridViewTextBoxColumn AUOUT;
        private System.Windows.Forms.DataGridViewTextBoxColumn CUSTIN;
        private System.Windows.Forms.DataGridViewTextBoxColumn CUSTOUT;
        private System.Windows.Forms.DataGridViewTextBoxColumn DOCYEAR;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dataGridView2;
    }
}