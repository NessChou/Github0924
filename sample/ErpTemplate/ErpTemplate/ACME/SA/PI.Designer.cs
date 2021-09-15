namespace ACME
{
    partial class PI
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
            System.Windows.Forms.Label cardCodeLabel;
            System.Windows.Forms.Label memoLabel;
            System.Windows.Forms.Label tYPELabel;
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.cardCodeTextBox = new System.Windows.Forms.TextBox();
            this.sACUSTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.sa = new ACME.ACMEDataSet.sa();
            this.idTextBox = new System.Windows.Forms.TextBox();
            this.memoTextBox = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.類型 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.tYPETextBox = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.sACUSTTableAdapter = new ACME.ACMEDataSet.saTableAdapters.SACUSTTableAdapter();
            cardCodeLabel = new System.Windows.Forms.Label();
            memoLabel = new System.Windows.Forms.Label();
            tYPELabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sACUSTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sa)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Size = new System.Drawing.Size(908, 591);
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            this.panel1.Controls.SetChildIndex(this.panel2, 0);
            this.panel1.Controls.SetChildIndex(this.panel3, 0);
            // 
            // cardCodeLabel
            // 
            cardCodeLabel.AutoSize = true;
            cardCodeLabel.Location = new System.Drawing.Point(14, 49);
            cardCodeLabel.Name = "cardCodeLabel";
            cardCodeLabel.Size = new System.Drawing.Size(53, 12);
            cardCodeLabel.TabIndex = 1;
            cardCodeLabel.Text = "帳戶名稱";
            // 
            // memoLabel
            // 
            memoLabel.AutoSize = true;
            memoLabel.Location = new System.Drawing.Point(25, 92);
            memoLabel.Name = "memoLabel";
            memoLabel.Size = new System.Drawing.Size(29, 12);
            memoLabel.TabIndex = 5;
            memoLabel.Text = "內容";
            // 
            // tYPELabel
            // 
            tYPELabel.AutoSize = true;
            tYPELabel.Location = new System.Drawing.Point(14, 19);
            tYPELabel.Name = "tYPELabel";
            tYPELabel.Size = new System.Drawing.Size(53, 12);
            tYPELabel.TabIndex = 7;
            tYPELabel.Text = "帳戶類型";
            // 
            // cardCodeTextBox
            // 
            this.cardCodeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sACUSTBindingSource, "CardCode", true));
            this.cardCodeTextBox.Location = new System.Drawing.Point(79, 46);
            this.cardCodeTextBox.Name = "cardCodeTextBox";
            this.cardCodeTextBox.Size = new System.Drawing.Size(407, 22);
            this.cardCodeTextBox.TabIndex = 2;
            // 
            // sACUSTBindingSource
            // 
            this.sACUSTBindingSource.DataMember = "SACUST";
            this.sACUSTBindingSource.DataSource = this.sa;
            // 
            // sa
            // 
            this.sa.DataSetName = "sa";
            this.sa.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // idTextBox
            // 
            this.idTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sACUSTBindingSource, "id", true));
            this.idTextBox.Location = new System.Drawing.Point(73, 39);
            this.idTextBox.Name = "idTextBox";
            this.idTextBox.Size = new System.Drawing.Size(0, 22);
            this.idTextBox.TabIndex = 4;
            // 
            // memoTextBox
            // 
            this.memoTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sACUSTBindingSource, "memo", true));
            this.memoTextBox.Location = new System.Drawing.Point(78, 76);
            this.memoTextBox.Multiline = true;
            this.memoTextBox.Name = "memoTextBox";
            this.memoTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.memoTextBox.Size = new System.Drawing.Size(408, 191);
            this.memoTextBox.TabIndex = 6;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.類型,
            this.Column1,
            this.Column2});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(908, 289);
            this.dataGridView1.TabIndex = 7;
            // 
            // 類型
            // 
            this.類型.DataPropertyName = "TYPE";
            this.類型.HeaderText = "類型";
            this.類型.Name = "類型";
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "cardcode";
            this.Column1.HeaderText = "帳戶名稱";
            this.Column1.Name = "Column1";
            this.Column1.Width = 150;
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column2.DataPropertyName = "memo";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Column2.DefaultCellStyle = dataGridViewCellStyle1;
            this.Column2.HeaderText = "內容";
            this.Column2.Name = "Column2";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "外幣帳戶",
            "台幣帳戶"});
            this.comboBox1.Location = new System.Drawing.Point(79, 18);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(118, 20);
            this.comboBox1.TabIndex = 9;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // tYPETextBox
            // 
            this.tYPETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sACUSTBindingSource, "TYPE", true));
            this.tYPETextBox.Location = new System.Drawing.Point(79, 16);
            this.tYPETextBox.Name = "tYPETextBox";
            this.tYPETextBox.Size = new System.Drawing.Size(100, 22);
            this.tYPETextBox.TabIndex = 10;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(tYPELabel);
            this.panel2.Controls.Add(this.tYPETextBox);
            this.panel2.Controls.Add(this.cardCodeTextBox);
            this.panel2.Controls.Add(this.comboBox1);
            this.panel2.Controls.Add(cardCodeLabel);
            this.panel2.Controls.Add(this.idTextBox);
            this.panel2.Controls.Add(this.memoTextBox);
            this.panel2.Controls.Add(memoLabel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(908, 280);
            this.panel2.TabIndex = 11;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dataGridView1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 280);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(908, 289);
            this.panel3.TabIndex = 12;
            // 
            // sACUSTTableAdapter
            // 
            this.sACUSTTableAdapter.ClearBeforeFill = true;
            // 
            // PI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(908, 630);
            this.Name = "PI";
            this.Text = "PI";
            this.Load += new System.EventHandler(this.PI_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sACUSTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sa)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACME.ACMEDataSet.sa sa;
        private System.Windows.Forms.BindingSource sACUSTBindingSource;
        private ACME.ACMEDataSet.saTableAdapters.SACUSTTableAdapter sACUSTTableAdapter;
        private System.Windows.Forms.TextBox cardCodeTextBox;
        private System.Windows.Forms.TextBox idTextBox;
        private System.Windows.Forms.TextBox memoTextBox;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 類型;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.TextBox tYPETextBox;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel2;
    }
}
