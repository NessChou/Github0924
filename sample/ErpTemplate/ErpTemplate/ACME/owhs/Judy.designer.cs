namespace ACME
{
    partial class Judy
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
            System.Windows.Forms.Label rMBLabel;
            System.Windows.Forms.Label hKLabel;
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button5 = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.ITEMCODE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Model = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Version = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Grade = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Moving = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.COST = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AddDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Qty_Stock = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Cost_Stock = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Date_Stock = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Cost_Average = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Price = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Date_Price = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.button6 = new System.Windows.Forms.Button();
            this.hKTextBox = new System.Windows.Forms.TextBox();
            this.sALES_DOCCURBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.wh = new ACME.ACMEDataSet.wh();
            this.rMBTextBox = new System.Windows.Forms.TextBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.dataGridView4 = new System.Windows.Forms.DataGridView();
            this.button11 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.sALES_DOCCURTableAdapter1 = new ACME.ACMEDataSet.whTableAdapters.SALES_DOCCURTableAdapter();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            rMBLabel = new System.Windows.Forms.Label();
            hKLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sALES_DOCCURBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.wh)).BeginInit();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).BeginInit();
            this.SuspendLayout();
            // 
            // rMBLabel
            // 
            rMBLabel.AutoSize = true;
            rMBLabel.Location = new System.Drawing.Point(23, 18);
            rMBLabel.Name = "rMBLabel";
            rMBLabel.Size = new System.Drawing.Size(34, 12);
            rMBLabel.TabIndex = 0;
            rMBLabel.Text = "RMB:";
            // 
            // hKLabel
            // 
            hKLabel.AutoSize = true;
            hKLabel.Location = new System.Drawing.Point(33, 60);
            hKLabel.Name = "hKLabel";
            hKLabel.Size = new System.Drawing.Size(24, 12);
            hKLabel.TabIndex = 2;
            hKLabel.Text = "HK:";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(3, 1);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(958, 679);
            this.tabControl1.TabIndex = 20;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.button5);
            this.tabPage2.Controls.Add(this.label11);
            this.tabPage2.Controls.Add(this.button4);
            this.tabPage2.Controls.Add(this.comboBox1);
            this.tabPage2.Controls.Add(this.button3);
            this.tabPage2.Controls.Add(this.button2);
            this.tabPage2.Controls.Add(this.dataGridView3);
            this.tabPage2.Location = new System.Drawing.Point(4, 21);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(950, 654);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "進貨成本";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.BackgroundImage = global::ACME.Properties.Resources.tw12_sp1b;
            this.button5.ForeColor = System.Drawing.Color.White;
            this.button5.Location = new System.Drawing.Point(282, 6);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(72, 22);
            this.button5.TabIndex = 8;
            this.button5.Text = "更新Price";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(6, 12);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(21, 12);
            this.label11.TabIndex = 7;
            this.label11.Text = "BU";
            // 
            // button4
            // 
            this.button4.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button4.BackgroundImage = global::ACME.Properties.Resources.tw12_sp1b;
            this.button4.ForeColor = System.Drawing.Color.White;
            this.button4.Location = new System.Drawing.Point(145, 5);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(53, 23);
            this.button4.TabIndex = 6;
            this.button4.Text = "查詢";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(33, 8);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(106, 20);
            this.comboBox1.TabIndex = 5;
            // 
            // button3
            // 
            this.button3.BackgroundImage = global::ACME.Properties.Resources.tw12_sp1b;
            this.button3.ForeColor = System.Drawing.Color.White;
            this.button3.Location = new System.Drawing.Point(204, 6);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(72, 22);
            this.button3.TabIndex = 4;
            this.button3.Text = "更新Cost";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.BackgroundImage = global::ACME.Properties.Resources.tw12_sp1b;
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Location = new System.Drawing.Point(360, 5);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(53, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToAddRows = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ITEMCODE,
            this.Model,
            this.Version,
            this.Grade,
            this.Moving,
            this.COST,
            this.AddDate,
            this.Qty_Stock,
            this.Cost_Stock,
            this.Date_Stock,
            this.Cost_Average,
            this.Price,
            this.Date_Price});
            this.dataGridView3.Location = new System.Drawing.Point(6, 34);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.RowTemplate.Height = 24;
            this.dataGridView3.Size = new System.Drawing.Size(938, 606);
            this.dataGridView3.TabIndex = 0;
            this.dataGridView3.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellValueChanged);
            this.dataGridView3.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView3_RowPostPaint);
            // 
            // ITEMCODE
            // 
            this.ITEMCODE.DataPropertyName = "ITEMCODE";
            this.ITEMCODE.HeaderText = "ACME P/N";
            this.ITEMCODE.Name = "ITEMCODE";
            this.ITEMCODE.Width = 130;
            // 
            // Model
            // 
            this.Model.DataPropertyName = "Model";
            this.Model.HeaderText = "Model";
            this.Model.Name = "Model";
            this.Model.Width = 80;
            // 
            // Version
            // 
            this.Version.DataPropertyName = "Version";
            this.Version.HeaderText = "Version";
            this.Version.Name = "Version";
            this.Version.Width = 50;
            // 
            // Grade
            // 
            this.Grade.DataPropertyName = "Grade";
            this.Grade.HeaderText = "Grade";
            this.Grade.Name = "Grade";
            this.Grade.Width = 50;
            // 
            // Moving
            // 
            this.Moving.DataPropertyName = "Moving";
            this.Moving.HeaderText = "Slow Moving";
            this.Moving.Name = "Moving";
            this.Moving.Width = 95;
            // 
            // COST
            // 
            this.COST.DataPropertyName = "COST";
            this.COST.HeaderText = "Cost_New";
            this.COST.Name = "COST";
            this.COST.Width = 60;
            // 
            // AddDate
            // 
            this.AddDate.DataPropertyName = "AddDate";
            this.AddDate.HeaderText = "Date_New";
            this.AddDate.Name = "AddDate";
            this.AddDate.Width = 60;
            // 
            // Qty_Stock
            // 
            this.Qty_Stock.DataPropertyName = "Qty_Stock";
            this.Qty_Stock.HeaderText = "Qty_Stock";
            this.Qty_Stock.Name = "Qty_Stock";
            this.Qty_Stock.Width = 60;
            // 
            // Cost_Stock
            // 
            this.Cost_Stock.DataPropertyName = "Cost_Stock";
            this.Cost_Stock.HeaderText = "Cost_Stock";
            this.Cost_Stock.Name = "Cost_Stock";
            this.Cost_Stock.Width = 65;
            // 
            // Date_Stock
            // 
            this.Date_Stock.DataPropertyName = "Date_Stock";
            this.Date_Stock.HeaderText = "Date_Stock";
            this.Date_Stock.Name = "Date_Stock";
            this.Date_Stock.Width = 70;
            // 
            // Cost_Average
            // 
            this.Cost_Average.DataPropertyName = "Cost_Average";
            this.Cost_Average.HeaderText = "Cost_Average";
            this.Cost_Average.Name = "Cost_Average";
            this.Cost_Average.Width = 80;
            // 
            // Price
            // 
            this.Price.DataPropertyName = "Price";
            this.Price.HeaderText = "Price";
            this.Price.Name = "Price";
            this.Price.Width = 60;
            // 
            // Date_Price
            // 
            this.Date_Price.DataPropertyName = "Date_Price";
            this.Date_Price.HeaderText = "Date_Price";
            this.Date_Price.Name = "Date_Price";
            this.Date_Price.Width = 70;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.button6);
            this.tabPage3.Controls.Add(hKLabel);
            this.tabPage3.Controls.Add(this.hKTextBox);
            this.tabPage3.Controls.Add(rMBLabel);
            this.tabPage3.Controls.Add(this.rMBTextBox);
            this.tabPage3.Location = new System.Drawing.Point(4, 21);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(950, 654);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "匯率維護";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(75, 97);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(75, 23);
            this.button6.TabIndex = 4;
            this.button6.Text = "更新匯率";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click_1);
            // 
            // hKTextBox
            // 
            this.hKTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sALES_DOCCURBindingSource, "HK", true));
            this.hKTextBox.Location = new System.Drawing.Point(63, 57);
            this.hKTextBox.Name = "hKTextBox";
            this.hKTextBox.Size = new System.Drawing.Size(100, 22);
            this.hKTextBox.TabIndex = 3;
            // 
            // sALES_DOCCURBindingSource
            // 
            this.sALES_DOCCURBindingSource.DataMember = "SALES_DOCCUR";
            this.sALES_DOCCURBindingSource.DataSource = this.wh;
            // 
            // wh
            // 
            this.wh.DataSetName = "wh";
            this.wh.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // rMBTextBox
            // 
            this.rMBTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sALES_DOCCURBindingSource, "RMB", true));
            this.rMBTextBox.Location = new System.Drawing.Point(63, 15);
            this.rMBTextBox.Name = "rMBTextBox";
            this.rMBTextBox.Size = new System.Drawing.Size(100, 22);
            this.rMBTextBox.TabIndex = 1;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.comboBox2);
            this.tabPage4.Controls.Add(this.dataGridView4);
            this.tabPage4.Controls.Add(this.button11);
            this.tabPage4.Controls.Add(this.button7);
            this.tabPage4.Location = new System.Drawing.Point(4, 21);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(950, 654);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "銷售月報";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "ASP",
            "QTY",
            "REV",
            "GP%"});
            this.comboBox2.Location = new System.Drawing.Point(176, 6);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(90, 20);
            this.comboBox2.TabIndex = 6;
            this.comboBox2.SelectedValueChanged += new System.EventHandler(this.comboBox2_SelectedValueChanged);
            // 
            // dataGridView4
            // 
            this.dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView4.Location = new System.Drawing.Point(3, 34);
            this.dataGridView4.Name = "dataGridView4";
            this.dataGridView4.RowTemplate.Height = 24;
            this.dataGridView4.Size = new System.Drawing.Size(477, 524);
            this.dataGridView4.TabIndex = 5;
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(84, 6);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(75, 22);
            this.button11.TabIndex = 4;
            this.button11.Text = "產生報表";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(6, 6);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(72, 22);
            this.button7.TabIndex = 0;
            this.button7.Text = "匯入EXCEL";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // sALES_DOCCURTableAdapter1
            // 
            this.sALES_DOCCURTableAdapter1.ClearBeforeFill = true;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "txt";
            this.openFileDialog1.Filter = "txt|*.txt|Excel|*.xls";
            this.openFileDialog1.Title = "請選取檔案";
            // 
            // Judy
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(973, 703);
            this.Controls.Add(this.tabControl1);
            this.Name = "Judy";
            this.Text = "費用計算";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sALES_DOCCURBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.wh)).EndInit();
            this.tabPage4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.TabPage tabPage3;
        private ACME.ACMEDataSet.wh wh;
        private ACME.ACMEDataSet.whTableAdapters.SALES_DOCCURTableAdapter sALES_DOCCURTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn ITEMCODE;
        private System.Windows.Forms.DataGridViewTextBoxColumn Model;
        private System.Windows.Forms.DataGridViewTextBoxColumn Version;
        private System.Windows.Forms.DataGridViewTextBoxColumn Grade;
        private System.Windows.Forms.DataGridViewTextBoxColumn Moving;
        private System.Windows.Forms.DataGridViewTextBoxColumn COST;
        private System.Windows.Forms.DataGridViewTextBoxColumn AddDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn Qty_Stock;
        private System.Windows.Forms.DataGridViewTextBoxColumn Cost_Stock;
        private System.Windows.Forms.DataGridViewTextBoxColumn Date_Stock;
        private System.Windows.Forms.DataGridViewTextBoxColumn Cost_Average;
        private System.Windows.Forms.DataGridViewTextBoxColumn Price;
        private System.Windows.Forms.DataGridViewTextBoxColumn Date_Price;
        private System.Windows.Forms.BindingSource sALES_DOCCURBindingSource;
        private ACME.ACMEDataSet.whTableAdapters.SALES_DOCCURTableAdapter sALES_DOCCURTableAdapter1;
        private System.Windows.Forms.TextBox hKTextBox;
        private System.Windows.Forms.TextBox rMBTextBox;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.DataGridView dataGridView4;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;

    }
}

