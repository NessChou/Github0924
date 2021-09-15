namespace ACME
{
    partial class CART2_LED
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
            System.Windows.Forms.Label iTEMCODELabel;
            System.Windows.Forms.Label iTEMNAMELabel;
            System.Windows.Forms.Label cT_QTYLabel;
            System.Windows.Forms.Label cT_NWLabel;
            System.Windows.Forms.Label cT_GWLabel;
            System.Windows.Forms.Label uNITLabel;
            System.Windows.Forms.Label memoLabel;
            this.wh = new ACME.ACMEDataSet.wh();
            this.cART_LEDBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.cART_LEDTableAdapter = new ACME.ACMEDataSet.whTableAdapters.CART_LEDTableAdapter();
            this.iTEMCODETextBox = new System.Windows.Forms.TextBox();
            this.iTEMNAMETextBox = new System.Windows.Forms.TextBox();
            this.cT_QTYTextBox = new System.Windows.Forms.TextBox();
            this.cT_NWTextBox = new System.Windows.Forms.TextBox();
            this.cT_GWTextBox = new System.Windows.Forms.TextBox();
            this.uNITTextBox = new System.Windows.Forms.TextBox();
            this.memoTextBox = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cREATE_USERTextBox = new System.Windows.Forms.TextBox();
            this.cREATE_DATETextBox = new System.Windows.Forms.TextBox();
            this.iDTextBox = new System.Windows.Forms.TextBox();
            this.uPDATE_DATETextBox = new System.Windows.Forms.TextBox();
            this.uPDATE_USERTextBox = new System.Windows.Forms.TextBox();
            iTEMCODELabel = new System.Windows.Forms.Label();
            iTEMNAMELabel = new System.Windows.Forms.Label();
            cT_QTYLabel = new System.Windows.Forms.Label();
            cT_NWLabel = new System.Windows.Forms.Label();
            cT_GWLabel = new System.Windows.Forms.Label();
            uNITLabel = new System.Windows.Forms.Label();
            memoLabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.wh)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cART_LEDBindingSource)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.uPDATE_USERTextBox);
            this.panel1.Controls.Add(this.uPDATE_DATETextBox);
            this.panel1.Controls.Add(this.iDTextBox);
            this.panel1.Controls.Add(this.cREATE_DATETextBox);
            this.panel1.Controls.Add(this.cREATE_USERTextBox);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Size = new System.Drawing.Size(727, 255);
            this.panel1.Controls.SetChildIndex(this.groupBox1, 0);
            this.panel1.Controls.SetChildIndex(this.groupBox2, 0);
            this.panel1.Controls.SetChildIndex(this.cREATE_USERTextBox, 0);
            this.panel1.Controls.SetChildIndex(this.cREATE_DATETextBox, 0);
            this.panel1.Controls.SetChildIndex(this.iDTextBox, 0);
            this.panel1.Controls.SetChildIndex(this.uPDATE_DATETextBox, 0);
            this.panel1.Controls.SetChildIndex(this.uPDATE_USERTextBox, 0);
            // 
            // iTEMCODELabel
            // 
            iTEMCODELabel.AutoSize = true;
            iTEMCODELabel.Location = new System.Drawing.Point(6, 19);
            iTEMCODELabel.Name = "iTEMCODELabel";
            iTEMCODELabel.Size = new System.Drawing.Size(53, 12);
            iTEMCODELabel.TabIndex = 1;
            iTEMCODELabel.Text = "項目號碼";
            // 
            // iTEMNAMELabel
            // 
            iTEMNAMELabel.AutoSize = true;
            iTEMNAMELabel.Location = new System.Drawing.Point(291, 19);
            iTEMNAMELabel.Name = "iTEMNAMELabel";
            iTEMNAMELabel.Size = new System.Drawing.Size(53, 12);
            iTEMNAMELabel.TabIndex = 3;
            iTEMNAMELabel.Text = "項目說明";
            // 
            // cT_QTYLabel
            // 
            cT_QTYLabel.AutoSize = true;
            cT_QTYLabel.Location = new System.Drawing.Point(21, 35);
            cT_QTYLabel.Name = "cT_QTYLabel";
            cT_QTYLabel.Size = new System.Drawing.Size(28, 12);
            cT_QTYLabel.TabIndex = 5;
            cT_QTYLabel.Text = "QTY";
            // 
            // cT_NWLabel
            // 
            cT_NWLabel.AutoSize = true;
            cT_NWLabel.Location = new System.Drawing.Point(197, 35);
            cT_NWLabel.Name = "cT_NWLabel";
            cT_NWLabel.Size = new System.Drawing.Size(24, 12);
            cT_NWLabel.TabIndex = 7;
            cT_NWLabel.Text = "NW";
            // 
            // cT_GWLabel
            // 
            cT_GWLabel.AutoSize = true;
            cT_GWLabel.Location = new System.Drawing.Point(375, 35);
            cT_GWLabel.Name = "cT_GWLabel";
            cT_GWLabel.Size = new System.Drawing.Size(24, 12);
            cT_GWLabel.TabIndex = 9;
            cT_GWLabel.Text = "GW";
            // 
            // uNITLabel
            // 
            uNITLabel.AutoSize = true;
            uNITLabel.Location = new System.Drawing.Point(30, 92);
            uNITLabel.Name = "uNITLabel";
            uNITLabel.Size = new System.Drawing.Size(29, 12);
            uNITLabel.TabIndex = 11;
            uNITLabel.Text = "單位";
            // 
            // memoLabel
            // 
            memoLabel.AutoSize = true;
            memoLabel.Location = new System.Drawing.Point(134, 92);
            memoLabel.Name = "memoLabel";
            memoLabel.Size = new System.Drawing.Size(29, 12);
            memoLabel.TabIndex = 13;
            memoLabel.Text = "備註";
            // 
            // wh
            // 
            this.wh.DataSetName = "wh";
            this.wh.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // cART_LEDBindingSource
            // 
            this.cART_LEDBindingSource.DataMember = "CART_LED";
            this.cART_LEDBindingSource.DataSource = this.wh;
            // 
            // cART_LEDTableAdapter
            // 
            this.cART_LEDTableAdapter.ClearBeforeFill = true;
            // 
            // iTEMCODETextBox
            // 
            this.iTEMCODETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "ITEMCODE", true));
            this.iTEMCODETextBox.Location = new System.Drawing.Point(65, 15);
            this.iTEMCODETextBox.Name = "iTEMCODETextBox";
            this.iTEMCODETextBox.Size = new System.Drawing.Size(219, 22);
            this.iTEMCODETextBox.TabIndex = 2;
            // 
            // iTEMNAMETextBox
            // 
            this.iTEMNAMETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "ITEMNAME", true));
            this.iTEMNAMETextBox.Location = new System.Drawing.Point(350, 15);
            this.iTEMNAMETextBox.Multiline = true;
            this.iTEMNAMETextBox.Name = "iTEMNAMETextBox";
            this.iTEMNAMETextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.iTEMNAMETextBox.Size = new System.Drawing.Size(344, 68);
            this.iTEMNAMETextBox.TabIndex = 4;
            // 
            // cT_QTYTextBox
            // 
            this.cT_QTYTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "CT_QTY", true));
            this.cT_QTYTextBox.Location = new System.Drawing.Point(76, 32);
            this.cT_QTYTextBox.Name = "cT_QTYTextBox";
            this.cT_QTYTextBox.Size = new System.Drawing.Size(100, 22);
            this.cT_QTYTextBox.TabIndex = 6;
            // 
            // cT_NWTextBox
            // 
            this.cT_NWTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "CT_NW", true));
            this.cT_NWTextBox.Location = new System.Drawing.Point(248, 32);
            this.cT_NWTextBox.Name = "cT_NWTextBox";
            this.cT_NWTextBox.Size = new System.Drawing.Size(100, 22);
            this.cT_NWTextBox.TabIndex = 8;
            // 
            // cT_GWTextBox
            // 
            this.cT_GWTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "CT_GW", true));
            this.cT_GWTextBox.Location = new System.Drawing.Point(426, 32);
            this.cT_GWTextBox.Name = "cT_GWTextBox";
            this.cT_GWTextBox.Size = new System.Drawing.Size(100, 22);
            this.cT_GWTextBox.TabIndex = 10;
            // 
            // uNITTextBox
            // 
            this.uNITTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "UNIT", true));
            this.uNITTextBox.Location = new System.Drawing.Point(65, 89);
            this.uNITTextBox.Name = "uNITTextBox";
            this.uNITTextBox.Size = new System.Drawing.Size(59, 22);
            this.uNITTextBox.TabIndex = 12;
            // 
            // memoTextBox
            // 
            this.memoTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "memo", true));
            this.memoTextBox.Location = new System.Drawing.Point(169, 89);
            this.memoTextBox.Multiline = true;
            this.memoTextBox.Name = "memoTextBox";
            this.memoTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.memoTextBox.Size = new System.Drawing.Size(525, 36);
            this.memoTextBox.TabIndex = 14;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(cT_NWLabel);
            this.groupBox1.Controls.Add(this.cT_NWTextBox);
            this.groupBox1.Controls.Add(this.cT_GWTextBox);
            this.groupBox1.Controls.Add(cT_GWLabel);
            this.groupBox1.Controls.Add(cT_QTYLabel);
            this.groupBox1.Controls.Add(this.cT_QTYTextBox);
            this.groupBox1.Location = new System.Drawing.Point(3, 140);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(709, 75);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "CARTON PACKAGE";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(iTEMCODELabel);
            this.groupBox2.Controls.Add(this.iTEMCODETextBox);
            this.groupBox2.Controls.Add(memoLabel);
            this.groupBox2.Controls.Add(this.memoTextBox);
            this.groupBox2.Controls.Add(this.iTEMNAMETextBox);
            this.groupBox2.Controls.Add(iTEMNAMELabel);
            this.groupBox2.Controls.Add(uNITLabel);
            this.groupBox2.Controls.Add(this.uNITTextBox);
            this.groupBox2.Location = new System.Drawing.Point(3, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(709, 131);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Packing MA";
            // 
            // cREATE_USERTextBox
            // 
            this.cREATE_USERTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "CREATE_USER", true));
            this.cREATE_USERTextBox.Location = new System.Drawing.Point(614, 180);
            this.cREATE_USERTextBox.Name = "cREATE_USERTextBox";
            this.cREATE_USERTextBox.Size = new System.Drawing.Size(0, 22);
            this.cREATE_USERTextBox.TabIndex = 17;
            // 
            // cREATE_DATETextBox
            // 
            this.cREATE_DATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "CREATE_DATE", true));
            this.cREATE_DATETextBox.Location = new System.Drawing.Point(430, 180);
            this.cREATE_DATETextBox.Name = "cREATE_DATETextBox";
            this.cREATE_DATETextBox.Size = new System.Drawing.Size(0, 22);
            this.cREATE_DATETextBox.TabIndex = 18;
            // 
            // iDTextBox
            // 
            this.iDTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "ID", true));
            this.iDTextBox.Location = new System.Drawing.Point(337, 181);
            this.iDTextBox.Name = "iDTextBox";
            this.iDTextBox.Size = new System.Drawing.Size(0, 22);
            this.iDTextBox.TabIndex = 19;
            // 
            // uPDATE_DATETextBox
            // 
            this.uPDATE_DATETextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "UPDATE_DATE", true));
            this.uPDATE_DATETextBox.Location = new System.Drawing.Point(87, 172);
            this.uPDATE_DATETextBox.Name = "uPDATE_DATETextBox";
            this.uPDATE_DATETextBox.Size = new System.Drawing.Size(0, 22);
            this.uPDATE_DATETextBox.TabIndex = 20;
            // 
            // uPDATE_USERTextBox
            // 
            this.uPDATE_USERTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cART_LEDBindingSource, "UPDATE_USER", true));
            this.uPDATE_USERTextBox.Location = new System.Drawing.Point(187, 172);
            this.uPDATE_USERTextBox.Name = "uPDATE_USERTextBox";
            this.uPDATE_USERTextBox.Size = new System.Drawing.Size(0, 22);
            this.uPDATE_USERTextBox.TabIndex = 21;
            // 
            // CART2_LED
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(727, 290);
            this.Name = "CART2_LED";
            this.Text = "包裝規格修改(LED)";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.wh)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cART_LEDBindingSource)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACME.ACMEDataSet.wh wh;
        private System.Windows.Forms.BindingSource cART_LEDBindingSource;
        private ACME.ACMEDataSet.whTableAdapters.CART_LEDTableAdapter cART_LEDTableAdapter;
        private System.Windows.Forms.TextBox iTEMNAMETextBox;
        private System.Windows.Forms.TextBox iTEMCODETextBox;
        private System.Windows.Forms.TextBox cT_QTYTextBox;
        private System.Windows.Forms.TextBox uNITTextBox;
        private System.Windows.Forms.TextBox cT_GWTextBox;
        private System.Windows.Forms.TextBox cT_NWTextBox;
        private System.Windows.Forms.TextBox memoTextBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox cREATE_DATETextBox;
        private System.Windows.Forms.TextBox cREATE_USERTextBox;
        private System.Windows.Forms.TextBox iDTextBox;
        private System.Windows.Forms.TextBox uPDATE_USERTextBox;
        private System.Windows.Forms.TextBox uPDATE_DATETextBox;
    }
}
