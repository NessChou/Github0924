namespace ACME.CRM
{
    partial class CrmMis
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
            System.Windows.Forms.Label docDateLabel;
            System.Windows.Forms.Label issueKindLabel;
            System.Windows.Forms.Label issueDescLabel;
            System.Windows.Forms.Label actionDescLabel;
            System.Windows.Forms.Label userCodeLabel;
            System.Windows.Forms.Label actionFlagLabel;
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.cRM = new ACME.CRM.CRM();
            this.aCME_MISBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.aCME_MISTableAdapter = new ACME.CRM.CRMTableAdapters.ACME_MISTableAdapter();
            this.aCME_MISDataGridView = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.docDateTextBox = new System.Windows.Forms.TextBox();
            this.issueDescTextBox = new System.Windows.Forms.TextBox();
            this.actionDescTextBox = new System.Windows.Forms.TextBox();
            this.userCodeTextBox = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.exBindingNavigator1 = new ACME.exBindingNavigator();
            docDateLabel = new System.Windows.Forms.Label();
            issueKindLabel = new System.Windows.Forms.Label();
            issueDescLabel = new System.Windows.Forms.Label();
            actionDescLabel = new System.Windows.Forms.Label();
            userCodeLabel = new System.Windows.Forms.Label();
            actionFlagLabel = new System.Windows.Forms.Label();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.cRM)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_MISBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_MISDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(587, 63);
            this.panel1.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.comboBox2);
            this.panel2.Controls.Add(this.comboBox1);
            this.panel2.Controls.Add(docDateLabel);
            this.panel2.Controls.Add(this.docDateTextBox);
            this.panel2.Controls.Add(issueKindLabel);
            this.panel2.Controls.Add(issueDescLabel);
            this.panel2.Controls.Add(this.issueDescTextBox);
            this.panel2.Controls.Add(actionDescLabel);
            this.panel2.Controls.Add(this.actionDescTextBox);
            this.panel2.Controls.Add(userCodeLabel);
            this.panel2.Controls.Add(this.userCodeTextBox);
            this.panel2.Controls.Add(actionFlagLabel);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.exBindingNavigator1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 440);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(587, 260);
            this.panel2.TabIndex = 1;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.aCME_MISDataGridView);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 63);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(587, 377);
            this.panel3.TabIndex = 2;
            // 
            // cRM
            // 
            this.cRM.DataSetName = "CRM";
            this.cRM.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // aCME_MISBindingSource
            // 
            this.aCME_MISBindingSource.DataMember = "ACME_MIS";
            this.aCME_MISBindingSource.DataSource = this.cRM;
            // 
            // aCME_MISTableAdapter
            // 
            this.aCME_MISTableAdapter.ClearBeforeFill = true;
            // 
            // aCME_MISDataGridView
            // 
            this.aCME_MISDataGridView.AllowUserToAddRows = false;
            this.aCME_MISDataGridView.AllowUserToDeleteRows = false;
            this.aCME_MISDataGridView.AutoGenerateColumns = false;
            this.aCME_MISDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7});
            this.aCME_MISDataGridView.DataSource = this.aCME_MISBindingSource;
            this.aCME_MISDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.aCME_MISDataGridView.Location = new System.Drawing.Point(0, 0);
            this.aCME_MISDataGridView.Name = "aCME_MISDataGridView";
            this.aCME_MISDataGridView.ReadOnly = true;
            this.aCME_MISDataGridView.RowTemplate.Height = 24;
            this.aCME_MISDataGridView.Size = new System.Drawing.Size(587, 377);
            this.aCME_MISDataGridView.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(490, 225);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "關閉";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // docDateLabel
            // 
            docDateLabel.AutoSize = true;
            docDateLabel.Location = new System.Drawing.Point(49, 50);
            docDateLabel.Name = "docDateLabel";
            docDateLabel.Size = new System.Drawing.Size(53, 12);
            docDateLabel.TabIndex = 4;
            docDateLabel.Text = "回報日期";
            // 
            // docDateTextBox
            // 
            this.docDateTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.aCME_MISBindingSource, "DocDate", true));
            this.docDateTextBox.Location = new System.Drawing.Point(106, 47);
            this.docDateTextBox.Name = "docDateTextBox";
            this.docDateTextBox.ReadOnly = true;
            this.docDateTextBox.Size = new System.Drawing.Size(100, 22);
            this.docDateTextBox.TabIndex = 5;
            // 
            // issueKindLabel
            // 
            issueKindLabel.AutoSize = true;
            issueKindLabel.Location = new System.Drawing.Point(226, 50);
            issueKindLabel.Name = "issueKindLabel";
            issueKindLabel.Size = new System.Drawing.Size(53, 12);
            issueKindLabel.TabIndex = 6;
            issueKindLabel.Text = "回報種類";
            // 
            // issueDescLabel
            // 
            issueDescLabel.AutoSize = true;
            issueDescLabel.Location = new System.Drawing.Point(49, 89);
            issueDescLabel.Name = "issueDescLabel";
            issueDescLabel.Size = new System.Drawing.Size(53, 12);
            issueDescLabel.TabIndex = 8;
            issueDescLabel.Text = "問題描述";
            // 
            // issueDescTextBox
            // 
            this.issueDescTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.aCME_MISBindingSource, "IssueDesc", true));
            this.issueDescTextBox.Location = new System.Drawing.Point(106, 86);
            this.issueDescTextBox.Multiline = true;
            this.issueDescTextBox.Name = "issueDescTextBox";
            this.issueDescTextBox.Size = new System.Drawing.Size(459, 43);
            this.issueDescTextBox.TabIndex = 9;
            // 
            // actionDescLabel
            // 
            actionDescLabel.AutoSize = true;
            actionDescLabel.Location = new System.Drawing.Point(47, 139);
            actionDescLabel.Name = "actionDescLabel";
            actionDescLabel.Size = new System.Drawing.Size(53, 12);
            actionDescLabel.TabIndex = 10;
            actionDescLabel.Text = "行案方案";
            // 
            // actionDescTextBox
            // 
            this.actionDescTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.aCME_MISBindingSource, "ActionDesc", true));
            this.actionDescTextBox.Location = new System.Drawing.Point(106, 136);
            this.actionDescTextBox.Multiline = true;
            this.actionDescTextBox.Name = "actionDescTextBox";
            this.actionDescTextBox.Size = new System.Drawing.Size(459, 38);
            this.actionDescTextBox.TabIndex = 11;
            // 
            // userCodeLabel
            // 
            userCodeLabel.AutoSize = true;
            userCodeLabel.Location = new System.Drawing.Point(59, 180);
            userCodeLabel.Name = "userCodeLabel";
            userCodeLabel.Size = new System.Drawing.Size(41, 12);
            userCodeLabel.TabIndex = 12;
            userCodeLabel.Text = "回報人";
            // 
            // userCodeTextBox
            // 
            this.userCodeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.aCME_MISBindingSource, "UserCode", true));
            this.userCodeTextBox.Location = new System.Drawing.Point(106, 180);
            this.userCodeTextBox.Name = "userCodeTextBox";
            this.userCodeTextBox.ReadOnly = true;
            this.userCodeTextBox.Size = new System.Drawing.Size(100, 22);
            this.userCodeTextBox.TabIndex = 13;
            // 
            // actionFlagLabel
            // 
            actionFlagLabel.AutoSize = true;
            actionFlagLabel.Location = new System.Drawing.Point(59, 218);
            actionFlagLabel.Name = "actionFlagLabel";
            actionFlagLabel.Size = new System.Drawing.Size(53, 12);
            actionFlagLabel.TabIndex = 14;
            actionFlagLabel.Text = "結案類別";
            // 
            // comboBox1
            // 
            this.comboBox1.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.aCME_MISBindingSource, "IssueKind", true));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(289, 47);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 20);
            this.comboBox1.TabIndex = 15;
            // 
            // comboBox2
            // 
            this.comboBox2.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.aCME_MISBindingSource, "ActionFlag", true));
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(106, 218);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(121, 20);
            this.comboBox2.TabIndex = 16;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "DocDate";
            this.dataGridViewTextBoxColumn2.HeaderText = "回報日期";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "IssueKind";
            this.dataGridViewTextBoxColumn3.HeaderText = "回報種類";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "IssueDesc";
            this.dataGridViewTextBoxColumn4.HeaderText = "問題描述";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "ActionDesc";
            this.dataGridViewTextBoxColumn5.HeaderText = "行動方案";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "UserCode";
            this.dataGridViewTextBoxColumn6.HeaderText = "回報人";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "ActionFlag";
            this.dataGridViewTextBoxColumn7.HeaderText = "結案類別";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.ReadOnly = true;
            // 
            // exBindingNavigator1
            // 
            this.exBindingNavigator1.AutoFillFlag = true;
            this.exBindingNavigator1.AutoSaveFlag = false;
            this.exBindingNavigator1.BindingSource = this.aCME_MISBindingSource;
            this.exBindingNavigator1.DataTable = this.cRM.ACME_MIS;
            this.exBindingNavigator1.DisplayMember = null;
            this.exBindingNavigator1.Dock = System.Windows.Forms.DockStyle.Top;
            this.exBindingNavigator1.IsDataDirty = false;
            this.exBindingNavigator1.Location = new System.Drawing.Point(0, 0);
            this.exBindingNavigator1.Name = "exBindingNavigator1";
            this.exBindingNavigator1.ParentBindingSource = null;
            this.exBindingNavigator1.Size = new System.Drawing.Size(587, 26);
            this.exBindingNavigator1.TabIndex = 0;
            this.exBindingNavigator1.BeforeDelete += new ACME.exBindingNavigator.BeforeDeleteEventHandler(this.exBindingNavigator1_BeforeDelete);
            this.exBindingNavigator1.AfterNew += new ACME.exBindingNavigator.AfterNewEventHandler(this.exBindingNavigator1_AfterNew);
            // 
            // CrmMis
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 700);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "CrmMis";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CrmMis";
            this.Load += new System.EventHandler(this.CrmMis_Load);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.cRM)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_MISBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_MISDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private exBindingNavigator exBindingNavigator1;
        private System.Windows.Forms.Panel panel3;
        private CRM cRM;
        private System.Windows.Forms.BindingSource aCME_MISBindingSource;
        private ACME.CRM.CRMTableAdapters.ACME_MISTableAdapter aCME_MISTableAdapter;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView aCME_MISDataGridView;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox docDateTextBox;
        private System.Windows.Forms.TextBox issueDescTextBox;
        private System.Windows.Forms.TextBox actionDescTextBox;
        private System.Windows.Forms.TextBox userCodeTextBox;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
    }
}