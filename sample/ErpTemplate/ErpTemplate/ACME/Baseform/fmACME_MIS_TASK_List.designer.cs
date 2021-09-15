 namespace ACME
 {
     partial class fmACME_MIS_TASK_List
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TextBoxAcDate2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TextBoxAcDate1 = new System.Windows.Forms.TextBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.TextBoxStartDate2 = new System.Windows.Forms.TextBox();
            this.lblStartDate = new System.Windows.Forms.Label();
            this.TextBoxStartDate1 = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.LinkLabel();
            this.btnQuery = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.gvData = new System.Windows.Forms.DataGridView();
            this.ColID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BU = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.UNIT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColKind = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EDIT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColTask = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColStartDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColEndDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColAcDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColUpdateUser = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colEdit = new System.Windows.Forms.DataGridViewImageColumn();
            this.dataGridViewImageColumn1 = new System.Windows.Forms.DataGridViewImageColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvData)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.TextBoxAcDate2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.TextBoxAcDate1);
            this.groupBox1.Controls.Add(this.linkLabel1);
            this.groupBox1.Controls.Add(this.groupBox4);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.TextBoxStartDate2);
            this.groupBox1.Controls.Add(this.lblStartDate);
            this.groupBox1.Controls.Add(this.TextBoxStartDate1);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.btnQuery);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(866, 137);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "查詢條件";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(510, 65);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 28;
            this.button1.Text = "EXCEL";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(513, 96);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(72, 16);
            this.checkBox1.TabIndex = 27;
            this.checkBox1.Text = "工作週報";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(224, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 12);
            this.label2.TabIndex = 25;
            this.label2.Text = "完成日期(迄)";
            // 
            // TextBoxAcDate2
            // 
            this.TextBoxAcDate2.Location = new System.Drawing.Point(303, 91);
            this.TextBoxAcDate2.Name = "TextBoxAcDate2";
            this.TextBoxAcDate2.Size = new System.Drawing.Size(100, 22);
            this.TextBoxAcDate2.TabIndex = 26;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 91);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 12);
            this.label3.TabIndex = 23;
            this.label3.Text = "完成日期(起)";
            // 
            // TextBoxAcDate1
            // 
            this.TextBoxAcDate1.Location = new System.Drawing.Point(106, 91);
            this.TextBoxAcDate1.Name = "TextBoxAcDate1";
            this.TextBoxAcDate1.Size = new System.Drawing.Size(100, 22);
            this.TextBoxAcDate1.TabIndex = 24;
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(766, 101);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(88, 12);
            this.linkLabel1.TabIndex = 22;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "工作週報(Word)";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.btnAdd_LinkClicked);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.radioButton3);
            this.groupBox4.Controls.Add(this.radioButton1);
            this.groupBox4.Controls.Add(this.radioButton2);
            this.groupBox4.Location = new System.Drawing.Point(224, 21);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(200, 38);
            this.groupBox4.TabIndex = 21;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "狀態";
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(122, 16);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(47, 16);
            this.radioButton3.TabIndex = 11;
            this.radioButton3.Text = "全部";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.Location = new System.Drawing.Point(12, 16);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(47, 16);
            this.radioButton1.TabIndex = 9;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "未結";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(65, 16);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(47, 16);
            this.radioButton2.TabIndex = 10;
            this.radioButton2.Text = "已結";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(224, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 12);
            this.label1.TabIndex = 19;
            this.label1.Text = "開始日期(迄)";
            // 
            // TextBoxStartDate2
            // 
            this.TextBoxStartDate2.Location = new System.Drawing.Point(303, 65);
            this.TextBoxStartDate2.Name = "TextBoxStartDate2";
            this.TextBoxStartDate2.Size = new System.Drawing.Size(100, 22);
            this.TextBoxStartDate2.TabIndex = 20;
            // 
            // lblStartDate
            // 
            this.lblStartDate.AutoSize = true;
            this.lblStartDate.Location = new System.Drawing.Point(27, 65);
            this.lblStartDate.Name = "lblStartDate";
            this.lblStartDate.Size = new System.Drawing.Size(73, 12);
            this.lblStartDate.TabIndex = 17;
            this.lblStartDate.Text = "開始日期(起)";
            // 
            // TextBoxStartDate1
            // 
            this.TextBoxStartDate1.Location = new System.Drawing.Point(106, 65);
            this.TextBoxStartDate1.Name = "TextBoxStartDate1";
            this.TextBoxStartDate1.Size = new System.Drawing.Size(100, 22);
            this.TextBoxStartDate1.TabIndex = 18;
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAdd.AutoSize = true;
            this.btnAdd.Location = new System.Drawing.Point(814, 21);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(29, 12);
            this.btnAdd.TabIndex = 16;
            this.btnAdd.TabStop = true;
            this.btnAdd.Text = "新增";
            this.btnAdd.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.btnAdd_LinkClicked);
            // 
            // btnQuery
            // 
            this.btnQuery.Image = global::ACME.Properties.Resources.bnQuery_Image;
            this.btnQuery.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnQuery.Location = new System.Drawing.Point(429, 65);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(61, 48);
            this.btnQuery.TabIndex = 15;
            this.btnQuery.Text = "查詢";
            this.btnQuery.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.gvData);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 137);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(866, 303);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "資料列表";
            // 
            // gvData
            // 
            this.gvData.AllowUserToAddRows = false;
            this.gvData.AllowUserToDeleteRows = false;
            this.gvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.gvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColID,
            this.BU,
            this.UNIT,
            this.ColKind,
            this.EDIT,
            this.ColTask,
            this.ColStartDate,
            this.ColEndDate,
            this.ColAcDate,
            this.ColUpdateUser,
            this.colEdit});
            this.gvData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gvData.Location = new System.Drawing.Point(3, 18);
            this.gvData.Name = "gvData";
            this.gvData.ReadOnly = true;
            this.gvData.RowTemplate.Height = 24;
            this.gvData.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gvData.Size = new System.Drawing.Size(860, 282);
            this.gvData.TabIndex = 0;
            this.gvData.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gvData_CellClick);
            // 
            // ColID
            // 
            this.ColID.DataPropertyName = "ID";
            this.ColID.HeaderText = "ID";
            this.ColID.Name = "ColID";
            this.ColID.ReadOnly = true;
            this.ColID.Width = 42;
            // 
            // BU
            // 
            this.BU.DataPropertyName = "BU";
            this.BU.HeaderText = "BU";
            this.BU.Name = "BU";
            this.BU.ReadOnly = true;
            this.BU.Width = 46;
            // 
            // UNIT
            // 
            this.UNIT.DataPropertyName = "UNIT";
            this.UNIT.HeaderText = "單位";
            this.UNIT.Name = "UNIT";
            this.UNIT.ReadOnly = true;
            this.UNIT.Width = 54;
            // 
            // ColKind
            // 
            this.ColKind.DataPropertyName = "Kind";
            this.ColKind.HeaderText = "種類";
            this.ColKind.Name = "ColKind";
            this.ColKind.ReadOnly = true;
            this.ColKind.Width = 54;
            // 
            // EDIT
            // 
            this.EDIT.DataPropertyName = "EDIT";
            this.EDIT.HeaderText = "動作";
            this.EDIT.Name = "EDIT";
            this.EDIT.ReadOnly = true;
            this.EDIT.Width = 54;
            // 
            // ColTask
            // 
            this.ColTask.DataPropertyName = "Task";
            this.ColTask.HeaderText = "工作項目";
            this.ColTask.Name = "ColTask";
            this.ColTask.ReadOnly = true;
            this.ColTask.Width = 78;
            // 
            // ColStartDate
            // 
            this.ColStartDate.DataPropertyName = "StartDate";
            this.ColStartDate.HeaderText = "開始日期";
            this.ColStartDate.Name = "ColStartDate";
            this.ColStartDate.ReadOnly = true;
            this.ColStartDate.Width = 78;
            // 
            // ColEndDate
            // 
            this.ColEndDate.DataPropertyName = "EndDate";
            this.ColEndDate.HeaderText = "預計完成日";
            this.ColEndDate.Name = "ColEndDate";
            this.ColEndDate.ReadOnly = true;
            this.ColEndDate.Width = 90;
            // 
            // ColAcDate
            // 
            this.ColAcDate.DataPropertyName = "AcDate";
            this.ColAcDate.HeaderText = "實際完成日";
            this.ColAcDate.Name = "ColAcDate";
            this.ColAcDate.ReadOnly = true;
            this.ColAcDate.Width = 90;
            // 
            // ColUpdateUser
            // 
            this.ColUpdateUser.DataPropertyName = "Owner";
            this.ColUpdateUser.HeaderText = "所有人";
            this.ColUpdateUser.Name = "ColUpdateUser";
            this.ColUpdateUser.ReadOnly = true;
            this.ColUpdateUser.Width = 66;
            // 
            // colEdit
            // 
            this.colEdit.HeaderText = "";
            this.colEdit.Image = global::ACME.Properties.Resources.addfile1;
            this.colEdit.Name = "colEdit";
            this.colEdit.ReadOnly = true;
            this.colEdit.Width = 5;
            // 
            // dataGridViewImageColumn1
            // 
            this.dataGridViewImageColumn1.HeaderText = "";
            this.dataGridViewImageColumn1.Name = "dataGridViewImageColumn1";
            this.dataGridViewImageColumn1.ReadOnly = true;
            this.dataGridViewImageColumn1.Width = 32;
            // 
            // fmACME_MIS_TASK_List
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(866, 440);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "fmACME_MIS_TASK_List";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "工作週報";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gvData)).EndInit();
            this.ResumeLayout(false);

         }
 
         #endregion
 
         private System.Windows.Forms.GroupBox groupBox1;
         private System.Windows.Forms.GroupBox groupBox2;
         private System.Windows.Forms.DataGridView gvData;
         private System.Windows.Forms.Button btnQuery;
         private System.Windows.Forms.DataGridViewImageColumn dataGridViewImageColumn1;
         private System.Windows.Forms.LinkLabel btnAdd;
         private System.Windows.Forms.Label lblStartDate;
         private System.Windows.Forms.TextBox TextBoxStartDate1;
         private System.Windows.Forms.Label label1;
         private System.Windows.Forms.TextBox TextBoxStartDate2;
         private System.Windows.Forms.GroupBox groupBox4;
         private System.Windows.Forms.RadioButton radioButton3;
         private System.Windows.Forms.RadioButton radioButton1;
         private System.Windows.Forms.RadioButton radioButton2;
         private System.Windows.Forms.LinkLabel linkLabel1;
         private System.Windows.Forms.Label label2;
         private System.Windows.Forms.TextBox TextBoxAcDate2;
         private System.Windows.Forms.Label label3;
         private System.Windows.Forms.TextBox TextBoxAcDate1;
         private System.Windows.Forms.CheckBox checkBox1;
         private System.Windows.Forms.Button button1;
         private System.Windows.Forms.DataGridViewTextBoxColumn ColID;
         private System.Windows.Forms.DataGridViewTextBoxColumn BU;
         private System.Windows.Forms.DataGridViewTextBoxColumn UNIT;
         private System.Windows.Forms.DataGridViewTextBoxColumn ColKind;
         private System.Windows.Forms.DataGridViewTextBoxColumn EDIT;
         private System.Windows.Forms.DataGridViewTextBoxColumn ColTask;
         private System.Windows.Forms.DataGridViewTextBoxColumn ColStartDate;
         private System.Windows.Forms.DataGridViewTextBoxColumn ColEndDate;
         private System.Windows.Forms.DataGridViewTextBoxColumn ColAcDate;
         private System.Windows.Forms.DataGridViewTextBoxColumn ColUpdateUser;
         private System.Windows.Forms.DataGridViewImageColumn colEdit;
     }
 }

