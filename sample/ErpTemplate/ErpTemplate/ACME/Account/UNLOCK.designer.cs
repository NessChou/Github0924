 namespace ACME
 {
     partial class UNLOCK
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
             System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UNLOCK));
             this.TextBoxDocEntry = new System.Windows.Forms.TextBox();
             this.lblID = new System.Windows.Forms.Label();
             this.TextBoxCardCode = new System.Windows.Forms.TextBox();
             this.btnCancel = new System.Windows.Forms.Button();
             this.btnSave = new System.Windows.Forms.Button();
             this.btnDelete = new System.Windows.Forms.Button();
             this.dpStartDate = new System.Windows.Forms.DateTimePicker();
             this.lblStartDate = new System.Windows.Forms.Label();
             this.TextBoxStartDate = new System.Windows.Forms.TextBox();
             this.dpEndDate = new System.Windows.Forms.DateTimePicker();
             this.lblAcDate = new System.Windows.Forms.Label();
             this.TextBoxEndDate = new System.Windows.Forms.TextBox();
             this.label2 = new System.Windows.Forms.Label();
             this.lblCreateDate = new System.Windows.Forms.Label();
             this.TextBoxCreateDate = new System.Windows.Forms.TextBox();
             this.lblCreateTime = new System.Windows.Forms.Label();
             this.TextBoxCreateTime = new System.Windows.Forms.TextBox();
             this.lblCreateUser = new System.Windows.Forms.Label();
             this.TextBoxHandler = new System.Windows.Forms.TextBox();
             this.TextBoxCardName = new System.Windows.Forms.TextBox();
             this.label1 = new System.Windows.Forms.Label();
             this.button1 = new System.Windows.Forms.Button();
             this.SuspendLayout();
             // 
             // TextBoxDocEntry
             // 
             this.TextBoxDocEntry.Location = new System.Drawing.Point(110, 20);
             this.TextBoxDocEntry.Name = "TextBoxDocEntry";
             this.TextBoxDocEntry.Size = new System.Drawing.Size(100, 22);
             this.TextBoxDocEntry.TabIndex = 12;
             // 
             // lblID
             // 
             this.lblID.AutoSize = true;
             this.lblID.Location = new System.Drawing.Point(87, 23);
             this.lblID.Name = "lblID";
             this.lblID.Size = new System.Drawing.Size(17, 12);
             this.lblID.TabIndex = 11;
             this.lblID.Text = "ID";
             // 
             // TextBoxCardCode
             // 
             this.TextBoxCardCode.Location = new System.Drawing.Point(110, 48);
             this.TextBoxCardCode.Name = "TextBoxCardCode";
             this.TextBoxCardCode.Size = new System.Drawing.Size(100, 22);
             this.TextBoxCardCode.TabIndex = 31;
             // 
             // btnCancel
             // 
             this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
             this.btnCancel.Image = global::ACME.Properties.Resources.bnExit_Image;
             this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
             this.btnCancel.Location = new System.Drawing.Point(502, 328);
             this.btnCancel.Name = "btnCancel";
             this.btnCancel.Size = new System.Drawing.Size(75, 23);
             this.btnCancel.TabIndex = 29;
             this.btnCancel.Text = "Cancel";
             this.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
             this.btnCancel.UseVisualStyleBackColor = true;
             this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
             // 
             // btnSave
             // 
             this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
             this.btnSave.Image = global::ACME.Properties.Resources.bnEndEdit_Image;
             this.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
             this.btnSave.Location = new System.Drawing.Point(421, 328);
             this.btnSave.Name = "btnSave";
             this.btnSave.Size = new System.Drawing.Size(75, 23);
             this.btnSave.TabIndex = 28;
             this.btnSave.Text = "Save";
             this.btnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
             this.btnSave.UseVisualStyleBackColor = true;
             this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
             // 
             // btnDelete
             // 
             this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
             this.btnDelete.Image = global::ACME.Properties.Resources.bnDelete_Image;
             this.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
             this.btnDelete.Location = new System.Drawing.Point(12, 328);
             this.btnDelete.Name = "btnDelete";
             this.btnDelete.Size = new System.Drawing.Size(75, 23);
             this.btnDelete.TabIndex = 27;
             this.btnDelete.Text = "Delete";
             this.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
             this.btnDelete.UseVisualStyleBackColor = true;
             this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
             // 
             // dpStartDate
             // 
             this.dpStartDate.CustomFormat = "yyyyMMdd";
             this.dpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
             this.dpStartDate.Location = new System.Drawing.Point(206, 104);
             this.dpStartDate.Name = "dpStartDate";
             this.dpStartDate.Size = new System.Drawing.Size(21, 22);
             this.dpStartDate.TabIndex = 43;
             this.dpStartDate.ValueChanged += new System.EventHandler(this.dpStartDate_ValueChanged);
             // 
             // lblStartDate
             // 
             this.lblStartDate.AutoSize = true;
             this.lblStartDate.Location = new System.Drawing.Point(51, 104);
             this.lblStartDate.Name = "lblStartDate";
             this.lblStartDate.Size = new System.Drawing.Size(53, 12);
             this.lblStartDate.TabIndex = 41;
             this.lblStartDate.Text = "開始日期";
             // 
             // TextBoxStartDate
             // 
             this.TextBoxStartDate.Location = new System.Drawing.Point(110, 104);
             this.TextBoxStartDate.Name = "TextBoxStartDate";
             this.TextBoxStartDate.Size = new System.Drawing.Size(100, 22);
             this.TextBoxStartDate.TabIndex = 42;
             // 
             // dpEndDate
             // 
             this.dpEndDate.CustomFormat = "yyyyMMdd";
             this.dpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
             this.dpEndDate.Location = new System.Drawing.Point(206, 130);
             this.dpEndDate.Name = "dpEndDate";
             this.dpEndDate.Size = new System.Drawing.Size(21, 22);
             this.dpEndDate.TabIndex = 46;
             this.dpEndDate.ValueChanged += new System.EventHandler(this.dpEndDate_ValueChanged);
             // 
             // lblAcDate
             // 
             this.lblAcDate.AutoSize = true;
             this.lblAcDate.Location = new System.Drawing.Point(51, 133);
             this.lblAcDate.Name = "lblAcDate";
             this.lblAcDate.Size = new System.Drawing.Size(53, 12);
             this.lblAcDate.TabIndex = 44;
             this.lblAcDate.Text = "結束日期";
             // 
             // TextBoxEndDate
             // 
             this.TextBoxEndDate.Location = new System.Drawing.Point(110, 130);
             this.TextBoxEndDate.Name = "TextBoxEndDate";
             this.TextBoxEndDate.Size = new System.Drawing.Size(100, 22);
             this.TextBoxEndDate.TabIndex = 45;
             // 
             // label2
             // 
             this.label2.AutoSize = true;
             this.label2.Location = new System.Drawing.Point(51, 51);
             this.label2.Name = "label2";
             this.label2.Size = new System.Drawing.Size(53, 12);
             this.label2.TabIndex = 47;
             this.label2.Text = "客戶編號";
             // 
             // lblCreateDate
             // 
             this.lblCreateDate.AutoSize = true;
             this.lblCreateDate.Location = new System.Drawing.Point(51, 172);
             this.lblCreateDate.Name = "lblCreateDate";
             this.lblCreateDate.Size = new System.Drawing.Size(53, 12);
             this.lblCreateDate.TabIndex = 50;
             this.lblCreateDate.Text = "建立日期";
             // 
             // TextBoxCreateDate
             // 
             this.TextBoxCreateDate.Enabled = false;
             this.TextBoxCreateDate.Location = new System.Drawing.Point(110, 172);
             this.TextBoxCreateDate.Name = "TextBoxCreateDate";
             this.TextBoxCreateDate.Size = new System.Drawing.Size(100, 22);
             this.TextBoxCreateDate.TabIndex = 52;
             // 
             // lblCreateTime
             // 
             this.lblCreateTime.AutoSize = true;
             this.lblCreateTime.Location = new System.Drawing.Point(221, 172);
             this.lblCreateTime.Name = "lblCreateTime";
             this.lblCreateTime.Size = new System.Drawing.Size(53, 12);
             this.lblCreateTime.TabIndex = 48;
             this.lblCreateTime.Text = "建立時間";
             // 
             // TextBoxCreateTime
             // 
             this.TextBoxCreateTime.Enabled = false;
             this.TextBoxCreateTime.Location = new System.Drawing.Point(280, 172);
             this.TextBoxCreateTime.Name = "TextBoxCreateTime";
             this.TextBoxCreateTime.Size = new System.Drawing.Size(100, 22);
             this.TextBoxCreateTime.TabIndex = 53;
             // 
             // lblCreateUser
             // 
             this.lblCreateUser.AutoSize = true;
             this.lblCreateUser.Location = new System.Drawing.Point(399, 172);
             this.lblCreateUser.Name = "lblCreateUser";
             this.lblCreateUser.Size = new System.Drawing.Size(41, 12);
             this.lblCreateUser.TabIndex = 49;
             this.lblCreateUser.Text = "建立者";
             // 
             // TextBoxHandler
             // 
             this.TextBoxHandler.Enabled = false;
             this.TextBoxHandler.Location = new System.Drawing.Point(446, 172);
             this.TextBoxHandler.Name = "TextBoxHandler";
             this.TextBoxHandler.Size = new System.Drawing.Size(100, 22);
             this.TextBoxHandler.TabIndex = 51;
             // 
             // TextBoxCardName
             // 
             this.TextBoxCardName.Location = new System.Drawing.Point(110, 76);
             this.TextBoxCardName.Name = "TextBoxCardName";
             this.TextBoxCardName.Size = new System.Drawing.Size(237, 22);
             this.TextBoxCardName.TabIndex = 54;
             // 
             // label1
             // 
             this.label1.AutoSize = true;
             this.label1.Location = new System.Drawing.Point(51, 79);
             this.label1.Name = "label1";
             this.label1.Size = new System.Drawing.Size(53, 12);
             this.label1.TabIndex = 55;
             this.label1.Text = "客戶名稱";
             // 
             // button1
             // 
             this.button1.ForeColor = System.Drawing.SystemColors.ActiveBorder;
             this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
             this.button1.Location = new System.Drawing.Point(216, 51);
             this.button1.Name = "button1";
             this.button1.Size = new System.Drawing.Size(26, 19);
             this.button1.TabIndex = 56;
             this.button1.Text = "...";
             this.button1.UseVisualStyleBackColor = true;
             this.button1.Click += new System.EventHandler(this.button1_Click);
             // 
             // UNLOCK
             // 
             this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
             this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
             this.ClientSize = new System.Drawing.Size(589, 363);
             this.Controls.Add(this.button1);
             this.Controls.Add(this.label1);
             this.Controls.Add(this.TextBoxCardName);
             this.Controls.Add(this.lblCreateDate);
             this.Controls.Add(this.TextBoxCreateDate);
             this.Controls.Add(this.lblCreateTime);
             this.Controls.Add(this.TextBoxCreateTime);
             this.Controls.Add(this.lblCreateUser);
             this.Controls.Add(this.TextBoxHandler);
             this.Controls.Add(this.label2);
             this.Controls.Add(this.dpEndDate);
             this.Controls.Add(this.lblAcDate);
             this.Controls.Add(this.TextBoxEndDate);
             this.Controls.Add(this.dpStartDate);
             this.Controls.Add(this.lblStartDate);
             this.Controls.Add(this.TextBoxStartDate);
             this.Controls.Add(this.TextBoxCardCode);
             this.Controls.Add(this.lblID);
             this.Controls.Add(this.TextBoxDocEntry);
             this.Controls.Add(this.btnCancel);
             this.Controls.Add(this.btnSave);
             this.Controls.Add(this.btnDelete);
             this.Name = "UNLOCK";
             this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
             this.Text = "修";
             this.Load += new System.EventHandler(this.fmACME_MIS_TASK_Load);
             this.ResumeLayout(false);
             this.PerformLayout();

         }
 
         #endregion
 
         private System.Windows.Forms.Label lblID;
         private System.Windows.Forms.TextBox TextBoxDocEntry;
         private System.Windows.Forms.Button btnDelete;
         private System.Windows.Forms.Button btnSave;
         private System.Windows.Forms.Button btnCancel;
         private System.Windows.Forms.TextBox TextBoxCardCode;
         private System.Windows.Forms.DateTimePicker dpStartDate;
         private System.Windows.Forms.Label lblStartDate;
         private System.Windows.Forms.TextBox TextBoxStartDate;
         private System.Windows.Forms.DateTimePicker dpEndDate;
         private System.Windows.Forms.Label lblAcDate;
         private System.Windows.Forms.TextBox TextBoxEndDate;
         private System.Windows.Forms.Label label2;
         private System.Windows.Forms.Label lblCreateDate;
         private System.Windows.Forms.TextBox TextBoxCreateDate;
         private System.Windows.Forms.Label lblCreateTime;
         private System.Windows.Forms.TextBox TextBoxCreateTime;
         private System.Windows.Forms.Label lblCreateUser;
         private System.Windows.Forms.TextBox TextBoxHandler;
         private System.Windows.Forms.TextBox TextBoxCardName;
         private System.Windows.Forms.Label label1;
         private System.Windows.Forms.Button button1;
 
     }
 }

