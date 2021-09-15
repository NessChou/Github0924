 namespace ACME
 {
     partial class AP_WHS
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
             this.TextBoxID = new System.Windows.Forms.TextBox();
             this.lblID = new System.Windows.Forms.Label();
             this.TextBoxWHSCODE = new System.Windows.Forms.TextBox();
             this.btnCancel = new System.Windows.Forms.Button();
             this.btnSave = new System.Windows.Forms.Button();
             this.btnDelete = new System.Windows.Forms.Button();
             this.label2 = new System.Windows.Forms.Label();
             this.TextBoxDESCRIPTION = new System.Windows.Forms.TextBox();
             this.label1 = new System.Windows.Forms.Label();
             this.TextBoxWHS = new System.Windows.Forms.TextBox();
             this.label3 = new System.Windows.Forms.Label();
             this.TextBoxLOCATION = new System.Windows.Forms.TextBox();
             this.LOCATION = new System.Windows.Forms.Label();
             this.SuspendLayout();
             // 
             // TextBoxID
             // 
             this.TextBoxID.Location = new System.Drawing.Point(110, 12);
             this.TextBoxID.Name = "TextBoxID";
             this.TextBoxID.Size = new System.Drawing.Size(100, 22);
             this.TextBoxID.TabIndex = 12;
             // 
             // lblID
             // 
             this.lblID.AutoSize = true;
             this.lblID.Location = new System.Drawing.Point(87, 15);
             this.lblID.Name = "lblID";
             this.lblID.Size = new System.Drawing.Size(17, 12);
             this.lblID.TabIndex = 11;
             this.lblID.Text = "ID";
             // 
             // TextBoxWHSCODE
             // 
             this.TextBoxWHSCODE.Location = new System.Drawing.Point(110, 72);
             this.TextBoxWHSCODE.Name = "TextBoxWHSCODE";
             this.TextBoxWHSCODE.Size = new System.Drawing.Size(308, 22);
             this.TextBoxWHSCODE.TabIndex = 31;
             // 
             // btnCancel
             // 
             this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
             this.btnCancel.Image = global::ACME.Properties.Resources.bnExit_Image;
             this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
             this.btnCancel.Location = new System.Drawing.Point(501, 340);
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
             this.btnSave.Location = new System.Drawing.Point(420, 340);
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
             this.btnDelete.Location = new System.Drawing.Point(12, 340);
             this.btnDelete.Name = "btnDelete";
             this.btnDelete.Size = new System.Drawing.Size(75, 23);
             this.btnDelete.TabIndex = 27;
             this.btnDelete.Text = "Delete";
             this.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
             this.btnDelete.UseVisualStyleBackColor = true;
             this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
             // 
             // label2
             // 
             this.label2.AutoSize = true;
             this.label2.Location = new System.Drawing.Point(63, 44);
             this.label2.Name = "label2";
             this.label2.Size = new System.Drawing.Size(41, 12);
             this.label2.TabIndex = 47;
             this.label2.Text = "倉庫別";
             // 
             // TextBoxDESCRIPTION
             // 
             this.TextBoxDESCRIPTION.Location = new System.Drawing.Point(110, 147);
             this.TextBoxDESCRIPTION.Multiline = true;
             this.TextBoxDESCRIPTION.Name = "TextBoxDESCRIPTION";
             this.TextBoxDESCRIPTION.Size = new System.Drawing.Size(386, 139);
             this.TextBoxDESCRIPTION.TabIndex = 54;
             // 
             // label1
             // 
             this.label1.AutoSize = true;
             this.label1.Location = new System.Drawing.Point(66, 150);
             this.label1.Name = "label1";
             this.label1.Size = new System.Drawing.Size(38, 12);
             this.label1.TabIndex = 55;
             this.label1.Text = "Ship to";
             // 
             // TextBoxWHS
             // 
             this.TextBoxWHS.Location = new System.Drawing.Point(110, 41);
             this.TextBoxWHS.Name = "TextBoxWHS";
             this.TextBoxWHS.Size = new System.Drawing.Size(100, 22);
             this.TextBoxWHS.TabIndex = 56;
             // 
             // label3
             // 
             this.label3.AutoSize = true;
             this.label3.Location = new System.Drawing.Point(51, 75);
             this.label3.Name = "label3";
             this.label3.Size = new System.Drawing.Size(53, 12);
             this.label3.TabIndex = 57;
             this.label3.Text = "倉庫名稱";
             // 
             // TextBoxLOCATION
             // 
             this.TextBoxLOCATION.Location = new System.Drawing.Point(110, 109);
             this.TextBoxLOCATION.Name = "TextBoxLOCATION";
             this.TextBoxLOCATION.Size = new System.Drawing.Size(308, 22);
             this.TextBoxLOCATION.TabIndex = 58;
             // 
             // LOCATION
             // 
             this.LOCATION.AutoSize = true;
             this.LOCATION.Location = new System.Drawing.Point(41, 112);
             this.LOCATION.Name = "LOCATION";
             this.LOCATION.Size = new System.Drawing.Size(63, 12);
             this.LOCATION.TabIndex = 59;
             this.LOCATION.Text = "LOCATION";
             // 
             // AP_WHS
             // 
             this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
             this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
             this.ClientSize = new System.Drawing.Size(588, 389);
             this.Controls.Add(this.LOCATION);
             this.Controls.Add(this.TextBoxLOCATION);
             this.Controls.Add(this.label3);
             this.Controls.Add(this.TextBoxWHS);
             this.Controls.Add(this.label1);
             this.Controls.Add(this.TextBoxDESCRIPTION);
             this.Controls.Add(this.label2);
             this.Controls.Add(this.TextBoxWHSCODE);
             this.Controls.Add(this.lblID);
             this.Controls.Add(this.TextBoxID);
             this.Controls.Add(this.btnCancel);
             this.Controls.Add(this.btnSave);
             this.Controls.Add(this.btnDelete);
             this.Name = "AP_WHS";
             this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
             this.Text = "修改";
             this.ResumeLayout(false);
             this.PerformLayout();

         }
 
         #endregion
 
         private System.Windows.Forms.Label lblID;
         private System.Windows.Forms.TextBox TextBoxID;
         private System.Windows.Forms.Button btnDelete;
         private System.Windows.Forms.Button btnSave;
         private System.Windows.Forms.Button btnCancel;
         private System.Windows.Forms.TextBox TextBoxWHSCODE;
         private System.Windows.Forms.Label label2;
         private System.Windows.Forms.TextBox TextBoxDESCRIPTION;
         private System.Windows.Forms.Label label1;
         private System.Windows.Forms.TextBox TextBoxWHS;
         private System.Windows.Forms.Label label3;
         private System.Windows.Forms.TextBox TextBoxLOCATION;
         private System.Windows.Forms.Label LOCATION;
 
     }
 }

