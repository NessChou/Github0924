namespace ACME
{
    partial class ACDsign
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
            System.Windows.Forms.Label temp4Label;
            System.Windows.Forms.Label temp5Label;
            this.temp4TextBox = new System.Windows.Forms.TextBox();
            this.tempBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ship = new ACME.ACMEDataSet.ship();
            this.temp5TextBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.tempTableAdapter = new ACME.ACMEDataSet.shipTableAdapters.tempTableAdapter();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            temp4Label = new System.Windows.Forms.Label();
            temp5Label = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.tempBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).BeginInit();
            this.SuspendLayout();
            // 
            // temp4Label
            // 
            temp4Label.AutoSize = true;
            temp4Label.Location = new System.Drawing.Point(12, 19);
            temp4Label.Name = "temp4Label";
            temp4Label.Size = new System.Drawing.Size(48, 12);
            temp4Label.TabIndex = 2;
            temp4Label.Text = "AR發票:";
            // 
            // temp5Label
            // 
            temp5Label.AutoSize = true;
            temp5Label.Location = new System.Drawing.Point(27, 227);
            temp5Label.Name = "temp5Label";
            temp5Label.Size = new System.Drawing.Size(46, 12);
            temp5Label.TabIndex = 4;
            temp5Label.Text = "AP發票:";
            temp5Label.Visible = false;
            // 
            // temp4TextBox
            // 
            this.temp4TextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.tempBindingSource, "temp4", true));
            this.temp4TextBox.Location = new System.Drawing.Point(66, 16);
            this.temp4TextBox.MaxLength = 8;
            this.temp4TextBox.Name = "temp4TextBox";
            this.temp4TextBox.Size = new System.Drawing.Size(100, 22);
            this.temp4TextBox.TabIndex = 3;
            // 
            // tempBindingSource
            // 
            this.tempBindingSource.DataMember = "temp";
            this.tempBindingSource.DataSource = this.ship;
            // 
            // ship
            // 
            this.ship.DataSetName = "ship";
            this.ship.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // temp5TextBox
            // 
            this.temp5TextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.tempBindingSource, "temp5", true));
            this.temp5TextBox.Location = new System.Drawing.Point(81, 224);
            this.temp5TextBox.MaxLength = 8;
            this.temp5TextBox.Name = "temp5TextBox";
            this.temp5TextBox.Size = new System.Drawing.Size(100, 22);
            this.temp5TextBox.TabIndex = 5;
            this.temp5TextBox.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(178, 16);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "存檔";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tempTableAdapter
            // 
            this.tempTableAdapter.ClearBeforeFill = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(205, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "限制SAP單據過帳日期不得大於設定值";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(415, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "例如:AR發票設定為20080701時,如果AR發票過帳日期為20080630則不允許存檔";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "如果是20080701則允許存檔";
            // 
            // ACDsign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(522, 325);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(temp5Label);
            this.Controls.Add(this.temp5TextBox);
            this.Controls.Add(temp4Label);
            this.Controls.Add(this.temp4TextBox);
            this.Controls.Add(this.button1);
            this.Name = "ACDsign";
            this.Text = "ACDsign";
            this.Load += new System.EventHandler(this.ACDsign_Load);
            ((System.ComponentModel.ISupportInitialize)(this.tempBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship ship;
            
        private System.Windows.Forms.BindingSource tempBindingSource;
        private ACME.ACMEDataSet.shipTableAdapters.tempTableAdapter tempTableAdapter;
        private System.Windows.Forms.TextBox temp4TextBox;
        private System.Windows.Forms.TextBox temp5TextBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}