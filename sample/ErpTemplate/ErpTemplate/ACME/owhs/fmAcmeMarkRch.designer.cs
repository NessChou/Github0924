namespace ACME
{
    partial class fmAcmeMarkRch
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.button63 = new System.Windows.Forms.Button();
            this.button28 = new System.Windows.Forms.Button();
            this.txtRCH = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.button30 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dgData = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button63);
            this.panel1.Controls.Add(this.button28);
            this.panel1.Controls.Add(this.txtRCH);
            this.panel1.Controls.Add(this.label26);
            this.panel1.Controls.Add(this.button30);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(944, 81);
            this.panel1.TabIndex = 0;
            // 
            // button63
            // 
            this.button63.Location = new System.Drawing.Point(329, 30);
            this.button63.Name = "button63";
            this.button63.Size = new System.Drawing.Size(75, 25);
            this.button63.TabIndex = 51;
            this.button63.Text = "2.Excel";
            this.button63.UseVisualStyleBackColor = true;
            this.button63.Click += new System.EventHandler(this.button63_Click);
            // 
            // button28
            // 
            this.button28.Location = new System.Drawing.Point(410, 30);
            this.button28.Name = "button28";
            this.button28.Size = new System.Drawing.Size(75, 25);
            this.button28.TabIndex = 50;
            this.button28.Text = "3.Mail";
            this.button28.UseVisualStyleBackColor = true;
            this.button28.Click += new System.EventHandler(this.button28_Click);
            // 
            // txtRCH
            // 
            this.txtRCH.Location = new System.Drawing.Point(94, 30);
            this.txtRCH.Name = "txtRCH";
            this.txtRCH.Size = new System.Drawing.Size(123, 25);
            this.txtRCH.TabIndex = 47;
            this.txtRCH.Text = "WH20200113002X";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(21, 35);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(67, 15);
            this.label26.TabIndex = 48;
            this.label26.Text = "工單號碼";
            // 
            // button30
            // 
            this.button30.Location = new System.Drawing.Point(248, 30);
            this.button30.Name = "button30";
            this.button30.Size = new System.Drawing.Size(75, 25);
            this.button30.TabIndex = 49;
            this.button30.Text = "1.查詢";
            this.button30.UseVisualStyleBackColor = true;
            this.button30.Click += new System.EventHandler(this.button30_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dgData);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 81);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(944, 285);
            this.panel2.TabIndex = 1;
            // 
            // dgData
            // 
            this.dgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgData.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgData.Location = new System.Drawing.Point(0, 0);
            this.dgData.Name = "dgData";
            this.dgData.RowTemplate.Height = 27;
            this.dgData.Size = new System.Drawing.Size(944, 285);
            this.dgData.TabIndex = 24;
            // 
            // fmAcmeMarkRch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(944, 366);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "fmAcmeMarkRch";
            this.Text = "fmAcmeMarkRch";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button63;
        private System.Windows.Forms.Button button28;
        private System.Windows.Forms.TextBox txtRCH;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.Button button30;
        private System.Windows.Forms.DataGridView dgData;
    }
}