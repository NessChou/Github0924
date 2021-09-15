namespace ACME
{
    partial class WHMULTI
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.Label shipping_OBULabel;
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.dualList2 = new ReflectionIT.Common.Windows.Forms.DualList(this.components);
            this.dualList3 = new ReflectionIT.Common.Windows.Forms.DualList(this.components);
            this.button5 = new System.Windows.Forms.Button();
            this.dualList4 = new ReflectionIT.Common.Windows.Forms.DualList(this.components);
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.button6 = new System.Windows.Forms.Button();
            this.dualList1 = new ReflectionIT.Common.Windows.Forms.DualList(this.components);
            this.button7 = new System.Windows.Forms.Button();
            shipping_OBULabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // shipping_OBULabel
            // 
            shipping_OBULabel.AutoSize = true;
            shipping_OBULabel.Location = new System.Drawing.Point(237, 25);
            shipping_OBULabel.Name = "shipping_OBULabel";
            shipping_OBULabel.Size = new System.Drawing.Size(29, 12);
            shipping_OBULabel.TabIndex = 149;
            shipping_OBULabel.Text = "倉庫";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(581, 64);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 60);
            this.button1.TabIndex = 112;
            this.button1.Text = "匯出Excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 12;
            this.listBox1.Location = new System.Drawing.Point(60, 64);
            this.listBox1.Name = "listBox1";
            this.listBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBox1.Size = new System.Drawing.Size(170, 280);
            this.listBox1.TabIndex = 114;
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 12;
            this.listBox2.Location = new System.Drawing.Point(366, 64);
            this.listBox2.Name = "listBox2";
            this.listBox2.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBox2.Size = new System.Drawing.Size(180, 280);
            this.listBox2.TabIndex = 115;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(262, 168);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 116;
            this.button2.Text = ">";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(262, 215);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 117;
            this.button3.Text = "<";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(71, 21);
            this.textBox1.MaxLength = 8;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(60, 22);
            this.textBox1.TabIndex = 118;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(154, 21);
            this.textBox2.MaxLength = 8;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(60, 22);
            this.textBox2.TabIndex = 119;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(137, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(11, 12);
            this.label2.TabIndex = 120;
            this.label2.Text = "~";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 121;
            this.label3.Text = "工單日期";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(446, 16);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 24);
            this.button4.TabIndex = 122;
            this.button4.Text = "查詢";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // dualList2
            // 
            this.dualList2.Button = this.button2;
            this.dualList2.ListBoxFrom = this.listBox1;
            this.dualList2.ListBoxTo = this.listBox2;
            // 
            // dualList3
            // 
            this.dualList3.Button = this.button3;
            this.dualList3.ListBoxFrom = this.listBox2;
            this.dualList3.ListBoxTo = this.listBox1;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(262, 260);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 123;
            this.button5.Text = "<<";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // dualList4
            // 
            this.dualList4.Action = ReflectionIT.Common.Windows.Forms.DualListAction.MoveAll;
            this.dualList4.Button = this.button5;
            this.dualList4.ListBoxFrom = this.listBox2;
            this.dualList4.ListBoxTo = this.listBox1;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(272, 20);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(168, 20);
            this.comboBox2.TabIndex = 150;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(262, 131);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(75, 23);
            this.button6.TabIndex = 151;
            this.button6.Text = ">>";
            this.button6.UseVisualStyleBackColor = true;
            // 
            // dualList1
            // 
            this.dualList1.Action = ReflectionIT.Common.Windows.Forms.DualListAction.CopyAll;
            this.dualList1.Button = this.button6;
            this.dualList1.ListBoxFrom = this.listBox1;
            this.dualList1.ListBoxTo = this.listBox2;
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(581, 168);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(75, 60);
            this.button7.TabIndex = 152;
            this.button7.Text = "EMAIL";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // WHMULTI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(716, 353);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(shipping_OBULabel);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button1);
            this.Name = "WHMULTI";
            this.Text = "合併收貨單";
            this.Load += new System.EventHandler(this.SHIPMULTI_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button4;
        private ReflectionIT.Common.Windows.Forms.DualList dualList2;
        private ReflectionIT.Common.Windows.Forms.DualList dualList3;
        private System.Windows.Forms.Button button5;
        private ReflectionIT.Common.Windows.Forms.DualList dualList4;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Button button6;
        private ReflectionIT.Common.Windows.Forms.DualList dualList1;
        private System.Windows.Forms.Button button7;
    }
}