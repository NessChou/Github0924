using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
namespace ACME
{
    public partial class CART2 : ACME.fmBase1
    {
        public CART2()
        {
            InitializeComponent();
        }

        private void CART2_Load(object sender, EventArgs e)
        {


        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;

            cARTTableAdapter.Connection = MyConnection;
        }
        public override void SetInit()
        {
            MyBS = cARTBindingSource;
            MyTableName = "CART";
            MyIDFieldName = "CardId";

            MasterTable = wh.CART;
        }
        public override void AfterCopy()
        {

            if (kyes == null)
            {
                string NumberName = "CA";
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                this.cardIdTextBox.Text = NumberName + AutoNum;
                kyes = this.cardIdTextBox.Text;
            }
        }
           public override void SetDefaultValue()
        {

            string NumberName = "CA" ;            
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
            comboBox1.Text = "TFT";
            this.cardIdTextBox.Text = NumberName + AutoNum;
            cREATE_DATETextBox.Text = DateTime.Now.ToString("yyyyMMdd");
            cREATE_USERTextBox.Text = fmLogin.LoginID.ToString();
            this.cARTBindingSource.EndEdit();
            kyes = null;
        }

        public override void STOP()
        {
            if (mODEL_NOTextBox.Text == "")
            {
                MessageBox.Show("請輸入MODEL");
                this.SSTOPID = "1";
                
            }
            if (mODEL_VerTextBox.Text == "")
            {
                MessageBox.Show("請輸入VER");
                this.SSTOPID = "1";

            }
            Boolean b1 = Information.IsNumeric(cT_LTextBox.Text);
            Boolean b2 = Information.IsNumeric(cT_WTextBox.Text);
            Boolean b3 = Information.IsNumeric(cT_HTextBox.Text);
            Boolean b4 = Information.IsNumeric(cT_GWTextBox.Text);
            Boolean b5 = Information.IsNumeric(cT_QTYTextBox.Text);
            Boolean b6 = Information.IsNumeric(cT_NWTextBox.Text);


            //CARTON
            if (b1.ToString() == "False" && cT_LTextBox.Text.Trim() != "")
            {
                MessageBox.Show("CARTON PACKAGE 'L' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (b2.ToString() == "False" && cT_WTextBox.Text.Trim() != "")
            {
                MessageBox.Show("CARTON PACKAGE 'W' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (b3.ToString() == "False" && cT_HTextBox.Text.Trim() != "")
            {
                MessageBox.Show("CARTON PACKAGE 'H' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (b4.ToString() == "False" && cT_GWTextBox.Text.Trim() != "")
            {
                MessageBox.Show("CARTON PACKAGE 'GW' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (b5.ToString() == "False" && cT_QTYTextBox.Text.Trim() != "")
            {
                MessageBox.Show("CARTON PACKAGE 'QTY' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (b6.ToString() == "False" && cT_NWTextBox.Text.Trim() != "")
            {
                MessageBox.Show("CARTON PACKAGE 'NW' 請輸入數字");
                this.SSTOPID = "1";
            }


            //PALLET

            Boolean c1 = Information.IsNumeric(pAL_LTextBox.Text);
            Boolean c2 = Information.IsNumeric(pAL_WTextBox.Text);
            Boolean c3 = Information.IsNumeric(pAL_HTextBox.Text);
            Boolean c4 = Information.IsNumeric(pAL_GWTextBox.Text);
            Boolean c5 = Information.IsNumeric(pAL_QTYTextBox.Text);
            Boolean c6 = Information.IsNumeric(pAL_NWTextBox.Text);
            Boolean c7 = Information.IsNumeric(pAL_CTNSTextBox.Text);

            if (c1.ToString() == "False" && pAL_LTextBox.Text.Trim() != "")
            {
                MessageBox.Show("PALLET PACKAGE 'L' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (c2.ToString() == "False" && pAL_WTextBox.Text.Trim() != "")
            {
                MessageBox.Show("PALLET PACKAGE 'W' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (c3.ToString() == "False" && pAL_HTextBox.Text.Trim() != "")
            {
                MessageBox.Show("PALLET PACKAGE 'H' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (c4.ToString() == "False" && pAL_GWTextBox.Text.Trim() != "")
            {
                MessageBox.Show("PALLET PACKAGE 'GW' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (c5.ToString() == "False" && pAL_LTextBox.Text.Trim() != "")
            {
                MessageBox.Show("PALLET PACKAGE 'QTY' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (c6.ToString() == "False" && pAL_NWTextBox.Text.Trim() != "")
            {
                MessageBox.Show("PALLET PACKAGE 'NW' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (c7.ToString() == "False" && pAL_CTNSTextBox.Text.Trim() != "")
            {
                MessageBox.Show("PALLET PACKAGE 'CTNS' 請輸入數字");
                this.SSTOPID = "1";
            }


            //20'CTNR

            Boolean d1 = Information.IsNumeric(cT20_LTextBox.Text);
            Boolean d2 = Information.IsNumeric(cT20_WTextBox.Text);
            Boolean d3 = Information.IsNumeric(cT20_HTextBox.Text);
            Boolean d5 = Information.IsNumeric(cT20_QTYTextBox.Text);
            Boolean d6 = Information.IsNumeric(cT20_PLTSTextBox.Text);
            if (d1.ToString() == "False" && cT20_LTextBox.Text.Trim() != "")
            {
                MessageBox.Show("20'CTNR 'L' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (d2.ToString() == "False" && cT20_WTextBox.Text.Trim() != "")
            {
                MessageBox.Show("20'CTNR 'W' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (d3.ToString() == "False" && cT20_HTextBox.Text.Trim() != "")
            {
                MessageBox.Show("20'CTNR 'H' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (d5.ToString() == "False" && cT20_QTYTextBox.Text.Trim() != "")
            {
                MessageBox.Show("20'CTNR 'QTY' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (d6.ToString() == "False" && cT20_PLTSTextBox.Text.Trim() != "")
            {
                MessageBox.Show("20'CTNR 'PLTS' 請輸入數字");
                this.SSTOPID = "1";
            }

            //40'CTNR

            Boolean e1 = Information.IsNumeric(cT40_LTextBox.Text);
            Boolean e2 = Information.IsNumeric(cT40_WTextBox.Text);
            Boolean e3 = Information.IsNumeric(cT40_HTextBox.Text);
            Boolean e5 = Information.IsNumeric(cT40_QTYTextBox.Text);
            Boolean e6 = Information.IsNumeric(cT40_PLTSTextBox.Text);
            if (e1.ToString() == "False" && cT40_LTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'CTNR 'L' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (e2.ToString() == "False" && cT40_WTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'CTNR 'W' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (e3.ToString() == "False" && cT40_HTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'CTNR 'H' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (e5.ToString() == "False" && cT40_QTYTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'CTNR 'QTY' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (e6.ToString() == "False" && cT40_PLTSTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'CTNR 'PLTS' 請輸入數字");
                this.SSTOPID = "1";
            }


            //40'HCTNR

            Boolean f1 = Information.IsNumeric(cT40_HLTextBox.Text);
            Boolean f2 = Information.IsNumeric(cT40_HWTextBox.Text);
            Boolean f3 = Information.IsNumeric(cT40_HHTextBox.Text);
            Boolean f5 = Information.IsNumeric(cT40_HQTYTextBox.Text);
            Boolean f6 = Information.IsNumeric(cT40_HPLTSTextBox.Text);
            if (f1.ToString() == "False" && cT40_HLTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'HCTNR 'L' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (f2.ToString() == "False" && cT40_HWTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'HCTNR 'W' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (f3.ToString() == "False" && cT40_HHTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'HCTNR 'H' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (f5.ToString() == "False" && cT40_HQTYTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'HCTNR 'QTY' 請輸入數字");
                this.SSTOPID = "1";
            }
            if (f6.ToString() == "False" && cT40_HPLTSTextBox.Text.Trim() != "")
            {
                MessageBox.Show("40'HCTNR 'PLTS' 請輸入數字");
                this.SSTOPID = "1";
            }
        }
        public override void AfterEdit()
        {

            string f1=fmLogin.LoginID.ToString().ToUpper();
            if (f1.ToString() != "LLEYTONCHEN")
            {
                uPDATE_USERTextBox.Text = fmLogin.LoginID.ToString();
                uPDATE_DATETextBox.Text = DateTime.Now.ToString("yyyyMMdd");
            }
        }
        public override void FillData()
        {
            try
            {

                cARTTableAdapter.Fill(wh.CART, MyID);
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public override bool UpdateData()
        {
            bool UpdateData;

            try
            {
                if (Convert.ToDouble(this.MyTableStatus) == 3)
                {

                    MyBS.RemoveAt(MyBS.Position);
          
                }
               

                cARTTableAdapter.Connection.Open();

                Validate();

                cARTBindingSource.EndEdit();
                

                cARTTableAdapter.Update(wh.CART);

                this.MyID = this.cardIdTextBox.Text;

                UpdateData = true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.cARTTableAdapter.Connection.Close();

            }
            return UpdateData;
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("WHPACK");

            comboBox1.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            dOCTYPETextBox.Text = comboBox1.Text;
        }


     
    }
}

