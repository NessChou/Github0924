using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class CART2_LED : ACME.fmBase1
    {
        public CART2_LED()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            cART_LEDTableAdapter.Connection = MyConnection;

        }

        public override void SetInit()
        {
            MyBS = cART_LEDBindingSource;
            MyTableName = "CART_LED";
            MyIDFieldName = "ID";
        }
        public override void SetDefaultValue()
        {

            string NumberName = "CAL";
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

            this.iDTextBox.Text = NumberName + AutoNum;
            cREATE_DATETextBox.Text = DateTime.Now.ToString("yyyyMMdd");
            cREATE_USERTextBox.Text = fmLogin.LoginID.ToString();
            this.cART_LEDBindingSource.EndEdit();
        }
        public override void STOP()
        {
            if (iTEMCODETextBox.Text == "")
            {
                MessageBox.Show("請輸入項目號碼");
                this.SSTOPID = "1";

            }
        }
        public override void AfterEdit()
        {
            uPDATE_USERTextBox.Text = fmLogin.LoginID.ToString();
            uPDATE_DATETextBox.Text = DateTime.Now.ToString("yyyyMMdd");
        }
        public override void FillData()
        {
            try
            {

                cART_LEDTableAdapter.Fill(wh.CART_LED, MyID);

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


                cART_LEDTableAdapter.Connection.Open();

                Validate();

                cART_LEDBindingSource.EndEdit();


                cART_LEDTableAdapter.Update(wh.CART_LED);

                this.MyID = this.iTEMCODETextBox.Text;

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
                this.cART_LEDTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
    }
}

