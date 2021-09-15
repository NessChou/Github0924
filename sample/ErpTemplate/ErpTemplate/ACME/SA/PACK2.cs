using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;

namespace ACME
{
    public partial class PACK2 : Form
    {

        System.Data.DataTable dtCostF = null;

        string FA = "acmesql02";

        public PACK2()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
         
        }





        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("�п���ɮ�");
                }
                else
                {
                    dtCostF = MakeTableF();
                    GD55(opdf.FileName);
                    dataGridView4.DataSource = dtCostF;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private System.Data.DataTable MakeTableF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�������", typeof(string));
            dt.Columns.Add("�^�����", typeof(string));
            dt.Columns.Add("�Τ@�s��", typeof(string));
            dt.Columns.Add("INVOICE���q����", typeof(string));
            dt.Columns.Add("INVOICE�a�}", typeof(string));
            dt.Columns.Add("INVOICE�q��", typeof(string));
            dt.Columns.Add("INVOICE�ǯu", typeof(string));
            dt.Columns.Add("INVOICE����H", typeof(string));
            dt.Columns.Add("SHIP���q����", typeof(string));
            dt.Columns.Add("SHIP�a�}", typeof(string));
            dt.Columns.Add("SHIP�q��", typeof(string));
            dt.Columns.Add("SHIP�ǯu", typeof(string));
            dt.Columns.Add("SHIP����H", typeof(string));
            dt.Columns.Add("�p���H¾��", typeof(string));
            dt.Columns.Add("�m�W", typeof(string));
            dt.Columns.Add("�q��", typeof(string));
            dt.Columns.Add("�ǯu", typeof(string));
            dt.Columns.Add("�l��", typeof(string));
            return dt;
        }

        private void GD55(string ExcelFile)
        {

            int h = 0;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);

            Microsoft.Office.Interop.Excel.Range range = null;


            DataRow dr = null;
            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            string DUP = "";
            int D = 0;
            int L = 0;


            string �Ȥ�s��;
                string �������;
                string �^�����;
                string �Τ@�s��;
                string INVOICE���q����;
                string INVOICE�a�};
                string INVOICE�q��;
                string INVOICE�ǯu;
                string INVOICE����H;
                string SHIP���q����;
                string SHIP�a�};
                string SHIP�q��;
                string SHIP�ǯu;
                string SHIP����H;
                string �p���H¾��;
                string �m�W;
                string �q��;
                string �ǯu;
                string �l��;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 2]);
                range.Select();
                ������� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 2]);
                range.Select();
                �^����� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[9, 2]);
                range.Select();
                �Τ@�s�� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[18, 2]);
                range.Select();
                INVOICE���q���� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[19, 2]);
                range.Select();
                INVOICE�a�} = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[20, 2]);
                range.Select();
                INVOICE�q�� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[21, 2]);
                range.Select();
                INVOICE�ǯu = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[22, 2]);
                range.Select();
                INVOICE����H = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[24, 2]);
                range.Select();
                SHIP���q���� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[25, 2]);
                range.Select();
                SHIP�a�} = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[26, 2]);
                range.Select();
                SHIP�q�� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[27, 2]);
                range.Select();
                SHIP�ǯu = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[28, 2]);
                range.Select();
                SHIP����H = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[30, 2]);
                range.Select();
                �p���H¾�� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[31, 2]);
                range.Select();
                �m�W = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[32, 2]);
                range.Select();
                �q�� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[33, 2]);
                range.Select();
                �ǯu = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[34, 2]);
                range.Select();
                �l�� = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[40, 1]);
                range.Select();
                �Ȥ�s�� = range.Text.ToString().Trim();
                try
                {
                

                    dr = dtCostF.NewRow();
                    dr["�Ȥ�s��"] = �Ȥ�s��;
                    dr["�������"] = �������;
                    dr["�^�����"] = �^�����;
                    dr["�Τ@�s��"] = �Τ@�s��;
                    dr["INVOICE���q����"] = INVOICE���q����;
                    dr["INVOICE�a�}"] = INVOICE�a�};
                    dr["INVOICE�q��"] = INVOICE�q��;
                    dr["INVOICE�ǯu"] = INVOICE�ǯu;
                    dr["INVOICE����H"] = INVOICE����H;
                    dr["SHIP���q����"] = SHIP���q����;
                    dr["SHIP�a�}"] = SHIP�a�};
                    dr["SHIP�q��"] = SHIP�q��;
                    dr["SHIP�ǯu"] = SHIP�ǯu;
                    dr["SHIP����H"] = SHIP����H;

                    dr["�p���H¾��"] = �p���H¾��;
                    dr["�m�W"] = �m�W;
                    dr["�q��"] = �q��;
                    dr["�ǯu"] = �ǯu;
                    dr["�l��"] = �l��;
                    dtCostF.Rows.Add(dr);

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            



            //Quit
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


            range = null;
            excelApp = null;
            excelBook = null;
            excelSheet = null;

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();


        }

        private void button12_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G2 = dtCostF;

            if (G2.Rows.Count > 0)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int i = 0; //  to be used as an index

                oCompany.CompanyDB = FA;
                oCompany.UserName = "A01";
                oCompany.Password = "89206602";
                int result = oCompany.Connect();
                if (result == 0)
                {
                    SAPbobsCOM.BusinessPartners  oCARD = null;
                    oCARD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                    string CARDCODE = dataGridView4.Rows[0].Cells["�Ȥ�s��"].Value.ToString();
                    oCARD.CardCode = CARDCODE;
                    oCARD.CardName =  dataGridView4.Rows[0].Cells["�������"].Value.ToString();
                    oCARD.CardType = SAPbobsCOM.BoCardTypes.cCustomer;

                    oCARD.CardForeignName = dataGridView4.Rows[0].Cells["�^�����"].Value.ToString();
                    oCARD.GroupCode = 103;
                    oCARD.Currency = "##";
                    oCARD.FederalTaxID =dataGridView4.Rows[0].Cells["�Τ@�s��"].Value.ToString();
                    oCARD.GetByKey(CARDCODE);
                    oCARD.VatGroup = "AR5%";
                    oCARD.Addresses.AddressName = dataGridView4.Rows[0].Cells["INVOICE���q����"].Value.ToString();
                    oCARD.Addresses.BuildingFloorRoom = dataGridView4.Rows[0].Cells["INVOICE���q����"].Value.ToString();
                    oCARD.Addresses.Street = dataGridView4.Rows[0].Cells["INVOICE�a�}"].Value.ToString();
                    oCARD.Addresses.Block = dataGridView4.Rows[0].Cells["INVOICE�q��"].Value.ToString();
                    oCARD.Addresses.City = dataGridView4.Rows[0].Cells["INVOICE�ǯu"].Value.ToString();
                    oCARD.Addresses.ZipCode = dataGridView4.Rows[0].Cells["INVOICE����H"].Value.ToString();
                    oCARD.Addresses.Country = "TW";
                    oCARD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                    oCARD.Addresses.Add();

                    oCARD.Addresses.AddressName = dataGridView4.Rows[0].Cells["SHIP���q����"].Value.ToString();
                    oCARD.Addresses.BuildingFloorRoom = dataGridView4.Rows[0].Cells["SHIP���q����"].Value.ToString();
                    oCARD.Addresses.Street = dataGridView4.Rows[0].Cells["SHIP�a�}"].Value.ToString();
                    oCARD.Addresses.Block = dataGridView4.Rows[0].Cells["SHIP�q��"].Value.ToString();
                    oCARD.Addresses.City = dataGridView4.Rows[0].Cells["SHIP�ǯu"].Value.ToString();
                    oCARD.Addresses.ZipCode = dataGridView4.Rows[0].Cells["SHIP����H"].Value.ToString();
                    oCARD.Addresses.Country = "TW";
                    oCARD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                    
                    oCARD.Addresses.UserFields.Fields.Item("U_Territory").Value = "�x�W";
                    oCARD.Addresses.UserFields.Fields.Item("U_Application").Value = "�x�W";
                    oCARD.Addresses.Add();

                    oCARD.ContactEmployees.Name =  dataGridView4.Rows[0].Cells["�p���H¾��"].Value.ToString()+dataGridView4.Rows[0].Cells["�m�W"].Value.ToString();
                    oCARD.ContactEmployees.Position = dataGridView4.Rows[0].Cells["�q��"].Value.ToString();
                    oCARD.ContactEmployees.Address = dataGridView4.Rows[0].Cells["�ǯu"].Value.ToString();
                    oCARD.ContactEmployees.E_Mail = dataGridView4.Rows[0].Cells["�l��"].Value.ToString();
                    oCARD.ContactEmployees.Add();

                    int res = oCARD.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("�W�ǿ��~ " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        MessageBox.Show("�פJ���\");

                    }




                }
            }
        }
    }
}