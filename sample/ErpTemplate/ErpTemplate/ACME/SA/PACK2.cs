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
                    MessageBox.Show("請選擇檔案");
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
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("中文全稱", typeof(string));
            dt.Columns.Add("英文全稱", typeof(string));
            dt.Columns.Add("統一編號", typeof(string));
            dt.Columns.Add("INVOICE公司全稱", typeof(string));
            dt.Columns.Add("INVOICE地址", typeof(string));
            dt.Columns.Add("INVOICE電話", typeof(string));
            dt.Columns.Add("INVOICE傳真", typeof(string));
            dt.Columns.Add("INVOICE收件人", typeof(string));
            dt.Columns.Add("SHIP公司全稱", typeof(string));
            dt.Columns.Add("SHIP地址", typeof(string));
            dt.Columns.Add("SHIP電話", typeof(string));
            dt.Columns.Add("SHIP傳真", typeof(string));
            dt.Columns.Add("SHIP收件人", typeof(string));
            dt.Columns.Add("聯絡人職務", typeof(string));
            dt.Columns.Add("姓名", typeof(string));
            dt.Columns.Add("電話", typeof(string));
            dt.Columns.Add("傳真", typeof(string));
            dt.Columns.Add("郵件", typeof(string));
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


            string 客戶編號;
                string 中文全稱;
                string 英文全稱;
                string 統一編號;
                string INVOICE公司全稱;
                string INVOICE地址;
                string INVOICE電話;
                string INVOICE傳真;
                string INVOICE收件人;
                string SHIP公司全稱;
                string SHIP地址;
                string SHIP電話;
                string SHIP傳真;
                string SHIP收件人;
                string 聯絡人職務;
                string 姓名;
                string 電話;
                string 傳真;
                string 郵件;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 2]);
                range.Select();
                中文全稱 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 2]);
                range.Select();
                英文全稱 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[9, 2]);
                range.Select();
                統一編號 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[18, 2]);
                range.Select();
                INVOICE公司全稱 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[19, 2]);
                range.Select();
                INVOICE地址 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[20, 2]);
                range.Select();
                INVOICE電話 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[21, 2]);
                range.Select();
                INVOICE傳真 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[22, 2]);
                range.Select();
                INVOICE收件人 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[24, 2]);
                range.Select();
                SHIP公司全稱 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[25, 2]);
                range.Select();
                SHIP地址 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[26, 2]);
                range.Select();
                SHIP電話 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[27, 2]);
                range.Select();
                SHIP傳真 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[28, 2]);
                range.Select();
                SHIP收件人 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[30, 2]);
                range.Select();
                聯絡人職務 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[31, 2]);
                range.Select();
                姓名 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[32, 2]);
                range.Select();
                電話 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[33, 2]);
                range.Select();
                傳真 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[34, 2]);
                range.Select();
                郵件 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[40, 1]);
                range.Select();
                客戶編號 = range.Text.ToString().Trim();
                try
                {
                

                    dr = dtCostF.NewRow();
                    dr["客戶編號"] = 客戶編號;
                    dr["中文全稱"] = 中文全稱;
                    dr["英文全稱"] = 英文全稱;
                    dr["統一編號"] = 統一編號;
                    dr["INVOICE公司全稱"] = INVOICE公司全稱;
                    dr["INVOICE地址"] = INVOICE地址;
                    dr["INVOICE電話"] = INVOICE電話;
                    dr["INVOICE傳真"] = INVOICE傳真;
                    dr["INVOICE收件人"] = INVOICE收件人;
                    dr["SHIP公司全稱"] = SHIP公司全稱;
                    dr["SHIP地址"] = SHIP地址;
                    dr["SHIP電話"] = SHIP電話;
                    dr["SHIP傳真"] = SHIP傳真;
                    dr["SHIP收件人"] = SHIP收件人;

                    dr["聯絡人職務"] = 聯絡人職務;
                    dr["姓名"] = 姓名;
                    dr["電話"] = 電話;
                    dr["傳真"] = 傳真;
                    dr["郵件"] = 郵件;
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
                    string CARDCODE = dataGridView4.Rows[0].Cells["客戶編號"].Value.ToString();
                    oCARD.CardCode = CARDCODE;
                    oCARD.CardName =  dataGridView4.Rows[0].Cells["中文全稱"].Value.ToString();
                    oCARD.CardType = SAPbobsCOM.BoCardTypes.cCustomer;

                    oCARD.CardForeignName = dataGridView4.Rows[0].Cells["英文全稱"].Value.ToString();
                    oCARD.GroupCode = 103;
                    oCARD.Currency = "##";
                    oCARD.FederalTaxID =dataGridView4.Rows[0].Cells["統一編號"].Value.ToString();
                    oCARD.GetByKey(CARDCODE);
                    oCARD.VatGroup = "AR5%";
                    oCARD.Addresses.AddressName = dataGridView4.Rows[0].Cells["INVOICE公司全稱"].Value.ToString();
                    oCARD.Addresses.BuildingFloorRoom = dataGridView4.Rows[0].Cells["INVOICE公司全稱"].Value.ToString();
                    oCARD.Addresses.Street = dataGridView4.Rows[0].Cells["INVOICE地址"].Value.ToString();
                    oCARD.Addresses.Block = dataGridView4.Rows[0].Cells["INVOICE電話"].Value.ToString();
                    oCARD.Addresses.City = dataGridView4.Rows[0].Cells["INVOICE傳真"].Value.ToString();
                    oCARD.Addresses.ZipCode = dataGridView4.Rows[0].Cells["INVOICE收件人"].Value.ToString();
                    oCARD.Addresses.Country = "TW";
                    oCARD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                    oCARD.Addresses.Add();

                    oCARD.Addresses.AddressName = dataGridView4.Rows[0].Cells["SHIP公司全稱"].Value.ToString();
                    oCARD.Addresses.BuildingFloorRoom = dataGridView4.Rows[0].Cells["SHIP公司全稱"].Value.ToString();
                    oCARD.Addresses.Street = dataGridView4.Rows[0].Cells["SHIP地址"].Value.ToString();
                    oCARD.Addresses.Block = dataGridView4.Rows[0].Cells["SHIP電話"].Value.ToString();
                    oCARD.Addresses.City = dataGridView4.Rows[0].Cells["SHIP傳真"].Value.ToString();
                    oCARD.Addresses.ZipCode = dataGridView4.Rows[0].Cells["SHIP收件人"].Value.ToString();
                    oCARD.Addresses.Country = "TW";
                    oCARD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                    
                    oCARD.Addresses.UserFields.Fields.Item("U_Territory").Value = "台灣";
                    oCARD.Addresses.UserFields.Fields.Item("U_Application").Value = "台灣";
                    oCARD.Addresses.Add();

                    oCARD.ContactEmployees.Name =  dataGridView4.Rows[0].Cells["聯絡人職務"].Value.ToString()+dataGridView4.Rows[0].Cells["姓名"].Value.ToString();
                    oCARD.ContactEmployees.Position = dataGridView4.Rows[0].Cells["電話"].Value.ToString();
                    oCARD.ContactEmployees.Address = dataGridView4.Rows[0].Cells["傳真"].Value.ToString();
                    oCARD.ContactEmployees.E_Mail = dataGridView4.Rows[0].Cells["郵件"].Value.ToString();
                    oCARD.ContactEmployees.Add();

                    int res = oCARD.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        MessageBox.Show("匯入成功");

                    }




                }
            }
        }
    }
}