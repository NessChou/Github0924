using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.SqlClient;

//namespace 修正為 Acme
//20080711 增加 零稅率銷售額資料檔 
//20090811 F1920 年月,國稅局修正為 3 碼,其餘不變
//20090811 申報年月與國稅局相同
//經海關 ->U_IN_BSREN
//非經海關 ->
//外銷方式=1 & 通關方式=1 不轉證明文件號碼
// TXT  28 開頭無法排序...
//格式+發票號碼
//零稅率排序 通關方式+外銷方式+日期
//寫到暫存檔 再匯出....


namespace ACME
{
    public partial class fmAcmeTax : Form
    {
        System.Data.DataTable dtGetAcmeStageG = null;
        string INVOICE = "";
        string strCn02 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn17 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public fmAcmeTax()
        {
            InitializeComponent();
        }

        private void button63_Click(object sender, EventArgs e)
        {

//            UPDATE COMINVOICE SET includetax = 1, AMOUNT = TOTAL, REMARK = TaxAmt, TaxAmt = 0
//WHERE TaxType = '0' and isnull(TaxRegNO,'') = '
//' and includetax='0'
//and applymonth = '202105'

            string CHICONN = "";
            if (comboBox1.Text == "聿豐" || comboBox1.Text == "忠孝")
            {
                CHICONN = strCn02;
            }
    
            if (comboBox1.Text == "韋峰")
            {
                CHICONN = strCn17;
            }
            string sMonth = DateToStr(StrToDate(textBox18.Text + "01").AddMonths(1)).Substring(0, 6);
            string sMonth2 = DateToStr(StrToDate(textBox18.Text + "01").AddMonths(-1)).Substring(0, 6);
            string sMonth3 = DateToStr(StrToDate(textBox18.Text + "01")).Substring(0, 6);
            int year = Convert.ToInt32(sMonth.Substring(0, 4));
            int month = Convert.ToInt32(sMonth.Substring(4, 2));

            int year2 = Convert.ToInt32(sMonth3.Substring(0, 4));
            int month2 = Convert.ToInt32(sMonth3.Substring(4, 2));

            int days = DateTime.DaysInMonth(year, month);
            int days2 = DateTime.DaysInMonth(year2, month2);
            System.Data.DataTable dt = null;
            System.Data.DataTable dt2 = null;
            System.Data.DataTable dt3 = null;
            if (comboBox1.Text == "聿豐" || comboBox1.Text == "忠孝")
            {
                dt = GetSAPInovice2CHIARMAS(sMonth2 + "01", sMonth3 + days2.ToString("00"), CHICONN, comboBox1.Text);
                dt2 = GetSAPInovice2(sMonth2 + "01", sMonth3 + days2.ToString("00"), comboBox1.Text);
                //dt2 = GetSAPInovice2(sMonth2 + "01", sMonth3 + days2.ToString("00"), comboBox1.Text);

                //dt3 = GetSAPInovice2F(sMonth2 + "01", sMonth3 + days2.ToString("00"));

                TOTAL2GARMAS(dt, dt2);
                bindingSource1.DataSource = dtGetAcmeStageG;
                dataGridView8.DataSource = dtGetAcmeStageG;
            }
  
            else if (comboBox1.Text == "韋峰")
            {
                dt = GetSAPInovice2CHI(sMonth2 + "01", sMonth3 + days2.ToString("00"), CHICONN, comboBox1.Text);
                dt2 = GetSAPInovice2(sMonth2 + "01", sMonth3 + days2.ToString("00"), comboBox1.Text);

                dt3 = GetSAPInovice2F(sMonth2 + "01", sMonth3 + days2.ToString("00"));

                TOTAL2GG(dt, dt2, dt3);
                bindingSource1.DataSource = dtGetAcmeStageG;
                dataGridView8.DataSource = dtGetAcmeStageG;
            }
            else if (comboBox1.Text == "宇豐")
            {
                dt = GetInoviceADLAB(sMonth2 + "01", sMonth3 + days2.ToString("00"));
                dt2 = GetSAPInoviceAD(sMonth2 + "01", sMonth3 + days2.ToString("00"));
                // dt3 = GetSAPInovice2F(sMonth2 + "01", sMonth3 + days2.ToString("00"));
                TOTAL2GGAD(dt, dt2);
                bindingSource1.DataSource = dtGetAcmeStageG;
                dataGridView8.DataSource = dtGetAcmeStageG;
            }


            else
            {
                dt = GetSAPInovice(sMonth + "01", sMonth + days.ToString("00"));
                bindingSource1.DataSource = dt;
                dataGridView8.DataSource = dt;
            }

        
        }

        private void button64_Click(object sender, EventArgs e)
        {
            INVOICE = "89206602";
            //try
            //{
                MsgLine.Text = "";
                MsgDocEntry.Text = "";
                OutputSAPInvoice(bindingSource1.DataSource as System.Data.DataTable);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
        //DocName	DocKind	DocNum	U_PC_BSTY1	U_PC_BSDAT	U_PC_BSINV	U_PC_BSAPP	U_PC_BSTY2	
        //U_PC_BSTY3	U_PC_BSTY4	U_PC_BSTY5	U_PC_BSTYC	U_PC_BSNOT	U_PC_BSAMN	U_PC_BSTAX	U_PC_BSAMT	
        //U_PC_BSCUS	U_IN_BSCLS	U_ACME_SHIPWORKDAY	U_IN_BSDTO	U_IN_BSTY7	U_IN_BSREN	SEQ	U_PC_BSTYI	DocTotal	VatSum	發票月份

        private void TOTAL2GG(System.Data.DataTable dt, System.Data.DataTable dt2, System.Data.DataTable dt3)
        {

          dtGetAcmeStageG = MakeTableJ1();
            System.Data.DataTable DT1 = dt;
            System.Data.DataTable DT2 = dt2;
            System.Data.DataTable DT3 = dt3;
            CHITABLE(DT1);
            if (DT2.Rows.Count > 0)
            {
                CHITABLE(DT2);
            }
            if (comboBox1.Text == "聿豐")
            {
                if (DT3.Rows.Count > 0)
                {
                    CHITABLE(DT3);
                }
            }
        
        }

        private void TOTAL2GARMAS(System.Data.DataTable dt, System.Data.DataTable DT2)
        {

            dtGetAcmeStageG = MakeTableJ1();
            System.Data.DataTable DT1 = dt;

            CHITABLE(DT1);

            if (DT2.Rows.Count > 0)
            {
                CHITABLE(DT2);
            }

        }
        private void TOTAL2GGAD(System.Data.DataTable dt, System.Data.DataTable dt2)
        {

            dtGetAcmeStageG = MakeTableJ1();
            System.Data.DataTable DT1 = dt;
            System.Data.DataTable DT2 = dt2;
        
            CHITABLE(DT1);
            if (DT2.Rows.Count > 0)
            {
                CHITABLE(DT2);
            }
      

        }
        private void CHITABLE(System.Data.DataTable DT2)
        {
            for (int i = 0; i <= DT2.Rows.Count - 1; i++)
            {
                DataRow dr = null;
                dr = dtGetAcmeStageG.NewRow();

                dr["DocName"] = DT2.Rows[i]["DocName"].ToString().Trim();
                dr["DocKind"] = DT2.Rows[i]["DocKind"].ToString().Trim();
                dr["DocNum"] = DT2.Rows[i]["DocNum"].ToString().Trim();
                dr["U_PC_BSTY1"] = DT2.Rows[i]["U_PC_BSTY1"].ToString().Trim();
                dr["U_PC_BSDAT"] = Convert.ToDateTime(DT2.Rows[i]["U_PC_BSDAT"]);
                dr["U_PC_BSINV"] = DT2.Rows[i]["U_PC_BSINV"].ToString().Trim();
     
                dr["U_PC_BSAPP"] = Convert.ToDateTime(DT2.Rows[i]["U_PC_BSAPP"]);
                dr["U_PC_BSTY2"] = DT2.Rows[i]["U_PC_BSTY2"].ToString().Trim();
                dr["U_PC_BSTY3"] = DT2.Rows[i]["U_PC_BSTY3"].ToString().Trim();
                dr["U_PC_BSTY4"] = DT2.Rows[i]["U_PC_BSTY4"].ToString().Trim();
                dr["U_PC_BSTY5"] = DT2.Rows[i]["U_PC_BSTY5"].ToString().Trim();
                dr["U_PC_BSTYC"] = DT2.Rows[i]["U_PC_BSTYC"].ToString().Trim();
                dr["U_PC_BSNOT"] = DT2.Rows[i]["U_PC_BSNOT"].ToString().Trim();
                dr["U_PC_BSAMN"] = Convert.ToDecimal(DT2.Rows[i]["U_PC_BSAMN"].ToString());
                dr["U_PC_BSTAX"] = DT2.Rows[i]["U_PC_BSTAX"].ToString().Trim();
                dr["U_PC_BSAMT"] = DT2.Rows[i]["U_PC_BSAMT"].ToString().Trim();
                dr["U_PC_BSCUS"] = DT2.Rows[i]["U_PC_BSCUS"].ToString().Trim();
                dr["U_IN_BSCLS"] = DT2.Rows[i]["U_IN_BSCLS"].ToString().Trim();
                dr["U_ACME_SHIPWORKDAY"] = DT2.Rows[i]["U_ACME_SHIPWORKDAY"].ToString().Trim();
                dr["U_IN_BSDTO"] = DT2.Rows[i]["U_IN_BSDTO"].ToString().Trim();
                dr["U_IN_BSTY7"] = DT2.Rows[i]["U_IN_BSTY7"].ToString().Trim();
                dr["U_IN_BSREN"] = DT2.Rows[i]["U_IN_BSREN"].ToString().Trim();
                dr["SEQ"] = DT2.Rows[i]["SEQ"].ToString().Trim();
                dr["U_PC_BSTYI"] = DT2.Rows[i]["U_PC_BSTYI"].ToString().Trim();
                dr["DocTotal"] = Convert.ToDecimal(DT2.Rows[i]["DocTotal"].ToString());
                dr["VatSum"] = Convert.ToDecimal(DT2.Rows[i]["VatSum"].ToString());
                dr["發票月份"] = DT2.Rows[i]["發票月份"].ToString().Trim();
                dr["APP"] = DT2.Rows[i]["APP"].ToString().Trim();
                dtGetAcmeStageG.Rows.Add(dr);
            }

        }
        private void OutputSAPInvoice(System.Data.DataTable dt)
        {
            //27557058
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string INV = "";
            //20080711
            if ( comboBox1.Text == "聿豐")
            {
                INV = "22468373.txt";
            }
            else if (comboBox1.Text == "忠孝")
            {
                INV = "73718819.txt";
            }
            else  if ( comboBox1.Text == "宇豐")
            {
                INV = "27557058.txt";
            }
            else if (comboBox1.Text == "韋峰")
            {
                INV = "50881943.txt";
            }
            else
            {
                INV = INVOICE + ".txt";
            }
            string fileName = lsAppDir + "\\Excel\\temp\\" + INV;


            string InvoiceDate="";
            FileStream fs = new FileStream(fileName, FileMode.Create);


            StreamWriter r = new StreamWriter(fs);

            //取前一個月
            //string PriorMonth = DateToStr(StrToDate(textBox18.Text + "01").AddMonths(-1)).Substring(0, 6);

            //改為申報月
            string PriorMonth = DateToStr(StrToDate(textBox18.Text + "01")).Substring(0, 6);
            //  r.WriteLine("12345");

            //格式  
            string F0102 = string.Empty;
            //稅籍編號
            string F0311 = string.Empty;
            //流水號
            string F1218 = string.Empty;
            //年度
            string F1920 = string.Empty;
            //月份
            string F2122 = string.Empty;
            //買受人統一編號
            string F2330 = string.Empty;
            //銷售人統一編號
            string F3138 = string.Empty;
            //發票字軌
            string F3940 = string.Empty;
            //發票號碼
            string F4148 = string.Empty;
            //銷售金額
            string F4960 = string.Empty;
            //課稅別
            string F6161 = string.Empty;
            //營業稅額
            string F6271 = string.Empty;
            //扣抵代號
            string F7272 = string.Empty;
            //空白
            string F7377 = string.Empty;
            //特種稅額稅率
            string F7878 = string.Empty;
            //彙加註記
            string F7979 = string.Empty;
            //通關方式註記
            string F8080 = string.Empty;

            //整列
            string sLine = string.Empty;

            //流水號
            int iF1218 = 1;


            string tmpF0102 = string.Empty;

            //發票號碼
            string tmpF3940 = string.Empty;

            //文件號碼
            string DocNum = string.Empty;
            //文件名稱
            string DocName = string.Empty;
            string DocKind = string.Empty;
            string APP = string.Empty;

            string F3134 = "    ";
            string F3548 = string.Empty;


            //20080711 零稅率 
            //銷售人統一編號
            string Z0108 = string.Empty;
            //銷售人縣市別
            string Z0909 = string.Empty;
            //銷售人稅籍編號
            string Z1018 = string.Empty;
            //資料所屬年月
            string Z1923 = string.Empty;
            //開立發票年月
            string Z2428 = string.Empty;
            //字軌號碼
            string Z2938 = string.Empty;
            //買受人統一編號
            string Z3946 = string.Empty;
            //外銷方式
            string Z4747 = string.Empty;
            //通關方式註記
            string Z4848 = string.Empty;
            //出口報單類別
            string Z4950 = string.Empty;
            //出口報單號碼
            string Z5164 = string.Empty;
            //金額
            string Z6576 = string.Empty;
            //輸出或結匯日期
            string Z7783 = string.Empty;

            //整列
            string ZLine = string.Empty;
            //20080711
            string INV2 = "";
            //20080711
            if (comboBox1.Text == "聿豐")
            {
                INV2 = "22468373.t02";
            }
            else if (comboBox1.Text == "忠孝")
            {
                INV2 = "73718819.t02";
            }
            else if (comboBox1.Text == "宇豐")
            {
                INV2 = "27557058.t02";
            }
            else if (comboBox1.Text == "韋峰")
            {
                INV2 = "50881943.t02";
            }
            else
            {
                INV2 = INVOICE + ".t02";
            }
        
            string ZfileName = lsAppDir + "\\Excel\\temp\\" + INV2;
   
            FileStream Zfs = new FileStream(ZfileName, FileMode.Create);
            StreamWriter Zr = new StreamWriter(Zfs);

            System.Data.DataTable dtZfs = MakeTable(); 
            DataRow  drNew;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
             
               
                DocNum = Convert.ToString(dt.Rows[i]["DocNum"]);
   
                DocName = Convert.ToString(dt.Rows[i]["DocName"]);
                DocKind = Convert.ToString(dt.Rows[i]["DocKind"]);

                tmpF0102 = Convert.ToString(dt.Rows[i]["U_PC_BSTY1"]);

                if (comboBox1.Text != "")
                {
                    APP = Convert.ToString(dt.Rows[i]["APP"]);
                }


                //if (DocNum == "56986")
                //{
                //    MessageBox.Show("aa");
                //}
                //格式

                //進項

                if (DocKind == "2")
                {

                    //免用時不轉

                    if (tmpF0102 == "4" || string.IsNullOrEmpty(tmpF0102))
                    {

                        continue;

                    }



                    //修正AP 貸

                    if (DocName == "AP貸項")
                    {

                        F0102 = "23";

                    }

                    else
                    {

                        F0102 = ConvertF0102_IN(DocNum, tmpF0102);

                    }

                }
                else
                {
                    if (DocName == "AR發票")
                    {
                        //20210205
                        if (tmpF0102 == "5" )
                        {
                            //免用時不轉
                            if (DocKind == "3" &&  ConviertF6161(Convert.ToString(dt.Rows[i]["U_PC_BSTY2"])) == "2" && DocName == "AR發票")
                            {

                            }
                            else
                            {
                                continue;
                            }
                        }
              


                        F0102 = ConvertF0102_Out(DocNum, tmpF0102);

                        if (comboBox1.Text == "聿豐" || comboBox1.Text == "忠孝" || comboBox1.Text == "宇豐" || comboBox1.Text == "韋峰")
                        {
                            F0102 = Convert.ToString(dt.Rows[i]["SEQ"]);

                        }
                        
                    }
                    //AR 貸項 34,35
                    if (DocName == "AR貸項")
                    {

                        F0102 = "33";
                    }

                }
                System.Data.DataTable G1 = GetID();
                string ID = G1.Rows[0][0].ToString();
                //稅籍編號140106837
                if (comboBox1.Text == "聿豐")
                {
                    F0311 = "721402154";
                }
                 else if (comboBox1.Text == "忠孝")
                {
                    F0311 = "140106837";
                }
                else if (comboBox1.Text == "宇豐")
                {
                    F0311 = "720225836";
                }
                else if (comboBox1.Text == "韋峰")
                {
                    F0311 = "720529618";
                }
                else
                {
                    F0311 = ID;
                }
                //流水號
                F1218 = iF1218.ToString("0000000");

                iF1218++;

                //年度
                F1920 = ConvertYear(PriorMonth.Substring(0, 4));

                //月份
                if (comboBox1.Text == "宇豐" || comboBox1.Text == "忠孝" || comboBox1.Text == "聿豐" || comboBox1.Text == "韋峰")
                {
                    F2122 = dt.Rows[i]["發票月份"].ToString().Trim();
                }
                else
                {
                    F2122 = PriorMonth.Substring(4, 2);
                }

                //進項的日期處理

                //20081113
                string t1 = dt.Rows[i]["U_PC_BSDAT"].ToString();
                InvoiceDate = DateToStr(Convert.ToDateTime(dt.Rows[i]["U_PC_BSDAT"]));

                if (DocKind == "2")
                {  //U_PC_BSDAT



                   // InvoiceDate = DateToStr(Convert.ToDateTime(dt.Rows[i]["U_PC_BSDAT"]));

                    //年度

                    F1920 = ConvertYear(InvoiceDate.Substring(0, 4));



                    //月份

                    F2122 = InvoiceDate.Substring(4, 2);

                }


                ////買受人統一編號
                if (DocKind == "2")
                {
                    if (comboBox1.Text == "聿豐")
                    {
                        F2330 = "22468373";
                    }
                    else if (comboBox1.Text == "忠孝")
                    {
                        F2330 = "73718819";
                    }
                   else  if (comboBox1.Text == "宇豐")
                    {
                        F2330 = "27557058";
                    }
                    else if (comboBox1.Text == "韋峰")
                    {
                        F2330 = "50881943";
                    }
                    else
                    {

                        F2330 = INVOICE;
                    }
                }
                else
                {
                    F2330 = Convert.ToString(dt.Rows[i]["U_PC_BSNOT"]);

                }




                F2330 = F2330.Replace("_", "");

                if (string.IsNullOrEmpty(F2330))
                {
                    MsgLine.Text += string.Format("{0}單號:{1}　買受人統一編號未輸入 " + "\r", DocName, DocNum);

                    F2330 = PadStrRight(F2330, 8, " ");
                }
                else if (F2330.Length > 8)
                {

                    MsgLine.Text += string.Format("{0}單號:{1}　買受人統一編號長度超過 8 " + "\r", DocName, DocNum);

                    F2330 = "00000000";
                }


                ////銷售人統一編號
                if (DocKind == "2")
                {
                    F3138 = Convert.ToString(dt.Rows[i]["U_PC_BSNOT"]);

                    F3138 = F3138.Replace("_", "");

                    if (string.IsNullOrEmpty(F3138))
                    {
                        MsgLine.Text += string.Format("{0}單號:{1}　銷售人統一編號未輸入 " + "\r", DocName, DocNum);

                        F3138 = PadStrRight(F3138, 8, " ");
                    }
                    else if (F3138.Length > 8)
                    {

                        MsgLine.Text += string.Format("{0}單號:{1}　銷售人統一編號長度超過 8 " + "\r", DocName, DocNum);

                        F3138 = "        ";
                    }
                }
                else
                {
                    if (comboBox1.Text == "聿豐")
                    {
                        F3138 = "22468373";
                    }
                    else if (comboBox1.Text == "忠孝")
                    {
                        F3138 = "73718819";
                    }
                    else if (comboBox1.Text == "宇豐")
                    {
                        F3138 = "27557058";
                    }
                    else if (comboBox1.Text == "韋峰")
                    {
                        F3138 = "50881943";
                    }
                    else
                    {
                        F3138 = INVOICE;
                    }
                }

                tmpF3940 = Convert.ToString(dt.Rows[i]["U_PC_BSINV"]);



               


                tmpF3940 = tmpF3940.Replace("_", "");


                if (string.IsNullOrEmpty(tmpF3940))
                {
                    MsgLine.Text += string.Format("{0}單號:{1}　發票號碼未輸入 " + "\r", DocName, DocNum);
                }
                else if (tmpF3940.Length != 10)
                {
                    MsgLine.Text += string.Format("{0}單號:{1}發票號碼:{2}　發票號碼長度不足 " + "\r", DocName, DocNum, tmpF3940);
                }

                try
                {
                    ////發票字軌
                    F3940 = tmpF3940.Substring(0, 2);
                }
                catch
                {
                    F3940 = "  ";
                }

                try
                {
                    ////發票號碼
                    F4148 = tmpF3940.Substring(2, 8);
                  
                }
                catch
                {
             
                    if (tmpF3940.Length - 2 >= 0)
                    {
                        F4148 = PadStrRight(tmpF3940.Substring(2, Convert.ToString(dt.Rows[i]["U_PC_BSINV"]).Length - 2),
                            8, " ");
                    }
                    else
                    {
                        F4148 = "        ";
                    }
                }

        


                ////課稅別
                F6161 = ConviertF6161(Convert.ToString(dt.Rows[i]["U_PC_BSTY2"]));



                //海關代徴
                //if (F0102 == "28" || F6161 == "2")
                if (F0102 == "28")
                {
                    F3548 = Convert.ToString(dt.Rows[i]["U_PC_BSCUS"]).Replace("_", "");

                    if (F3548.Length != 14)
                    {
                        MsgLine.Text += string.Format("{0}單號:{1}　海關代徵營業稅繳納證號碼長度錯誤 " + "\r", DocName, DocNum);
                    }

                }
                else
                {

                    F3548 = Convert.ToString(dt.Rows[i]["U_PC_BSCUS"]).Replace("_", "");

                    if (F3548.Length == 14)
                    {
                        MsgLine.Text += string.Format("{0}單號:{1}　憑證類別錯誤有海關代徵營業稅繳納證號碼" + "\r", DocName, DocNum);
                    }

                }

                //if (F4148 == "63538853")
                //{
                //    MessageBox.Show("");
                //}



                ////銷售金額

                string g1 = dt.Rows[i]["U_PC_BSAMN"].ToString();

                string DocTotal = dt.Rows[i]["DocTotal"].ToString();
                int T1 = Convert.ToInt32(dt.Rows[i]["DocTotal"]);
                int T2 = Convert.ToInt32(dt.Rows[i]["VatSum"]);
                F4960 = Convert.ToInt32(dt.Rows[i]["U_PC_BSAMN"]).ToString("000000000000");

                if (Convert.ToInt32(dt.Rows[i]["U_PC_BSAMN"]) !=
                    Convert.ToInt32(dt.Rows[i]["DocTotal"]) - Convert.ToInt32(dt.Rows[i]["VatSum"]))
                {
                    MsgLine.Text += string.Format("{0}單號:{1}　銷售金額與文件金額不符" + "\r", DocName, DocNum);
                }


                //if (Convert.ToInt32(dt.Rows[i]["U_PC_BSTAX"]) != Convert.ToInt32(dt.Rows[i]["VatSum"]))
                //{
                //    MsgLine.Text += string.Format("{0}單號:{1}　稅額與文件稅額不符" + "\r", DocName, DocNum);
                //}



                //零稅
                if (F6161 == "2")
                {
                    if (Convert.ToInt32(dt.Rows[i]["U_PC_BSTAX"]) != 0)
                    {
                        MsgLine.Text += string.Format("{0}單號:{1}　零稅率之稅額應為零" + "\r", DocName, DocNum);
                    }

                }
                else if (F6161 == "1")
                {
                    if (Convert.ToInt32(dt.Rows[i]["U_PC_BSTAX"]) == 0)
                    {
                        MsgLine.Text += string.Format("{0}單號:{1}　應稅之稅額不應為零" + "\r", DocName, DocNum);
                    }

                }
            

                ////營業稅額
                F6271 = Convert.ToInt32(dt.Rows[i]["U_PC_BSTAX"]).ToString("0000000000");

                //if (F2330.Trim() == "")
                //{
                //    F4960 = (Convert.ToInt32(dt.Rows[i]["U_PC_BSAMN"]) + Convert.ToInt32(dt.Rows[i]["U_PC_BSTAX"])).ToString("000000000000");

                //    //20090908 修正
                //    //F6271 = "0000000000";
                //}

                if (F0102 == "22")
                {
                    if (F4148 != "")
                    {
                        //    if (isNumber(F4148.Substring(0, 2)))
                        //{

                        //            F4960 = (Convert.ToInt32(dt.Rows[i]["U_PC_BSAMN"]) + Convert.ToInt32(dt.Rows[i]["U_PC_BSTAX"])).ToString("000000000000");

                        //            F6271 = "0000000000";
                        //        }
                        //if (isNumber(F3940))
                        //{
                        //    F4960 = (Convert.ToInt32(dt.Rows[i]["U_PC_BSAMN"]) + Convert.ToInt32(dt.Rows[i]["U_PC_BSTAX"])).ToString("000000000000");
                        //    //20090908
                        //    //F6271 = "0000000000";

                        //}

                    }

                }

                ////扣抵代號
                F7272 = ConviertF6161(Convert.ToString(dt.Rows[i]["U_PC_BSTY5"]),
                                      Convert.ToString(dt.Rows[i]["U_PC_BSTY4"]));

                if (F0102 == "23")
                {
                    F7272 = "1";
                }
                //1234
                if (APP == "輔助")
                {
                    F7272 = dt.Rows[i]["U_PC_BSTY4"].ToString();

                }
                ////空白
                F7377 = "     ";
                ////特種稅額稅率
                F7878 = " ";
                ////彙加註記
                F7979 = " ";
                ////通關方式註記
                string GG = "2";
                //if (comboBox1.Text == "宇豐")
                //{
                //    GG = "3";
                //}
                //零稅率 - 2011/10/13 進項 不要管通關方式
                if (F6161 == GG)
                {

                    if (F0102.Substring(0, 1) == "2")
                    {
                        F8080 = " ";
                    }
                    else
                    {

                        if (Convert.ToString(dt.Rows[i]["U_PC_BSTY3"]) == "0")
                        {

                              F8080 = "2";
                           
                        }
                        else if (Convert.ToString(dt.Rows[i]["U_PC_BSTY3"]) == "1")
                        {
                             F8080 = "1";
                           
                        }
                        else
                        {
                            F8080 = " ";
                        }
                    }

                }
                else
                {

                    F8080 = " ";
                }



                //海關代徴
                //if (F0102 == "28" || F6161 == "2")
                if (F0102 == "28")
                {
                    sLine =
                        //格式
                                        F0102 +
                        //稅籍編號
                                        F0311 +
                        //流水號
                                        F1218 +
                        //年度
                                        F1920 +
                        //月份
                                        F2122 +
                        //買受人統一編號
                                        F2330 +

                                        F3134 +

                                        F3548 +
                        //銷售金額
                                        F4960 +
                        //課稅別
                                        F6161 +
                        //營業稅額
                                        F6271 +
                        //扣抵代號
                                        F7272 +
                        //空白
                                        F7377 +
                        //特種稅額稅率
                                        F7878 +
                        //彙加註記
                                        F7979 +
                        //通關方式註記
                                        F8080;

                }
                else
                {
                    sLine =
                        //格式
                    F0102 +
                        //稅籍編號
                    F0311 +
                        //流水號
                    F1218 +
                        //年度
                    F1920 +
                        //月份
                    F2122 +
                        //買受人統一編號
                    F2330 +
                        //銷售人統一編號
                    F3138 +
                        //發票字軌
                    F3940 +
                        //發票號碼
                    F4148 +
                        //銷售金額
                    F4960 +
                        //課稅別
                    F6161 +
                        //營業稅額
                    F6271 +
                        //扣抵代號
                    F7272 +
                        //空白
                    F7377 +
                        //特種稅額稅率
                    F7878 +
                        //彙加註記
                    F7979 +
                        //通關方式註記
                    F8080;
                }



                //20090811
                if (sLine.Length != 81)
                {
                    //MsgDocEntry.Lines
                    MsgDocEntry.Text += string.Format("進銷項->{0} 單號:{1} 發票號碼:{2} 資料長度錯誤 " + "\r", DocName, DocNum, F3940 + F4148);
                }

                if (F0102.Trim() != "")
                {
                    r.WriteLine(sLine);
                }
                //20090812
      

               

                //if (F4148 == "63538853")
                //{
                //    MessageBox.Show("63538853");
                //}

                //20080711 銷售零稅率轉出
                //條件銷項 & 課稅別 = 2 


       
                if (DocKind == "3" && F6161 == GG && DocName == "AR發票")
                {

                    //MessageBox.Show(F4148);
                    //銷售人統一編號
                    Z0108 = F3138;
                    //銷售人縣市別
                    Z0909 = "A";//台北市
                    //銷售人稅籍編號
                    Z1018 = F0311;
                    //資料所屬年月
                    //注意民國 100 年
                    //if (F1920 == "00")
                    //{
                    //    Z1923 = "1" + F1920 + F2122;
                    //}
                    //else
                    //{
                    //20090811
                    //Z1923 = "0" + F1920 + F2122;
                    // }
                    Z1923 = F1920 + F2122;
                    //開立發票年月
                    Z2428 = Z1923;
                    //字軌號碼
                    Z2938 = F3940 + F4148;

                    if (Z2938.Length < 10)
                    {
                        Z2938 = PadStrRight(Z2938, 10, " ");
                        MsgLine.Text += string.Format("{0}單號:{1}發票號碼:{2}　發票號碼長度不足" + "\r", DocName, DocNum, Z2938);

                    }
                    //買受人統一編號
                    Z3946 = F2330;

                    if (Z3946.Length < 8)
                    {
                        Z3946 = PadStrRight(Z3946, 8, " ");
                        MsgLine.Text += string.Format("{0}單號:{1}　買受人統一編號長度不足" + "\r", DocName, DocNum);

                    }

                    //外銷方式
                    Z4747 = Convert.ToString(dt.Rows[i]["U_IN_BSTY7"]);
                    //通關方式註記
                    Z4848 = F8080;







                    //出口報單類別
                    Z4950 = Convert.ToString(dt.Rows[i]["U_IN_BSCLS"]);

                    if (Z4950.Length < 2)
                    {
                        Z4950 = PadStrRight(Z4950, 2, " ");
                        MsgLine.Text += string.Format("{0}單號:{1}　出口報單類別錯誤" + "\r", DocName, DocNum);

                    }


                    //20090811
                    //以通通關方式註記 1 ->非經海關  2->經海關 
                    if (Z4848 == "1")
                    {
                        //出口報單號碼
                        Z5164 = Convert.ToString(dt.Rows[i]["U_ACME_SHIPWORKDAY"]);
                        if (Z5164.Length < 14)
                        {
                            Z5164 = PadStrRight(Z5164, 14, " ");
                            MsgLine.Text += string.Format("{0}單號:{1}　出口報單號碼" + "\r", DocName, DocNum);

                        }


                    }
                    else
                    {
                        //U_IN_BSREN
                        //出口報單號碼
                        Z5164 = Convert.ToString(dt.Rows[i]["U_IN_BSREN"]);
                        if (Z5164.Length < 14)
                        {
                            Z5164 = PadStrRight(Z5164, 14, " ");
                            MsgLine.Text += string.Format("{0}單號:{1}　出口報單號碼" + "\r", DocName, DocNum);

                        }

                    }





                    //金額
                    Z6576 = F4960;

                    if (Z6576.Length < 12)
                    {
                        Z6576 = PadStrRight(Z6576, 12, " ");
                        MsgLine.Text += string.Format("{0}單號:{1}　出口報單號碼" + "\r", DocName, DocNum);

                    }

                    //輸出或結匯日期




                    string outputDate = "       ";



                    try
                    {

                        if (Z4747 == "4")
                        {
                            outputDate = ConvertYear100(InvoiceDate.Substring(0, 4)) + InvoiceDate.Substring(4, 4);
                        }
                        else
                        {

                            string FG = dt.Rows[i]["U_IN_BSDTO"].ToString();
                        
                            outputDate = DateToStr(Convert.ToDateTime(dt.Rows[i]["U_IN_BSDTO"]));
                        }
                    }
                    catch
                    {
                        MsgLine.Text += string.Format("{0}單號:{1}　輸出或結匯日期錯誤" + "\r", DocName, DocNum);

                    };

                    //年度

                    if (outputDate == "       ")
                    {
                        Z7783 = outputDate;

                    }
                    else
                    {
                        try
                        {
                            Z7783 = ConvertYear100(outputDate.Substring(0, 4)) + outputDate.Substring(4, 4);

                        }
                        catch
                        {
                            Z7783 = outputDate;
                            MsgLine.Text += string.Format("{0}單號:{1}　輸出或結匯日期錯誤" + "\r", DocName, DocNum);

                        }

                    }


                    //20090811 這個不需要
                    ///外銷方式+通關方式註記
                    if (Z4747 + Z4848 == "11")
                    {

                        Z4950 = "  ";

                        Z5164 = "              ";

                    }

                    if (Z4747 == "1")
                    {
                        Z3946 = "        ";
                    }

                    ZLine =
                        //銷售人統一編號
                    Z0108 +
                        //銷售人縣市別
                     Z0909 +
                        //銷售人稅籍編號
                     Z1018 +
                        //資料所屬年月
                     Z1923 +
                        //開立發票年月
                     Z2428 +
                        //字軌號碼
                     Z2938 +
                        //買受人統一編號
                     Z3946 +
                        //外銷方式
                     Z4747 +
                        //通關方式註記
                     Z4848 +
                        //出口報單類別
                     Z4950 +
                        //出口報單號碼
                     Z5164 +
                        //金額
                     Z6576 +
                        //輸出或結匯日期
                     Z7783;



                    //20090811
                    if (ZLine.Length != 83)
                    {
                        //MsgDocEntry.Lines
                        MsgDocEntry.Text += string.Format("零稅率->{0}單號:{1} 發票號碼:{2} 資料長度錯誤 " + "\r", DocName, DocNum, Z2938);
                    }


                    // Zr.WriteLine(ZLine);


                    drNew = dtZfs.NewRow();


                    drNew["通關方式"] = Z4848;
                    drNew["外銷方式"] = Z4747;
                    drNew["日期"] = Z7783;
                    drNew["資料"] = ZLine;

                    dtZfs.Rows.Add(drNew);



                }




            }// for 迴圈



            DataView dv = dtZfs.DefaultView;

            dv.Sort = "通關方式,外銷方式,日期";

            System.Data.DataTable dtView = dv.ToTable();

            for (int i = 0; i <= dtView.Rows.Count - 1; i++)
            {

                Zr.WriteLine(Convert.ToString(dtView.Rows[i]["資料"]));

            }




            fs.Flush();
            r.Close();


            //20080711 銷售零稅率轉出
            Zfs.Flush();
            Zr.Close();


            System.Diagnostics.Process.Start(fileName);
            if (comboBox1.Text == "" || comboBox1.Text == "宇豐")
            {
                //利用 NotePad 開啟 副檔 .t02
                System.Diagnostics.Process.Start("notepad.exe", ZfileName);
            }

        }



        //日期處理--------------------------------------------------------------------------------------------
        private DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }


        private string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }

        //日期處理--------------------------------------------------------------------------------------------


        //取得發票明細

        //20080430 增加AP 貸項

        private System.Data.DataTable GetSAPInovice(string DocDate1, string DocDate2)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();



            //AP 發票

            sb.Append(" SELECT 'AP發票' as DocName, '2' as DocKind , T0.DocNum,T0.[U_PC_BSTY1], T0.[U_PC_BSDAT], T0.[U_PC_BSINV], T0.[U_PC_BSAPP], T0.[U_PC_BSTY2],");
            sb.Append("  T0.[U_PC_BSTY3], T0.[U_PC_BSTY4], T0.[U_PC_BSTY5], T0.[U_PC_BSTYC], T0.[U_PC_BSNOT], ");
            sb.Append("  T0.[U_PC_BSAMN], ROUND(T0.[U_PC_BSTAX],0) U_PC_BSTAX, T0.[U_PC_BSAMT], T0.[U_PC_BSCUS],");
            sb.Append("  '' AS U_IN_BSCLS,'' AS U_ACME_SHIPWORKDAY,'' AS U_IN_BSDTO,'' AS U_IN_BSTY7, ");
            sb.Append("  '' AS  U_IN_BSREN, ");
            sb.Append(" CASE ");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=0  THEN '21'");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=1  THEN '25'");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=2  THEN '22'");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=3  THEN '28'");
            sb.Append(" END  AS SEQ, ");
            sb.Append("  T0.[U_PC_BSTYI],T0.DocTotal,T0.VatSum FROM OPCH T0");
            sb.Append(" WHERE  T0.[U_PC_BSAPP] >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    T0.[U_PC_BSAPP] <= Convert(varchar(8),@DocDate2,112)");
            sb.Append(" AND    T0.[U_PC_BSTYC] ='0' ");

            //sb.Append(" ORDER BY U_PC_BSAPP");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'AP貸項' as DocName,'2' as DocKind , T1.[U_BSREN] as DocNum,T2.[U_RP_BSTY1], T2.[U_RP_BSDAT], T0.[U_BSINV], T2.[U_RP_BSAPP], T0.[U_BSTY2],");
            sb.Append(" '' U_PC_BSTY3, '' U_PC_BSTY4,'' U_PC_BSTY5, T2.[U_RP_BSTYC], T2.[U_RP_BSNOT], ");
            sb.Append("  T0.U_BSAMN2 as U_PC_BSAMN , ROUND(T0.U_BSTAX2,0) as U_PC_BSTAX,T0.U_BSAMN2+ROUND(T0.U_BSTAX2,0) [U_BSAMT],'' AS U_PC_BSCUS,   ");
            sb.Append("  '' AS U_IN_BSCLS,'' AS U_ACME_SHIPWORKDAY,'' AS U_IN_BSDTO,'' AS U_IN_BSTY7, ");
            sb.Append("  '' AS  U_IN_BSREN, ");
            sb.Append(" CASE ");
            sb.Append(" WHEN T2.[U_RP_BSTY1]=0  THEN '21'");
            sb.Append(" WHEN T2.[U_RP_BSTY1]=1  THEN '25'");
            sb.Append(" WHEN T2.[U_RP_BSTY1]=2  THEN '22'");
            sb.Append(" WHEN T2.[U_RP_BSTY1]=3  THEN '28'");
            sb.Append(" END  AS SEQ, ");
            sb.Append("  T0.[U_BSTYI],T0.U_BSAMN2+ROUND(T0.U_BSTAX2,0)  as DocTotal, ROUND(T0.U_BSTAX2,0)  as VatSum ");
            sb.Append("  FROM [@CADMEN_PMD1] T0");
            sb.Append("  left join [@CADMEN_PMD]  T1 on T0.DOCENTRY=T1.DOCENTRY");
            sb.Append("  left join [ORPC]  T2 on T1.U_BSREN=T2.DOCENTRY");
            sb.Append(" WHERE  T2.[U_RP_BSAPP] >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    T2.[U_RP_BSAPP] <= Convert(varchar(8),@DocDate2,112)");
            sb.Append(" UNION ALL");

            sb.Append(" SELECT 'AP分錄' as DocName,'2' as DocKind , T1.[U_BSREN] as DocNum,T0.[U_PC_BSTY1],CASE WHEN  T1.[U_BSREN]=411940 THEN T0.U_PC_BSDT2 ELSE  T0.[U_PC_BSDAT] END, T0.[U_PC_BSINV], T0.[U_PC_BSAPP], T0.[U_PC_BSTY2],");
            sb.Append("  T0.[U_PC_BSTY3], T0.[U_PC_BSTY4], T0.[U_PC_BSTY5], T0.[U_PC_BSTYC], T0.[U_PC_BSNOT], ");
            sb.Append("  T0.[U_PC_BSAMN], ROUND(T0.[U_PC_BSTAX],0) U_PC_BSTAX, T0.[U_PC_BSAMT],T0.[U_PC_BSCUS],   ");
            sb.Append("  '' AS U_IN_BSCLS,'' AS U_ACME_SHIPWORKDAY,'' AS U_IN_BSDTO,'' AS U_IN_BSTY7, ");
            sb.Append("  '' AS  U_IN_BSREN, ");
            sb.Append(" CASE ");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=0  THEN '21'");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=1  THEN '25'");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=2  THEN '22'");
            sb.Append(" WHEN T0.[U_PC_BSTY1]=3  THEN '28'");
            sb.Append(" END  AS SEQ, ");
            sb.Append("  T0.[U_PC_BSTYI], T0.U_PC_BSAMT as DocTotal,T0.U_PC_BSTAX as VatSum FROM [@CADMEN_FMD1] T0");
            sb.Append(" left join [@CADMEN_FMD]  T1 on T0.DOCENTRY=T1.DOCENTRY");
            sb.Append(" WHERE  (T0.[U_PC_BSAPP] >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    T0.[U_PC_BSAPP] <= Convert(varchar(8),@DocDate2,112))  AND (T1.[U_BSREN] <> '359509')");
            sb.Append(" UNION ALL");

            sb.Append("               SELECT 'AP貸項' as DocName,'2' as DocKind , T1.[U_BSREN] as DocNum,T0.[U_PC_BSTY1], T0.[U_PC_BSDT2], T0.[U_PC_BSINV], T0.[U_PC_BSAPP], T0.[U_PC_BSTY2], ");
            sb.Append("                ''  [U_PC_BSTY3], '' [U_PC_BSTY4], '' [U_PC_BSTY5], T0.[U_PC_BSTYC], T0.[U_PC_BSNOT],  ");
            sb.Append("                T0.[U_PC_BSAMN], ROUND(T0.[U_PC_BSTAX],0) U_PC_BSTAX, T0.[U_PC_BSAMT],T0.[U_PC_BSCUS],    ");
            sb.Append("                '' AS U_IN_BSCLS,'' AS U_ACME_SHIPWORKDAY,'' AS U_IN_BSDTO,'' AS U_IN_BSTY7,  ");
            sb.Append("                '' AS  U_IN_BSREN,'25'  AS SEQ,  ");
            sb.Append("                T0.[U_PC_BSTYI], T0.U_PC_BSAMT as DocTotal,T0.U_PC_BSTAX as VatSum FROM [@CADMEN_FMD1] T0 ");
            sb.Append("               left join [@CADMEN_FMD]  T1 on T0.DOCENTRY=T1.DOCENTRY ");
            sb.Append(" WHERE  (T0.[U_PC_BSAPP] >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    T0.[U_PC_BSAPP] <= Convert(varchar(8),@DocDate2,112))  AND (T1.[U_BSREN] = '359509')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'AR發票' as DocName, '3' as DocKind , DocNum,U_IN_BSTY1 U_PC_BSTY1,U_IN_BSDAT U_PC_BSDAT,U_IN_BSINV U_PC_BSINV,U_IN_BSAPP U_PC_BSAPP,U_IN_BSTY2 U_PC_BSTY2,");
            sb.Append("  U_IN_BSTY3 U_PC_BSTY3,U_IN_BSTY4 U_PC_BSTY4,'' U_PC_BSTY5,U_IN_BSTYC U_PC_BSTYC,U_IN_BSNOT U_PC_BSNOT, ");
            sb.Append("  U_IN_BSAMN U_PC_BSAMN,U_IN_BSTAX U_PC_BSTAX,U_IN_BSAMT U_PC_BSAMT,U_PC_BSCUS AS U_PC_BSCUS, ");
            sb.Append("  U_IN_BSCLS AS U_IN_BSCLS,U_ACME_SHIPWORKDAY AS U_ACME_SHIPWORKDAY,U_IN_BSDTO AS U_IN_BSDTO,CASE U_IN_BSTY7 WHEN 0 THEN 1 ELSE U_IN_BSTY7 END   AS U_IN_BSTY7, ");
            sb.Append("  U_IN_BSREN AS  U_IN_BSREN, ");
            sb.Append(" CASE ");
            sb.Append(" WHEN U_IN_BSTY1=0  THEN '31'");
            sb.Append(" WHEN U_IN_BSTY1=1  THEN '35'");
            sb.Append(" WHEN U_IN_BSTY1=2  THEN '32'");
            sb.Append(" WHEN U_IN_BSTY1=3  THEN '32'");
            sb.Append(" WHEN U_IN_BSTY1=4  THEN '31'");
            sb.Append(" END  AS SEQ, ");
            sb.Append("  U_IN_BSTYI U_PC_BSTYI,DocTotal,VatSum FROM OINV T0 ");
            sb.Append(" WHERE  T0.[U_IN_BSAPP] >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    T0.[U_IN_BSAPP] <= Convert(varchar(8),@DocDate2,112)");
            sb.Append(" AND    T0.[U_IN_BSTYC] ='0' ");


            //20080409
            //AR 貸項勾稽
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'AR貸項' as DocName,'3' as DocKind , T1.[U_BSREN] as DocNum,T2.[U_RI_BSTY1], T2.[U_RI_BSDAT], T0.[U_BSINV], T2.[U_RI_BSAPP], T0.[U_BSTY2],");
            sb.Append(" '' U_PC_BSTY3, '' U_PC_BSTY4,'' U_PC_BSTY5, T2.[U_RI_BSTYC], T2.[U_RI_BSNOT], ");
            sb.Append("       T0.U_BSAMN2 as U_PC_BSAMN , ROUND(T0.U_BSTAX2,0) as U_PC_BSTAX, T0.U_BSAMN2+ROUND(T0.U_BSTAX2,0) [U_BSAMT],U_PC_BSCUS AS U_PC_BSCUS,       ");
            sb.Append("  '' AS U_IN_BSCLS,'' AS U_ACME_SHIPWORKDAY,'' AS U_IN_BSDTO,'' AS U_IN_BSTY7, ");
            sb.Append("  '' AS  U_IN_BSREN, ");
            sb.Append(" '33'  AS SEQ, ");
            sb.Append("  T0.[U_BSTYI], T0.U_BSAMN2+ROUND(T0.U_BSTAX2,0) as DocTotal,ROUND(T0.U_BSTAX2,0)  as VatSum ");
            sb.Append("  FROM [@CADMEN_CMD1] T0");
            sb.Append("  left join [@CADMEN_CMD]  T1 on T0.DOCENTRY=T1.DOCENTRY");
            sb.Append("  left join [ORIN]  T2 on T1.U_BSREN=T2.DOCENTRY");
            sb.Append(" WHERE  T2.[U_RI_BSAPP] >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    T2.[U_RI_BSAPP] <= Convert(varchar(8),@DocDate2,112)");

            sb.Append(" ORDER BY SEQ,U_PC_BSINV");




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSAPInovice2(string DocDate1, string DocDate2, string COMPANY)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("          SELECT 'AP發票' as DocName, '2' as DocKind,T0.ID DocNum,SUBSTRING(U_PC_BSTY1,1,1) U_PC_BSTY1,");
            sb.Append("         CAST(DOCDATE AS DATETIME) U_PC_BSDAT,CASE ISNULL(SHIPDATE,'') WHEN '' THEN '__________' ELSE  SHIPDATE  END   U_PC_BSINV");
            sb.Append("         ,CAST(DOCDATE AS DATETIME) U_PC_BSAPP,'0' U_PC_BSTY2");
            sb.Append("         ,CASE WHEN U_IN_BSTAX<500 THEN 1 ELSE 0 END  U_PC_BSTY3,SUBSTRING(U_PC_BSTY4,1,1) U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC");
            sb.Append("         ,CASE ISNULL(UNIT,'') WHEN '' THEN '________' ELSE UNIT END U_PC_BSNOT,");
            sb.Append("            U_IN_BSAMN U_PC_BSAMN,U_IN_BSTAX U_PC_BSTAX,U_IN_BSAMT U_PC_BSAMT,'______________' AS U_PC_BSCUS,");
            sb.Append("         ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,");
            sb.Append(" CASE ");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=0  THEN '21'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=1  THEN '25'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=2  THEN '22'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=3  THEN '28'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=5  THEN '23'");
            sb.Append(" END  AS SEQ, ");
            sb.Append("         '0' U_PC_BSTYI,U_IN_BSAMT DocTotal, U_IN_BSTAX VatSum,SUBSTRING(DelRemark,5,2) 發票月份,'輔助' APP");
            sb.Append("         FROM dbo.GB_POTATOAR T0");
            sb.Append("         WHERE 1=1 and ISNULL(SHIPDATE,'') <> ''  AND COMPANY =@COMPANY AND SUBSTRING(U_PC_BSTY1,1,1) BETWEEN 0 AND 5 ");
            sb.Append(" AND  DelRemark >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    DelRemark <= Convert(varchar(8),@DocDate2,112)  ORDER BY DocName,DOCNUM");
        



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];

        }


        private System.Data.DataTable GetSAPInoviceAD(string DocDate1, string DocDate2)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("          SELECT 'AP發票' as DocName, '2' as DocKind,T0.ID DocNum,SUBSTRING(U_PC_BSTY1,1,1) U_PC_BSTY1,");
            sb.Append("         CAST(DOCDATE AS DATETIME) U_PC_BSDAT,CASE ISNULL(SHIPDATE,'') WHEN '' THEN '__________' ELSE  SHIPDATE  END   U_PC_BSINV");
            sb.Append("         ,CAST(DOCDATE AS DATETIME) U_PC_BSAPP,'0' U_PC_BSTY2");
            sb.Append("         ,CASE WHEN U_IN_BSTAX<500 THEN 1 ELSE 0 END  U_PC_BSTY3,SUBSTRING(U_PC_BSTY4,1,1) U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC");
            sb.Append("         ,CASE ISNULL(UNIT,'') WHEN '' THEN '________' ELSE UNIT END U_PC_BSNOT,");
            sb.Append("            U_IN_BSAMN U_PC_BSAMN,U_IN_BSTAX U_PC_BSTAX,U_IN_BSAMT U_PC_BSAMT,CASE  WHEN SUBSTRING(U_PC_BSTY1,1,1)=3 THEN SHIPDATE ELSE  '______________' END AS U_PC_BSCUS,");
            sb.Append("         ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,");
            sb.Append(" CASE ");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=0  THEN '21'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=1  THEN '25'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=2  THEN '22'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=3  THEN '28'");
            sb.Append(" WHEN SUBSTRING(U_PC_BSTY1,1,1)=5  THEN '23'");
            sb.Append(" END  AS SEQ, ");
            sb.Append("         '0' U_PC_BSTYI,U_IN_BSAMT DocTotal, U_IN_BSTAX VatSum,SUBSTRING(DelRemark,5,2) 發票月份,'輔助' APP");
            sb.Append("         FROM dbo.AD_INVOICEAP T0");
            sb.Append("         WHERE 1=1 and ISNULL(SHIPDATE,'') <> ''  AND SUBSTRING(U_PC_BSTY1,1,1) BETWEEN 0 AND 5 ");
            sb.Append(" AND  DelRemark >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    DelRemark <= Convert(varchar(8),@DocDate2,112)  ORDER BY DocName,DOCNUM");




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];

        }

        private System.Data.DataTable GetSAPInovice2F(string DocDate1, string DocDate2)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("                        SELECT DelRemark,'AR發票' as DocName, '3' as DocKind,T0.ID DocNum,3 U_PC_BSTY1, ");
            sb.Append("                       CAST(DOCDATE AS DATETIME) U_PC_BSDAT,CASE ISNULL(SHIPDATE,'') WHEN '' THEN '__________' ELSE  SHIPDATE  END   U_PC_BSINV ");
            sb.Append("                       ,CAST(DOCDATE AS DATETIME) U_PC_BSAPP,'F'  U_PC_BSTY2  ");
            sb.Append("                       ,CASE WHEN U_IN_BSTAX<500 THEN 1 ELSE 0 END  U_PC_BSTY3,'' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ");
            sb.Append("                       ,'________' U_PC_BSNOT, ");
            sb.Append("                          0 U_PC_BSAMN,0 U_PC_BSTAX,0 U_PC_BSAMT,'______________' AS U_PC_BSCUS, ");
            sb.Append("                       ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN, ");
            sb.Append("          32 SEQ,  ");
            sb.Append("                       '0' U_PC_BSTYI,U_IN_BSAMT DocTotal, U_IN_BSTAX VatSum,SUBSTRING(DelRemark,5,2) 發票月份,'輔助' APP ");
            sb.Append("                       FROM dbo.GB_POTATOARF T0 ");
            sb.Append(" WHERE  DelRemark >= Convert(varchar(8),@DocDate1,112)");
            sb.Append(" AND    DelRemark <= Convert(varchar(8),@DocDate2,112)  ORDER BY DocName,DOCNUM");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetInoviceADLAB(string DocDate1, string DocDate2)
        {
            SqlConnection connection = new SqlConnection(strCn16);
            StringBuilder sb = new StringBuilder();

  
            sb.Append(" SELECT DISTINCT  I.InvoiceType,'AR發票' as DocName, '3' as DocKind,CAST(I.SrcBillNO  AS VARCHAR)+'1' DocNum,CASE I.InvoiceType WHEN 31 THEN 4 ELSE 3 END U_PC_BSTY1,    ");
            sb.Append(" CAST(CAST(I.InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV    ");
            sb.Append(" , CAST(CAST(I.InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' when '2' then '1' else cast(i.taxtype as varchar) end U_PC_BSTY2,Z.ExportStyle U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,I.TaxRegNO U_PC_BSNOT,    ");
            sb.Append(" ISNULL(I.AMOUNT,0)U_PC_BSAMN,     ");
            sb.Append(" CASE I.InvoiceType WHEN 31 THEN CAST(ISNULL(I.TaxAmt,0) AS INT) WHEN 35 THEN CASE WHEN ISNULL(I.TaxRegNO,'')<> '' THEN CAST(ISNULL(I.TaxAmt,0) AS INT) ELSE 0 END ELSE 0 END  U_PC_BSTAX,     ");
            sb.Append(" ISNULL(I.AMOUNT,0)+(CASE I.InvoiceType WHEN 31 THEN CAST(ISNULL(I.TaxAmt,0) AS INT) WHEN 35 THEN CASE WHEN ISNULL(I.TaxRegNO,'')<> '' THEN CAST(ISNULL(I.TaxAmt,0) AS INT) ELSE 0 END ELSE 0 END) U_PC_BSAMT,'______________' AS U_PC_BSCUS     ");
            sb.Append(" ,  ExportType  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,Convert(varchar(10), CAST(CAST(Z.OutDate AS VARCHAR) AS DATETIME),111)  U_IN_BSDTO ,OutSaleStyle+1   U_IN_BSTY7,Z.ApplyNO  U_IN_BSREN,I.InvoiceType SEQ,     ");
            sb.Append(" '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum,    ");
            sb.Append(" SUBSTRING(CAST( CASE WHEN I.InvoiceType=33 THEN I.DistInvoDate ELSE I.InvoiceDate END  AS VARCHAR),5,2) 發票月份,'正航' APP     ");
            sb.Append(" FROM  comInvoice I   ");
            sb.Append(" left Join ComProdRec O On O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1     ");
            sb.Append(" left join OrdBillSub G On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO     ");
            sb.Append(" left join OrdBillMain A On   G.Flag=A.Flag  And G.BillNO=A.BillNO     ");
            sb.Append(" left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)    ");
            sb.Append(" left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1    ");
            sb.Append(" Left Join comProduct B On B.ProdID =G.ProdID     ");
            sb.Append(" LEFT JOIN StkBillMain V ON (O.BillNO =V.BillNO)     ");
            sb.Append(" LEFT JOIN StkZeroTax Z ON (I.SrcBillNO =Z.SrcBillNO AND Z.BillDate <>20190311 )   ");
            sb.Append(" WHERE I.InvoiceNO <> ''  AND InvoiceDate  > 20161231   ");
            sb.Append(" AND I.InvoiceType <> 36   AND IsCancel <> 1 AND I.InvoiceType <> 21   ");
            sb.Append(" AND CASE WHEN I.InvoiceType=33 THEN  CAST(I.DistInvoDate AS VARCHAR)   ELSE  CAST(I.InvoiceDate AS VARCHAR) END BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112)");
            sb.Append(" AND  I.InvoiceType <>23");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT DISTINCT  I.InvoiceType,'AP發票' as DocName, '2' as DocKind,CAST(I.SrcBillNO  AS VARCHAR)+'1' DocNum,5 U_PC_BSTY1,    ");
            sb.Append(" CAST(CAST(I.InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV    ");
            sb.Append(" , CAST(CAST(I.InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' when '2' then '1' else cast(i.taxtype as varchar) end U_PC_BSTY2,Z.ExportStyle U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,I.TaxRegNO U_PC_BSNOT,    ");
            sb.Append(" ISNULL(I.AMOUNT,0)U_PC_BSAMN,     ");
            sb.Append("  CAST(ISNULL(I.TaxAmt,0) AS INT)  U_PC_BSTAX,     ");
            sb.Append(" ISNULL(I.AMOUNT,0)+ CAST(ISNULL(I.TaxAmt,0) AS INT) U_PC_BSAMT,'______________' AS U_PC_BSCUS     ");
            sb.Append(" ,  ExportType  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,Convert(varchar(10), CAST(CAST(Z.OutDate AS VARCHAR) AS DATETIME),111)  U_IN_BSDTO ,OutSaleStyle+1   U_IN_BSTY7,Z.ApplyNO  U_IN_BSREN,I.InvoiceType SEQ,     ");
            sb.Append(" '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum,    ");
            sb.Append(" SUBSTRING(CAST(I.InvoiceDate   AS VARCHAR),5,2) 發票月份,'正航' APP     ");
            sb.Append(" FROM  comInvoice I   ");
            sb.Append(" left Join ComProdRec O On O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1     ");
            sb.Append(" left join OrdBillSub G On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO     ");
            sb.Append(" left join OrdBillMain A On   G.Flag=A.Flag  And G.BillNO=A.BillNO     ");
            sb.Append(" left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)    ");
            sb.Append(" left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1    ");
            sb.Append(" Left Join comProduct B On B.ProdID =G.ProdID     ");
            sb.Append(" LEFT JOIN StkBillMain V ON (O.BillNO =V.BillNO)     ");
            sb.Append(" LEFT JOIN StkZeroTax Z ON (I.SrcBillNO =Z.SrcBillNO AND Z.BillDate <>20190311 )   ");
            sb.Append(" WHERE I.InvoiceNO <> ''  AND InvoiceDate  > 20161231   ");
            sb.Append(" AND I.InvoiceType <> 36   AND IsCancel <> 1 AND I.InvoiceType <> 21   ");
            sb.Append(" AND CASE WHEN I.InvoiceType=33 THEN  CAST(I.DistInvoDate AS VARCHAR)   ELSE  CAST(I.InvoiceDate AS VARCHAR) END  BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112)");
            sb.Append(" AND  I.InvoiceType =23");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


  

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSAPInovice2CHIARMAS(string DocDate1, string DocDate2, string strCn, string COMPANY)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT DISTINCT  'AR發票' as DocName, '3' as DocKind,CAST(O.BillNO  AS VARCHAR)+'1' DocNum,CASE I.InvoiceType WHEN 31 THEN 4 ELSE 3 END U_PC_BSTY1, ");
            sb.Append("              CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV ");
            sb.Append("                  , CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' else cast(i.taxtype as varchar) end U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,I.TaxRegNO U_PC_BSNOT, ");
            sb.Append("                                   ISNULL(I.AMOUNT,0)U_PC_BSAMN,  ");
            sb.Append("                                                CASE I.InvoiceType WHEN 31 THEN CAST(ISNULL(I.TaxAmt,0) AS INT) WHEN 35 THEN CASE WHEN ISNULL(I.TaxRegNO,'')<> '' THEN CAST(ISNULL(I.TaxAmt,0) AS INT) ELSE 0 END ELSE 0 END  U_PC_BSTAX,  ");
            sb.Append("                                        ISNULL(I.AMOUNT,0)+(CASE I.InvoiceType WHEN 31 THEN CAST(ISNULL(I.TaxAmt,0) AS INT) WHEN 35 THEN CASE WHEN ISNULL(I.TaxRegNO,'')<> '' THEN CAST(ISNULL(I.TaxAmt,0) AS INT) ELSE 0 END ELSE 0 END) U_PC_BSAMT,'______________' AS U_PC_BSCUS  ");
            sb.Append("                                  ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType SEQ,  ");
            sb.Append("                                     '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum, ");
            sb.Append("                                     SUBSTRING(CAST(O.BillDate AS VARCHAR),5,2)  發票月份,'正航' APP  ");
            sb.Append("               FROM ComProdRec O  ");
            sb.Append("                  left join COMBILLACCOUNTS S ON (O.BillNO =S.CustID AND S.Flag =500) ");
            sb.Append("                   left join comCustomer U On  U.ID=S.CustID AND U.Flag =1 ");
            sb.Append("                   Left Join comProduct B On B.ProdID =O.ProdID  ");
            sb.Append("                       Left Join comInvoice I On O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1  ");
            sb.Append("                     LEFT JOIN StkBillMain V ON (O.BillNO =V.BillNO)  ");
            sb.Append("					             INNER JOIN COMINVO I2 ON ( I.InvoiceType =I2.InvoType AND SUBSTRING(I.InvoiceNO,1,2)=I2.HEAD)  ");
            sb.Append("              WHERE  O.Flag =500   ");
            sb.Append("               AND I.InvoiceNO <> ''  ");
            sb.Append(" AND  CAST(InvoiceDate AS VARCHAR)   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");

            if (COMPANY == "聿豐")
            {
                sb.Append("			  AND I2.ReMark ='22468373'");
            }
            if (COMPANY == "忠孝")
            {
                sb.Append("			  AND I2.ReMark ='73718819'");
            }

            if (COMPANY == "聿豐")
            {
                sb.Append("           UNION ALL");


                sb.Append("                                                    select DISTINCT  'AR發票' as DocName,'3' as DocKind,CAST(InvoBillNo AS VARCHAR) DocNum,CASE I.InvoiceType WHEN 31 THEN 4 ELSE 3 END U_PC_BSTY1,  ");
                sb.Append("                                                              CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV  ");
                sb.Append("                                    , CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' else cast(i.taxtype as varchar) end U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,TaxRegNO U_PC_BSNOT,  ");
                sb.Append("                                      ISNULL(I.AMOUNT,0) U_PC_BSAMN,CAST(ISNULL(I.TaxAmt,0) AS INT) U_PC_BSTAX,ISNULL(I.AMOUNT,0)+CAST(ISNULL(I.TaxAmt,0) AS INT) U_PC_BSAMT,'______________' AS U_PC_BSCUS  ");
                sb.Append("                                      ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType  SEQ,  ");
                sb.Append("                                         '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum,  ");
                sb.Append("                                         SUBSTRING(CAST(InvoiceDate AS VARCHAR),5,2)  發票月份,'正航' APP  ");
                sb.Append("                                                          from comInvoice I ");
                sb.Append("														  			");
                sb.Append("					             INNER JOIN COMINVO I2 ON ( I.InvoiceType =I2.InvoType AND SUBSTRING(I.InvoiceNO,1,2)=I2.HEAD)  ");
                sb.Append("														  where I.Flag =2 AND I.IsCancel <> 1   ");
                sb.Append(" AND  CAST(InvoiceDate AS VARCHAR)   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");
                sb.Append("                                                       and SUBSTRING(SrcBillNO,1,2) <> 'DN'  ");
                sb.Append("			  AND I2.ReMark ='22468373'");
            }
            sb.Append("           UNION ALL");
            //作廢
            //sb.Append("                                                select DISTINCT  'AR發票' as DocName,'3' as DocKind,'AA00000000' DocNum,4 U_PC_BSTY1,  ");
            //sb.Append("                                                          CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV  ");
            //sb.Append("                                , CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,'F' U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,'' U_PC_BSNOT,  ");
            //sb.Append("                                  0 U_PC_BSAMN,0 U_PC_BSTAX,0 U_PC_BSAMT,'______________' AS U_PC_BSCUS  ");
            //sb.Append("                                  ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType SEQ,  ");
            //sb.Append("                                     '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum,  ");
            //sb.Append("                                     SUBSTRING(CAST(InvoiceDate AS VARCHAR),5,2)  發票月份,'正航' APP  ");
            //sb.Append("               FROM ComProdRec O  ");
            //sb.Append("                  left join COMBILLACCOUNTS S ON (O.BillNO =S.CustID AND S.Flag =500) ");
            //sb.Append("                   left join comCustomer U On  U.ID=S.CustID AND U.Flag =1 ");
            //sb.Append("                   Left Join comProduct B On B.ProdID =O.ProdID  ");
            //sb.Append("                       Left Join comInvoice I On O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel = 1  ");
            //sb.Append("                     LEFT JOIN StkBillMain V ON (O.BillNO =V.BillNO)  ");
            //sb.Append("					 			             INNER JOIN COMINVO I2 ON (I.InvoiceType =I2.InvoType AND SUBSTRING(I.InvoiceNO,1,2)=I2.HEAD)  ");
            //sb.Append("              WHERE  O.Flag =500   ");
            //sb.Append("               AND I.InvoiceNO <> ''  ");
            sb.Append("                                                          select DISTINCT  'AR發票' as DocName,'3' as DocKind,'AA00000000' DocNum,4 U_PC_BSTY1,   ");
            sb.Append("                                                                       CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV   ");
            sb.Append("                                             , CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,'F' U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,'' U_PC_BSNOT,   ");
            sb.Append("                                               0 U_PC_BSAMN,0 U_PC_BSTAX,0 U_PC_BSAMT,'______________' AS U_PC_BSCUS   ");
            sb.Append("                                               ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType SEQ,   ");
            sb.Append("                                                  '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum,   ");
            sb.Append("                                                  SUBSTRING(CAST(InvoiceDate AS VARCHAR),5,2)  發票月份,'正航' APP   ");
            sb.Append("                            FROM  comInvoice I  ");
            sb.Append("                          INNER JOIN COMINVO I2 ON (I.InvoiceType =I2.InvoType AND SUBSTRING(I.InvoiceNO,1,2)=I2.HEAD)   ");
            sb.Append("                           WHERE  I.IsCancel = 1");

            sb.Append(" AND  CAST(APPLYMONTH AS VARCHAR)+'01'   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");
            if (COMPANY == "聿豐")
            {
                sb.Append("			  AND I2.ReMark ='22468373'");
            }
            if (COMPANY == "忠孝")
            {
                sb.Append("			  AND I2.ReMark ='73718819'");
            }
            sb.Append("           UNION ALL");
            //折讓
            sb.Append("                   SELECT DISTINCT  'AR貸項' as DocName, '3' as DocKind,CAST(O.BillNO  AS VARCHAR)+'1' DocNum,CASE I.InvoiceType WHEN 31 THEN 4 ELSE 3 END U_PC_BSTY1,  ");
            sb.Append("                            CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV  ");
            sb.Append("                                , CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' else cast(i.taxtype as varchar) end U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,I.TaxRegNO U_PC_BSNOT,  ");
            sb.Append("                                                 ISNULL(I.AMOUNT,0) U_PC_BSAMN,   ");
            sb.Append("                                                            CAST(ISNULL(I.TaxAmt,0) AS INT)   U_PC_BSTAX,   ");
            sb.Append("                                                      ISNULL(I.AMOUNT,0)+CAST(ISNULL(I.TaxAmt,0) AS INT) U_PC_BSAMT,'______________' AS U_PC_BSCUS   ");
            sb.Append("                                                ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType SEQ,   ");
            sb.Append("                                                   '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum,  ");
            sb.Append("                                                   SUBSTRING(CAST(O.BillDate AS VARCHAR),5,2)  發票月份,'正航' APP   ");
            sb.Append("                             FROM ComProdRec O   ");
            sb.Append("                LEFT join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo  AND  CASE O.Flag WHEN 701 THEN 698 ELSE O.Flag END=S.Flag)  ");
            sb.Append("                              LEFT Join comInvoice I On  O.BillNO=I.SrcBillNO AND I.Flag =4   ");
            sb.Append("							  		             INNER JOIN COMINVO I2 ON ( SUBSTRING(I.InvoiceNO,1,2)=I2.HEAD)  ");
            sb.Append("                            WHERE  O.Flag IN  (600,701) AND I.IsCancel <> 1        ");
            sb.Append(" AND  CAST(O.BillDate AS VARCHAR)   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");
            if (COMPANY == "聿豐")
            {
                sb.Append("			  AND I2.ReMark ='22468373'");
            }
            if (COMPANY == "忠孝")
            {
                sb.Append("			  AND I2.ReMark ='73718819'");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }



            return ds.Tables[0];

        }
        private System.Data.DataTable GetSAPInovice2CHI(string DocDate1, string DocDate2, string strCn,string COMPANY)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT  'AR發票' as DocName, '3' as DocKind,CAST(O.BillNO  AS VARCHAR)+'1' DocNum,CASE I.InvoiceType WHEN 31 THEN 4 ELSE 3 END U_PC_BSTY1,");
            sb.Append(" CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV");
            sb.Append("     , CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' else cast(i.taxtype as varchar) end U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,I.TaxRegNO U_PC_BSNOT,");
            sb.Append("                      ISNULL(I.AMOUNT,0)U_PC_BSAMN, ");
            sb.Append("                                   CASE I.InvoiceType WHEN 31 THEN CAST(ISNULL(I.TaxAmt,0) AS INT) WHEN 35 THEN CASE WHEN ISNULL(I.TaxRegNO,'')<> '' THEN CAST(ISNULL(I.TaxAmt,0) AS INT) ELSE 0 END ELSE 0 END  U_PC_BSTAX, ");
            sb.Append("                           ISNULL(I.AMOUNT,0)+(CASE I.InvoiceType WHEN 31 THEN CAST(ISNULL(I.TaxAmt,0) AS INT) WHEN 35 THEN CASE WHEN ISNULL(I.TaxRegNO,'')<> '' THEN CAST(ISNULL(I.TaxAmt,0) AS INT) ELSE 0 END ELSE 0 END) U_PC_BSAMT,'______________' AS U_PC_BSCUS ");
            sb.Append("                     ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType SEQ, ");
            sb.Append("                        '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum,");
            sb.Append("                        SUBSTRING(CAST(O.BillDate AS VARCHAR),5,2)  發票月份,'正航' APP ");
            sb.Append("  FROM ComProdRec O ");
            sb.Append("     left join COMBILLACCOUNTS S ON (O.BillNO =S.CustID AND S.Flag =500)");
            sb.Append("      left join comCustomer U On  U.ID=S.CustID AND U.Flag =1");
            sb.Append("      Left Join comProduct B On B.ProdID =O.ProdID ");
            sb.Append("          Left Join comInvoice I On O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1 ");
            sb.Append("        LEFT JOIN StkBillMain V ON (O.BillNO =V.BillNO) ");
            sb.Append(" WHERE  O.Flag =500  ");
            sb.Append("  AND I.InvoiceNO <> '' ");
            sb.Append(" AND  CAST(O.BillDate AS VARCHAR)   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");

            if (COMPANY == "聿豐")
            {
                sb.Append("           UNION ALL");
                sb.Append("                                   select DISTINCT  'AR發票' as DocName,'3' as DocKind,CAST(InvoBillNo AS VARCHAR) DocNum,CASE I.InvoiceType WHEN 31 THEN 4 ELSE 3 END U_PC_BSTY1, ");
                sb.Append("                                             CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV ");
                sb.Append("                   , CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' else cast(i.taxtype as varchar) end U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,TaxRegNO U_PC_BSNOT, ");
                sb.Append("                     ISNULL(I.AMOUNT,0) U_PC_BSAMN,CAST(ISNULL(I.TaxAmt,0) AS INT) U_PC_BSTAX,ISNULL(I.AMOUNT,0)+CAST(ISNULL(I.TaxAmt,0) AS INT) U_PC_BSAMT,'______________' AS U_PC_BSCUS ");
                sb.Append("                     ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType  SEQ, ");
                sb.Append("                        '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum, ");
                sb.Append("                        SUBSTRING(CAST(InvoiceDate AS VARCHAR),5,2)  發票月份,'正航' APP ");
                sb.Append("                                         from comInvoice I where I.Flag =2 AND I.IsCancel <> 1  ");
                sb.Append(" AND  CAST(InvoiceDate AS VARCHAR)   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");
                sb.Append("                                      and SUBSTRING(SrcBillNO,1,2) <> 'DN' ");
                if (COMPANY == "忠孝")
                {
                    sb.Append("                                     AND I.InvoiceType<> 36 ");
                }
            }
            sb.Append("           UNION ALL");
            //作廢
            sb.Append("                                   select DISTINCT  'AR發票' as DocName,'3' as DocKind,'AA00000000' DocNum,4 U_PC_BSTY1, ");
            sb.Append("                                             CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV ");
            sb.Append("                   , CAST(CAST(InvoiceDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,'F' U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,'' U_PC_BSNOT, ");
            sb.Append("                     0 U_PC_BSAMN,0 U_PC_BSTAX,0 U_PC_BSAMT,'______________' AS U_PC_BSCUS ");
            sb.Append("                     ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType SEQ, ");
            sb.Append("                        '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum, ");
            sb.Append("                        SUBSTRING(CAST(InvoiceDate AS VARCHAR),5,2)  發票月份,'正航' APP ");
            sb.Append("  FROM ComProdRec O ");

            sb.Append("     left join COMBILLACCOUNTS S ON (O.BillNO =S.CustID AND S.Flag =500)");
            sb.Append("      left join comCustomer U On  U.ID=S.CustID AND U.Flag =1");
            sb.Append("      Left Join comProduct B On B.ProdID =O.ProdID ");
            sb.Append("          Left Join comInvoice I On O.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel = 1 ");
            sb.Append("        LEFT JOIN StkBillMain V ON (O.BillNO =V.BillNO) ");
            sb.Append(" WHERE  O.Flag =500  ");
            sb.Append("  AND I.InvoiceNO <> '' ");
            sb.Append(" AND  CAST(O.BillDate AS VARCHAR)   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");
            sb.Append("           UNION ALL");
            //折讓
            sb.Append("               SELECT DISTINCT  'AR貸項' as DocName, '3' as DocKind,CAST(O.BillNO  AS VARCHAR)+'1' DocNum,CASE I.InvoiceType WHEN 31 THEN 4 ELSE 3 END U_PC_BSTY1, ");
            sb.Append("               CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSDAT,CASE ISNULL(I.InvoiceNO,'') WHEN '' THEN '__________' ELSE  I.InvoiceNO  END   U_PC_BSINV ");
            sb.Append("                   , CAST(CAST(O.BillDate AS VARCHAR)  AS DATETIME) U_PC_BSAPP,case cast(i.taxtype as varchar) when '1' then '2' else cast(i.taxtype as varchar) end U_PC_BSTY2,'1' U_PC_BSTY3,'0' U_PC_BSTY4,'' U_PC_BSTY5,0 U_PC_BSTYC ,I.TaxRegNO U_PC_BSNOT, ");
            sb.Append("                                    ISNULL(I.AMOUNT,0) U_PC_BSAMN,  ");
            sb.Append("                                               CAST(ISNULL(I.TaxAmt,0) AS INT)   U_PC_BSTAX,  ");
            sb.Append("                                         ISNULL(I.AMOUNT,0)+CAST(ISNULL(I.TaxAmt,0) AS INT) U_PC_BSAMT,'______________' AS U_PC_BSCUS  ");
            sb.Append("                                   ,  ''  U_IN_BSCLS,'' U_ACME_SHIPWORKDAY,'' U_IN_BSDTO ,'0'  U_IN_BSTY7,'' U_IN_BSREN,I.InvoiceType SEQ,  ");
            sb.Append("                                      '0' U_PC_BSTYI,  ISNULL(I.AMOUNT,0) DocTotal, ISNULL(I.TaxAmt,0) VatSum, ");
            sb.Append("                                      SUBSTRING(CAST(O.BillDate AS VARCHAR),5,2)  發票月份,'正航' APP  ");
            sb.Append("                FROM ComProdRec O  ");
            sb.Append("   LEFT join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo  AND  CASE O.Flag WHEN 701 THEN 698 ELSE O.Flag END=S.Flag) ");
            sb.Append("                 LEFT Join comInvoice I On  O.BillNO=I.SrcBillNO AND I.Flag =4  ");
            sb.Append("               WHERE  O.Flag IN  (600,701) AND I.IsCancel <> 1  ");
            sb.Append(" AND  CAST(O.BillDate AS VARCHAR)   BETWEEN Convert(varchar(8),@DocDate1,112) AND Convert(varchar(8),@DocDate2,112) ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];

        }

        //進項
        private string ConvertF0102_IN(string DocNum, string F0102)
        {
            switch (F0102)
            {
                //1 給系統 -1 使用

                case "0":
                    return "21";
                    break;
                case "1":
                    return "25";
                    break;
                case "2":
                    return "22";
                    break;
                case "3":
                    return "28";
                case "8":
                    return "25";
                    //ADLAB
                case "5":
                    return "23";
                    break;
                default:
                    // MessageBox.Show("OTHER ..進項類別:" + F0102);
                    MsgLine.Text += string.Format("單號:{1} 憑證類別未處理 " + "\r", DocNum, F0102);
                    return "21";
                    break;

            }
        }


        //銷項
        private string ConvertF0102_Out(string DocNum, string F0102)
        {
            switch (F0102)
            {
                //1 給系統 -1 使用

                case "0":
                    return "31";
                    break;
                case "1":
                    return "35";
                    break;
                case "2":
                    return "32";
                    break;
                case "3":
                    return "32";
                    break;
                case "4":
                    return "31";
                    break;
                case "5":
                    return "36";
                    break;
                case "8":
                    return "35";
                    break;

                default:
                    MsgLine.Text += string.Format("單號:{1} 憑證類別未處理 " + "\r", DocNum, F0102);

                    return "  ";
                    break;

            }
        }


        //F6161
        //課稅別
        private string ConviertF6161(string F6161)
        {
            switch (F6161)
            {

                case "0":
                    return "1";　//應稅
                    break;
                case "1":
                    return "2"; //零稅率
                    break;
                case "2":
                    return "3";//免稅
                    break;
                case "F":
                    return "F";//免稅
                    break;
                default:
                    return "D"; //空白作廢
                    break;

            }
        }

        //F7272
        //得扣抵
        private string ConviertF6161(string BSTY5, string BSTY4)
        {
            if (BSTY5 == "0" && (BSTY4 == "0" || BSTY4 == "1"))
            {
                return "1";
            }
            else if (BSTY5 == "0" && BSTY4 == "2")
            {
                return "2";
            }
            else if (BSTY5 == "1" && (BSTY4 == "0" || BSTY4 == "1"))
            {
                return "3";
            }
            else if (BSTY5 == "1" && BSTY4 == "2")
            {
                return "4";
            }
            else
            {
                return " ";
            }

        }

        /// <summary>
        /// 20090811
        /// </summary>
        /// <param name="Year"></param>
        /// <returns></returns>
        private string ConvertYear(string Year)
        {
            int iYear = Convert.ToInt32(Year) - 1911;
            //return iYear.ToString("00");
            return iYear.ToString("000");

        }

        //民國 100 年
        private string ConvertYear100(string Year)
        {
            int iYear = Convert.ToInt32(Year) - 1911;
            return iYear.ToString("000");

        }



        public static string PadStrRight(string BufString, int num, string RString)
        {
            string r = "";

            byte[] bytStr = System.Text.Encoding.Default.GetBytes(BufString);

            // int NeedNum = num - BufString.Length;
            int NeedNum = num - bytStr.Length;

            for (int i = 1; i <= NeedNum; i++)
            {
                r = r + RString;
            }


            return BufString + r;
        }
        private bool isNumber(string s)
        {
            int Flag = 0;
            char[] str = s.ToCharArray();
            for (int i = 0; i < str.Length; i++)
            {
                if (Char.IsNumber(str[i]))
                {
                    Flag++;
                }
                else
                {
                    Flag = -1;
                    break;
                }
            }
            if (Flag > 0)
            {
                return false ;
            }
            else
            {

              
                return true  ;
            }
        }

        private void fmAcmeTax_Load(object sender, EventArgs e)
        {
            //取前一個月...申報月
            textBox18.Text = DateToStr(DateTime.Today.AddMonths(-1)).Substring(0, 6);
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView8);
        }

        private void linkLabel1_Click(object sender, EventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\媒體申報.doc";
            System.Diagnostics.Process.Start(aa);
        }

        //動態產生資料結構
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("通關方式", typeof(string));
            dt.Columns.Add("外銷方式", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("資料", typeof(string));

            //最後一個總計
            //  dt.Columns.Add("Qty", typeof(int));


            //DataColumn[] colPk = new DataColumn[4];
            //colPk[0] = dt.Columns["通關方式"];
            //colPk[0] = dt.Columns["外銷方式"];
            //colPk[0] = dt.Columns["日期"];
            //colPk[0] = dt.Columns["資料"];
            //dt.PrimaryKey = colPk;

            return dt;
        }

        private System.Data.DataTable MakeTableJ1()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("DocName", typeof(string));
            dt.Columns.Add("DocKind", typeof(string));
            dt.Columns.Add("DocNum", typeof(string));
            dt.Columns.Add("U_PC_BSTY1", typeof(string));
            dt.Columns.Add("U_PC_BSDAT", typeof(DateTime));
            dt.Columns.Add("U_PC_BSINV", typeof(string));
            dt.Columns.Add("U_PC_BSAPP", typeof(DateTime));
            dt.Columns.Add("U_PC_BSTY2", typeof(string));
            dt.Columns.Add("U_PC_BSTY3", typeof(string));
            dt.Columns.Add("U_PC_BSTY4", typeof(string));
            dt.Columns.Add("U_PC_BSTY5", typeof(string));
            dt.Columns.Add("U_PC_BSTYC", typeof(string));
            dt.Columns.Add("U_PC_BSNOT", typeof(string));
            dt.Columns.Add("U_PC_BSAMN", typeof(decimal));
            dt.Columns.Add("U_PC_BSTAX", typeof(string));
            dt.Columns.Add("U_PC_BSAMT", typeof(string));
            dt.Columns.Add("U_PC_BSCUS", typeof(string));
            dt.Columns.Add("U_IN_BSCLS", typeof(string));
            dt.Columns.Add("U_ACME_SHIPWORKDAY", typeof(string));
            dt.Columns.Add("U_IN_BSDTO", typeof(string));
            dt.Columns.Add("U_IN_BSTY7", typeof(string));
            dt.Columns.Add("U_IN_BSREN", typeof(string));
            dt.Columns.Add("SEQ", typeof(string));
            dt.Columns.Add("U_PC_BSTYI", typeof(string));
            dt.Columns.Add("DocTotal", typeof(decimal));
            dt.Columns.Add("VatSum", typeof(decimal));
            dt.Columns.Add("發票月份", typeof(string));
            dt.Columns.Add("APP", typeof(string));
            return dt;
        }
        //DocName	DocKind	DocNum	U_PC_BSTY1	U_PC_BSDAT	U_PC_BSINV	U_PC_BSAPP	U_PC_BSTY2	U_PC_BSTY3	
        //U_PC_BSTY4	U_PC_BSTY5	U_PC_BSTYC	U_PC_BSNOT	U_PC_BSAMN	U_PC_BSTAX	U_PC_BSAMT	
        //U_PC_BSCUS	U_IN_BSCLS	U_ACME_SHIPWORKDAY	U_IN_BSDTO	U_IN_BSTY7	U_IN_BSREN	SEQ	
        //U_PC_BSTYI	DocTotal	VatSum	發票月份

        public System.Data.DataTable GetID()
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "select TAXIDNUM2 稅籍編號 from OADM";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

  
    }
}
