using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ACME
{
    public partial class fmSiCompare : Form
    {

        private static string ConnectiongString = "server=10.10.1.47;pwd=NewType;uid=NewType;database=EIPTest";
        private string ShipConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
        private string SapConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";


        public fmSiCompare()
        {
            InitializeComponent();
        }


        private void AddCompareRow(string DocType, string InvoiceNo, System.Data.DataTable dt, string Item, string Acme, string Auo,
          string LineNum, string AcmeSiNo, string Msg, string PoNo)
        {
            DataRow dr = dt.NewRow();
            dr["DocType"] = DocType;
            dr["InvoiceNo"] = InvoiceNo;
            dr["LineNum"] = LineNum;
            dr["Item"] = Item;
            dr["Acme"] = Acme;
            dr["Auo"] = Auo;

            dr["PO"] = PoNo;


            int X = String.Compare(
                Acme,
                Auo, true);

            if (X == 0)
            {
                dr["Result"] = "True";
            }
            else
            {
                //dr["Result"] = "False-" + X.ToString();
                dr["Result"] = "False";
            }

            //Always false
            if (Item == "品名")
            {
                dr["Result"] = "";
            }
            else if (Item == "PartNo")
            {
                dr["Result"] = "";
            }
            else if (Item == "取貨地" || Item == "裝船港" || Item == "目的地" || Item == "卸貨港")
            {

                if (Acme == "SONGJIANG, CHINA" & Auo == "SHANGHAI")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SHENZHEN, CHINA" & Auo == "深圳機場")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SUZHOU, CHINA" & Auo == "SUZHOU")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "TAOYUAN,TAIWAN" & Auo == "進金生/蘆竹八股")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SHENZHEN, CHINA" & Auo == "深圳/巨航(機場)")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SHENZHEN, CHINA" & Auo == "深圳/巨航(坪山)")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "XIAMEN, CHINA" & Auo == "廈門/宏高")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SUZHOU, CHINA" & Auo == "Suzhou/宏高(new)")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SUZHOU, CHINA" & Auo == "偉創蘇州倉")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "HONG KONG" & Auo == "HK/宏高")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "XIAMEN, CHINA" & Auo == "Xamen")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "TAOYUAN,TAIWAN" & Auo == "新得利倉儲")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "TAOYUAN,TAIWAN" & Auo == "聯倉桃園")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "HONG KONG" & Auo == "香港宇思物流")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SHENZHEN, CHINA" & Auo == "深圳機場.")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "TAIPEI, TAIWAN" & Auo == "TAIPEI")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "XIAMEN, CHINA" & Auo == "Xamen")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "TAOYUAN, TAIWAN" & Auo == "TAOYUAN PORT")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "HEFEI, CHINA" & Auo == "HEFEI")
                {
                    dr["Result"] = "True";
                }
                else if (Acme == "SHANGHAI, CHINA" & Auo == "SHANGHAI")
                {
                    dr["Result"] = "True";
                }





                //對照表檔
                //類別 ACME AUO
                //地點 "HONG KONG" "HK/宏高"
            }

            dr["AcmeSI"] = AcmeSiNo;
            dr["Message"] = Msg;
            dt.Rows.Add(dr);
        }


        public System.Data.DataTable GetData(string Sql)
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);


            SqlCommand command = new SqlCommand();
            command.Connection = connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(Sql);



            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();

            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_Stage");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_Stage"];
        }

        //20180528
        private System.Data.DataTable MakeTableCompare()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("DocType", typeof(string));
            dt.Columns.Add("InvoiceNo", typeof(string));
            dt.Columns.Add("AcmeSI", typeof(string));
            //20180608
            dt.Columns.Add("PO", typeof(string));

            dt.Columns.Add("LineNum", typeof(string));
            dt.Columns.Add("Item", typeof(string));
            dt.Columns.Add("Acme", typeof(string));

            dt.Columns.Add("Auo", typeof(string));

            dt.Columns.Add("Result", typeof(string));

            dt.Columns.Add("Message", typeof(string));




            DataColumn[] colPk = new DataColumn[4];
            colPk[0] = dt.Columns["InvoiceNo"];
            colPk[1] = dt.Columns["DocType"];
            colPk[2] = dt.Columns["Item"];
            colPk[3] = dt.Columns["LineNum"];
            dt.PrimaryKey = colPk;

            dt.TableName = "RPA_Compare";

            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }


        public System.Data.DataTable GetData_SAP(string Sql)
        {
            SqlConnection connection = new SqlConnection(SapConnectiongString);


            SqlCommand command = new SqlCommand();
            command.Connection = connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(Sql);



            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();

            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_Stage");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_Stage"];
        }

        private void DoCompare(System.Data.DataTable dt, string AuoInvoiceNo)
        {
            //CpNo SINO 可能是多筆 SH20180515012X/SH20180517005X
            //SH20180521016X ACME-T
            //1PCS SI 被拆單
            //可能併單


            string SqlInvH = "select * from RPA_InvoiceH where InvoiceNo='{0}'";

            SqlInvH = string.Format(SqlInvH, AuoInvoiceNo);
            System.Data.DataTable dtInvH = GetData(SqlInvH);
            //dgHeader.DataSource = dtInvH;


            if (dtInvH.Rows.Count == 0)
            {
                MessageBox.Show("查無此 Invoice#" + AuoInvoiceNo);
                return;
            }


            string SqlInvD = "select * from RPA_InvoiceD where InvoiceNo='{0}'";

            SqlInvD = string.Format(SqlInvD, AuoInvoiceNo);
            System.Data.DataTable dtInvD = GetData(SqlInvD);
            //dgDetail.DataSource = dtInvD;


            string InvoiceNo = AuoInvoiceNo;
            string DocType = "主檔";
            string LineNum = "";

            string IsLC = "";
            string LCNo = "";


            //主檔 
            string AuoSiNo = "";
            string Msg = "";
            if (Convert.IsDBNull(dtInvH.Rows[0]["CpNo"]))
            {
                AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
               "主檔沒有SI號碼",
               "0",
               Convert.ToString(dtInvH.Rows.Count),
               LineNum,
               AuoSiNo,
               Msg, "");
                return;
            }
            else
            {
                //確認長度
                AuoSiNo = Convert.ToString(dtInvH.Rows[0]["CpNo"]);
            }


            string Sql = "select tradeCondition 貿易條件," +
"add3 付款方式 ," +
"closeDay 結關日," +
"forecastDay 預計開航日," +
"arriveDay 預計抵達日," +
"receivePlace 取貨地," +
"goalPlace 目的地," +
"shipment 裝船港," +
"unloadCargo 卸貨港 " +
"from SHIPPING_MAIN " +
"where ShippingCode ='{0}' ";

            string Sql1 = "select Docentry,Dscription,Quantity,T1.ItemPrice," +
            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
            "where T0.shippingcode='{0}'";

            Sql = string.Format(Sql, AuoSiNo);
            System.Data.DataTable dtAcmeH = GetData(Sql);

          //  dataGridView1.DataSource = dtAcmeH;


            Sql1 = string.Format(Sql1, AuoSiNo);
            System.Data.DataTable dtAcmeD = GetData(Sql1);
          //  dataGridView2.DataSource = dtAcmeD;


            string AuoSiNo_Multi = "";

            if (dtAcmeH.Rows.Count == 0)
            {
                //檢查長度
                if (AuoSiNo.Length != 14)
                {
                    Msg = "SI號碼:" + AuoSiNo;
                }
                //機率小 -> 人工判斷
                //LSZ1355824		主檔筆數為零	0	1	False	SI號碼:SH20180515012X/SH20180517005X
                if (AuoSiNo.Length > 14 & AuoSiNo.IndexOf("/") > 0)
                {
                    string[] s = AuoSiNo.Split('/');
                    //先拆一個 或 合併成 in ('s1','s2')
                    for (int k = 0; k <= s.Length - 1; k++)
                    {
                        AuoSiNo_Multi += "'" + s[k] + "',";
                    }

                    AuoSiNo_Multi = AuoSiNo_Multi.Substring(0, AuoSiNo_Multi.Length - 1);

                    //先假設取第二筆
                    AuoSiNo = s[1];
                }

                AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
               "主檔筆數為零",
               "0",
               Convert.ToString(dtInvH.Rows.Count),
               LineNum,
               AuoSiNo,
               Msg, "");

                Msg = "";

                if (string.IsNullOrEmpty(AuoSiNo_Multi))
                {
                    return;
                }
                else
                {

                    Sql = "select tradeCondition 貿易條件," +
"add3 付款方式 ," +
"closeDay 結關日," +
"forecastDay 預計開航日," +
"arriveDay 預計抵達日," +
"receivePlace 取貨地," +
"goalPlace 目的地," +
"shipment 裝船港," +
"unloadCargo 卸貨港 " +
"from SHIPPING_MAIN " +
"where ShippingCode  in ({0}) ";

                    Sql = string.Format(Sql, AuoSiNo_Multi);
                    dtAcmeH = GetData(Sql);
                    if (dtAcmeH.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        AddCompareRow(
                 DocType,
                 InvoiceNo,
                 dt,
                 "主檔筆數多筆",
                 Convert.ToString(dtAcmeH.Rows.Count),
                 "",
                 LineNum,
                 AuoSiNo_Multi,
                 Msg, "");
                    }

                }
            }




            //  System.Data.DataTable dt = MakeTableCompare();

            AddCompareRow(
                DocType,
                InvoiceNo,
                dt,
                "付款方式",
                Convert.ToString(dtAcmeH.Rows[0]["付款方式"]),
                Convert.ToString(dtInvH.Rows[0]["Payment"]),
                LineNum,
               AuoSiNo,
               Msg, "");

            if (Convert.ToString(dtInvH.Rows[0]["Payment"]).IndexOf("L/C") >= 0)
            {
                IsLC = "Y";

                try
                {
                    LCNo = Convert.ToString(dtInvH.Rows[0]["LC"]);
                }
                catch
                {
                }

            }

            AddCompareRow(
               DocType,
                InvoiceNo,
                dt,
                "貿易條件",
                Convert.ToString(dtAcmeH.Rows[0]["貿易條件"]),
                Convert.ToString(dtInvH.Rows[0]["TradeTerm"]).TrimStart(' '),
                LineNum,
               AuoSiNo,
               Msg, "");

            //結關日
            AddCompareRow(
              DocType,
                InvoiceNo,
                dt,
              "結關日",
              Convert.ToString(dtAcmeH.Rows[0]["結關日"]),
              Convert.ToString(dtInvH.Rows[0]["InvoiceDate"]).Replace("/", ""),
              LineNum,
               AuoSiNo,
               Msg, "");

            //預計開航日
            AddCompareRow(
             DocType,
                InvoiceNo,
                dt,
             "預計開航日",
             Convert.ToString(dtAcmeH.Rows[0]["預計開航日"]),
             Convert.ToString(dtInvH.Rows[0]["ETD"]).Replace("/", ""),
             LineNum,
               AuoSiNo,
               Msg, "");

            //預計抵達日
            AddCompareRow(
            DocType,
                InvoiceNo,
                dt,
            "預計抵達日",
            Convert.ToString(dtAcmeH.Rows[0]["預計抵達日"]),
            Convert.ToString(dtInvH.Rows[0]["ETA"]).Replace("/", ""),
            LineNum,
               AuoSiNo,
               Msg, "");

            //取貨地
            AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "取貨地",
           Convert.ToString(dtAcmeH.Rows[0]["取貨地"]),
           Convert.ToString(dtInvH.Rows[0]["ShipFrom"]),
           LineNum,
               AuoSiNo,
               Msg, "");

            //目的地
            AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "目的地",
           Convert.ToString(dtAcmeH.Rows[0]["目的地"]),
           Convert.ToString(dtInvH.Rows[0]["ShipTo"]).TrimStart(' '),
           LineNum,
               AuoSiNo,
               Msg, "");

            //裝船港
            AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "裝船港",
           Convert.ToString(dtAcmeH.Rows[0]["裝船港"]),
           Convert.ToString(dtInvH.Rows[0]["ShipFrom"]).TrimStart(' '),
           LineNum,
               AuoSiNo,
               Msg, "");


            //卸貨港
            AddCompareRow(
             DocType,
             InvoiceNo,
             dt,
         "卸貨港",
         Convert.ToString(dtAcmeH.Rows[0]["卸貨港"]),
         Convert.ToString(dtInvH.Rows[0]["ShipTo"]).TrimStart(' '),
         LineNum,
               AuoSiNo,
               Msg, "");

            //表頭額外檢查規則
            //where TradeTerm like 'CIF TAIWAN' -> 5% TaxRate
            string TradeTerm = "";
            string TaxRate = "";
            try
            {
                TradeTerm = Convert.ToString(dtInvH.Rows[0]["TradeTerm"]).TrimStart(' ');
            }
            catch
            {

            }
            if (!string.IsNullOrEmpty(TradeTerm))
            {
                if (TradeTerm.IndexOf("CIF TAIWAN") >= 0)
                {
                    try
                    {
                        TaxRate = Convert.ToString(dtInvH.Rows[0]["TaxRate"]).TrimStart(' ').Trim();
                    }
                    catch
                    {
                    }
                    AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "CIF TAIWAN Tax",
          "5%",
           TaxRate,
           LineNum,
               AuoSiNo,
               Msg, "");


                }

            }


            //比對多筆明細檔-------------------------------------------------------------------------

            DataRow dr;
            DocType = "明細";
            string SiNo;
            string PoNo;
            System.Data.DataTable dtSiD = null;

            string Price = "";
            string Qty = "";
            string AcmeLC = "";
            string Item = "";
            string ItemName = "";
            string AcmeItemNo = "";
            string AcmeModel = "";



            for (int i = 0; i <= dtInvD.Rows.Count - 1; i++)
            {
                dr = dtInvD.Rows[i];

                Msg = "";


                Msg = "";

                try
                {
                    PoNo = Convert.ToString(dtInvD.Rows[i]["PoNo"]);
                }
                catch
                {
                    PoNo = "";
                }


                try
                {
                    SiNo = Convert.ToString(dtInvD.Rows[i]["SiNo"]);

                    //只取 14 碼 
                    if (SiNo.Length > 14)
                    {
                        //人工比對
                        Msg += "SI號碼 > 14:" + SiNo + "\r\n";
                        SiNo = SiNo.Substring(0, 14);

                    }
                }
                catch
                {
                    SiNo = "";
                }

                //假設法
                //因為等級產地未知
                //可能多筆
                AcmeItemNo = GetAcmeItemCode(Convert.ToString(dtInvD.Rows[i]["ItemNO"]));

                if (!string.IsNullOrEmpty(AcmeItemNo))
                {
                    AcmeModel = AcmeItemNo.Substring(0, 9);
                }

                Msg += "Acme Model:" + AcmeModel;
                AddCompareRow(
             DocType,
             InvoiceNo,
             dt,
         "PartNo",
         "",
         Convert.ToString(dtInvD.Rows[i]["ItemNO"]),
         LineNum,
             SiNo,
             Msg, PoNo);

                if (!string.IsNullOrEmpty(SiNo) & !string.IsNullOrEmpty(PoNo))
                {
                    //dtSiD = GetAcmeSiD(SiNo, PoNo);
                    // dtSiD = GetAcmeSiD_ItemCode(SiNo, PoNo,AcmeItemNo);

                    //使用 Model 法
                    dtSiD = GetAcmeSiD_Model(SiNo, PoNo, AcmeModel);
                }
                else
                {
                    dtSiD = null;
                }


                Price = "";
                AcmeLC = "";
                Qty = "";

                Item = "";
                ItemName = "";
                if (dtSiD != null)
                {
                    if (dtSiD.Rows.Count == 1)
                    {
                        Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
                        Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
                        ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
                        AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

                    }
                    else if (dtSiD.Rows.Count == 0)
                    {

                        Msg = "Model法-SI筆數 = " + dtSiD.Rows.Count.ToString()
                            + "|" + Convert.ToString(dtInvD.Rows[i]["ItemNO"])
                            + "|" + AcmeModel;


                        AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "SI筆數異常#1",
           dtSiD.Rows.Count.ToString(),
           "",
           LineNum,
               SiNo,
               Msg, PoNo);
                        Msg = "";
                    }
                    else if (dtSiD.Rows.Count > 1)
                    {

                        //             Msg = "Model法-SI筆數 = " + dtSiD.Rows.Count.ToString();
                        //             AddCompareRow(
                        //    DocType,
                        //    InvoiceNo,
                        //    dt,
                        //"SI筆數異常#1",
                        //dtSiD.Rows.Count.ToString(),
                        //"",
                        //LineNum,
                        //    SiNo,
                        //    Msg, PoNo);
                        //             Msg = "";

                        dtSiD = GetAcmeSiD_Qty(SiNo, PoNo, AcmeModel, Convert.ToString(dtInvD.Rows[i]["Qty"]));
                        //PO +SiNo 一樣..可能要繼續判斷 -> 使用 數量法
                        //品名法
                        //單價法
                        //數量法

                        if (dtSiD.Rows.Count == 1)
                        {
                            //數量可能被拆成兩筆
                            //LSZ1357851
                            //100 -> 90 + 10

                            Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
                            Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
                            ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
                            AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

                        }
                        else
                        {

                            string PartNo = Convert.ToString(dtInvD.Rows[i]["ItemNO"]).Substring(9, 3);
                            dtSiD = GetAcmeSiD_Qty_VerNo(SiNo, PoNo, AcmeModel, Convert.ToString(dtInvD.Rows[i]["Qty"]), PartNo);

                            //多筆數量相同 取第一筆
                            if (dtSiD.Rows.Count >= 1)
                            {

                                Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
                                Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
                                ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
                                AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

                            }
                            else
                            {


                                System.Data.DataTable dtCheck = GetAcmeSiD_VerNo(SiNo, PoNo, AcmeModel, PartNo);

                                if (dtCheck.Rows.Count > 0)
                                {
                                    //多筆相同 取第一筆
                                    Price = Convert.ToString(dtCheck.Rows[0]["ItemPrice"]);
                                    Qty = Convert.ToString(dtCheck.Rows[0]["Quantity"]);
                                    ItemName = Convert.ToString(dtCheck.Rows[0]["Dscription"]);
                                    AcmeLC = Convert.ToString(dtCheck.Rows[0]["LC"]);

                                    Msg = "Model+Ver-SI筆數異常()" + dtCheck.Rows.Count.ToString()
                                   + "|" + Convert.ToString(dtInvD.Rows[i]["ItemNO"])
                                   + "|" + AcmeModel
                                   + "|" + PartNo;

                                    AddCompareRow(
                   DocType,
                   InvoiceNo,
                   dt,
               "SI筆數異常#3",
               dtCheck.Rows.Count.ToString(),
               "",
               LineNum,
                   SiNo,
                   Msg, PoNo);
                                    Msg = "";





                                }
                                else
                                {

                                    Msg = "Model+Ver+數量法-SI筆數異常" + dtSiD.Rows.Count.ToString()
                                   + "|" + Convert.ToString(dtInvD.Rows[i]["ItemNO"])
                                   + "|" + AcmeModel
                                   + "|" + PartNo;

                                    AddCompareRow(
                       DocType,
                       InvoiceNo,
                       dt,
                   "SI筆數異常#2",
                   dtSiD.Rows.Count.ToString(),
                   "",
                   LineNum,
                       SiNo,
                       Msg, PoNo);
                                    Msg = "";
                                }





                            }






                        }
                    }
                }


                LineNum = (i + 1).ToString();
                //SH20180518017X
                //LSJ1357321
                //PO
                //比法
                AddCompareRow(
                 DocType,
                 InvoiceNo,
                 dt,
             "PO",
             PoNo,
             Convert.ToString(dtInvD.Rows[i]["PoNo"]).TrimStart(' '),
             LineNum,
               SiNo,
               Msg, PoNo);

                //    //PartNo 
                //    AddCompareRow(
                //    DocType,
                //    InvoiceNo,
                //    dt,
                //"Model",
                //AcmeItemNo,
                //Convert.ToString(dtInvD.Rows[i]["ItemNo"]).TrimStart(' '),
                //LineNum,
                //SiNo,
                //   Msg);



                //ItemNo
                AddCompareRow(
                DocType,
                InvoiceNo,
                dt,
            "品名",
            ItemName,
            Convert.ToString(dtInvD.Rows[i]["ItemName"]).TrimStart(' '),
            LineNum,
               SiNo,
               Msg, PoNo);


                //Price
                AddCompareRow(
                DocType,
                InvoiceNo,
                dt,
            "單價",
            Price,
            Convert.ToString(dtInvD.Rows[i]["Price"]).TrimStart(' '),
            LineNum,
               SiNo,
               Msg, PoNo);

                //Qty
                AddCompareRow(
                DocType,
                InvoiceNo,
                dt,
            "數量",
            Qty,
            Convert.ToString(dtInvD.Rows[i]["Qty"]).TrimStart(' '),
            LineNum,
               SiNo,
               Msg, PoNo);

                //    //Amount
                //    AddCompareRow(
                //    DocType,
                //    InvoiceNo,
                //    dt,
                //"金額",
                //"",
                //Convert.ToString(dtInvD.Rows[i]["Amount"]).TrimStart(' '),
                //LineNum);

                //20180607
                //LC //先用主檔法//明細檔未知 L/C NO 放在何處
                if (IsLC == "Y")
                {
                    AddCompareRow(
                    DocType,
                    InvoiceNo,
                    dt,
                "L/C",
                AcmeLC,
                LCNo,
                LineNum,
                   SiNo,
                   Msg, PoNo);
                }


                //
                System.Data.DataTable dtGetCheckSum = GetCheckSum(InvoiceNo, SiNo, PoNo, Convert.ToString(dtInvD.Rows[i]["ItemNO"]));
                System.Data.DataTable dtGetCheckSumAcme = GetCheckSumAcme(SiNo, PoNo, Convert.ToString(dtInvD.Rows[i]["ItemNO"]));

                Int32 iGetCheckSumAcme = 0;
                Int32 iGetCheckSumAcmeQty = 0;
                Int32 iGetCheckSum = 0;
                Int32 iGetCheckSumQty = 0;


                try
                {
                    iGetCheckSumAcme = Convert.ToInt32(dtGetCheckSumAcme.Rows[0]["明細筆數"]);
                    iGetCheckSumAcmeQty = Convert.ToInt32(dtGetCheckSumAcme.Rows[0]["Qty"]);

                }
                catch
                {

                }
                try
                {
                    iGetCheckSum = Convert.ToInt32(dtGetCheckSum.Rows[0]["明細筆數"]);
                    iGetCheckSumQty = Convert.ToInt32(dtGetCheckSum.Rows[0]["Qty"]);
                }
                catch
                {
                }



                AddCompareRow(
                   DocType,
                   InvoiceNo,
                   dt,
               "彙總筆數",
               Convert.ToString(iGetCheckSumAcme),
               Convert.ToString(iGetCheckSum),
               LineNum,
                  SiNo,
                  Msg, PoNo);

                AddCompareRow(
                 DocType,
                 InvoiceNo,
                 dt,
             "彙總數量",
             Convert.ToString(iGetCheckSumAcmeQty),
               Convert.ToString(iGetCheckSumQty),
             LineNum,
                SiNo,
                Msg, PoNo);


            }

        }

        private string GetAcmeItemCode(string ItemCode)
        {
            string sql = "select  ItemCode from oitm where U_partno='{0}'";
            sql = string.Format(sql, ItemCode);
            System.Data.DataTable dt = GetData_SAP(sql);

            if (dt.Rows.Count > 0)
            {
                return Convert.ToString(dt.Rows[0]["ItemCode"]);
            }
            else
            {
                return string.Empty;
            }

        }


        private void button10_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("確定執行嗎？ ", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            //表頭 LC
            //明細 LC 多對多

            //LSZ1357851
            string SqlInvH = "select * from RPA_InvoiceH ";

            //異常記錄

            //string SqlInvH = "select * from RPA_InvoiceH where InvoiceNo='LSZ1358075' ";//
            //
            //Z681362055
            //LSZ1355824
            //M111362201
            //LSJ1360467 ->多筆 SI
            //LSZ1358079
            //LSZ1355824 -> 多筆 SI SI號碼:SH20180515012X/SH20180517005X
            //LSZ1358077
            //LSJ1357321
            //LSZ1358067
            //Z191356452
            //LSJ1359205 -> SI Line 被拆成兩筆

            //select * from RPA_InvoiceD where sino is null

            //string SqlInvH = "select * from RPA_InvoiceH where InvoiceNo='Z191356452' ";

            //string SqlInvH = "select * from RPA_InvoiceH where InvoiceNo='LSZ1357851' ";
            System.Data.DataTable dtInvH = GetData(SqlInvH);

            System.Data.DataTable dt = MakeTableCompare();
            string AuoInvoiceNo = "";

            for (int i = 0; i <= dtInvH.Rows.Count - 1; i++)
            {
                AuoInvoiceNo = Convert.ToString(dtInvH.Rows[i]["InvoiceNo"]);
                DoCompare(dt, AuoInvoiceNo);

            }
            dgCompare.DataSource = dt;
            dgCompare.DataSource = dt;
           
            GridViewAutoSize(dgCompare);

            //過濾

            DataView dv = dt.DefaultView;
            System.Data.DataTable dtFalse = dv.ToTable();

            dv = dtFalse.DefaultView;
            dv.RowFilter = "result ='False'";


            dgFalse.DataSource = dtFalse;

            GridViewAutoSize(dgFalse);
        }

//        private void DoCompare(System.Data.DataTable dt, string AuoInvoiceNo)
//        {
//            //CpNo SINO 可能是多筆 SH20180515012X/SH20180517005X
//            //SH20180521016X ACME-T
//            //1PCS SI 被拆單
//            //可能併單


//            string SqlInvH = "select * from RPA_InvoiceH where InvoiceNo='{0}'";

//            SqlInvH = string.Format(SqlInvH, AuoInvoiceNo);
//            System.Data.DataTable dtInvH = GetData(SqlInvH);
//            //dgHeader.DataSource = dtInvH;

//            string SqlInvD = "select * from RPA_InvoiceD where InvoiceNo='{0}'";

//            SqlInvD = string.Format(SqlInvD, AuoInvoiceNo);
//            System.Data.DataTable dtInvD = GetData(SqlInvD);
//           // dgDetail.DataSource = dtInvD;


//            string InvoiceNo = AuoInvoiceNo;
//            string DocType = "主檔";
//            string LineNum = "";

//            string IsLC = "";
//            string LCNo = "";


//            //主檔 
//            string AuoSiNo = "";
//            string Msg = "";
//            if (Convert.IsDBNull(dtInvH.Rows[0]["CpNo"]))
//            {
//                AddCompareRow(
//               DocType,
//               InvoiceNo,
//               dt,
//               "主檔沒有SI號碼",
//               "0",
//               Convert.ToString(dtInvH.Rows.Count),
//               LineNum,
//               AuoSiNo,
//               Msg, "");
//                return;
//            }
//            else
//            {
//                //確認長度
//                AuoSiNo = Convert.ToString(dtInvH.Rows[0]["CpNo"]);
//            }


//            string Sql = "select tradeCondition 貿易條件," +
//"add3 付款方式 ," +
//"closeDay 結關日," +
//"forecastDay 預計開航日," +
//"arriveDay 預計抵達日," +
//"receivePlace 取貨地," +
//"goalPlace 目的地," +
//"shipment 裝船港," +
//"unloadCargo 卸貨港 " +
//"from SHIPPING_MAIN " +
//"where ShippingCode ='{0}' ";

//            string Sql1 = "select Docentry,Dscription,Quantity,T1.ItemPrice," +
//            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
//            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
//            "where T0.shippingcode='{0}'";

//            Sql = string.Format(Sql, AuoSiNo);
//            System.Data.DataTable dtAcmeH = GetData(Sql);

//           // dataGridView1.DataSource = dtAcmeH;


//            Sql1 = string.Format(Sql1, AuoSiNo);
//            System.Data.DataTable dtAcmeD = GetData(Sql1);
//          //  dataGridView2.DataSource = dtAcmeD;


//            string AuoSiNo_Multi = "";

//            if (dtAcmeH.Rows.Count == 0)
//            {
//                //檢查長度
//                if (AuoSiNo.Length != 14)
//                {
//                    Msg = "SI號碼:" + AuoSiNo;
//                }
//                //機率小 -> 人工判斷
//                //LSZ1355824		主檔筆數為零	0	1	False	SI號碼:SH20180515012X/SH20180517005X
//                if (AuoSiNo.Length > 14 & AuoSiNo.IndexOf("/") > 0)
//                {
//                    string[] s = AuoSiNo.Split('/');
//                    //先拆一個 或 合併成 in ('s1','s2')
//                    for (int k = 0; k <= s.Length - 1; k++)
//                    {
//                        AuoSiNo_Multi += "'" + s[k] + "',";
//                    }

//                    AuoSiNo_Multi = AuoSiNo_Multi.Substring(0, AuoSiNo_Multi.Length - 1);

//                    //先假設取第二筆
//                    AuoSiNo = s[1];
//                }

//                AddCompareRow(
//               DocType,
//               InvoiceNo,
//               dt,
//               "主檔筆數為零",
//               "0",
//               Convert.ToString(dtInvH.Rows.Count),
//               LineNum,
//               AuoSiNo,
//               Msg, "");

//                Msg = "";

//                if (string.IsNullOrEmpty(AuoSiNo_Multi))
//                {
//                    return;
//                }
//                else
//                {

//                    Sql = "select tradeCondition 貿易條件," +
//"add3 付款方式 ," +
//"closeDay 結關日," +
//"forecastDay 預計開航日," +
//"arriveDay 預計抵達日," +
//"receivePlace 取貨地," +
//"goalPlace 目的地," +
//"shipment 裝船港," +
//"unloadCargo 卸貨港 " +
//"from SHIPPING_MAIN " +
//"where ShippingCode  in ({0}) ";

//                    Sql = string.Format(Sql, AuoSiNo_Multi);
//                    dtAcmeH = GetData(Sql);
//                    if (dtAcmeH.Rows.Count == 0)
//                    {
//                        return;
//                    }
//                    else
//                    {
//                        AddCompareRow(
//                 DocType,
//                 InvoiceNo,
//                 dt,
//                 "主檔筆數多筆",
//                 Convert.ToString(dtAcmeH.Rows.Count),
//                 "",
//                 LineNum,
//                 AuoSiNo_Multi,
//                 Msg, "");
//                    }

//                }
//            }




//            //  System.Data.DataTable dt = MakeTableCompare();

//            AddCompareRow(
//                DocType,
//                InvoiceNo,
//                dt,
//                "付款方式",
//                Convert.ToString(dtAcmeH.Rows[0]["付款方式"]),
//                Convert.ToString(dtInvH.Rows[0]["Payment"]),
//                LineNum,
//               AuoSiNo,
//               Msg, "");

//            if (Convert.ToString(dtInvH.Rows[0]["Payment"]).IndexOf("L/C") >= 0)
//            {
//                IsLC = "Y";

//                try
//                {
//                    LCNo = Convert.ToString(dtInvH.Rows[0]["LC"]);
//                }
//                catch
//                {
//                }

//            }

//            AddCompareRow(
//               DocType,
//                InvoiceNo,
//                dt,
//                "貿易條件",
//                Convert.ToString(dtAcmeH.Rows[0]["貿易條件"]),
//                Convert.ToString(dtInvH.Rows[0]["TradeTerm"]).TrimStart(' '),
//                LineNum,
//               AuoSiNo,
//               Msg, "");

//            //結關日
//            AddCompareRow(
//              DocType,
//                InvoiceNo,
//                dt,
//              "結關日",
//              Convert.ToString(dtAcmeH.Rows[0]["結關日"]),
//              Convert.ToString(dtInvH.Rows[0]["InvoiceDate"]).Replace("/", ""),
//              LineNum,
//               AuoSiNo,
//               Msg, "");

//            //預計開航日
//            AddCompareRow(
//             DocType,
//                InvoiceNo,
//                dt,
//             "預計開航日",
//             Convert.ToString(dtAcmeH.Rows[0]["預計開航日"]),
//             Convert.ToString(dtInvH.Rows[0]["ETD"]).Replace("/", ""),
//             LineNum,
//               AuoSiNo,
//               Msg, "");

//            //預計抵達日
//            AddCompareRow(
//            DocType,
//                InvoiceNo,
//                dt,
//            "預計抵達日",
//            Convert.ToString(dtAcmeH.Rows[0]["預計抵達日"]),
//            Convert.ToString(dtInvH.Rows[0]["ETA"]).Replace("/", ""),
//            LineNum,
//               AuoSiNo,
//               Msg, "");

//            //取貨地
//            AddCompareRow(
//               DocType,
//               InvoiceNo,
//               dt,
//           "取貨地",
//           Convert.ToString(dtAcmeH.Rows[0]["取貨地"]),
//           Convert.ToString(dtInvH.Rows[0]["ShipFrom"]),
//           LineNum,
//               AuoSiNo,
//               Msg, "");

//            //目的地
//            AddCompareRow(
//               DocType,
//               InvoiceNo,
//               dt,
//           "目的地",
//           Convert.ToString(dtAcmeH.Rows[0]["目的地"]),
//           Convert.ToString(dtInvH.Rows[0]["ShipTo"]).TrimStart(' '),
//           LineNum,
//               AuoSiNo,
//               Msg, "");

//            //裝船港
//            AddCompareRow(
//               DocType,
//               InvoiceNo,
//               dt,
//           "裝船港",
//           Convert.ToString(dtAcmeH.Rows[0]["裝船港"]),
//           Convert.ToString(dtInvH.Rows[0]["ShipFrom"]).TrimStart(' '),
//           LineNum,
//               AuoSiNo,
//               Msg, "");


//            //卸貨港
//            AddCompareRow(
//             DocType,
//             InvoiceNo,
//             dt,
//         "卸貨港",
//         Convert.ToString(dtAcmeH.Rows[0]["卸貨港"]),
//         Convert.ToString(dtInvH.Rows[0]["ShipTo"]).TrimStart(' '),
//         LineNum,
//               AuoSiNo,
//               Msg, "");

//            //表頭額外檢查規則
//            //where TradeTerm like 'CIF TAIWAN' -> 5% TaxRate
//            string TradeTerm = "";
//            string TaxRate = "";
//            try
//            {
//                TradeTerm = Convert.ToString(dtInvH.Rows[0]["TradeTerm"]).TrimStart(' ');
//            }
//            catch
//            {

//            }
//            if (!string.IsNullOrEmpty(TradeTerm))
//            {
//                if (TradeTerm.IndexOf("CIF TAIWAN") >= 0)
//                {
//                    try
//                    {
//                        TaxRate = Convert.ToString(dtInvH.Rows[0]["TaxRate"]).TrimStart(' ').Trim();
//                    }
//                    catch
//                    {
//                    }
//                    AddCompareRow(
//               DocType,
//               InvoiceNo,
//               dt,
//           "CIF TAIWAN Tax",
//          "5%",
//           TaxRate,
//           LineNum,
//               AuoSiNo,
//               Msg, "");


//                }

//            }


//            //比對多筆明細檔-------------------------------------------------------------------------

//            DataRow dr;
//            DocType = "明細";
//            string SiNo;
//            string PoNo;
//            System.Data.DataTable dtSiD = null;

//            string Price = "";
//            string Qty = "";
//            string AcmeLC = "";
//            string Item = "";
//            string ItemName = "";
//            string AcmeItemNo = "";
//            string AcmeModal = "";



//            for (int i = 0; i <= dtInvD.Rows.Count - 1; i++)
//            {
//                dr = dtInvD.Rows[i];

//                Msg = "";


//                Msg = "";

//                try
//                {
//                    PoNo = Convert.ToString(dtInvD.Rows[i]["PoNo"]);
//                }
//                catch
//                {
//                    PoNo = "";
//                }


//                try
//                {
//                    SiNo = Convert.ToString(dtInvD.Rows[i]["SiNo"]);

//                    //只取 14 碼 
//                    if (SiNo.Length > 14)
//                    {
//                        //人工比對
//                        Msg += "SI號碼 > 14:" + SiNo + "\r\n";
//                        SiNo = SiNo.Substring(0, 14);

//                    }
//                }
//                catch
//                {
//                    SiNo = "";
//                }

//                //假設法
//                //因為等級產地未知
//                //可能多筆
//                AcmeItemNo = GetAcmeItemCode(Convert.ToString(dtInvD.Rows[i]["ItemNO"]));

//                if (!string.IsNullOrEmpty(AcmeItemNo))
//                {
//                    AcmeModal = AcmeItemNo.Substring(0, 9);
//                }

//                Msg += "Acme Modal:" + AcmeModal;
//                AddCompareRow(
//             DocType,
//             InvoiceNo,
//             dt,
//         "PartNo",
//         "",
//         Convert.ToString(dtInvD.Rows[i]["ItemNO"]),
//         LineNum,
//             SiNo,
//             Msg, PoNo);

//                if (!string.IsNullOrEmpty(SiNo) & !string.IsNullOrEmpty(PoNo))
//                {
//                    //dtSiD = GetAcmeSiD(SiNo, PoNo);
//                    // dtSiD = GetAcmeSiD_ItemCode(SiNo, PoNo,AcmeItemNo);

//                    //使用 Modal 法
//                    dtSiD = GetAcmeSiD_Modal(SiNo, PoNo, AcmeModal);
//                }
//                else
//                {
//                    dtSiD = null;
//                }


//                Price = "";
//                AcmeLC = "";
//                Qty = "";

//                Item = "";
//                ItemName = "";
//                if (dtSiD != null)
//                {
//                    if (dtSiD.Rows.Count == 1)
//                    {
//                        Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
//                        Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
//                        ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
//                        AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

//                    }
//                    else if (dtSiD.Rows.Count == 0)
//                    {

//                        Msg = "Model法-SI筆數 = " + dtSiD.Rows.Count.ToString()
//                            + "|" + Convert.ToString(dtInvD.Rows[i]["ItemNO"])
//                            + "|" + AcmeModal;


//                        AddCompareRow(
//               DocType,
//               InvoiceNo,
//               dt,
//           "SI筆數異常#1",
//           dtSiD.Rows.Count.ToString(),
//           "",
//           LineNum,
//               SiNo,
//               Msg, PoNo);
//                        Msg = "";
//                    }
//                    else if (dtSiD.Rows.Count > 1)
//                    {

//                        //             Msg = "Model法-SI筆數 = " + dtSiD.Rows.Count.ToString();
//                        //             AddCompareRow(
//                        //    DocType,
//                        //    InvoiceNo,
//                        //    dt,
//                        //"SI筆數異常#1",
//                        //dtSiD.Rows.Count.ToString(),
//                        //"",
//                        //LineNum,
//                        //    SiNo,
//                        //    Msg, PoNo);
//                        //             Msg = "";

//                        dtSiD = GetAcmeSiD_Qty(SiNo, PoNo, AcmeModal, Convert.ToString(dtInvD.Rows[i]["Qty"]));
//                        //PO +SiNo 一樣..可能要繼續判斷 -> 使用 數量法
//                        //品名法
//                        //單價法
//                        //數量法

//                        if (dtSiD.Rows.Count == 1)
//                        {
//                            //數量可能被拆成兩筆
//                            //LSZ1357851
//                            //100 -> 90 + 10

//                            Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
//                            Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
//                            ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
//                            AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

//                        }
//                        else
//                        {

//                            string PartNo = Convert.ToString(dtInvD.Rows[i]["ItemNO"]).Substring(9, 3);
//                            dtSiD = GetAcmeSiD_Qty_VerNo(SiNo, PoNo, AcmeModal, Convert.ToString(dtInvD.Rows[i]["Qty"]), PartNo);

//                            //多筆數量相同 取第一筆
//                            if (dtSiD.Rows.Count >= 1)
//                            {

//                                Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
//                                Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
//                                ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
//                                AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

//                            }
//                            else
//                            {


//                                System.Data.DataTable dtCheck = GetAcmeSiD_VerNo(SiNo, PoNo, AcmeModal, PartNo);

//                                if (dtCheck.Rows.Count > 0)
//                                {
//                                    //多筆相同 取第一筆
//                                    Price = Convert.ToString(dtCheck.Rows[0]["ItemPrice"]);
//                                    Qty = Convert.ToString(dtCheck.Rows[0]["Quantity"]);
//                                    ItemName = Convert.ToString(dtCheck.Rows[0]["Dscription"]);
//                                    AcmeLC = Convert.ToString(dtCheck.Rows[0]["LC"]);

//                                    Msg = "Model+Ver-SI筆數異常()" + dtCheck.Rows.Count.ToString()
//                                   + "|" + Convert.ToString(dtInvD.Rows[i]["ItemNO"])
//                                   + "|" + AcmeModal
//                                   + "|" + PartNo;

//                                    AddCompareRow(
//                   DocType,
//                   InvoiceNo,
//                   dt,
//               "SI筆數異常#3",
//               dtCheck.Rows.Count.ToString(),
//               "",
//               LineNum,
//                   SiNo,
//                   Msg, PoNo);
//                                    Msg = "";





//                                }
//                                else
//                                {

//                                    Msg = "Model+Ver+數量法-SI筆數異常" + dtSiD.Rows.Count.ToString()
//                                   + "|" + Convert.ToString(dtInvD.Rows[i]["ItemNO"])
//                                   + "|" + AcmeModal
//                                   + "|" + PartNo;

//                                    AddCompareRow(
//                       DocType,
//                       InvoiceNo,
//                       dt,
//                   "SI筆數異常#2",
//                   dtSiD.Rows.Count.ToString(),
//                   "",
//                   LineNum,
//                       SiNo,
//                       Msg, PoNo);
//                                    Msg = "";
//                                }





//                            }






//                        }
//                    }
//                }


//                LineNum = (i + 1).ToString();
//                //SH20180518017X
//                //LSJ1357321
//                //PO
//                //比法
//                AddCompareRow(
//                 DocType,
//                 InvoiceNo,
//                 dt,
//             "PO",
//             PoNo,
//             Convert.ToString(dtInvD.Rows[i]["PoNo"]).TrimStart(' '),
//             LineNum,
//               SiNo,
//               Msg, PoNo);

//                //    //PartNo 
//                //    AddCompareRow(
//                //    DocType,
//                //    InvoiceNo,
//                //    dt,
//                //"Model",
//                //AcmeItemNo,
//                //Convert.ToString(dtInvD.Rows[i]["ItemNo"]).TrimStart(' '),
//                //LineNum,
//                //SiNo,
//                //   Msg);



//                //ItemNo
//                AddCompareRow(
//                DocType,
//                InvoiceNo,
//                dt,
//            "品名",
//            ItemName,
//            Convert.ToString(dtInvD.Rows[i]["ItemName"]).TrimStart(' '),
//            LineNum,
//               SiNo,
//               Msg, PoNo);


//                //Price
//                AddCompareRow(
//                DocType,
//                InvoiceNo,
//                dt,
//            "單價",
//            Price,
//            Convert.ToString(dtInvD.Rows[i]["Price"]).TrimStart(' '),
//            LineNum,
//               SiNo,
//               Msg, PoNo);

//                //Qty
//                AddCompareRow(
//                DocType,
//                InvoiceNo,
//                dt,
//            "數量",
//            Qty,
//            Convert.ToString(dtInvD.Rows[i]["Qty"]).TrimStart(' '),
//            LineNum,
//               SiNo,
//               Msg, PoNo);

//                //    //Amount
//                //    AddCompareRow(
//                //    DocType,
//                //    InvoiceNo,
//                //    dt,
//                //"金額",
//                //"",
//                //Convert.ToString(dtInvD.Rows[i]["Amount"]).TrimStart(' '),
//                //LineNum);

//                //20180607
//                //LC //先用主檔法//明細檔未知 L/C NO 放在何處
//                if (IsLC == "Y")
//                {
//                    AddCompareRow(
//                    DocType,
//                    InvoiceNo,
//                    dt,
//                "L/C",
//                AcmeLC,
//                LCNo,
//                LineNum,
//                   SiNo,
//                   Msg, PoNo);
//                }



//            }

//        }


        private void GridViewAutoSize(DataGridView dgv)
        {

            for (int i = 0; i <= dgv.Columns.Count - 1; i++)
            {
                dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            // dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            for (int i = 0; i <= dgv.Columns.Count - 1; i++)
            {
                int colw = dgv.Columns[i].Width;
                dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgv.Columns[i].Width = colw;
            }
        }


        private System.Data.DataTable GetAcmeSiD_Modal(string SiNo, string PoNo, string ItemCode)
        {

            string Sql1 = "select T1.Docentry PoNo,Dscription,Quantity,T1.ItemPrice," +
            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
            "where T0.shippingcode='{0}' and T1.Docentry='{1}' and Substring(T1.ItemCode,1,9)='{2}'";

            Sql1 = string.Format(Sql1, SiNo, PoNo, ItemCode);
            System.Data.DataTable dtAcmeD = GetData(Sql1);

            return dtAcmeD;
        }


        private System.Data.DataTable GetAcmeSiD_Qty(string SiNo, string PoNO, string AcmeModal, string Qty)
        {

            string Sql1 = "select T1.Docentry PoNo,Dscription,Quantity,T1.ItemPrice," +
            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
            "where T0.shippingcode='{0}' and T1.Docentry='{1}' and Substring(T1.ItemCode,1,9)='{2}' and T1.Quantity='{3}' ";

            Sql1 = string.Format(Sql1, SiNo, PoNO, AcmeModal, Qty);
            System.Data.DataTable dtAcmeD = GetData(Sql1);

            return dtAcmeD;
        }


        private System.Data.DataTable GetAcmeSiD_VerNo(string SiNo, string PoNO, string AcmeModal, string PartNo)
        {

            string Sql1 = "select T1.Docentry PoNo,Dscription,Quantity,T1.ItemPrice," +
            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
            "where T0.shippingcode='{0}' and T1.Docentry='{1}' and Substring(T1.ItemCode,1,9)='{2}' and Substring(T1.ItemCode,12,3)='{3}'";

            Sql1 = string.Format(Sql1, SiNo, PoNO, AcmeModal, PartNo);
            System.Data.DataTable dtAcmeD = GetData(Sql1);

            return dtAcmeD;
        }

        private System.Data.DataTable GetAcmeSiD_Model(string SiNo, string PoNo, string ItemCode)
        {

            string Sql1 = "select T1.Docentry PoNo,Dscription,Quantity,T1.ItemPrice," +
            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
            "where T0.shippingcode='{0}' and T1.Docentry='{1}' and Substring(T1.ItemCode,1,9)='{2}'";

            Sql1 = string.Format(Sql1, SiNo, PoNo, ItemCode);
            System.Data.DataTable dtAcmeD = GetData(Sql1);

            return dtAcmeD;
        }


        private System.Data.DataTable GetAcmeSiD_Qty_VerNo(string SiNo, string PoNO, string AcmeModal, string Qty, string PartNo)
        {

            string Sql1 = "select T1.Docentry PoNo,Dscription,Quantity,T1.ItemPrice," +
            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
            "where T0.shippingcode='{0}' and T1.Docentry='{1}' and Substring(T1.ItemCode,1,9)='{2}' and T1.Quantity='{3}' and Substring(T1.ItemCode,12,3)='{4}'";

            Sql1 = string.Format(Sql1, SiNo, PoNO, AcmeModal, Qty, PartNo);
            System.Data.DataTable dtAcmeD = GetData(Sql1);

            return dtAcmeD;
        }


        //20180912
        private void DoCompareNew(System.Data.DataTable dt, string AuoInvoiceNo)
        {
            //CpNo SINO 可能是多筆 SH20180515012X/SH20180517005X
            //SH20180521016X ACME-T
            //1PCS SI 被拆單
            //可能併單


            // string SqlInvH = "select * from RPA_InvoiceH where InvoiceNo='{0}'";
            string SqlInvH = "select top 1* from RPA_Grade where InvoiceNo='{0}'";

            SqlInvH = string.Format(SqlInvH, AuoInvoiceNo);
            System.Data.DataTable dtInvH = GetData(SqlInvH);
           // dgHeader.DataSource = dtInvH;


            if (dtInvH.Rows.Count == 0)
            {
                MessageBox.Show("查無此 Invoice#" + AuoInvoiceNo);
                return;
            }


            //string SqlInvD = "select * from RPA_InvoiceD where InvoiceNo='{0}'";
            string SqlInvD = "select * from RPA_Grade where InvoiceNo='{0}'";

            SqlInvD = string.Format(SqlInvD, AuoInvoiceNo);
            System.Data.DataTable dtInvD = GetData(SqlInvD);
           // dgDetail.DataSource = dtInvD;


            string InvoiceNo = AuoInvoiceNo;
            string DocType = "主檔";
            string LineNum = "";

            string IsLC = "";
            string LCNo = "";


            //主檔 
            string AuoSiNo = "";
            string Msg = "";
            if (Convert.IsDBNull(dtInvH.Rows[0]["SiNo"]))
            {
                AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
               "主檔沒有SI號碼",
               "0",
               Convert.ToString(dtInvH.Rows.Count),
               LineNum,
               AuoSiNo,
               Msg, "");
                return;
            }
            else
            {
                //確認長度
                AuoSiNo = Convert.ToString(dtInvH.Rows[0]["SiNo"]);
            }


            string Sql = "select tradeCondition 貿易條件," +
"add3 付款方式 ," +
"closeDay 結關日," +
"forecastDay 預計開航日," +
"arriveDay 預計抵達日," +
"receivePlace 取貨地," +
"goalPlace 目的地," +
"shipment 裝船港," +
"unloadCargo 卸貨港 " +
"from SHIPPING_MAIN " +
"where ShippingCode ='{0}' ";

            string Sql1 = "select Docentry,Dscription,Quantity,T1.ItemPrice," +
            "T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum " +
            "from LcInstro T0 INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) " +
            "where T0.shippingcode='{0}'";

            Sql = string.Format(Sql, AuoSiNo);
            System.Data.DataTable dtAcmeH = GetData(Sql);

           // dataGridView1.DataSource = dtAcmeH;


            Sql1 = string.Format(Sql1, AuoSiNo);
            System.Data.DataTable dtAcmeD = GetData(Sql1);
           // dataGridView2.DataSource = dtAcmeD;


            string AuoSiNo_Multi = "";

            if (dtAcmeH.Rows.Count == 0)
            {
                //檢查長度
                if (AuoSiNo.Length != 14)
                {
                    Msg = "SI號碼:" + AuoSiNo;
                }
                //機率小 -> 人工判斷
                //LSZ1355824		主檔筆數為零	0	1	False	SI號碼:SH20180515012X/SH20180517005X
                if (AuoSiNo.Length > 14 & AuoSiNo.IndexOf("/") > 0)
                {
                    string[] s = AuoSiNo.Split('/');
                    //先拆一個 或 合併成 in ('s1','s2')
                    for (int k = 0; k <= s.Length - 1; k++)
                    {
                        AuoSiNo_Multi += "'" + s[k] + "',";
                    }

                    AuoSiNo_Multi = AuoSiNo_Multi.Substring(0, AuoSiNo_Multi.Length - 1);

                    //先假設取第二筆
                    AuoSiNo = s[1];
                }

                AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
               "主檔筆數為零",
               "0",
               Convert.ToString(dtInvH.Rows.Count),
               LineNum,
               AuoSiNo,
               Msg, "");

                Msg = "";

                if (string.IsNullOrEmpty(AuoSiNo_Multi))
                {
                    return;
                }
                else
                {

                    Sql = "select tradeCondition 貿易條件," +
"add3 付款方式 ," +
"closeDay 結關日," +
"forecastDay 預計開航日," +
"arriveDay 預計抵達日," +
"receivePlace 取貨地," +
"goalPlace 目的地," +
"shipment 裝船港," +
"unloadCargo 卸貨港 " +
"from SHIPPING_MAIN " +
"where ShippingCode  in ({0}) ";

                    Sql = string.Format(Sql, AuoSiNo_Multi);
                    dtAcmeH = GetData(Sql);
                    if (dtAcmeH.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        AddCompareRow(
                 DocType,
                 InvoiceNo,
                 dt,
                 "主檔筆數多筆",
                 Convert.ToString(dtAcmeH.Rows.Count),
                 "",
                 LineNum,
                 AuoSiNo_Multi,
                 Msg, "");
                    }

                }
            }




            //  System.Data.DataTable dt = MakeTableCompare();

            AddCompareRow(
                DocType,
                InvoiceNo,
                dt,
                "付款方式",
                Convert.ToString(dtAcmeH.Rows[0]["付款方式"]),
                Convert.ToString(dtInvH.Rows[0]["Payment"]),
                LineNum,
               AuoSiNo,
               Msg, "");

            if (Convert.ToString(dtInvH.Rows[0]["Payment"]).IndexOf("L/C") >= 0)
            {
                IsLC = "Y";

                try
                {
                    LCNo = Convert.ToString(dtInvH.Rows[0]["LC"]);
                }
                catch
                {
                }

            }

            AddCompareRow(
               DocType,
                InvoiceNo,
                dt,
                "貿易條件",
                Convert.ToString(dtAcmeH.Rows[0]["貿易條件"]),
                Convert.ToString(dtInvH.Rows[0]["TradeTerm"]).TrimStart(' '),
                LineNum,
               AuoSiNo,
               Msg, "");

            //結關日
            AddCompareRow(
              DocType,
                InvoiceNo,
                dt,
              "結關日",
              Convert.ToString(dtAcmeH.Rows[0]["結關日"]),
                //Convert.ToString(dtInvH.Rows[0]["InvoiceDate"]).Replace("/", ""),
              Convert.ToString(dtInvH.Rows[0]["ShipDate"]).Replace("/", ""),
              LineNum,
               AuoSiNo,
               Msg, "");

            //預計開航日
            AddCompareRow(
             DocType,
                InvoiceNo,
                dt,
             "預計開航日",
             Convert.ToString(dtAcmeH.Rows[0]["預計開航日"]),
             Convert.ToString(dtInvH.Rows[0]["ETD"]).Replace("/", ""),
             LineNum,
               AuoSiNo,
               Msg, "");

            //預計抵達日
            AddCompareRow(
            DocType,
                InvoiceNo,
                dt,
            "預計抵達日",
            Convert.ToString(dtAcmeH.Rows[0]["預計抵達日"]),
            Convert.ToString(dtInvH.Rows[0]["ETA"]).Replace("/", ""),
            LineNum,
               AuoSiNo,
               Msg, "");

            // //取貨地
            // AddCompareRow(
            //    DocType,
            //    InvoiceNo,
            //    dt,
            //"取貨地",
            //Convert.ToString(dtAcmeH.Rows[0]["取貨地"]),
            //Convert.ToString(dtInvH.Rows[0]["ShipCity"]),
            //LineNum,
            //    AuoSiNo,
            //    Msg, "");

            //目的地
            AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "目的地",
           Convert.ToString(dtAcmeH.Rows[0]["目的地"]),
           Convert.ToString(dtInvH.Rows[0]["ShipCity"]).TrimStart(' '),
           LineNum,
               AuoSiNo,
               Msg, "");

            // //裝船港
            // AddCompareRow(
            //    DocType,
            //    InvoiceNo,
            //    dt,
            //"裝船港",
            //Convert.ToString(dtAcmeH.Rows[0]["裝船港"]),
            //Convert.ToString(dtInvH.Rows[0]["ShipFrom"]).TrimStart(' '),
            //LineNum,
            //    AuoSiNo,
            //    Msg, "");


            //   //卸貨港
            //   AddCompareRow(
            //    DocType,
            //    InvoiceNo,
            //    dt,
            //"卸貨港",
            //Convert.ToString(dtAcmeH.Rows[0]["卸貨港"]),
            //Convert.ToString(dtInvH.Rows[0]["ShipTo"]).TrimStart(' '),
            //LineNum,
            //      AuoSiNo,
            //      Msg, "");

            //表頭額外檢查規則
            //where TradeTerm like 'CIF TAIWAN' -> 5% TaxRate
            string TradeTerm = "";
            string TaxRate = "";
            try
            {
                TradeTerm = Convert.ToString(dtInvH.Rows[0]["TradeTerm"]).TrimStart(' ');
            }
            catch
            {

            }
            if (!string.IsNullOrEmpty(TradeTerm))
            {
                if (TradeTerm.IndexOf("CIF TAIWAN") >= 0)
                {
                    try
                    {
                        TaxRate = Convert.ToString(dtInvH.Rows[0]["TaxRate"]).TrimStart(' ').Trim();
                    }
                    catch
                    {
                    }
                    AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "CIF TAIWAN Tax",
          "5%",
           TaxRate,
           LineNum,
               AuoSiNo,
               Msg, "");


                }

            }


            //比對多筆明細檔-------------------------------------------------------------------------

            DataRow dr;
            DocType = "明細";
            string SiNo;
            string PoNo;
            System.Data.DataTable dtSiD = null;

            string Price = "";
            string Qty = "";
            string AcmeLC = "";
            string Item = "";
            string ItemName = "";
            string AcmeItemNo = "";
            string AcmeModel = "";



            for (int i = 0; i <= dtInvD.Rows.Count - 1; i++)
            {
                dr = dtInvD.Rows[i];

                Msg = "";


                Msg = "";

                try
                {
                    PoNo = Convert.ToString(dtInvD.Rows[i]["PoNo"]);
                }
                catch
                {
                    PoNo = "";
                }


                try
                {
                    SiNo = Convert.ToString(dtInvD.Rows[i]["SiNo"]);

                    //只取 14 碼 
                    if (SiNo.Length > 14)
                    {
                        //人工比對
                        Msg += "SI號碼 > 14:" + SiNo + "\r\n";
                        SiNo = SiNo.Substring(0, 14);

                    }
                }
                catch
                {
                    SiNo = "";
                }

                //假設法
                //因為等級產地未知
                //可能多筆
                //AcmeItemNo = GetAcmeItemCode(Convert.ToString(dtInvD.Rows[i]["ItemNO"]));

                AcmeItemNo = GetAcmeItemCode(Convert.ToString(dtInvD.Rows[i]["PartNo"]));

                //20180912
                //91.21T07.10A
                //select * from RPA_Grade 
                //where InvoiceNo='LSZ1411119'
                //order by partno

                //select * from RPA_InvoiceD
                //where InvoiceNo='LSZ1411119'
                //order by ItemNo


                if (!string.IsNullOrEmpty(AcmeItemNo))
                {
                    AcmeModel = AcmeItemNo.Substring(0, 9);
                }

                Msg += "Acme Model:" + AcmeModel;
                AddCompareRow(
             DocType,
             InvoiceNo,
             dt,
         "PartNo",
         "",
         Convert.ToString(dtInvD.Rows[i]["PartNo"]),
         LineNum,
             SiNo,
             Msg, PoNo);

                if (!string.IsNullOrEmpty(SiNo) & !string.IsNullOrEmpty(PoNo))
                {
                    //dtSiD = GetAcmeSiD(SiNo, PoNo);
                    // dtSiD = GetAcmeSiD_ItemCode(SiNo, PoNo,AcmeItemNo);

                    //使用 Model 法
                    dtSiD = GetAcmeSiD_Model(SiNo, PoNo, AcmeModel);
                }
                else
                {
                    dtSiD = null;
                }


                Price = "";
                AcmeLC = "";
                Qty = "";

                Item = "";
                ItemName = "";
                if (dtSiD != null)
                {
                    if (dtSiD.Rows.Count == 1)
                    {
                        Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
                        Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
                        ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
                        AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

                    }
                    else if (dtSiD.Rows.Count == 0)
                    {

                        Msg = "Model法-SI筆數 = " + dtSiD.Rows.Count.ToString()
                            + "|" + Convert.ToString(dtInvD.Rows[i]["PartNo"])
                            + "|" + AcmeModel;


                        AddCompareRow(
               DocType,
               InvoiceNo,
               dt,
           "SI筆數異常#1",
           dtSiD.Rows.Count.ToString(),
           "",
           LineNum,
               SiNo,
               Msg, PoNo);
                        Msg = "";
                    }
                    else if (dtSiD.Rows.Count > 1)
                    {

                        //             Msg = "Model法-SI筆數 = " + dtSiD.Rows.Count.ToString();
                        //             AddCompareRow(
                        //    DocType,
                        //    InvoiceNo,
                        //    dt,
                        //"SI筆數異常#1",
                        //dtSiD.Rows.Count.ToString(),
                        //"",
                        //LineNum,
                        //    SiNo,
                        //    Msg, PoNo);
                        //             Msg = "";

                        dtSiD = GetAcmeSiD_Qty(SiNo, PoNo, AcmeModel, Convert.ToString(dtInvD.Rows[i]["Qty"]));
                        //PO +SiNo 一樣..可能要繼續判斷 -> 使用 數量法
                        //品名法
                        //單價法
                        //數量法

                        if (dtSiD.Rows.Count == 1)
                        {
                            //數量可能被拆成兩筆
                            //LSZ1357851
                            //100 -> 90 + 10

                            Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
                            Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
                            ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
                            AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

                        }
                        else
                        {

                            string PartNo = Convert.ToString(dtInvD.Rows[i]["PartNo"]).Substring(9, 3);
                            dtSiD = GetAcmeSiD_Qty_VerNo(SiNo, PoNo, AcmeModel, Convert.ToString(dtInvD.Rows[i]["Qty"]), PartNo);

                            //多筆數量相同 取第一筆
                            if (dtSiD.Rows.Count >= 1)
                            {

                                Price = Convert.ToString(dtSiD.Rows[0]["ItemPrice"]);
                                Qty = Convert.ToString(dtSiD.Rows[0]["Quantity"]);
                                ItemName = Convert.ToString(dtSiD.Rows[0]["Dscription"]);
                                AcmeLC = Convert.ToString(dtSiD.Rows[0]["LC"]);

                            }
                            else
                            {


                                System.Data.DataTable dtCheck = GetAcmeSiD_VerNo(SiNo, PoNo, AcmeModel, PartNo);

                                if (dtCheck.Rows.Count > 0)
                                {
                                    //多筆相同 取第一筆
                                    Price = Convert.ToString(dtCheck.Rows[0]["ItemPrice"]);
                                    Qty = Convert.ToString(dtCheck.Rows[0]["Quantity"]);
                                    ItemName = Convert.ToString(dtCheck.Rows[0]["Dscription"]);
                                    AcmeLC = Convert.ToString(dtCheck.Rows[0]["LC"]);

                                    Msg = "Model+Ver-SI筆數異常()" + dtCheck.Rows.Count.ToString()
                                   + "|" + Convert.ToString(dtInvD.Rows[i]["PartNo"])
                                   + "|" + AcmeModel
                                   + "|" + PartNo;

                                    AddCompareRow(
                   DocType,
                   InvoiceNo,
                   dt,
               "SI筆數異常#3",
               dtCheck.Rows.Count.ToString(),
               "",
               LineNum,
                   SiNo,
                   Msg, PoNo);
                                    Msg = "";





                                }
                                else
                                {

                                    Msg = "Model+Ver+數量法-SI筆數異常" + dtSiD.Rows.Count.ToString()
                                   + "|" + Convert.ToString(dtInvD.Rows[i]["ItemNO"])
                                   + "|" + AcmeModel
                                   + "|" + PartNo;

                                    AddCompareRow(
                       DocType,
                       InvoiceNo,
                       dt,
                   "SI筆數異常#2",
                   dtSiD.Rows.Count.ToString(),
                   "",
                   LineNum,
                       SiNo,
                       Msg, PoNo);
                                    Msg = "";
                                }





                            }






                        }
                    }
                }


                LineNum = (i + 1).ToString();
                //SH20180518017X
                //LSJ1357321
                //PO
                //比法
                AddCompareRow(
                 DocType,
                 InvoiceNo,
                 dt,
             "PO",
             PoNo,
             Convert.ToString(dtInvD.Rows[i]["PoNo"]).TrimStart(' '),
             LineNum,
               SiNo,
               Msg, PoNo);

                //    //PartNo 
                //    AddCompareRow(
                //    DocType,
                //    InvoiceNo,
                //    dt,
                //"Model",
                //AcmeItemNo,
                //Convert.ToString(dtInvD.Rows[i]["ItemNo"]).TrimStart(' '),
                //LineNum,
                //SiNo,
                //   Msg);



                //    //ItemNo
                //    AddCompareRow(
                //    DocType,
                //    InvoiceNo,
                //    dt,
                //"品名",
                //ItemName,
                //Convert.ToString(dtInvD.Rows[i]["ItemName"]).TrimStart(' '),
                //LineNum,
                //   SiNo,
                //   Msg, PoNo);


                //Price
                AddCompareRow(
                DocType,
                InvoiceNo,
                dt,
            "單價",
            Price,
            Convert.ToString(dtInvD.Rows[i]["Price"]).TrimStart(' '),
            LineNum,
               SiNo,
               Msg, PoNo);

                //Qty
                AddCompareRow(
                DocType,
                InvoiceNo,
                dt,
            "數量",
            Qty,
            Convert.ToString(dtInvD.Rows[i]["Qty"]).TrimStart(' '),
            LineNum,
               SiNo,
               Msg, PoNo);

                //    //Amount
                //    AddCompareRow(
                //    DocType,
                //    InvoiceNo,
                //    dt,
                //"金額",
                //"",
                //Convert.ToString(dtInvD.Rows[i]["Amount"]).TrimStart(' '),
                //LineNum);

                //20180607
                //LC //先用主檔法//明細檔未知 L/C NO 放在何處
                if (IsLC == "Y")
                {
                    AddCompareRow(
                    DocType,
                    InvoiceNo,
                    dt,
                "L/C",
                AcmeLC,
                LCNo,
                LineNum,
                   SiNo,
                   Msg, PoNo);
                }


                //
                System.Data.DataTable dtGetCheckSum = GetCheckSum(InvoiceNo, SiNo, PoNo, Convert.ToString(dtInvD.Rows[i]["PartNo"]));
                System.Data.DataTable dtGetCheckSumAcme = GetCheckSumAcme(SiNo, PoNo, Convert.ToString(dtInvD.Rows[i]["PartNo"]));

                Int32 iGetCheckSumAcme = 0;
                Int32 iGetCheckSumAcmeQty = 0;
                Int32 iGetCheckSum = 0;
                Int32 iGetCheckSumQty = 0;


                try
                {
                    iGetCheckSumAcme = Convert.ToInt32(dtGetCheckSumAcme.Rows[0]["明細筆數"]);
                    iGetCheckSumAcmeQty = Convert.ToInt32(dtGetCheckSumAcme.Rows[0]["Qty"]);

                }
                catch
                {

                }
                try
                {
                    iGetCheckSum = Convert.ToInt32(dtGetCheckSum.Rows[0]["明細筆數"]);
                    iGetCheckSumQty = Convert.ToInt32(dtGetCheckSum.Rows[0]["Qty"]);
                }
                catch
                {
                }



                AddCompareRow(
                   DocType,
                   InvoiceNo,
                   dt,
               "彙總筆數",
               Convert.ToString(iGetCheckSumAcme),
               Convert.ToString(iGetCheckSum),
               LineNum,
                  SiNo,
                  Msg, PoNo);

                AddCompareRow(
                 DocType,
                 InvoiceNo,
                 dt,
             "彙總數量",
             Convert.ToString(iGetCheckSumAcmeQty),
               Convert.ToString(iGetCheckSumQty),
             LineNum,
                SiNo,
                Msg, PoNo);


            }

        }

        private System.Data.DataTable GetCheckSum(string InvoiceNo, string SiNo, string PoNo, string ItemNo)
        {
            //判斷 同一單號中 有 多筆重覆料號
            string Sql = "select InvoiceNo,ItemNo,Count(*) 明細筆數,Sum(Qty) Qty from RPA_InvoiceD " +
                "Where InvoiceNo ='{0}' and  SiNo='{1}' and PoNo='{2}' and ItemNo='{3}' " +
                      "group by  InvoiceNo,ItemNo";
            //having Count(*) > 1

            Sql = string.Format(Sql, InvoiceNo, SiNo, PoNo, ItemNo);
            System.Data.DataTable dt = GetData(Sql);

            return dt;


        }


        private System.Data.DataTable GetCheckSumAcme(string SiNo, string PoNo, string PartNo)
        {
            //判斷 同一單號中 有 多筆重覆料號
            string Sql = "select T0.shippingcode,T1.Docentry,TT.U_partno,Count(*) 明細筆數,Sum(Convert(int,T1.Quantity)) Qty " +
"from LcInstro T0 " +
"INNER JOIN LcInstro1 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE " +
"AND T0.DOCNUM=T1.DOCNUM ) " +
"Inner join acmesql02..oitm TT on TT.ItemCode=t1.ItemCode  collate Chinese_Taiwan_Stroke_CS_AS " +
"where T0.shippingcode='{0}' " +
"and T1.Docentry='{1}' " +
"and TT.U_partno='{2}' " +
"group by T0.shippingcode,T1.Docentry,TT.U_partno";
            //having Count(*) > 1

            Sql = string.Format(Sql, SiNo, PoNo, PartNo);
            System.Data.DataTable dt = GetData(Sql);

            return dt;


        }


        private void dgCompare_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {

            DataGridView dgv = (DataGridView)sender;
            DataGridViewRow dgr = dgv.Rows[e.RowIndex];
            DataRowView row = (DataRowView)dgv.Rows[e.RowIndex].DataBoundItem;

            try
            {
                if (!Convert.IsDBNull(row["Result"]))
                {
                    if (Convert.ToString(row["Result"]) == "False")
                    {
                        dgr.DefaultCellStyle.BackColor = Color.Yellow;
                    }

                }
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {


            System.Data.DataTable dt = MakeTableCompare();


            if (radioButton1.Checked)
            {

                DoCompareNew(dt, txtAuo.Text);
            }
            else
            {
                DoCompare(dt, txtAuo.Text);
            }

            dgCompare.DataSource = dt;
            
            GridViewAutoSize(dgCompare);


            //過濾

            DataView dv = dt.DefaultView;
            System.Data.DataTable dtFalse = dv.ToTable();

            dv = dtFalse.DefaultView;
            dv.RowFilter = "result ='False'";


            dgFalse.DataSource = dtFalse;

            
            
            //string SqlInvH = "select * from RPA_InvoiceH Where InvoiceNo='{0}'";
            //SqlInvH = string.Format(SqlInvH, textBox1.Text);

            //System.Data.DataTable dtInvH = GetData(SqlInvH);

            //System.Data.DataTable dt = MakeTableCompare();
            //string AuoInvoiceNo = "";

            
            //    AuoInvoiceNo = Convert.ToString(dtInvH.Rows[0]["InvoiceNo"]);
            //    DoCompare(dt, AuoInvoiceNo);

           
            //dgCompare.DataSource = dt;
           
            //GridViewAutoSize(dgCompare);

            ////過濾

            //DataView dv = dt.DefaultView;
            //System.Data.DataTable dtFalse = dv.ToTable();

            //dv = dtFalse.DefaultView;
            //dv.RowFilter = "result ='False'";


            //dgFalse.DataSource = dtFalse;

            //GridViewAutoSize(dgFalse);
          

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dgCompare);
        }


    }
}
